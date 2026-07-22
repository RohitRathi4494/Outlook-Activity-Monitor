# Outlook Activity Monitor — Project Documentation

> Internal reference doc. Read this first to understand the whole project quickly.
> Last updated: 2026-07-22.

---

## 1. What this project is

A small **FastAPI web app** that lets a Microsoft 365 user sign in with their
Outlook/Microsoft account and then:

1. **Syncs their mailbox** (received + sent mail metadata) from Microsoft Graph
   every 30 minutes.
2. **Detects forwarded mail** — for each received message it works out whether
   the user later forwarded it, and to whom.
3. **Emails the user a daily Excel report** of the previous day's received
   activity (to their own mailbox), and also lets them download any day's
   report on demand from a dashboard.

**Strict per-user isolation** is the central design constraint: every Graph
call uses that user's *own* delegated access token against `/me` endpoints, and
every database query is filtered by that user's id. One user can never see or
act on another user's mailbox. This is repeated deliberately across the code
(`auth.py`, `collector.py`, `report.py`, `mailer.py`).

- **Tenant:** single-tenant Entra ID app for `salwangurgaon.com`.
- **Repo:** https://github.com/RohitRathi4494/Outlook-Activity-Monitor
- **Live deployment:** https://outlook-activity-monitor.onrender.com (Render, free plan)

---

## 2. Tech stack

| Concern | Choice |
|---|---|
| Web framework | FastAPI + Uvicorn |
| Auth | MSAL (OAuth 2.0 Authorization Code flow, **delegated** permissions only) |
| Microsoft API | Microsoft Graph v1.0 (`https://graph.microsoft.com/v1.0`) |
| HTTP client | httpx (async) |
| ORM / DB | SQLAlchemy 2.x — Postgres in prod (Render), SQLite locally |
| Excel | openpyxl |
| Token encryption at rest | cryptography **Fernet** |
| Sessions | Starlette `SessionMiddleware` (signed cookie, holds only a user id) |
| In-process scheduling (local only) | APScheduler |
| Production scheduling | **GitHub Actions cron** hitting `/tasks/*` |
| Hosting | Render (free web service + free Postgres) |

Pinned versions live in `requirements.txt` (FastAPI 0.115.0, SQLAlchemy 2.0.35,
msal 1.31.0, psycopg2-binary 2.9.12, openpyxl 3.1.5, etc.).

---

## 3. Repository layout

```
outlook-activity-monitor/
├── main.py              FastAPI app + all HTTP routes
├── auth.py              MSAL OAuth: build auth URL, exchange code, mint access tokens
├── collector.py         Fetch/dedupe/store a user's mail; forward detection
├── graph_client.py      Thin Graph HTTP client: pagination + 429 retry/backoff
├── report.py            Build the per-user daily Excel (.xlsx) report
├── mailer.py            Email a user's report to their own mailbox via Graph sendMail
├── scheduler.py         Job bodies: poll_all_users() and send_daily_reports()
├── models.py            SQLAlchemy models (User, Message) + engine/DB selection
├── crypto_utils.py      Fernet encrypt/decrypt of refresh tokens
├── templates/
│   ├── index.html       Landing / "Sign in with Microsoft"
│   └── dashboard.html   Signed-in user's report controls
├── .github/workflows/
│   └── scheduled-tasks.yml   GitHub Actions cron that drives /tasks/*
├── render.yaml          Render blueprint (web service + Postgres, env vars)
├── requirements.txt
├── .env / .env.example  Local config (secrets; .env is gitignored)
├── outlook_activity.db  Local SQLite DB (gitignored; dev only)
└── README.md
```

---

## 4. Data model (`models.py`)

The DB engine is chosen at import time:
- If `DATABASE_URL` is set → use it (Render Postgres). A legacy `postgres://`
  scheme is rewritten to `postgresql://` for SQLAlchemy's psycopg2 dialect.
- Otherwise → local SQLite file `outlook_activity.db` (dev).

**`User`** — one row per signed-in mailbox owner.
- `id` (PK) = Entra **object id (`oid`)**, stable per user in the tenant.
- `email`, `display_name`
- `encrypted_refresh_token` (Fernet ciphertext, **never** the plaintext)
- `created_at`, `updated_at`
- `messages` relationship (cascade delete)

**`Message`** — one row per received *or* sent mail item, always tagged with
`user_id`.
- `direction` = `"received"` or `"sent"`
- Sender/subject/time fields, `to_recipients`/`cc_recipients` (rendered as
  `"Name <email>; …"`), `has_attachments`, `importance`, `conversation_id`,
  `internet_message_id`
- **Forward metadata (on received rows only):** `forwarded` (bool),
  `forwarded_to`, `forwarded_time` — computed once at collection time.
- Unique constraint: `(user_id, message_id, direction)` → dedupe key.

`init_db()` runs `create_all` on startup (no migrations framework; schema is
create-only).

---

## 5. Authentication flow (`auth.py`)

Delegated OAuth 2.0 Authorization Code flow via MSAL
`ConfidentialClientApplication`. **No passwords, no application permissions** —
every Graph call is scoped to the signed-in user via `/me`.

- **Scopes:** `Mail.Read`, `Mail.Send`, `User.Read` (MSAL auto-appends
  `offline_access`/`openid`/`profile`, which is what makes a **refresh token**
  come back).
- `build_auth_url(state)` → Microsoft login URL (state = CSRF token).
- `exchange_code_for_user(code)` → exchanges the code, extracts `oid`/email/name
  from id-token claims, **encrypts and stores the refresh token**, upserts the
  `User` row. Raises if no refresh token comes back (usually means
  `offline_access`/admin consent missing).
- `get_access_token_for_user(user_id)` → loads + decrypts that user's refresh
  token, mints a fresh access token, and **persists the rotated refresh token**
  if Microsoft returns a new one. Returns `None` on failure (revoked/expired
  consent) — callers must treat `None` as "user must sign in again" and must
  **never** fall back to another user's credentials.

Refresh tokens are encrypted at rest with Fernet (`crypto_utils.py`); the
`ENCRYPTION_KEY` is a URL-safe base64 32-byte key.

---

## 6. HTTP routes (`main.py`)

| Method + path | Purpose | Auth |
|---|---|---|
| `GET /` | Landing page / redirect to dashboard if signed in | session |
| `GET /login` | Redirect to Microsoft OAuth authorize | — |
| `GET /auth/callback` | Exchange code, start session, run first sync | OAuth state |
| `GET /logout` | Clear session | session |
| `GET /dashboard` | Report controls for the signed-in user | session |
| `GET /report?date=YYYY-MM-DD` | Download that user's Excel report | session |
| `GET /healthz` | Cheap liveness check (used to warm the service) | — |
| `POST /tasks/sync` | Run the mailbox sync for **every** user | `X-Task-Token` |
| `POST /tasks/daily-report` | Sync, then email every user their prior-day report | `X-Task-Token` |

Key implementation details:
- **Current user is resolved strictly from the signed session cookie**
  (`_current_user`), never from a query/body param. The session cookie
  (`oam_session`) is signed with a secret derived from `ENCRYPTION_KEY`; it
  holds only the user id, no tokens.
- Cookie is marked `Secure` only in production (detected via Render's `RENDER`
  env var) so local `http://localhost` dev still works.
- `/tasks/*` are guarded by `_verify_task_token()`: constant-time compare
  against `TASK_TRIGGER_TOKEN`. If that env var is unset on the server → **503**
  "Task triggers are not configured"; if the header is missing/wrong → **401**.
  *(This 503-vs-401 distinction is a useful diagnostic — see §10.)*
- On startup: `init_db()` then, **only if `ENABLE_INTERNAL_SCHEDULER=true`**,
  start APScheduler. In production this is `false` (external cron drives things).

---

## 7. Collection & forward detection (`collector.py`, `graph_client.py`)

`collect_for_user(user_id)` is the single entry point (called after first
sign-in and by the 30-min sync):

1. Get the user's access token (no-op return `0` if none).
2. Pull **received** mail from `/me/messages` and **sent** mail from
   `/me/mailFolders/sentitems/messages`, paginating via `@odata.nextLink`
   (`graph_client.get_all_pages`). Graph 429s are retried with exponential
   backoff honoring `Retry-After` (max 5 retries).
3. **Forward detection:** sent items whose subject starts with `FW:`/`Fwd:` are
   candidates. A received message is considered forwarded if a candidate
   **shares its `conversationId`** OR **references its `internetMessageId`**
   (via `References`/`In-Reply-To` headers), and the forward wasn't sent before
   the message was received. The earliest matching forward wins; its recipients
   + time are stored on the received row.
4. **Dedupe** against existing `(user_id, message_id, direction)` rows and insert
   only new ones. Commits atomically (rollback on error).

`sentDateTime` and `internetMessageHeaders` are only `$select`ed for sent items
(needed for forward matching).

---

## 8. Reporting & email (`report.py`, `mailer.py`)

**`report.py::generate_report(user_id, report_date)`** builds an in-memory
`.xlsx` (BytesIO) of that user's **received** messages for the given day
(`received_datetime` in `[day_start, day_start+1)`), ordered ascending. Columns:
Received From, Subject, Received Time, To/CC Recipients, Has Attachments,
Importance, Forwarded (Y/N), Forwarded To, Forwarded Time, Conversation ID,
Message ID. Header row bold + frozen; columns auto-sized (capped at width 60).
**Every query is filtered by `user_id`** — the isolation enforcement point on
the download path.

**`mailer.py::send_daily_report(user_id, email, report_date)`** generates the
report and sends it as a base64 attachment via Graph
`POST /me/sendMail` using **that user's own access token**, so the mail is both
sent from and delivered to the same mailbox. Returns `False` (skips, no retry
with other creds) if no access token can be obtained. `saveToSentItems=true`.

---

## 9. Scheduling — the important production detail

There are **two** scheduling mechanisms; only one is active per environment.

### (a) In-process APScheduler (`scheduler.py`) — LOCAL DEV ONLY
`start_scheduler()` registers:
- `poll_all_users` — every 30 min.
- `send_daily_reports` — cron at 06:00 **server-local** time, covering the
  previous full calendar day.

Both loop over all users and **swallow per-user exceptions** so one user's
failure can't block the others. Enabled when `ENABLE_INTERNAL_SCHEDULER=true`
(the default). **Turned off in production** because a free Render instance
sleeps between requests, so an in-process timer can't be relied on.

### (b) GitHub Actions cron (`.github/workflows/scheduled-tasks.yml`) — PRODUCTION
This is what actually runs the jobs in production. GitHub's free cron wakes the
sleeping Render service and POSTs a trigger endpoint:

- `*/30 * * * *` → task `sync` → `POST /tasks/sync`
- `30 0 * * *` (00:30 UTC = **06:00 IST**) → task `daily-report` →
  `POST /tasks/daily-report`
- Also runnable manually via **workflow_dispatch** (choose `sync` or
  `daily-report`).

Job steps:
1. **Decide which task** (based on which cron / dispatch input fired).
2. **Check required secrets** — validates `APP_URL` is an absolute `http(s)` URL
   and `TASK_TRIGGER_TOKEN` is set; fails fast with a clear `::error::`
   otherwise. *(Added 2026-07-22 — see §10.)*
3. **Warm up the service** — `curl .../healthz` with retries (free host
   cold-starts in ~30–60 s). Emits a `::warning::` if it doesn't succeed but
   does **not** fail the job (the trigger step is the real gate).
4. **Trigger task** — `curl -fsS -X POST` to `${APP_URL}/tasks/<task>` with the
   `X-Task-Token` header. `-f` makes curl fail the job on any HTTP ≥ 400.

`concurrency: group: scheduled-tasks, cancel-in-progress: false` prevents a slow
sync and the daily run from overlapping.

> **Note:** GitHub scheduled crons are best-effort and often delayed/skipped on
> free runners, so the exact 00:30 UTC daily slot may drift. Manual dispatch is
> the reliable way to force a run.

---

## 10. Known incident + fixes (history)

**2026-07-22 — "All jobs have failed" / no daily summary email.**
- **Symptom:** every scheduled run (11/11) failed; no report email sent.
- **Diagnosis:** the failing step was *Trigger task*, and the CI annotation was
  `Process completed with exit code 3` = **curl "URL malformed."** The
  `APP_URL` repository secret was unset, so the request URL collapsed to
  `/tasks/sync` (no scheme/host). The warm-up step's `|| true` had been masking
  the same failure, making it look like only the trigger broke.
- **Confirmed healthy at the time:** `GET /healthz` → 200 (32 s cold start);
  `POST /tasks/sync` with no token → **401** (not 503), proving
  `TASK_TRIGGER_TOKEN` was correctly set on Render and the guard worked. The
  app/deployment were fine — only GitHub didn't know the URL.
- **Fix:** (1) user added the `APP_URL` repo **Secret** =
  `https://outlook-activity-monitor.onrender.com` (must be a **Secret**, not a
  **Variable**; no trailing slash/quotes). (2) Hardened the workflow with the
  "Check required secrets" step and replaced the warm-up `|| true` with a
  warning (commit `c7e7b6e`).
- **Verified:** manual `workflow_dispatch` of `daily-report` (run #13) succeeded
  and the report email arrived.

**Diagnostic playbook for future CI failures (no repo-admin token needed):**
Job logs need admin rights, but these public API calls are enough to localize a
failure:
```
# list recent runs + conclusions
curl -fsS ".../actions/runs?per_page=15"
# which step failed
curl -fsS ".../actions/runs/<RUN_ID>/jobs"
# the failure annotation (often carries the real error, e.g. curl exit code)
curl -fsS ".../check-suites/<CHECK_SUITE_ID>/check-runs"
curl -fsS ".../check-runs/<CHECK_RUN_ID>/annotations"
```
Then probe the live service directly (`curl .../healthz`, `POST .../tasks/sync`
with no token) to separate app-side vs config-side causes. Curl exit codes to
know: **3** = malformed URL, **6** = DNS, **7** = connect refused, **22** =
HTTP ≥ 400 (with `-f`), **28** = timeout.

---

## 11. Configuration (environment variables)

| Var | Where | Purpose |
|---|---|---|
| `CLIENT_ID` | Render + local | Entra app (client) id |
| `TENANT_ID` | Render + local | Entra tenant id (single tenant) |
| `CLIENT_SECRET` | Render + local | Entra client secret **value** |
| `REDIRECT_URI` | Render + local | Must exactly match the app registration's Web redirect URI. Local: `http://localhost:8000/auth/callback` |
| `ENCRYPTION_KEY` | Render + local | Fernet key (base64 32-byte). Also derives the session-cookie signing secret. |
| `TASK_TRIGGER_TOKEN` | Render + **GitHub secret** | Shared secret guarding `/tasks/*`. **The GitHub value must equal the Render value.** |
| `APP_URL` | **GitHub secret** | Base URL of the deployed service, e.g. `https://outlook-activity-monitor.onrender.com` (no trailing slash). **Must be a Secret, not a Variable.** |
| `ENABLE_INTERNAL_SCHEDULER` | Render = `false`, local = `true` | Toggles APScheduler. |
| `DATABASE_URL` | Render (from Postgres) | Postgres connection string; absent locally → SQLite. |
| `RENDER` | set by Render | Presence flips the session cookie to `Secure`. |

`render.yaml` declares all of these (most `sync: false` = set manually in the
Render dashboard), plus the free Postgres database
`outlook-activity-monitor-db`.

> **Watch-out:** Render free Postgres and free web services have lifecycle
> limits (services sleep; free databases can expire). If `/tasks/*` starts
> returning 500s where it used to work, suspect DB availability — note that the
> top-level `db.query(...)` in `poll_all_users`/`send_daily_reports` is *not*
> inside the per-user try/except, so a DB outage there surfaces as a 500 and
> fails the CI job.

---

## 12. Running locally

```bash
cd outlook-activity-monitor
python -m venv venv && source venv/Scripts/activate   # Windows Git Bash
pip install -r requirements.txt
cp .env.example .env        # then fill in real values
#   - generate ENCRYPTION_KEY: python -c "from cryptography.fernet import Fernet; print(Fernet.generate_key().decode())"
#   - leave TASK_TRIGGER_TOKEN blank locally (rely on the in-process scheduler)
uvicorn main:app --reload --port 8000
# open http://localhost:8000  ->  "Sign in with Microsoft"
```

With `ENABLE_INTERNAL_SCHEDULER=true` (default), APScheduler runs the 30-min
sync and the 06:00 daily email locally — no GitHub Actions needed.

To exercise the production trigger path locally, set `TASK_TRIGGER_TOKEN` and:
```bash
curl -X POST -H "X-Task-Token: <token>" http://localhost:8000/tasks/daily-report
```

---

## 13. Gotchas / things to remember

- **Isolation is sacred:** never introduce a code path that reads a user id from
  a request param, or that reuses one user's token for another. Every Graph call
  = that user's token + `/me`; every `Message` query = filtered by `user_id`.
- **Two schedulers, one active:** APScheduler (local) vs GitHub Actions (prod),
  gated by `ENABLE_INTERNAL_SCHEDULER`. Don't "fix" prod by turning APScheduler
  back on — a sleeping free host can't run it.
- **The daily report covers *yesterday*** (previous full calendar day), sent at
  06:00 IST (00:30 UTC).
- **Schema is create-only** (`create_all`) — there are no migrations; a model
  change to an existing column won't auto-apply to an existing DB.
- **Secrets: Secret vs Variable.** The workflow reads `secrets.APP_URL` /
  `secrets.TASK_TRIGGER_TOKEN`. Adding them under the Actions *Variables* tab
  leaves `secrets.*` empty and reproduces the exit-3 failure.
- **No automated tests** currently exist in the repo.
