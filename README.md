# Outlook Activity Monitor

A multi-user web app where each person signs in with their own Microsoft 365 /
Outlook work account (delegated OAuth 2.0, authorization code flow). From
then on, every day they're automatically emailed an Excel report of **their
own** email activity — no need to log back in. Strict per-user isolation:
every stored row and every report query is filtered by the signed-in user's
id, and all Graph calls use that user's own token against `/me` endpoints —
the app never reads (or emails) another user's mailbox.

## How it works

- **Auth**: MSAL confidential client, OAuth 2.0 authorization code flow,
  delegated scopes only (`Mail.Read`, `Mail.Send`, `offline_access`,
  `User.Read`). No passwords, no application permissions.
- **Storage**: SQLAlchemy, SQLite by default for local dev. If a
  `DATABASE_URL` env var is set (e.g. Render's attached Postgres), it's used
  instead — see [Deploying to Render](#deploying-to-render) below.
  `User.encrypted_refresh_token` is encrypted at rest with Fernet before it
  ever touches disk either way.
- **Collection**: `collector.py` fetches a user's received mail
  (`/me/messages`) and sent mail (`/me/mailFolders/sentitems/messages`),
  paginating through `@odata.nextLink` and retrying on HTTP 429 with
  exponential backoff honoring `Retry-After`. A received message is flagged
  as forwarded when a sent item with a `FW:`/`Fwd:` subject shares its
  `conversationId`, or references its `internetMessageId` via the
  `References`/`In-Reply-To` headers.
- **Scheduler** (`scheduler.py`): two background jobs, both looping over
  every stored user independently (one user's failure never blocks another):
  - `poll_all_users` re-syncs every user's mailbox every 30 minutes, using
    that user's own refresh token.
  - `send_daily_reports` runs once a day at **06:00 server-local time** and
    emails each user their own complete report for the previous calendar day
    — via `mailer.py`, which calls Graph's delegated `/me/sendMail` with the
    `.xlsx` as an attachment, sent to (and from) that user's own mailbox.
- **Report**: `GET /report?date=YYYY-MM-DD` streams an `.xlsx` built with
  openpyxl for the currently logged-in user only — still available for an
  on-demand download of any date, in addition to the automatic daily email.

---

## 1. Azure Portal setup

The app registration lives in the **salwangurgaon.com** Entra ID tenant
(`TENANT_ID` given below) — this is what restricts sign-in to accounts on
that domain only. If you ever need to recreate it (e.g. for a different
domain/tenant), do the following in the Azure Portal
(portal.azure.com, switched into the correct tenant → **Microsoft Entra ID**
→ **App registrations**):

1. **New registration**
   - Name: anything, e.g. `Outlook Activity Monitor`.
   - Supported account types: **"Accounts in this organizational directory
     only (Single tenant)"** — this is what limits sign-in to that tenant's
     own accounts.
   - Redirect URI: Platform **Web**, URI `http://localhost:8000/auth/callback`.
   - Copy the **Application (client) ID** and **Directory (tenant) ID** from
     the Overview page into `.env` as `CLIENT_ID` / `TENANT_ID`.

2. **Add a client secret**
   - **Certificates & secrets** → **New client secret** → give it a
     description and expiry → **Add**.
   - Copy the secret's **Value** immediately (it's hidden after you leave the
     page) and put it in `.env` as `CLIENT_SECRET`.

3. **Add delegated API permissions**
   - **API permissions** → **Add a permission** → **Microsoft Graph** →
     **Delegated permissions** → add:
     - `Mail.Read`
     - `Mail.Send`
     - `offline_access`
     - `User.Read`
   - Click **Grant admin consent** so any user on the domain can sign in
     without an individual consent prompt (needs admin rights on the tenant;
     otherwise each user just consents once on first sign-in).
   - If `Mail.Send` is being added to an **existing** app registration that
     users already signed into (rather than a brand new one), re-clicking
     **Grant admin consent** is enough — already-stored refresh tokens pick
     up the newly consented scope automatically the next time the app
     refreshes them, with no need for anyone to sign in again.

## 2. Generate an encryption key

Refresh tokens are encrypted at rest with [Fernet](https://cryptography.io/en/latest/fernet/).
Generate a key with:

```bash
python -c "from cryptography.fernet import Fernet; print(Fernet.generate_key().decode())"
```

Paste the output into `.env` as `ENCRYPTION_KEY`.

## 3. Configure `.env`

Copy the example and fill in the two values from steps 1 and 2:

```bash
cp .env.example .env
```

```
CLIENT_ID=17c09f6d-dfb3-451f-85f1-96e40d048038
TENANT_ID=4fae5689-bd93-4e1a-99ec-4da629fd0416
CLIENT_SECRET=<value from step 1.2>
REDIRECT_URI=http://localhost:8000/auth/callback
ENCRYPTION_KEY=<value from step 2>
```

## 4. Run locally

```bash
python -m venv venv
# Windows:
venv\Scripts\activate
# macOS/Linux:
source venv/bin/activate

pip install -r requirements.txt

uvicorn main:app --reload
```

Open **http://localhost:8000**, click **Sign in with Microsoft**, and
authenticate with any @salwangurgaon.com work account. You'll land on a
dashboard showing your email, a date picker, and a **Download report**
button. From then on, that user is automatically emailed their own report
every day at 06:00 — no further action needed on their part.

A SQLite file `outlook_activity.db` is created automatically on first run —
no migrations needed.

## 5. Multi-user usage

Any @salwangurgaon.com user can sign in independently at the same URL. Each gets
their own row in the `users` table and their own set of `messages` rows; the
session cookie (signed, not encrypted — it holds only a user id, never a
token) determines whose data a request can see, and `/report`, the 30-minute
sync, and the daily report email all filter/act on data strictly for one
user's id at a time.

## Project structure

```
outlook-activity-monitor/
├── .env.example        # all required env vars, no real secrets
├── .gitignore
├── requirements.txt
├── main.py             # FastAPI app, routes, session handling
├── auth.py             # MSAL login/callback/refresh
├── graph_client.py     # Graph HTTP calls, pagination, 429 retry
├── collector.py        # fetch + forward detection + dedupe + store
├── models.py            # SQLAlchemy models: User, Message
├── crypto_utils.py      # Fernet encrypt/decrypt of refresh tokens
├── report.py             # openpyxl Excel report generation
├── mailer.py              # emails the daily report via Graph /me/sendMail
├── scheduler.py           # APScheduler: 30-min sync + daily report email
├── templates/
│   ├── index.html
│   └── dashboard.html
└── README.md
```

## Deploying for free (Render + GitHub Actions cron)

**The scheduling problem, and how this is solved for free.** The app's value
is its automatic work — a 30-minute mailbox sync and a daily report email.
Every free host (Render, Fly, Koyeb, Cloud Run) *sleeps* the process when no
HTTP request is arriving, and an in-process scheduler like APScheduler cannot
fire while asleep. So instead of keeping a process awake 24/7 (which costs
money), the scheduling is moved **out** of the app:

- The app exposes two token-protected endpoints — `POST /tasks/sync` and
  `POST /tasks/daily-report` — that run exactly the same work the old
  `scheduler.py` jobs did.
- A **GitHub Actions** workflow (`.github/workflows/scheduled-tasks.yml`),
  which is free, runs on a cron. Each run first hits `/healthz` to wake the
  sleeping service, then POSTs the trigger. The HTTP request itself is what
  wakes the host, so the job runs reliably regardless of sleep.
- The in-process scheduler is turned off in production via
  `ENABLE_INTERNAL_SCHEDULER=false` (set in `render.yaml`), so nothing runs
  twice. Locally it stays on, so `uvicorn --reload` behaves as before.

This keeps the whole thing on free tiers with no credit card and no
never-sleep instance.

### Step 1 — Deploy the web service on Render

1. **Push this repo to GitHub** (already done if you're reading this) and
   sign up at [render.com](https://render.com) with that GitHub account.
2. **New → Blueprint**, pick this repo. Render reads `render.yaml` and
   provisions a free **web service** plus a free **Postgres** auto-wired as
   `DATABASE_URL`.
3. **Set the environment variables** on the web service (Render dashboard →
   your service → **Environment**) — `render.yaml` leaves the secrets blank
   on purpose:
   - `CLIENT_ID`, `TENANT_ID`, `CLIENT_SECRET` — from the Entra app registration.
   - `ENCRYPTION_KEY` — generate with the command in section 2 above.
   - `TASK_TRIGGER_TOKEN` — generate a long random string, e.g.
     `python -c "import secrets; print(secrets.token_urlsafe(32))"`. Keep it
     handy; GitHub needs the **same** value in step 3.
   - `REDIRECT_URI` — Render assigns a URL like
     `https://outlook-activity-monitor.onrender.com`; set this to that URL
     with `/auth/callback` appended.
   - `ENABLE_INTERNAL_SCHEDULER` is already `"false"` from `render.yaml` — leave it.

### Step 2 — Register the redirect URI in Azure

App registration → **Authentication** → **Add URI** under the existing Web
platform → `https://<your-render-subdomain>.onrender.com/auth/callback` →
**Save**. (Keep the `localhost` one too for local dev.)

### Step 3 — Turn on the free cron (GitHub Actions)

The workflow is already in the repo. Give it the two secrets it needs:
GitHub repo → **Settings → Secrets and variables → Actions → New repository
secret**:
- `APP_URL` — your Render URL, no trailing slash, e.g.
  `https://outlook-activity-monitor.onrender.com`
- `TASK_TRIGGER_TOKEN` — the **exact same** value you set on Render in step 1.

The workflow then runs automatically: a sync every 30 minutes and the daily
report at **00:30 UTC = 06:00 IST**. To test immediately, go to the repo's
**Actions** tab → **Scheduled tasks** → **Run workflow**, pick `sync` or
`daily-report`, and run it manually.

> **Note on GitHub Actions minutes:** scheduled workflows are *unlimited and
> free on public repos*. On a **private** repo they draw from the free
> 2,000 minutes/month — the every-30-min sync alone is ~1,440 runs/month, so
> either keep the repo public, lengthen the sync interval in the workflow, or
> move the frequent sync to [cron-job.org](https://cron-job.org) (also free).
> Also, GitHub cron can be delayed several minutes under load — fine for this
> app, but don't expect 06:00:00 to the second.

Render auto-deploys on every push to `main`. Once deployed, `RENDER=true` is
set automatically, which marks the session cookie `Secure` — no manual toggle.

**One caveat still worth a reminder:** Render's free Postgres is deleted 90
days after the *database's* creation. For a trial that's fine; before it
lapses, either upgrade the DB or point `DATABASE_URL` at a free
[Neon](https://neon.tech) or [Supabase](https://supabase.com) Postgres, which
don't expire on that clock — otherwise stored refresh tokens and mail history
are lost.
