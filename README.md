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

## Deploying to Render

**Why not Vercel:** Vercel only runs code as short-lived serverless
functions with no persistent disk. This app needs a process that stays
alive continuously (the 30-minute sync + daily 06:00 report-email jobs in
`scheduler.py`) and a database that survives between requests and restarts
— neither works on Vercel. Render's free web service runs a real persistent
process, which is what this needs.

1. **Push this repo to GitHub** (already done if you're reading this from
   the repo) and sign up at [render.com](https://render.com) with that
   GitHub account.

2. **New → Blueprint**, pick this repo. Render reads `render.yaml` in the
   repo root and provisions:
   - a free **web service** running `uvicorn main:app --host 0.0.0.0 --port $PORT`
   - a free **Postgres database**, auto-wired into the web service as
     `DATABASE_URL` (this is what makes storage survive restarts/redeploys
     — SQLite would not).

3. **Set the remaining environment variables** on the web service (Render
   dashboard → your service → Environment), since `render.yaml` intentionally
   leaves these blank for you to fill in rather than committing them to git:
   - `CLIENT_ID`, `TENANT_ID`, `CLIENT_SECRET` — from your Entra ID app
     registration (see step 1 above).
   - `ENCRYPTION_KEY` — generate with the command in step 2 above (use a
     **different** key than your local `.env` if you want local and
     production refresh tokens to be independently encrypted).
   - `REDIRECT_URI` — Render gives your service a URL like
     `https://outlook-activity-monitor.onrender.com`; set this to
     `https://outlook-activity-monitor.onrender.com/auth/callback` (use
     whatever subdomain Render actually assigned).

4. **Add that same redirect URI in Azure Portal** — App registration →
   **Authentication** → **Add URI** under the existing Web platform →
   `https://<your-render-subdomain>.onrender.com/auth/callback` → **Save**.
   (Keep the `localhost` one too if you still want to run locally.)

5. Render auto-deploys on every push to `main`. Once deployed, `RENDER=true`
   is set automatically in Render's environment, which the app uses to mark
   the session cookie `Secure` (see `main.py`) — no manual toggle needed.

**Free tier caveats worth knowing:**
- Render's free Postgres is deleted after 90 days of the *database's* age
  unless upgraded to a paid plan — fine for testing, not for indefinite
  production use. Set a reminder to upgrade before then, or your users'
  refresh tokens and mail history will be lost.
- Render's free web service spins down after 15 minutes of no incoming
  HTTP traffic and takes ~30-60s to wake on the next request. APScheduler
  jobs only fire while the process is actually running, so the 06:00 daily
  email and 30-minute sync **will be silently skipped while spun down**,
  unless something pings the service to keep it awake (e.g. an external
  uptime monitor hitting it every few minutes) or you upgrade to a paid
  instance that never sleeps.
