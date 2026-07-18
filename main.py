"""Outlook Activity Monitor — FastAPI app.

Routes:
  GET  /               landing page / "Sign in with Microsoft"
  GET  /login          redirect to Microsoft's OAuth authorize endpoint
  GET  /auth/callback  exchange code for tokens, start a session, run a first sync
  GET  /logout         clear the session
  GET  /dashboard      shows the signed-in user's email + report controls
  GET  /report         download that user's Excel report for a given date

Every route that needs the current user resolves it strictly from the signed
session cookie (never from a query/body parameter), and every DB query is
filtered by that user's id — see _current_user() and report.py.
"""

import hashlib
import logging
import os
import secrets
from datetime import date as date_cls
from datetime import datetime

from dotenv import load_dotenv

load_dotenv()

from fastapi import FastAPI, Request
from fastapi.responses import HTMLResponse, PlainTextResponse, RedirectResponse, StreamingResponse
from fastapi.templating import Jinja2Templates
from starlette.middleware.sessions import SessionMiddleware

from auth import build_auth_url, exchange_code_for_user
from collector import collect_for_user
from models import SessionLocal, User, init_db
from report import generate_report
from scheduler import scheduler, start_scheduler

logging.basicConfig(level=logging.INFO, format="%(asctime)s %(levelname)s %(name)s: %(message)s")
logger = logging.getLogger("main")

BASE_DIR = os.path.dirname(os.path.abspath(__file__))

app = FastAPI(title="Outlook Activity Monitor")

# Sessions hold only a user id (the Entra 'oid'), never a token or secret.
# Signed (not encrypted) is fine for that — tampering is prevented by the
# signature, and there's nothing confidential in the payload. We derive the
# signing secret from ENCRYPTION_KEY so no extra value is needed in .env.
_session_secret = hashlib.sha256(os.environ["ENCRYPTION_KEY"].encode()).hexdigest()
app.add_middleware(
    SessionMiddleware,
    secret_key=_session_secret,
    session_cookie="oam_session",
    https_only=False,  # set True when served over HTTPS in production
)

templates = Jinja2Templates(directory=os.path.join(BASE_DIR, "templates"))


@app.on_event("startup")
def on_startup():
    init_db()
    start_scheduler()


@app.on_event("shutdown")
def on_shutdown():
    if scheduler.running:
        scheduler.shutdown(wait=False)


def _current_user(request: Request) -> User | None:
    """Resolve the signed-in user strictly from the session cookie. Returns
    None if there is no session or the user no longer exists."""
    user_id = request.session.get("user_id")
    if not user_id:
        return None
    db = SessionLocal()
    try:
        return db.get(User, user_id)
    finally:
        db.close()


@app.get("/", response_class=HTMLResponse)
def index(request: Request):
    if _current_user(request):
        return RedirectResponse(url="/dashboard")
    return templates.TemplateResponse("index.html", {"request": request})


@app.get("/login")
def login(request: Request):
    # CSRF protection for the OAuth redirect: verified against the session in /auth/callback.
    state = secrets.token_urlsafe(16)
    request.session["oauth_state"] = state
    return RedirectResponse(url=build_auth_url(state))


@app.get("/auth/callback")
async def auth_callback(
    request: Request,
    code: str | None = None,
    state: str | None = None,
    error: str | None = None,
    error_description: str | None = None,
):
    if error:
        return HTMLResponse(
            f"<h3>Sign-in failed</h3><p>{error}: {error_description}</p><p><a href='/'>Back</a></p>",
            status_code=400,
        )

    expected_state = request.session.pop("oauth_state", None)
    if not code or not state or state != expected_state:
        return HTMLResponse(
            "<h3>Invalid authentication response.</h3><p><a href='/'>Back</a></p>", status_code=400
        )

    try:
        user_info = exchange_code_for_user(code)
    except RuntimeError as exc:
        return HTMLResponse(f"<h3>Sign-in failed</h3><p>{exc}</p><p><a href='/'>Back</a></p>", status_code=400)

    request.session["user_id"] = user_info["user_id"]
    request.session["email"] = user_info["email"]

    # Best-effort initial sync so the dashboard has data right away; the
    # scheduler will keep this fresh every 30 minutes regardless of outcome.
    try:
        await collect_for_user(user_info["user_id"])
    except Exception:
        logger.exception("Initial mail collection failed for user %s", user_info["user_id"])

    return RedirectResponse(url="/dashboard")


@app.get("/logout")
def logout(request: Request):
    request.session.clear()
    return RedirectResponse(url="/")


@app.get("/dashboard", response_class=HTMLResponse)
def dashboard(request: Request):
    user = _current_user(request)
    if not user:
        return RedirectResponse(url="/")
    return templates.TemplateResponse(
        "dashboard.html",
        {"request": request, "user": user, "today": date_cls.today().isoformat()},
    )


@app.get("/report")
def report(request: Request, date: str | None = None):
    user = _current_user(request)
    if not user:
        return RedirectResponse(url="/")

    report_date = date or date_cls.today().isoformat()
    try:
        datetime.strptime(report_date, "%Y-%m-%d")
    except ValueError:
        return PlainTextResponse("Invalid date format, expected YYYY-MM-DD.", status_code=400)

    buffer = generate_report(user.id, report_date)
    filename = f"outlook_activity_{report_date}.xlsx"
    return StreamingResponse(
        buffer,
        media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        headers={"Content-Disposition": f'attachment; filename="{filename}"'},
    )
