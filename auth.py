"""MSAL-based OAuth 2.0 Authorization Code flow (delegated permissions only).

Flow:
  1. build_auth_url()            -> redirect the browser to Microsoft's login page
  2. exchange_code_for_user(code) -> called from /auth/callback, stores the
     user's ENCRYPTED refresh token and returns basic profile info
  3. get_access_token_for_user(user_id) -> mints a fresh access token from the
     stored refresh token whenever the app (route or scheduler) needs to call
     Graph on that user's behalf. Never touches another user's token.

No passwords are ever handled by this app (Basic Auth is deprecated / disabled
by Microsoft); no application permissions are used — every Graph call is made
with a token scoped to exactly the signed-in user via /me.
"""

import os
from datetime import datetime
from typing import Optional

import msal
from dotenv import load_dotenv

from crypto_utils import decrypt_token, encrypt_token
from models import SessionLocal, User

load_dotenv()

CLIENT_ID = os.environ["CLIENT_ID"]
TENANT_ID = os.environ["TENANT_ID"]
CLIENT_SECRET = os.environ["CLIENT_SECRET"]
REDIRECT_URI = os.environ["REDIRECT_URI"]

AUTHORITY = f"https://login.microsoftonline.com/{TENANT_ID}"

# Delegated scopes only. offline_access + openid + profile are reserved scopes
# that MSAL's confidential client automatically appends for the auth-code
# flow, which is what makes a refresh_token come back in the token response.
# Mail.Send is required to email the daily report to the user's own mailbox.
SCOPES = ["Mail.Read", "Mail.Send", "User.Read"]


def _msal_app() -> msal.ConfidentialClientApplication:
    return msal.ConfidentialClientApplication(
        client_id=CLIENT_ID,
        client_credential=CLIENT_SECRET,
        authority=AUTHORITY,
    )


def build_auth_url(state: str) -> str:
    """Build the Microsoft identity platform authorization URL to redirect the user to."""
    app = _msal_app()
    return app.get_authorization_request_url(
        scopes=SCOPES,
        state=state,
        redirect_uri=REDIRECT_URI,
    )


def exchange_code_for_user(code: str) -> dict:
    """Exchange the authorization code for tokens, persist the user's ENCRYPTED
    refresh token, and return {user_id, email, display_name, access_token}.

    user_id is the token's 'oid' claim (Entra object id) — this is the primary
    key everything else in the app filters on to enforce per-user isolation.
    """
    app = _msal_app()
    result = app.acquire_token_by_authorization_code(
        code=code,
        scopes=SCOPES,
        redirect_uri=REDIRECT_URI,
    )

    if "error" in result:
        raise RuntimeError(f"{result.get('error')}: {result.get('error_description')}")

    refresh_token = result.get("refresh_token")
    if not refresh_token:
        raise RuntimeError(
            "Microsoft did not return a refresh token. Make sure 'offline_access' "
            "is consented for this app and that admin consent has been granted "
            "for Mail.Read / User.Read in Azure Portal -> API permissions."
        )

    claims = result.get("id_token_claims", {}) or {}
    user_id = claims.get("oid") or claims.get("sub")
    email = claims.get("preferred_username") or claims.get("email") or ""
    display_name = claims.get("name") or email

    if not user_id:
        raise RuntimeError("Could not determine user identity from ID token claims.")

    db = SessionLocal()
    try:
        encrypted = encrypt_token(refresh_token)
        user = db.get(User, user_id)
        if user:
            user.email = email
            user.display_name = display_name
            user.encrypted_refresh_token = encrypted
            user.updated_at = datetime.utcnow()
        else:
            user = User(
                id=user_id,
                email=email,
                display_name=display_name,
                encrypted_refresh_token=encrypted,
            )
            db.add(user)
        db.commit()
    finally:
        db.close()

    return {
        "user_id": user_id,
        "email": email,
        "display_name": display_name,
        "access_token": result["access_token"],
    }


def get_access_token_for_user(user_id: str) -> Optional[str]:
    """Mint a fresh Graph access token for user_id using THEIR stored refresh
    token. Returns None if there is no user / refresh fails (e.g. revoked
    consent, expired refresh token) — callers should treat None as
    "this user needs to sign in again" and must not fall back to any other
    user's credentials.
    """
    db = SessionLocal()
    try:
        user = db.get(User, user_id)
        if not user or not user.encrypted_refresh_token:
            return None
        refresh_token = decrypt_token(user.encrypted_refresh_token)
    finally:
        db.close()

    app = _msal_app()
    result = app.acquire_token_by_refresh_token(refresh_token, scopes=SCOPES)

    if "error" in result:
        return None

    # Microsoft may rotate the refresh token; persist the latest one.
    new_refresh_token = result.get("refresh_token", refresh_token)
    db = SessionLocal()
    try:
        user = db.get(User, user_id)
        if user:
            user.encrypted_refresh_token = encrypt_token(new_refresh_token)
            user.updated_at = datetime.utcnow()
            db.commit()
    finally:
        db.close()

    return result["access_token"]
