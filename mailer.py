"""Email a user's own daily Excel activity report to their own mailbox.

Uses Microsoft Graph's delegated /me/sendMail endpoint with THAT user's own
access token, so the report always lands in (and is sent from) the same
mailbox it was generated from — never a different user's. Requires the
Mail.Send delegated scope in addition to Mail.Read (see auth.SCOPES).
"""

import base64
import logging

import httpx

from auth import get_access_token_for_user
from graph_client import GRAPH_BASE
from report import generate_report

logger = logging.getLogger("mailer")

ATTACHMENT_CONTENT_TYPE = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
SEND_TIMEOUT = 30.0


async def send_daily_report(user_id: str, email: str, report_date: str) -> bool:
    """Generate report_date's Excel report for user_id and send it as an
    attachment to their own mailbox (email). Returns True if sent, False if
    skipped because no valid access token could be obtained for this user
    (e.g. revoked consent — the caller should not retry with any other
    user's credentials).
    """
    access_token = get_access_token_for_user(user_id)
    if not access_token:
        logger.warning("No valid access token for user %s; skipping report email.", user_id)
        return False

    buffer = generate_report(user_id, report_date)
    attachment_b64 = base64.b64encode(buffer.read()).decode("utf-8")
    filename = f"outlook_activity_{report_date}.xlsx"

    payload = {
        "message": {
            "subject": f"Your Outlook Activity Report - {report_date}",
            "body": {
                "contentType": "Text",
                "content": f"Attached is your Outlook activity report for {report_date}.",
            },
            "toRecipients": [{"emailAddress": {"address": email}}],
            "attachments": [
                {
                    "@odata.type": "#microsoft.graph.fileAttachment",
                    "name": filename,
                    "contentType": ATTACHMENT_CONTENT_TYPE,
                    "contentBytes": attachment_b64,
                }
            ],
        },
        "saveToSentItems": "true",
    }

    async with httpx.AsyncClient(timeout=SEND_TIMEOUT) as client:
        response = await client.post(
            f"{GRAPH_BASE}/me/sendMail",
            headers={"Authorization": f"Bearer {access_token}"},
            json=payload,
        )
        response.raise_for_status()

    logger.info("Emailed user %s's report for %s to %s.", user_id, report_date, email)
    return True
