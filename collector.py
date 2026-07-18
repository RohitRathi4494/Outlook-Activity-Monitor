"""Fetch a single user's own received + sent mail from Graph, detect forwards,
deduplicate, and store everything tagged with that user's id.

collect_for_user() is the only entry point. It is called:
  - once, right after a user first signs in (main.py)
  - every 30 minutes for every stored user (scheduler.py)

It always uses THAT user's own access token (via auth.get_access_token_for_user)
and Graph's /me endpoints, so it can never read another user's mailbox.
"""

import logging
from datetime import datetime
from typing import Optional

from auth import get_access_token_for_user
from graph_client import GRAPH_BASE, get_all_pages
from models import Message, SessionLocal

logger = logging.getLogger("collector")

# Fields needed to populate the report + do forward matching.
_BASE_SELECT = (
    "id,from,subject,receivedDateTime,toRecipients,ccRecipients,"
    "hasAttachments,importance,conversationId,internetMessageId"
)
RECEIVED_SELECT = _BASE_SELECT
# Sent items additionally need sentDateTime, and internetMessageHeaders so we
# can check the References / In-Reply-To headers for forward detection.
SENT_SELECT = _BASE_SELECT + ",sentDateTime,internetMessageHeaders"

FORWARD_PREFIXES = ("fw:", "fwd:")
PAGE_SIZE = 50


def _parse_graph_datetime(value: Optional[str]) -> Optional[datetime]:
    """Parse a Graph ISO-8601 UTC timestamp like '2024-05-01T12:34:56Z' or
    '2024-05-01T12:34:56.1234567Z' into a naive UTC datetime."""
    if not value:
        return None
    value = value.rstrip("Z")
    if "." in value:
        date_part, frac = value.split(".", 1)
        value = f"{date_part}.{frac[:6].ljust(6, '0')}"
        return datetime.strptime(value, "%Y-%m-%dT%H:%M:%S.%f")
    return datetime.strptime(value, "%Y-%m-%dT%H:%M:%S")


def _format_recipients(recipients: Optional[list]) -> str:
    """Render a Graph recipient list as 'Name <email>; Name2 <email2>'."""
    if not recipients:
        return ""
    parts = []
    for recipient in recipients:
        addr = (recipient or {}).get("emailAddress", {}) or {}
        name = addr.get("name") or ""
        email = addr.get("address") or ""
        parts.append(f"{name} <{email}>" if name else email)
    return "; ".join(p for p in parts if p)


def _get_header_value(headers: Optional[list], header_name: str) -> str:
    """Look up a header value (case-insensitive) from Graph's internetMessageHeaders."""
    if not headers:
        return ""
    target = header_name.lower()
    for h in headers:
        if (h.get("name") or "").lower() == target:
            return h.get("value") or ""
    return ""


def _is_forward_subject(subject: Optional[str]) -> bool:
    return (subject or "").strip().lower().startswith(FORWARD_PREFIXES)


def _find_matching_forward(received_msg: dict, forward_candidates: list[dict]) -> Optional[dict]:
    """Find the earliest sent 'forward' item that corresponds to received_msg,
    per the spec: shares conversationId OR references the received message's
    internetMessageId (via References/In-Reply-To headers), AND has a
    FW:/Fwd: subject. Returns the matching sent-item dict, or None.
    """
    conversation_id = received_msg.get("conversationId")
    internet_message_id = received_msg.get("internetMessageId") or ""
    received_time = _parse_graph_datetime(received_msg.get("receivedDateTime"))

    best_candidate = None
    best_sent_time = None

    for candidate in forward_candidates:
        shares_conversation = bool(conversation_id) and candidate.get("conversationId") == conversation_id

        references = _get_header_value(candidate.get("internetMessageHeaders"), "References")
        in_reply_to = _get_header_value(candidate.get("internetMessageHeaders"), "In-Reply-To")
        references_message = bool(internet_message_id) and (
            internet_message_id in references or internet_message_id in in_reply_to
        )

        if not (shares_conversation or references_message):
            continue

        sent_time = _parse_graph_datetime(candidate.get("sentDateTime"))
        # A forward can't happen before the original was received.
        if received_time and sent_time and sent_time < received_time:
            continue

        if best_candidate is None or (
            sent_time and (best_sent_time is None or sent_time < best_sent_time)
        ):
            best_candidate = candidate
            best_sent_time = sent_time

    return best_candidate


async def collect_for_user(user_id: str) -> int:
    """Fetch, dedupe, and store new mail activity for a single user.
    Returns the number of new rows inserted. Silently no-ops (returns 0) if a
    valid access token cannot be obtained for the user (e.g. they revoked
    consent) — it does NOT fall back to any other credentials.
    """
    access_token = get_access_token_for_user(user_id)
    if not access_token:
        logger.warning("No valid access token for user %s; skipping collection.", user_id)
        return 0

    received_raw = await get_all_pages(
        access_token,
        f"{GRAPH_BASE}/me/messages",
        params={"$select": RECEIVED_SELECT, "$top": str(PAGE_SIZE)},
    )
    sent_raw = await get_all_pages(
        access_token,
        f"{GRAPH_BASE}/me/mailFolders/sentitems/messages",
        params={"$select": SENT_SELECT, "$top": str(PAGE_SIZE)},
    )

    forward_candidates = [m for m in sent_raw if _is_forward_subject(m.get("subject"))]

    inserted = 0
    db = SessionLocal()
    try:
        existing_received_ids = {
            row.message_id
            for row in db.query(Message.message_id)
            .filter(Message.user_id == user_id, Message.direction == "received")
            .all()
        }
        existing_sent_ids = {
            row.message_id
            for row in db.query(Message.message_id)
            .filter(Message.user_id == user_id, Message.direction == "sent")
            .all()
        }

        for m in received_raw:
            if m["id"] in existing_received_ids:
                continue

            forward = _find_matching_forward(m, forward_candidates)
            from_addr = (m.get("from") or {}).get("emailAddress", {}) or {}

            db.add(
                Message(
                    user_id=user_id,
                    message_id=m["id"],
                    direction="received",
                    from_name=from_addr.get("name"),
                    from_email=from_addr.get("address"),
                    subject=m.get("subject"),
                    received_datetime=_parse_graph_datetime(m.get("receivedDateTime")),
                    to_recipients=_format_recipients(m.get("toRecipients")),
                    cc_recipients=_format_recipients(m.get("ccRecipients")),
                    has_attachments=bool(m.get("hasAttachments")),
                    importance=m.get("importance"),
                    conversation_id=m.get("conversationId"),
                    internet_message_id=m.get("internetMessageId"),
                    forwarded=bool(forward),
                    forwarded_to=_format_recipients(forward.get("toRecipients")) if forward else None,
                    forwarded_time=_parse_graph_datetime(forward.get("sentDateTime")) if forward else None,
                )
            )
            existing_received_ids.add(m["id"])
            inserted += 1

        for m in sent_raw:
            if m["id"] in existing_sent_ids:
                continue

            from_addr = (m.get("from") or {}).get("emailAddress", {}) or {}

            db.add(
                Message(
                    user_id=user_id,
                    message_id=m["id"],
                    direction="sent",
                    from_name=from_addr.get("name"),
                    from_email=from_addr.get("address"),
                    subject=m.get("subject"),
                    sent_datetime=_parse_graph_datetime(m.get("sentDateTime")),
                    to_recipients=_format_recipients(m.get("toRecipients")),
                    cc_recipients=_format_recipients(m.get("ccRecipients")),
                    has_attachments=bool(m.get("hasAttachments")),
                    importance=m.get("importance"),
                    conversation_id=m.get("conversationId"),
                    internet_message_id=m.get("internetMessageId"),
                )
            )
            existing_sent_ids.add(m["id"])
            inserted += 1

        db.commit()
    except Exception:
        db.rollback()
        raise
    finally:
        db.close()

    logger.info("User %s: inserted %d new message row(s).", user_id, inserted)
    return inserted
