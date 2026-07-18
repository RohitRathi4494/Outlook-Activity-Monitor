"""Background jobs, started/stopped from main.py's FastAPI startup/shutdown
events:

  - poll_all_users:      every 30 minutes, refresh every stored user's
                         mailbox using THAT user's own refresh token.
  - send_daily_reports:  once a day, email each user their own previous
                         day's Excel activity report to their own mailbox.
"""

import logging
from datetime import date, timedelta

from apscheduler.schedulers.asyncio import AsyncIOScheduler

from collector import collect_for_user
from mailer import send_daily_report
from models import SessionLocal, User

logger = logging.getLogger("scheduler")

scheduler = AsyncIOScheduler()

POLL_INTERVAL_MINUTES = 30

# Daily report email: sent at 06:00 server-local time, covering the
# previous full calendar day (00:00-23:59) so the report is always complete.
DAILY_REPORT_HOUR = 6
DAILY_REPORT_MINUTE = 0


async def poll_all_users():
    """Loop over every stored user and refresh their mailbox independently.
    A failure for one user (e.g. revoked consent) must not block the others.
    """
    db = SessionLocal()
    try:
        user_ids = [u.id for u in db.query(User.id).all()]
    finally:
        db.close()

    logger.info("Scheduled poll starting for %d user(s).", len(user_ids))
    for user_id in user_ids:
        try:
            await collect_for_user(user_id)
        except Exception:
            logger.exception("Scheduled collection failed for user %s", user_id)


async def send_daily_reports():
    """Once a day, make sure yesterday's mail is fully synced, then email
    every stored user their own report for yesterday. A failure for one user
    must not block the others.
    """
    db = SessionLocal()
    try:
        users = [(u.id, u.email) for u in db.query(User.id, User.email).all()]
    finally:
        db.close()

    report_date = (date.today() - timedelta(days=1)).isoformat()
    logger.info("Daily report run starting for %d user(s), date=%s.", len(users), report_date)
    for user_id, email in users:
        try:
            await collect_for_user(user_id)
            await send_daily_report(user_id, email, report_date)
        except Exception:
            logger.exception("Daily report email failed for user %s", user_id)


def start_scheduler():
    if not scheduler.running:
        scheduler.add_job(
            poll_all_users,
            trigger="interval",
            minutes=POLL_INTERVAL_MINUTES,
            id="poll_all_users",
            replace_existing=True,
        )
        scheduler.add_job(
            send_daily_reports,
            trigger="cron",
            hour=DAILY_REPORT_HOUR,
            minute=DAILY_REPORT_MINUTE,
            id="send_daily_reports",
            replace_existing=True,
        )
        scheduler.start()
        logger.info(
            "Scheduler started: polling every %d minutes, daily report email at %02d:%02d.",
            POLL_INTERVAL_MINUTES,
            DAILY_REPORT_HOUR,
            DAILY_REPORT_MINUTE,
        )
