"""SQLAlchemy models: User (one row per signed-in mailbox owner) and Message
(one row per received or sent mail item, always tagged with the owning user_id).

Every query in this app MUST filter Message by user_id — that is the entire
isolation boundary between users. See report.py and collector.py.
"""

import os
from datetime import datetime

from sqlalchemy import (
    Boolean,
    Column,
    DateTime,
    ForeignKey,
    Integer,
    String,
    Text,
    UniqueConstraint,
    create_engine,
)
from sqlalchemy.orm import declarative_base, relationship, sessionmaker

DB_PATH = os.path.join(os.path.dirname(os.path.abspath(__file__)), "outlook_activity.db")
engine = create_engine(f"sqlite:///{DB_PATH}", connect_args={"check_same_thread": False})
SessionLocal = sessionmaker(bind=engine, autoflush=False, autocommit=False)
Base = declarative_base()


class User(Base):
    """A signed-in mailbox owner. Primary key is the Entra ID object id (oid) —
    stable across sessions and unique per user in the tenant."""

    __tablename__ = "users"

    id = Column(String, primary_key=True)
    email = Column(String, nullable=False, index=True)
    display_name = Column(String)
    encrypted_refresh_token = Column(Text, nullable=False)
    created_at = Column(DateTime, default=datetime.utcnow)
    updated_at = Column(DateTime, default=datetime.utcnow, onupdate=datetime.utcnow)

    messages = relationship("Message", back_populates="user", cascade="all, delete-orphan")


class Message(Base):
    """A single received or sent mail item, always scoped to user_id.

    'direction' distinguishes received items (which populate the Excel report)
    from sent items (which are stored for completeness and used to detect
    forwards). Forward metadata is computed once at collection time and
    stored directly on the received row.
    """

    __tablename__ = "messages"

    id = Column(Integer, primary_key=True, autoincrement=True)
    user_id = Column(String, ForeignKey("users.id"), nullable=False, index=True)
    message_id = Column(String, nullable=False, index=True)  # Graph message id
    direction = Column(String, nullable=False)  # 'received' or 'sent'

    from_name = Column(String)
    from_email = Column(String)
    subject = Column(String)

    received_datetime = Column(DateTime, index=True)
    sent_datetime = Column(DateTime, index=True)

    to_recipients = Column(Text)  # "Name <email>; Name2 <email2>"
    cc_recipients = Column(Text)
    has_attachments = Column(Boolean, default=False)
    importance = Column(String)
    conversation_id = Column(String, index=True)
    internet_message_id = Column(String)

    forwarded = Column(Boolean, default=False)
    forwarded_to = Column(Text)
    forwarded_time = Column(DateTime)

    user = relationship("User", back_populates="messages")

    __table_args__ = (
        UniqueConstraint("user_id", "message_id", "direction", name="uq_user_message_direction"),
    )


def init_db():
    Base.metadata.create_all(bind=engine)
