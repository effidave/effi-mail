"""Data models for email entities."""

from dataclasses import dataclass, field
from datetime import datetime
from typing import Optional, List
from enum import Enum


class EmailCategory(Enum):
    """Category for domain classification."""
    CLIENT = "Client"
    INTERNAL = "Internal"
    MARKETING = "Marketing"
    PERSONAL = "Personal"
    UNCATEGORIZED = "Uncategorized"


class TriageStatus(Enum):
    """Triage status for emails."""
    PENDING = "pending"
    PROCESSED = "processed"
    DEFERRED = "deferred"
    ARCHIVED = "archived"


@dataclass
class Email:
    """Represents an email message."""
    id: str  # Outlook EntryID
    subject: str
    sender_name: str
    sender_email: str
    domain: str
    received_time: datetime
    body_preview: str = ""
    has_attachments: bool = False
    attachment_names: List[str] = field(default_factory=list)
    categories: str = ""
    conversation_id: Optional[str] = None
    folder_path: str = "Inbox"
    
    # Triage fields
    triage_status: TriageStatus = TriageStatus.PENDING
    client_id: Optional[str] = None
    matter_id: Optional[str] = None
    processed_at: Optional[datetime] = None
    notes: str = ""


@dataclass
class Domain:
    """Represents a sender domain with its category."""
    name: str
    category: EmailCategory = EmailCategory.UNCATEGORIZED
    email_count: int = 0
    last_seen: Optional[datetime] = None
    sample_senders: List[str] = field(default_factory=list)


@dataclass
class Client:
    """Represents a client."""
    id: str
    name: str
    domains: List[str] = field(default_factory=list)
    folder_path: Optional[str] = None  # Path in effi-work


@dataclass
class Matter:
    """Represents a matter/case for a client."""
    id: str
    client_id: str
    name: str
    description: str = ""
    folder_path: Optional[str] = None  # Path in effi-work
    active: bool = True
