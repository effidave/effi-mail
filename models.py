"""Data models for email entities."""

from dataclasses import dataclass, field
from datetime import datetime
from typing import Optional, List
from enum import Enum


# Domains that should never be auto-categorized (require specific contact_email matching)
GENERIC_DOMAINS = {
    # Personal email providers
    'gmail.com', 'googlemail.com',
    'outlook.com', 'hotmail.com', 'hotmail.co.uk', 'live.com', 'live.co.uk',
    'yahoo.com', 'yahoo.co.uk',
    'icloud.com', 'me.com', 'mac.com',
    'protonmail.com', 'proton.me',
    'aol.com',
    # UK ISP email
    'btinternet.com', 'btopenworld.com',
    'sky.com', 'skynet.be',
    'virginmedia.com', 'virgin.net',
    'talktalk.net',
    'plusnet.com',
}


class EmailCategory(Enum):
    """Category for domain classification."""
    CLIENT = "Client"
    INTERNAL = "Internal"
    MARKETING = "Marketing"
    PERSONAL = "Personal"
    THIRD_PARTY = "Third-Party"
    UNCATEGORIZED = "Uncategorized"


class TriageStatus(Enum):
    """Triage status for emails."""
    PENDING = "pending"
    PROCESSED = "processed"
    DEFERRED = "deferred"
    ARCHIVED = "archived"


class EmailDirection(Enum):
    """Direction of email (inbound or outbound)."""
    INBOUND = "inbound"
    OUTBOUND = "outbound"


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
    direction: str = "inbound"  # 'inbound' or 'outbound'
    recipients_to: List[str] = field(default_factory=list)  # JSON array of To addresses
    recipients_cc: List[str] = field(default_factory=list)  # JSON array of CC addresses
    recipient_domains: str = ""  # Comma-separated domains from To/CC (computed at sync)
    internet_message_id: Optional[str] = None  # RFC2822 Message-ID (permanent identifier)
    
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
    """Represents a client (metadata only - domains come from effi-clients)."""
    id: str
    name: str
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


@dataclass
class Counterparty:
    """Represents a counterparty in a matter."""
    id: str
    matter_id: str
    name: str
    contact_name: Optional[str] = None
    contact_email: Optional[str] = None
    domains: List[str] = field(default_factory=list)  # Email domains associated with this counterparty
    notes: str = ""
