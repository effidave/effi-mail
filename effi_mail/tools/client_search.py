"""Client search tools for effi-mail MCP server."""

import json
from datetime import datetime, time
from typing import Optional

from effi_mail.helpers import outlook, format_email_summary
from effi_work_client import get_client_identifiers_from_effi_work


async def get_emails_by_client(
    client_id: str,
    days: int = 30,
    date_from: Optional[str] = None,
    date_to: Optional[str] = None,
    limit: int = 100
) -> str:
    """Get client correspondence. client_id is case-insensitive."""
    # Parse dates
    # date_from: start of day (00:00:00)
    # date_to: end of day (23:59:59) to include all emails on that date
    date_from_dt = datetime.strptime(date_from, "%Y-%m-%d") if date_from else None
    date_to_dt = datetime.combine(
        datetime.strptime(date_to, "%Y-%m-%d").date(), time(23, 59, 59)
    ) if date_to else None
    
    # Get client identifiers from effi-core (fresh data)
    identifiers = await get_client_identifiers_from_effi_work(client_id)
    if not identifiers.get("domains"):
        return json.dumps({
            "error": f"Client not found: {client_id}",
            "source": identifiers.get("source"),
            "hint": "Use list_dms_clients to find the exact client name. Client names often include 'Ltd', 'Limited', etc."
        })
    
    # Search Outlook directly
    emails = outlook.search_outlook_by_identifiers(
        domains=identifiers["domains"],
        contact_emails=identifiers.get("contact_emails", []),
        days=days,
        date_from=date_from_dt,
        date_to=date_to_dt,
        limit=limit,
    )
    
    return json.dumps({
        "client_id": client_id,
        "identifiers": identifiers,
        "count": len(emails),
        "emails": [format_email_summary(e, include_preview=True, include_recipients=True) for e in emails]
    }, indent=2)


def search_outlook_direct(
    sender_domain: Optional[str] = None,
    sender_email: Optional[str] = None,
    recipient_domain: Optional[str] = None,
    recipient_email: Optional[str] = None,
    subject_contains: Optional[str] = None,
    body_contains: Optional[str] = None,
    date_from: Optional[str] = None,
    date_to: Optional[str] = None,
    days: int = 30,
    folder: str = "Inbox",
    limit: int = 50
) -> str:
    """Search Outlook with filters. folder: 'Inbox' or 'Sent Items'. Dates: YYYY-MM-DD."""
    # Parse dates
    # date_from: start of day (00:00:00)
    # date_to: end of day (23:59:59) to include all emails on that date
    date_from_dt = datetime.strptime(date_from, "%Y-%m-%d") if date_from else None
    date_to_dt = datetime.combine(
        datetime.strptime(date_to, "%Y-%m-%d").date(), time(23, 59, 59)
    ) if date_to else None
    
    emails = outlook.search_outlook(
        sender_domain=sender_domain,
        sender_email=sender_email,
        recipient_domain=recipient_domain,
        recipient_email=recipient_email,
        subject_contains=subject_contains,
        body_contains=body_contains,
        date_from=date_from_dt,
        date_to=date_to_dt,
        days=days,
        folder=folder,
        limit=limit,
    )
    
    return json.dumps({
        "folder": folder,
        "count": len(emails),
        "emails": [format_email_summary(e, include_preview=True, include_recipients=True) for e in emails]
    }, indent=2)
