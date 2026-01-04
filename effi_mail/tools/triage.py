"""Triage tools for effi-mail MCP server."""

import json
from typing import List

from effi_mail.helpers import outlook


def triage_email(email_id: str, status: str) -> str:
    """Set triage status ('processed', 'deferred', 'archived') on an email."""
    success = outlook.set_triage_status(email_id, status)
    if success:
        return json.dumps({"success": True, "email_id": email_id, "status": status})
    return json.dumps({"error": f"Failed to set triage status on {email_id}"})


def batch_triage(email_ids: List[str], status: str) -> str:
    """Triage multiple emails with the same status."""
    results = outlook.batch_set_triage_status(email_ids, status)
    
    return json.dumps({
        "success": results["failed"] == 0,
        "triaged": results["success"],
        "failed": results["failed"],
        "status": status
    })


def batch_archive_domain(domain: str, days: int = 30) -> str:
    """Archive all pending emails from a domain. Useful for marketing cleanup."""
    # Get pending emails from this domain
    pending_emails = outlook.get_pending_emails_from_domain(domain, days=days)
    
    # Archive them all
    archived = 0
    for email in pending_emails:
        if outlook.set_triage_status(email.id, "archived"):
            archived += 1
    
    return json.dumps({
        "success": True,
        "domain": domain,
        "archived_count": archived
    })
