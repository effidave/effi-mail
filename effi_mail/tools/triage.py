"""Triage tools for effi-mail MCP server."""

import json
from typing import List

from effi_mail.helpers import outlook


def triage_email(email_id: str, status: str) -> str:
    """Assign triage status to an email using Outlook categories.
    
    Status is stored in the email itself as an Outlook category.
    
    Args:
        email_id: Email EntryID
        status: Triage status - 'processed', 'deferred', or 'archived'
        
    Returns:
        JSON string with success/error status
    """
    success = outlook.set_triage_status(email_id, status)
    if success:
        return json.dumps({"success": True, "email_id": email_id, "status": status})
    return json.dumps({"error": f"Failed to set triage status on {email_id}"})


def batch_triage(email_ids: List[str], status: str) -> str:
    """Triage multiple emails at once with the same status.
    
    Args:
        email_ids: List of email EntryIDs
        status: Triage status to apply - 'processed', 'deferred', or 'archived'
        
    Returns:
        JSON string with count of successful/failed operations
    """
    results = outlook.batch_set_triage_status(email_ids, status)
    
    return json.dumps({
        "success": results["failed"] == 0,
        "triaged": results["success"],
        "failed": results["failed"],
        "status": status
    })


def batch_archive_domain(domain: str, days: int = 30) -> str:
    """Archive all pending emails from a specific domain.
    
    Useful for marketing emails. Gets pending emails from Outlook and archives them.
    
    Args:
        domain: Domain to archive all emails from
        days: Days to look back (default: 30)
        
    Returns:
        JSON string with count of archived emails
    """
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
