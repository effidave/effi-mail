"""Triage tools for effi-mail MCP server."""

import json
from typing import List

from effi_mail.helpers import outlook


def triage_email(email_id: str, status: str) -> str:
    """Set triage status ('action', 'waiting', 'processed', 'archived') on an email."""
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


def archive_email(email_id: str, folder: str = "Archive", create_path: bool = False) -> str:
    r"""Move an email to a folder (default: Archive).
    
    This physically moves the email from Inbox to the specified folder.
    Note: The email's EntryID will change after the move.
    
    Supports folder paths with subfolders:
      - "Archive" (root-level folder)
      - "Inbox\~Zero\Growth Engineering" (subfolder path)
    
    Args:
        email_id: Outlook EntryID of the email to archive
        folder: Folder name or path (default: "Archive")
        create_path: If True, create missing folders in the path (default: False)
        
    Returns:
        JSON with success status, old_id, new_id, folder path, and folders_created if any
    """
    result = outlook.move_to_archive(email_id, folder_path=folder, create_path=create_path)
    return json.dumps(result)


def batch_archive_emails(email_ids: List[str], folder: str = "Archive", create_path: bool = False) -> str:
    r"""Move multiple emails to a folder (default: Archive).
    
    This physically moves emails from Inbox to the specified folder.
    Note: Email EntryIDs will change after the move.
    
    Supports folder paths with subfolders:
      - "Archive" (root-level folder)
      - "Inbox\~Zero\Growth Engineering" (subfolder path)
    
    Args:
        email_ids: List of Outlook EntryIDs to archive
        folder: Folder name or path (default: "Archive")
        create_path: If True, create missing folders in the path (default: False)
        
    Returns:
        JSON with success/failed counts, moved email details, and folders_created if any
    """
    result = outlook.batch_move_to_archive(email_ids, folder_path=folder, create_path=create_path)
    return json.dumps(result)
