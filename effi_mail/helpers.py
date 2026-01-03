"""Helper functions for effi-mail MCP server."""

import json
from typing import Any
from outlook_client import OutlookClient


# Shared Outlook client instance
outlook = OutlookClient()


def truncate_text(text: str, max_length: int = 500) -> str:
    """Truncate text to max length with indicator.
    
    Args:
        text: Text to truncate
        max_length: Maximum length before truncation
        
    Returns:
        Original text if under max_length, otherwise truncated with indicator
    """
    if len(text) <= max_length:
        return text
    return text[:max_length] + f"... [{len(text) - max_length} more chars]"


def format_email_summary(email: Any, include_preview: bool = False, include_recipients: bool = False) -> dict:
    """Format email for MCP response.
    
    Args:
        email: Email object from OutlookClient
        include_preview: Include body preview in response
        include_recipients: Include recipient lists in response
        
    Returns:
        Dict with email summary fields
    """
    result = {
        "id": email.id,
        "subject": email.subject,
        "sender": f"{email.sender_name} <{email.sender_email}>",
        "domain": email.domain,
        "received": email.received_time.isoformat(),
        "has_attachments": email.has_attachments,
        "direction": email.direction,
    }
    # Get triage status from Outlook categories
    triage = outlook.get_triage_status(email.id)
    if triage:
        result["triage_status"] = triage
    if include_preview:
        result["preview"] = truncate_text(email.body_preview, 200)
    if include_recipients:
        result["recipients_to"] = email.recipients_to
        result["recipients_cc"] = email.recipients_cc
    return result
