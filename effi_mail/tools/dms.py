"""DMS (DMSforLegal) tools for effi-mail MCP server."""

import json
from datetime import datetime, time
from typing import Optional

from effi_mail.helpers import outlook, format_email_summary


def list_dms_clients() -> str:
    """List all client folders in DMSforLegal."""
    clients = outlook.list_dms_clients()
    return json.dumps({
        "count": len(clients),
        "clients": clients
    }, indent=2)


def list_dms_matters(client: str) -> str:
    """List matter folders for a client in DMSforLegal."""
    if not client:
        return json.dumps({"error": "client parameter is required"})
    
    matters = outlook.list_dms_matters(client)
    return json.dumps({
        "client": client,
        "count": len(matters),
        "matters": matters
    }, indent=2)


def get_dms_emails(client: str, matter: str, limit: int = 50) -> str:
    """Get emails filed under a client/matter in DMSforLegal."""
    if not client or not matter:
        return json.dumps({"error": "client and matter parameters are required"})
    
    emails = outlook.get_dms_emails(client, matter, limit=limit)
    return json.dumps({
        "client": client,
        "matter": matter,
        "count": len(emails),
        "emails": [format_email_summary(e, include_preview=True) for e in emails]
    }, indent=2)


def search_dms(
    client: Optional[str] = None,
    matter: Optional[str] = None,
    subject_contains: Optional[str] = None,
    date_from: Optional[str] = None,
    date_to: Optional[str] = None,
    limit: int = 50
) -> str:
    """Search emails in DMSforLegal. Dates are YYYY-MM-DD format."""
    # Parse dates
    # date_from: start of day (00:00:00)
    # date_to: end of day (23:59:59) to include all emails on that date
    date_from_dt = datetime.strptime(date_from, "%Y-%m-%d") if date_from else None
    date_to_dt = datetime.combine(
        datetime.strptime(date_to, "%Y-%m-%d").date(), time(23, 59, 59)
    ) if date_to else None
    
    emails = outlook.search_dms_emails(
        client=client,
        matter=matter,
        subject_contains=subject_contains,
        date_from=date_from_dt,
        date_to=date_to_dt,
        limit=limit,
    )
    
    return json.dumps({
        "client": client,
        "matter": matter,
        "count": len(emails),
        "emails": [format_email_summary(e, include_preview=True) for e in emails]
    }, indent=2)


def file_email_to_dms(
    email_id: str,
    client: str,
    matter: str,
) -> str:
    """File an email to a DMS client/matter folder.
    
    Copies the email to the matter's Emails folder, adds "Filed" category
    to the original, and marks it as effi:processed.
    
    Args:
        email_id: EntryID of the email to file
        client: Client folder name in DMS
        matter: Matter folder name under the client
        
    Returns:
        JSON with success status, filed email details, or error message.
    """
    if not email_id:
        return json.dumps({"success": False, "error": "email_id parameter is required"})
    if not client:
        return json.dumps({"success": False, "error": "client parameter is required"})
    if not matter:
        return json.dumps({"success": False, "error": "matter parameter is required"})
    
    # Validate client exists
    clients = outlook.list_dms_clients()
    if client not in clients:
        return json.dumps({
            "success": False,
            "error": f"Client '{client}' not found in DMS. Available clients: {clients}"
        })
    
    # Validate matter exists
    matters = outlook.list_dms_matters(client)
    if matter not in matters:
        return json.dumps({
            "success": False,
            "error": f"Matter '{matter}' not found for client '{client}'. Available matters: {matters}"
        })
    
    result = outlook.file_email_to_dms(
        email_id=email_id,
        client_name=client,
        matter_name=matter,
    )
    
    return json.dumps(result, indent=2)


def batch_file_emails_to_dms(
    email_ids: list,
    client: str,
    matter: str,
) -> str:
    """File multiple emails to a DMS client/matter folder.
    
    Validates the DMS folder once upfront, then files each email.
    Continues processing if individual emails fail.
    
    Args:
        email_ids: List of EntryIDs to file
        client: Client folder name in DMS
        matter: Matter folder name under the client
        
    Returns:
        JSON with filed_count, failed_count, and details for each.
    """
    if not email_ids:
        return json.dumps({
            "success": True,
            "filed_count": 0,
            "failed_count": 0,
            "filed_emails": [],
            "failed_emails": [],
            "message": "No emails provided to file"
        })
    if not client:
        return json.dumps({"success": False, "error": "client parameter is required"})
    if not matter:
        return json.dumps({"success": False, "error": "matter parameter is required"})
    
    # Validate client exists
    clients = outlook.list_dms_clients()
    if client not in clients:
        return json.dumps({
            "success": False,
            "error": f"Client '{client}' not found in DMS. Available clients: {clients}"
        })
    
    # Validate matter exists
    matters = outlook.list_dms_matters(client)
    if matter not in matters:
        return json.dumps({
            "success": False,
            "error": f"Matter '{matter}' not found for client '{client}'. Available matters: {matters}"
        })
    
    result = outlook.batch_file_emails_to_dms(
        email_ids=email_ids,
        client_name=client,
        matter_name=matter,
    )
    
    return json.dumps(result, indent=2)

