"""DMS (DMSforLegal) tools for effi-mail MCP server."""

import json
from datetime import datetime, time
from typing import Optional

from effi_mail.helpers import outlook, format_email_summary, build_response_with_auto_file


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


def get_dms_emails(
    client: str,
    matter: str,
    limit: int = 50,
    output_file: str = "",
    force_inline: bool = False,
    auto_file_threshold: int = 20
) -> str:
    """Get emails filed under a client/matter in DMSforLegal.
    
    Large results (>{auto_file_threshold} emails) are auto-saved to a cache file.
    Use force_inline=True to return full payload inline regardless of size.
    Use output_file to save results to a specific path.
    
    ⚠️ Results are LIMITED. Check 'results_truncated' in response to determine if more records exist.
    """
    if not client or not matter:
        return json.dumps({"error": "client and matter parameters are required"})
    
    # Fetch limit+1 to detect truncation
    emails = outlook.get_dms_emails(client, matter, limit=limit + 1)
    was_truncated = len(emails) > limit
    emails = emails[:limit]
    
    formatted = [format_email_summary(e, include_preview=True) for e in emails]
    return build_response_with_auto_file(
        data={
            "client": client,
            "matter": matter,
            "emails": formatted
        },
        items_key="emails",
        count=len(formatted),
        limit=limit,
        was_truncated=was_truncated,
        total_available=None,
        output_file=output_file,
        force_inline=force_inline,
        auto_file_threshold=auto_file_threshold,
        cache_prefix=f"dms_{client}_{matter}"
    )


def get_dms_admin_emails(
    client: str,
    matter: str,
    limit: int = 50,
    output_file: str = "",
    force_inline: bool = False,
    auto_file_threshold: int = 20
) -> str:
    """Get admin/system emails filed under a client/matter Admin folder in DMSforLegal.
    
    Large results (>{auto_file_threshold} emails) are auto-saved to a cache file.
    Use force_inline=True to return full payload inline regardless of size.
    Use output_file to save results to a specific path.
    
    ⚠️ Results are LIMITED. Check 'results_truncated' in response to determine if more records exist.
    """
    if not client or not matter:
        return json.dumps({"error": "client and matter parameters are required"})
    
    # Fetch limit+1 to detect truncation
    emails = outlook.get_dms_admin_emails(client, matter, limit=limit + 1)
    was_truncated = len(emails) > limit
    emails = emails[:limit]
    
    formatted = [format_email_summary(e, include_preview=True) for e in emails]
    return build_response_with_auto_file(
        data={
            "client": client,
            "matter": matter,
            "folder": "Admin",
            "emails": formatted
        },
        items_key="emails",
        count=len(formatted),
        limit=limit,
        was_truncated=was_truncated,
        total_available=None,
        output_file=output_file,
        force_inline=force_inline,
        auto_file_threshold=auto_file_threshold,
        cache_prefix=f"dms_admin_{client}_{matter}"
    )


def search_dms(
    client: Optional[str] = None,
    matter: Optional[str] = None,
    subject_contains: Optional[str] = None,
    date_from: Optional[str] = None,
    date_to: Optional[str] = None,
    limit: int = 50,
    output_file: str = "",
    force_inline: bool = False,
    auto_file_threshold: int = 20
) -> str:
    """Search emails in DMSforLegal. Dates are YYYY-MM-DD format.
    
    Large results (>{auto_file_threshold} emails) are auto-saved to a cache file.
    Use force_inline=True to return full payload inline regardless of size.
    Use output_file to save results to a specific path.
    
    ⚠️ Results are LIMITED. Check 'results_truncated' in response to determine if more records exist.
    """
    # Parse dates
    # date_from: start of day (00:00:00)
    # date_to: end of day (23:59:59) to include all emails on that date
    date_from_dt = datetime.strptime(date_from, "%Y-%m-%d") if date_from else None
    date_to_dt = datetime.combine(
        datetime.strptime(date_to, "%Y-%m-%d").date(), time(23, 59, 59)
    ) if date_to else None
    
    # Fetch limit+1 to detect truncation
    emails = outlook.search_dms_emails(
        client=client,
        matter=matter,
        subject_contains=subject_contains,
        date_from=date_from_dt,
        date_to=date_to_dt,
        limit=limit + 1,
    )
    was_truncated = len(emails) > limit
    emails = emails[:limit]
    
    formatted = [format_email_summary(e, include_preview=True) for e in emails]
    return build_response_with_auto_file(
        data={
            "client": client,
            "matter": matter,
            "emails": formatted
        },
        items_key="emails",
        count=len(formatted),
        limit=limit,
        was_truncated=was_truncated,
        total_available=None,
        output_file=output_file,
        force_inline=force_inline,
        auto_file_threshold=auto_file_threshold,
        cache_prefix="dms_search"
    )


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


def file_admin_email_to_dms(
    email_id: str,
    client: str,
    matter: str,
) -> str:
    """File an admin/system email to a DMS client/matter Admin folder.
    
    Use for internal notifications related to a matter (e.g. 'New Project created' emails).
    Files to the Admin subfolder instead of Emails.
    
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
    
    result = outlook.file_email_to_dms_admin(
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

