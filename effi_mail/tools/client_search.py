"""Client search tools for effi-mail MCP server."""

import json
from datetime import datetime, time
from typing import Optional

from effi_mail.helpers import search, retrieval, folders, format_email_summary, build_response_with_auto_file
from effi_work_client import get_client_identifiers_from_effi_work


async def get_emails_by_client(
    client_id: str,
    days: int = 30,
    date_from: Optional[str] = None,
    date_to: Optional[str] = None,
    limit: int = 100,
    output_file: str = "",
    force_inline: bool = False,
    auto_file_threshold: int = 20
) -> str:
    """Get client correspondence. client_id is case-insensitive.
    
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
    
    # Get client identifiers from effi-core (fresh data)
    identifiers = await get_client_identifiers_from_effi_work(client_id)
    if not identifiers.get("domains"):
        return json.dumps({
            "error": f"Client not found: {client_id}",
            "source": identifiers.get("source"),
            "hint": "Use list_dms_clients to find the exact client name. Client names often include 'Ltd', 'Limited', etc."
        })
    
    # Search Outlook with limit+1 to detect truncation
    emails = search.search_outlook_by_identifiers(
        domains=identifiers["domains"],
        contact_emails=identifiers.get("contact_emails", []),
        days=days,
        date_from=date_from_dt,
        date_to=date_to_dt,
        limit=limit + 1,
    )
    was_truncated = len(emails) > limit
    emails = emails[:limit]
    
    formatted = [format_email_summary(e, include_preview=True, include_recipients=True) for e in emails]
    return build_response_with_auto_file(
        data={
            "client_id": client_id,
            "identifiers": identifiers,
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
        cache_prefix=f"client_{client_id}"
    )


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
    limit: int = 50,
    output_file: str = "",
    force_inline: bool = False,
    auto_file_threshold: int = 20
) -> str:
    """Search Outlook with filters. folder: 'Inbox' or 'Sent Items'. Dates: YYYY-MM-DD.
    
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
    
    # Search with limit+1 to detect truncation
    emails = search.search_outlook(
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
        limit=limit + 1,
    )
    was_truncated = len(emails) > limit
    emails = emails[:limit]
    
    formatted = [format_email_summary(e, include_preview=True, include_recipients=True) for e in emails]
    # Sanitize folder name for cache prefix (remove path separators)
    safe_folder = folder.lower().replace(' ', '_').replace('\\', '_').replace('/', '_')
    return build_response_with_auto_file(
        data={
            "folder": folder,
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
        cache_prefix=f"search_{safe_folder}"
    )


# ============================================================================
# Commitment Scanning Tools
# ============================================================================

SCANNED_CATEGORY = "effi:scanned"


def scan_for_commitments(
    days: int = 14,
    limit: int = 100,
    output_file: str = "",
    force_inline: bool = False,
    auto_file_threshold: int = 20
) -> str:
    """Scan sent emails for commitment detection. Returns unscanned emails with full body.
    
    Fetches emails from Sent Items that don't have the effi:scanned category,
    returning full email body and recipient information for commitment parsing.
    
    Large results (>{auto_file_threshold} emails) are auto-saved to a cache file.
    Use force_inline=True to return full payload inline regardless of size.
    Use output_file to save results to a specific path.
    
    ⚠️ Results are LIMITED. Check 'results_truncated' in response to determine if more records exist.
    
    Args:
        days: Number of days to look back (default 14)
        limit: Maximum emails to return (default 100)
        output_file: Path to save results to (optional)
        force_inline: Return full payload inline regardless of size (default False)
        auto_file_threshold: Auto-file results above this count (default 20)
    
    Returns:
        JSON with emails including full body content for commitment scanning
    """
    # Get sent emails - fetch more to account for filtering
    emails = search.search_outlook(
        folder="Sent Items",
        days=days,
        limit=limit * 2,  # Fetch extra since we filter out scanned
    )
    
    # Filter out already scanned emails
    unscanned = [e for e in emails if SCANNED_CATEGORY not in (e.categories or "")]
    
    # Get full email details for each unscanned email
    result_emails = []
    for email in unscanned:
        try:
            full_email = retrieval.get_email_full(email.id)
            result_emails.append({
                "id": email.id,
                "subject": full_email.get("subject", email.subject),
                "sent_time": full_email.get("received_time", email.received_time.isoformat() if email.received_time else None),
                "body": full_email.get("body", ""),
                "recipients_to": full_email.get("recipients_to", []),
                "recipients_cc": full_email.get("recipients_cc", []),
            })
        except Exception:
            # Fallback to basic info if full email fetch fails
            result_emails.append({
                "id": email.id,
                "subject": email.subject,
                "sent_time": email.received_time.isoformat() if email.received_time else None,
                "body": email.body_preview or "",
                "recipients_to": [],
                "recipients_cc": [],
            })
    
    # Track truncation
    total_unscanned = len(unscanned)
    was_truncated = total_unscanned > limit
    result_emails = result_emails[:limit]
    
    return build_response_with_auto_file(
        data={"emails": result_emails},
        items_key="emails",
        count=len(result_emails),
        limit=limit,
        was_truncated=was_truncated,
        total_available=total_unscanned if was_truncated else None,
        output_file=output_file,
        force_inline=force_inline,
        auto_file_threshold=auto_file_threshold,
        cache_prefix="commitments"
    )


def mark_scanned(email_id: str) -> str:
    """Mark an email as scanned for commitments.
    
    Adds the effi:scanned category to the email to prevent re-scanning.
    
    Args:
        email_id: Outlook EntryID of the email to mark
    
    Returns:
        JSON with success status
    """
    success = folders.set_category(email_id, SCANNED_CATEGORY)
    
    if success:
        return json.dumps({
            "success": True,
            "email_id": email_id,
            "category": SCANNED_CATEGORY,
        })
    else:
        return json.dumps({
            "success": False,
            "email_id": email_id,
            "error": "Failed to set category on email",
        })


def batch_mark_scanned(email_ids: list[str]) -> str:
    """Mark multiple emails as scanned for commitments.
    
    Args:
        email_ids: List of Outlook EntryIDs to mark
    
    Returns:
        JSON with counts of successful and failed markings
    """
    marked_count = 0
    failed_count = 0
    failed_ids = []
    
    for email_id in email_ids:
        success = folders.set_category(email_id, SCANNED_CATEGORY)
        if success:
            marked_count += 1
        else:
            failed_count += 1
            failed_ids.append(email_id)
    
    return json.dumps({
        "marked_count": marked_count,
        "failed_count": failed_count,
        "failed_ids": failed_ids,
    })
