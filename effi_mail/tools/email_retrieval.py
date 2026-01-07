"""Email retrieval tools for effi-mail MCP server."""

import json
from typing import Optional

from effi_mail.helpers import outlook, format_email_summary, truncate_text, build_response_with_auto_file
from domain_categories import get_domain_category


def get_pending_emails(
    days: int = 30,
    limit: int = 100,
    category_filter: Optional[str] = None,
    output_file: str = "",
    force_inline: bool = False,
    auto_file_threshold: int = 20
) -> str:
    """Get untriaged emails grouped by domain. Filter by category: Client, Internal, Marketing, Personal, Uncategorized.
    
    Large results (>{auto_file_threshold} emails) are auto-saved to a cache file.
    Use force_inline=True to return full payload inline regardless of size.
    Use output_file to save results to a specific path.
    
    ⚠️ Results are LIMITED. Check 'results_truncated' in response to determine if more records exist.
    """
    # Query Outlook with limit+1 to detect truncation
    result = outlook.get_pending_emails(days=days, limit=limit + 1, group_by_domain=True)
    
    # Check if results were truncated
    total_available = result.get("total", 0)
    was_truncated = total_available > limit
    
    # Trim to requested limit
    all_domains = result.get("domains", [])
    email_count = 0
    trimmed_domains = []
    for domain_data in all_domains:
        if email_count >= limit:
            break
        remaining = limit - email_count
        if len(domain_data.get("emails", [])) <= remaining:
            trimmed_domains.append(domain_data)
            email_count += len(domain_data.get("emails", []))
        else:
            # Partial domain - take only what we need
            domain_copy = domain_data.copy()
            domain_copy["emails"] = domain_data["emails"][:remaining]
            trimmed_domains.append(domain_copy)
            email_count += remaining
    
    # Filter by domain category if specified
    if category_filter:
        filtered_domains = []
        for domain_data in trimmed_domains:
            domain_name = domain_data["domain"]
            domain_cat = get_domain_category(domain_name)
            if domain_cat == category_filter:
                domain_data["category"] = domain_cat
                filtered_domains.append(domain_data)
        
        # Format emails for response
        for domain_data in filtered_domains:
            domain_data["emails"] = [
                format_email_summary(e, include_preview=True) 
                for e in domain_data["emails"]
            ]
        
        count = sum(len(d["emails"]) for d in filtered_domains)
        return build_response_with_auto_file(
            data={"domains": filtered_domains},
            items_key="domains",
            count=count,
            limit=limit,
            was_truncated=was_truncated,
            total_available=total_available if was_truncated else None,
            output_file=output_file,
            force_inline=force_inline,
            auto_file_threshold=auto_file_threshold,
            cache_prefix="pending_emails"
        )
    
    # No filter - add category info and format
    for domain_data in trimmed_domains:
        domain_name = domain_data["domain"]
        domain_data["category"] = get_domain_category(domain_name)
        domain_data["emails"] = [
            format_email_summary(e, include_preview=True) 
            for e in domain_data["emails"]
        ]
    
    count = sum(len(d["emails"]) for d in trimmed_domains)
    return build_response_with_auto_file(
        data={"domains": trimmed_domains},
        items_key="domains",
        count=count,
        limit=limit,
        was_truncated=was_truncated,
        total_available=total_available if was_truncated else None,
        output_file=output_file,
        force_inline=force_inline,
        auto_file_threshold=auto_file_threshold,
        cache_prefix="pending_emails"
    )


def get_inbox_emails_by_domain(
    domain: str,
    limit: int = 20,
    output_file: str = "",
    force_inline: bool = False,
    auto_file_threshold: int = 20
) -> str:
    """Get Inbox emails from a sender domain.
    
    Large results (>{auto_file_threshold} emails) are auto-saved to a cache file.
    Use force_inline=True to return full payload inline regardless of size.
    Use output_file to save results to a specific path.
    
    ⚠️ Results are LIMITED. Check 'results_truncated' in response to determine if more records exist.
    """
    # Search Outlook with limit+1 to detect truncation
    emails = outlook.search_outlook(sender_domain=domain, limit=limit + 1)
    was_truncated = len(emails) > limit
    emails = emails[:limit]
    
    formatted = [format_email_summary(e, include_preview=True) for e in emails]
    return build_response_with_auto_file(
        data={
            "domain": domain,
            "category": get_domain_category(domain),
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
        cache_prefix=f"inbox_{domain}"
    )


def get_sent_emails_by_domain(
    domain: str,
    days: int = 30,
    limit: int = 20,
    output_file: str = "",
    force_inline: bool = False,
    auto_file_threshold: int = 20
) -> str:
    """Get Sent Items emails to a recipient domain. Use to verify if emails have been responded to.
    
    Large results (>{auto_file_threshold} emails) are auto-saved to a cache file.
    Use force_inline=True to return full payload inline regardless of size.
    Use output_file to save results to a specific path.
    
    ⚠️ Results are LIMITED. Check 'results_truncated' in response to determine if more records exist.
    """
    # Ensure recipient domains are set on recent sent items
    outlook._set_recipient_domains()
    
    # Search Outlook Sent Items with limit+1 to detect truncation
    emails = outlook.search_outlook(recipient_domain=domain, folder="Sent Items", days=days, limit=limit + 1)
    was_truncated = len(emails) > limit
    emails = emails[:limit]
    
    formatted = [format_email_summary(e, include_preview=True) for e in emails]
    return build_response_with_auto_file(
        data={
            "domain": domain,
            "category": get_domain_category(domain),
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
        cache_prefix=f"sent_{domain}"
    )


def get_email_by_id(
    email_id: str,
    include_body: bool = True,
    include_attachments: bool = True,
    max_body_length: Optional[int] = None
) -> str:
    """Get full email by EntryID or internet_message_id (auto-detected)."""
    # Get email directly from Outlook
    full_email = outlook.get_email_full(email_id)
    
    if full_email:
        result = full_email.copy()
        if not include_body:
            result.pop("body", None)
            result.pop("html_body", None)
        elif max_body_length and "body" in result:
            result["body"] = truncate_text(result["body"], max_body_length)
        if not include_attachments:
            result.pop("attachments", None)
        return json.dumps(result, indent=2)
    else:
        return json.dumps({"error": f"Email not found: {email_id}"})


def download_attachment(
    email_id: str,
    attachment_name: str,
    save_path: Optional[str] = None
) -> str:
    """Download attachment to save_path or ./attachments/{domain}/{date}/{filename}."""
    result = outlook.download_attachment(
        email_id=email_id,
        attachment_name=attachment_name,
        save_path=save_path
    )
    return json.dumps(result, indent=2)


def search_inbox_by_subject(
    subject_starts_with: str,
    days: int = 30,
    limit: int = 50,
    output_file: str = "",
    force_inline: bool = False,
    auto_file_threshold: int = 20
) -> str:
    """Search Inbox for emails where subject starts with specified text. Returns email IDs for further processing.
    
    Large results (>{auto_file_threshold} emails) are auto-saved to a cache file.
    Use force_inline=True to return full payload inline regardless of size.
    Use output_file to save results to a specific path.
    
    ⚠️ Results are LIMITED. Check 'results_truncated' in response to determine if more records exist.
    """
    emails = outlook.search_outlook(
        subject_contains=subject_starts_with,
        folder="Inbox",
        days=days,
        limit=limit + 1  # Fetch one extra to detect truncation
    )
    # Filter to only those where subject actually starts with the text
    matching = [
        {"id": e.id, "subject": e.subject, "sender": e.sender_email, "received": e.received_time.isoformat() if e.received_time else None}
        for e in emails
        if e.subject and e.subject.lower().startswith(subject_starts_with.lower())
    ]
    was_truncated = len(matching) > limit
    matching = matching[:limit]
    
    return build_response_with_auto_file(
        data={"emails": matching},
        items_key="emails",
        count=len(matching),
        limit=limit,
        was_truncated=was_truncated,
        total_available=None,
        output_file=output_file,
        force_inline=force_inline,
        auto_file_threshold=auto_file_threshold,
        cache_prefix="search_subject"
    )
