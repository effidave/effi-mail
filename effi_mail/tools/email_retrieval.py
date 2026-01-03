"""Email retrieval tools for effi-mail MCP server."""

import json
from typing import Optional

from effi_mail.helpers import outlook, format_email_summary, truncate_text
from domain_categories import get_domain_category


def get_pending_emails(
    days: int = 30,
    limit: int = 100,
    category_filter: Optional[str] = None
) -> str:
    """Get emails pending triage (no effi: category), grouped by domain.
    
    Queries Outlook directly for emails without triage status.
    
    Args:
        days: Days to look back (default: 30)
        limit: Maximum emails to return (default: 100)
        category_filter: Filter by domain category: Client, Internal, Marketing, Personal, Uncategorized
        
    Returns:
        JSON string with pending emails grouped by domain
    """
    # Query Outlook directly for pending emails (no effi: category)
    result = outlook.get_pending_emails(days=days, limit=limit, group_by_domain=True)
    
    # Filter by domain category if specified
    if category_filter:
        filtered_domains = []
        for domain_data in result.get("domains", []):
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
        
        total = sum(len(d["emails"]) for d in filtered_domains)
        return json.dumps({
            "total_pending": total,
            "domains": filtered_domains
        }, indent=2)
    
    # No filter - add category info and format
    for domain_data in result.get("domains", []):
        domain_name = domain_data["domain"]
        domain_data["category"] = get_domain_category(domain_name)
        domain_data["emails"] = [
            format_email_summary(e, include_preview=True) 
            for e in domain_data["emails"]
        ]
    
    return json.dumps({
        "total_pending": result["total"],
        "domains": result.get("domains", [])
    }, indent=2)


def get_inbox_emails_by_domain(domain: str, limit: int = 20) -> str:
    """Get emails from a specific sender domain (Inbox only).
    
    Args:
        domain: Domain name to filter by (e.g., 'gmail.com')
        limit: Maximum emails to return (default: 20)
        
    Returns:
        JSON string with emails from the domain
    """
    # Search Outlook directly for emails from this domain
    emails = outlook.search_outlook(sender_domain=domain, limit=limit)
    return json.dumps({
        "domain": domain,
        "category": get_domain_category(domain),
        "count": len(emails),
        "emails": [format_email_summary(e, include_preview=True) for e in emails]
    }, indent=2)


def get_email_by_id(
    email_id: str,
    include_body: bool = True,
    include_attachments: bool = True,
    max_body_length: Optional[int] = None
) -> str:
    """Get full email details by ID.
    
    Accepts EntryID or internet_message_id (auto-detected).
    
    Args:
        email_id: Email ID - either Outlook EntryID or internet_message_id (format: <...@...>)
        include_body: Include full email body (default: True)
        include_attachments: Include attachment metadata (default: True)
        max_body_length: Truncate body to this length (default: None = no truncation)
        
    Returns:
        JSON string with email details
    """
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
