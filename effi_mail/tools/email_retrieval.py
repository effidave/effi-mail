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
    """Get untriaged emails grouped by domain. Filter by category: Client, Internal, Marketing, Personal, Uncategorized."""
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
    """Get Inbox emails from a sender domain."""
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
