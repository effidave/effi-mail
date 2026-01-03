"""Domain categorization tools for effi-mail MCP server."""

import json
from typing import Optional

from effi_mail.helpers import outlook
from domain_categories import (
    get_domain_category,
    set_domain_category,
    get_all_domain_categories,
)


def get_uncategorized_domains(days: int = 30, limit: int = 20) -> str:
    """Get list of domains that haven't been categorized yet.
    
    Args:
        days: Days to look back (default: 30)
        limit: Maximum domains to return (default: 20)
        
    Returns:
        JSON string with uncategorized domains and email counts
    """
    # Scan ALL emails in date range for uncategorized domains (no limit)
    result = outlook.get_domain_counts(days=days, limit=None, pending_only=False)
    
    uncategorized = []
    for domain_data in result.get("domains", []):
        domain_name = domain_data["domain"]
        category = get_domain_category(domain_name)
        if category == "Uncategorized":
            uncategorized.append({
                "name": domain_name,
                "email_count": domain_data["count"],
                "sample_subjects": domain_data["sample_subjects"]
            })
            if len(uncategorized) >= limit:
                break
    
    return json.dumps({
        "count": len(uncategorized),
        "days_scanned": days,
        "domains": uncategorized
    }, indent=2)


def categorize_domain(domain: str, category: str) -> str:
    """Set the category for a sender domain.
    
    Saves to domain_categories.json.
    
    Args:
        domain: Domain name
        category: Category to assign - 'Client', 'Internal', 'Marketing', 'Personal', or 'Spam'
        
    Returns:
        JSON string with success status
    """
    set_domain_category(domain, category)
    return json.dumps({"success": True, "domain": domain, "category": category})


def get_domain_summary() -> str:
    """Get summary of all domains grouped by category.
    
    Reads from domain_categories.json.
    
    Returns:
        JSON string with domains grouped by category
    """
    all_categories = get_all_domain_categories()
    
    # Group by category
    result = {}
    for domain, category in all_categories.items():
        if category not in result:
            result[category] = {"count": 0, "domains": []}
        result[category]["count"] += 1
        if len(result[category]["domains"]) < 10:
            result[category]["domains"].append(domain)
    
    return json.dumps(result, indent=2)
