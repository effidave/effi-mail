"""Domain categorization tools for effi-mail MCP server."""

import json
from typing import Optional

from effi_mail.helpers import outlook, build_response_with_auto_file
from domain_categories import (
    get_domain_category,
    set_domain_category,
    get_all_domain_categories,
)


def get_uncategorized_domains(
    days: int = 3650,
    limit: int = 20,
    output_file: str = "",
    force_inline: bool = False,
    auto_file_threshold: int = 20
) -> str:
    """Get domains without a category, with email counts.
    
    Large results (>{auto_file_threshold} domains) are auto-saved to a cache file.
    Use force_inline=True to return full payload inline regardless of size.
    Use output_file to save results to a specific path.
    
    ⚠️ Results are LIMITED. Check 'results_truncated' in response to determine if more records exist.
    
    Args:
        days: Days to scan back (default 3650 = ~10 years to cover full inbox history)
        limit: Maximum domains to return (default 20)
        output_file: Path to save results to (optional)
        force_inline: Return full payload inline regardless of size (default False)
        auto_file_threshold: Auto-file results above this count (default 20)
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
    
    # Track total before truncating
    total_uncategorized = len(uncategorized)
    was_truncated = total_uncategorized > limit
    uncategorized = uncategorized[:limit]
    
    return build_response_with_auto_file(
        data={
            "days_scanned": days,
            "domains": uncategorized
        },
        items_key="domains",
        count=len(uncategorized),
        limit=limit,
        was_truncated=was_truncated,
        total_available=total_uncategorized if was_truncated else None,
        output_file=output_file,
        force_inline=force_inline,
        auto_file_threshold=auto_file_threshold,
        cache_prefix="uncategorized_domains"
    )


def categorize_domain(domain: str, category: str) -> str:
    """Set domain category (case-insensitive): Client, Internal, Marketing, Personal, or Spam."""
    set_domain_category(domain, category)
    return json.dumps({"success": True, "domain": domain, "category": category})


def get_domain_summary() -> str:
    """Get all domains grouped by category."""
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
