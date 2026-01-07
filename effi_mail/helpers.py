"""Helper functions for effi-mail MCP server."""

import json
import os
from datetime import datetime
from pathlib import Path
from typing import Any, Optional
from outlook_client import OutlookClient


# Shared Outlook client instance
outlook = OutlookClient()

# Cache directory for large responses
CACHE_DIR = Path.home() / ".effi" / "cache"


def get_cache_path(prefix: str) -> Path:
    """Generate a timestamped cache file path.
    
    Args:
        prefix: Prefix for the cache file name (e.g., 'emails', 'search')
        
    Returns:
        Path to the cache file
    """
    CACHE_DIR.mkdir(parents=True, exist_ok=True)
    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    return CACHE_DIR / f"{prefix}_{timestamp}.json"


def write_cache_file(items: list, prefix: str, source_tool: str = "") -> str:
    """Write items to a cache file with metadata and tracking flags.
    
    Creates a structured cache file with:
    - metadata: created timestamp, source tool, counts
    - items: list with _retrieved and _processed flags added
    
    Args:
        items: List of items to cache
        prefix: Prefix for the cache file name
        source_tool: Name of the tool that generated this data
        
    Returns:
        Absolute path to the cache file as string
    """
    cache_path = get_cache_path(prefix)
    
    # Add tracking flags to each item
    tracked_items = []
    for item in items:
        tracked_item = item.copy() if isinstance(item, dict) else {"value": item}
        tracked_item["_retrieved"] = False
        tracked_item["_processed"] = False
        tracked_items.append(tracked_item)
    
    # Build cache structure
    cache_data = {
        "metadata": {
            "created": datetime.now().isoformat(),
            "source_tool": source_tool or prefix,
            "total_items": len(tracked_items),
            "retrieved_count": 0,
            "processed_count": 0
        },
        "items": tracked_items
    }
    
    with open(cache_path, "w", encoding="utf-8") as f:
        json.dump(cache_data, f, indent=2, default=str)
    return str(cache_path)


def build_response_with_auto_file(
    data: dict,
    items_key: str,
    count: int,
    limit: int,
    was_truncated: bool,
    total_available: Optional[int] = None,
    output_file: str = "",
    force_inline: bool = False,
    auto_file_threshold: int = 20,
    preview_count: int = 5,
    cache_prefix: str = "results"
) -> str:
    """Build a JSON response with optional auto-filing for large results.
    
    Args:
        data: Full response dict including items
        items_key: Key in data dict containing the large list (e.g., 'emails', 'domains')
        count: Number of items
        limit: Limit that was applied
        was_truncated: Whether results were truncated at source
        total_available: Total available if truncated
        output_file: Agent-specified output path (takes priority)
        force_inline: Override auto-filing and return everything inline
        auto_file_threshold: Item count above which to auto-file
        preview_count: Number of items to include as preview when auto-filing
        cache_prefix: Prefix for auto-generated cache file names
        
    Returns:
        JSON string response
    """
    items = data.get(items_key, [])
    
    # Build base response metadata
    response = {
        "count": count,
        "limit_applied": limit,
        "results_truncated": was_truncated,
    }
    if was_truncated and total_available:
        response["total_available"] = total_available
    
    # Copy any additional metadata from data (except the items)
    for key, value in data.items():
        if key != items_key and key not in response:
            response[key] = value
    
    # Explicit file path takes priority
    if output_file:
        file_path = os.path.expanduser(output_file)
        with open(file_path, "w", encoding="utf-8") as f:
            json.dump(items, f, indent=2, default=str)
        response["written_to"] = file_path
        response["items_saved"] = count
        return json.dumps(response, indent=2, default=str)
    
    # Agent override - return everything inline
    if force_inline:
        response[items_key] = items
        return json.dumps(response, indent=2, default=str)
    
    # Auto-file for large results
    if count > auto_file_threshold:
        cache_path = write_cache_file(items, cache_prefix, source_tool=cache_prefix)
        response["preview"] = items[:preview_count]
        response["full_data_file"] = cache_path
        response["auto_filed"] = True
        response["auto_file_note"] = f"Results ({count}) exceeded threshold ({auto_file_threshold}). Full data saved to file. Use read_cache_file to paginate through results."
        return json.dumps(response, indent=2, default=str)
    
    # Small result - inline
    response[items_key] = items
    return json.dumps(response, indent=2, default=str)


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


def build_conversation_filter(conversation_id: str) -> str:
    """Build Outlook Restrict filter for ConversationTopic.
    
    Note: ConversationID is NOT a filterable property in Outlook's Restrict().
    We use ConversationTopic instead, which is the normalized subject line.
    
    Args:
        conversation_id: Ignored - kept for API compatibility
        
    Returns:
        This function is deprecated. Use build_conversation_topic_filter instead.
        
    Raises:
        ValueError: Always - ConversationID filtering is not supported
    """
    raise ValueError(
        "ConversationID is not filterable via Outlook Restrict(). "
        "Use ConversationTopic filtering instead."
    )


def build_conversation_topic_filter(conversation_topic: str) -> str:
    """Build Outlook Restrict filter for ConversationTopic.
    
    ConversationTopic is the normalized subject line (without RE:/FW: prefixes).
    This is the only reliable way to filter for conversation threads.
    
    Args:
        conversation_topic: The ConversationTopic to filter by
        
    Returns:
        Jet-style filter string for use with Items.Restrict()
        
    Raises:
        ValueError: If conversation_topic is empty or None
    """
    if not conversation_topic:
        raise ValueError("conversation_topic cannot be empty or None")
    
    # Escape single quotes by doubling them (Jet syntax)
    escaped_topic = conversation_topic.replace("'", "''")
    
    return f"[ConversationTopic] = '{escaped_topic}'"
