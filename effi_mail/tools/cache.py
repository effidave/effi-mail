"""Cache file tools for effi-mail MCP server.

Provides paginated access to auto-filed large results with tracking.
"""

import json
import os
from typing import Optional, List

from effi_mail.helpers import CACHE_DIR


def read_cache_file(
    file_path: str,
    start: int = 0,
    limit: int = 20,
    filter_field: Optional[str] = None,
    filter_value: Optional[str] = None,
    fields: Optional[List[str]] = None,
    include_retrieved: bool = False,
    unprocessed_only: bool = False
) -> str:
    """Read items from a cache file with pagination and filtering.
    
    Items are automatically marked as _retrieved=True when returned.
    Use this to page through large auto-filed results without loading all into memory.
    
    Args:
        file_path: Path to the cache file
        start: Offset for pagination (skip this many unretrieved items)
        limit: Maximum items to return (default 20)
        filter_field: Field name to filter on (e.g., "domain", "sender")
        filter_value: Value to match (substring, case-insensitive)
        fields: Only return these fields per item (reduces payload size)
        include_retrieved: Include already-retrieved items (default False - only unretrieved)
        unprocessed_only: Only return items that are retrieved but not processed
        
    Returns:
        JSON with items and updated status counts
    """
    file_path = os.path.expanduser(file_path)
    
    if not os.path.exists(file_path):
        return json.dumps({"error": f"Cache file not found: {file_path}"})
    
    try:
        with open(file_path, "r", encoding="utf-8") as f:
            cache_data = json.load(f)
    except json.JSONDecodeError as e:
        return json.dumps({"error": f"Invalid JSON in cache file: {e}"})
    
    # Handle legacy cache files (just a list)
    if isinstance(cache_data, list):
        return json.dumps({
            "error": "Legacy cache file format. Re-run the original query to create a tracked cache file."
        })
    
    metadata = cache_data.get("metadata", {})
    items = cache_data.get("items", [])
    
    # Filter based on retrieval/processing status
    if unprocessed_only:
        # Items that have been retrieved but not processed
        candidates = [item for item in items if item.get("_retrieved") and not item.get("_processed")]
    elif include_retrieved:
        # All items
        candidates = items
    else:
        # Only unretrieved items
        candidates = [item for item in items if not item.get("_retrieved")]
    
    # Apply field filter if specified
    if filter_field and filter_value:
        filter_value_lower = filter_value.lower()
        candidates = [
            item for item in candidates
            if filter_field in item and filter_value_lower in str(item[filter_field]).lower()
        ]
    
    # Apply pagination
    selected = candidates[start:start + limit]
    
    # Mark selected items as retrieved in the original items list
    selected_indices = []
    for sel_item in selected:
        for i, item in enumerate(items):
            if item is sel_item or item.get("id") == sel_item.get("id"):
                items[i]["_retrieved"] = True
                selected_indices.append(i)
                break
    
    # Update metadata counts
    metadata["retrieved_count"] = sum(1 for item in items if item.get("_retrieved"))
    metadata["processed_count"] = sum(1 for item in items if item.get("_processed"))
    
    # Save updated cache file
    cache_data["metadata"] = metadata
    cache_data["items"] = items
    with open(file_path, "w", encoding="utf-8") as f:
        json.dump(cache_data, f, indent=2, default=str)
    
    # Prepare response items (optionally filter fields, remove tracking flags)
    response_items = []
    for item in selected:
        if fields:
            response_item = {k: v for k, v in item.items() if k in fields or k == "id"}
        else:
            response_item = {k: v for k, v in item.items() if not k.startswith("_")}
        response_items.append(response_item)
    
    return json.dumps({
        "count": len(response_items),
        "total_in_file": metadata.get("total_items", len(items)),
        "retrieved_count": metadata["retrieved_count"],
        "processed_count": metadata["processed_count"],
        "remaining_unretrieved": metadata.get("total_items", len(items)) - metadata["retrieved_count"],
        "filter_applied": f"{filter_field}={filter_value}" if filter_field else None,
        "items": response_items
    }, indent=2)


def mark_cache_processed(
    file_path: str,
    ids: List[str]
) -> str:
    """Mark specific items as processed in a cache file.
    
    Use this after taking action on items (e.g., archiving emails) to track progress.
    
    Args:
        file_path: Path to the cache file
        ids: List of item IDs to mark as processed
        
    Returns:
        JSON with updated counts
    """
    file_path = os.path.expanduser(file_path)
    
    if not os.path.exists(file_path):
        return json.dumps({"error": f"Cache file not found: {file_path}"})
    
    try:
        with open(file_path, "r", encoding="utf-8") as f:
            cache_data = json.load(f)
    except json.JSONDecodeError as e:
        return json.dumps({"error": f"Invalid JSON in cache file: {e}"})
    
    if isinstance(cache_data, list):
        return json.dumps({"error": "Legacy cache file format. Cannot track processing."})
    
    metadata = cache_data.get("metadata", {})
    items = cache_data.get("items", [])
    
    # Mark matching items as processed
    marked_count = 0
    ids_set = set(ids)
    for item in items:
        if item.get("id") in ids_set:
            item["_processed"] = True
            item["_retrieved"] = True  # Also mark as retrieved
            marked_count += 1
    
    # Update metadata counts
    metadata["retrieved_count"] = sum(1 for item in items if item.get("_retrieved"))
    metadata["processed_count"] = sum(1 for item in items if item.get("_processed"))
    
    # Save updated cache file
    cache_data["metadata"] = metadata
    cache_data["items"] = items
    with open(file_path, "w", encoding="utf-8") as f:
        json.dump(cache_data, f, indent=2, default=str)
    
    return json.dumps({
        "marked_count": marked_count,
        "ids_provided": len(ids),
        "total_in_file": metadata.get("total_items", len(items)),
        "retrieved_count": metadata["retrieved_count"],
        "processed_count": metadata["processed_count"],
        "remaining_unprocessed": metadata.get("total_items", len(items)) - metadata["processed_count"]
    }, indent=2)


def get_cache_status(file_path: str) -> str:
    """Get status and counts for a cache file.
    
    Use this to check progress on processing a large result set.
    
    Args:
        file_path: Path to the cache file
        
    Returns:
        JSON with metadata and counts
    """
    file_path = os.path.expanduser(file_path)
    
    if not os.path.exists(file_path):
        return json.dumps({"error": f"Cache file not found: {file_path}"})
    
    try:
        with open(file_path, "r", encoding="utf-8") as f:
            cache_data = json.load(f)
    except json.JSONDecodeError as e:
        return json.dumps({"error": f"Invalid JSON in cache file: {e}"})
    
    if isinstance(cache_data, list):
        return json.dumps({
            "format": "legacy",
            "total_items": len(cache_data),
            "note": "Legacy format without tracking. Re-run query to create tracked cache."
        })
    
    metadata = cache_data.get("metadata", {})
    items = cache_data.get("items", [])
    
    # Recalculate counts from actual data
    retrieved_count = sum(1 for item in items if item.get("_retrieved"))
    processed_count = sum(1 for item in items if item.get("_processed"))
    total = len(items)
    
    return json.dumps({
        "file_path": file_path,
        "created": metadata.get("created"),
        "source_tool": metadata.get("source_tool"),
        "total_items": total,
        "retrieved_count": retrieved_count,
        "processed_count": processed_count,
        "remaining_unretrieved": total - retrieved_count,
        "remaining_unprocessed": total - processed_count,
        "percent_retrieved": round(100 * retrieved_count / total, 1) if total else 0,
        "percent_processed": round(100 * processed_count / total, 1) if total else 0
    }, indent=2)


def reset_cache_flags(
    file_path: str,
    reset_retrieved: bool = True,
    reset_processed: bool = True
) -> str:
    """Reset tracking flags on all items in a cache file.
    
    Use this to re-process a cache file from the beginning.
    
    Args:
        file_path: Path to the cache file
        reset_retrieved: Reset _retrieved flags to False (default True)
        reset_processed: Reset _processed flags to False (default True)
        
    Returns:
        JSON with confirmation and updated counts
    """
    file_path = os.path.expanduser(file_path)
    
    if not os.path.exists(file_path):
        return json.dumps({"error": f"Cache file not found: {file_path}"})
    
    try:
        with open(file_path, "r", encoding="utf-8") as f:
            cache_data = json.load(f)
    except json.JSONDecodeError as e:
        return json.dumps({"error": f"Invalid JSON in cache file: {e}"})
    
    if isinstance(cache_data, list):
        return json.dumps({"error": "Legacy cache file format. Cannot reset flags."})
    
    metadata = cache_data.get("metadata", {})
    items = cache_data.get("items", [])
    
    # Reset flags
    retrieved_reset = 0
    processed_reset = 0
    for item in items:
        if reset_retrieved and item.get("_retrieved"):
            item["_retrieved"] = False
            retrieved_reset += 1
        if reset_processed and item.get("_processed"):
            item["_processed"] = False
            processed_reset += 1
    
    # Update metadata counts
    metadata["retrieved_count"] = sum(1 for item in items if item.get("_retrieved"))
    metadata["processed_count"] = sum(1 for item in items if item.get("_processed"))
    
    # Save updated cache file
    cache_data["metadata"] = metadata
    cache_data["items"] = items
    with open(file_path, "w", encoding="utf-8") as f:
        json.dump(cache_data, f, indent=2, default=str)
    
    return json.dumps({
        "success": True,
        "retrieved_flags_reset": retrieved_reset,
        "processed_flags_reset": processed_reset,
        "total_items": len(items),
        "retrieved_count": metadata["retrieved_count"],
        "processed_count": metadata["processed_count"]
    }, indent=2)


def list_cache_files(days: int = 7) -> str:
    """List cache files created in the last N days.
    
    Args:
        days: Look back period (default 7 days)
        
    Returns:
        JSON with list of cache files and their status
    """
    from datetime import datetime, timedelta
    
    if not CACHE_DIR.exists():
        return json.dumps({"count": 0, "files": []})
    
    cutoff = datetime.now() - timedelta(days=days)
    files = []
    
    for file_path in CACHE_DIR.glob("*.json"):
        try:
            stat = file_path.stat()
            modified = datetime.fromtimestamp(stat.st_mtime)
            
            if modified < cutoff:
                continue
            
            # Try to read metadata
            with open(file_path, "r", encoding="utf-8") as f:
                cache_data = json.load(f)
            
            if isinstance(cache_data, dict) and "metadata" in cache_data:
                metadata = cache_data["metadata"]
                files.append({
                    "path": str(file_path),
                    "name": file_path.name,
                    "created": metadata.get("created"),
                    "source_tool": metadata.get("source_tool"),
                    "total_items": metadata.get("total_items"),
                    "retrieved_count": metadata.get("retrieved_count"),
                    "processed_count": metadata.get("processed_count"),
                    "size_kb": round(stat.st_size / 1024, 1)
                })
            else:
                # Legacy format
                files.append({
                    "path": str(file_path),
                    "name": file_path.name,
                    "format": "legacy",
                    "total_items": len(cache_data) if isinstance(cache_data, list) else None,
                    "modified": modified.isoformat(),
                    "size_kb": round(stat.st_size / 1024, 1)
                })
        except Exception:
            continue
    
    # Sort by creation time, newest first
    files.sort(key=lambda x: x.get("created") or x.get("modified") or "", reverse=True)
    
    return json.dumps({
        "count": len(files),
        "cache_dir": str(CACHE_DIR),
        "days_scanned": days,
        "files": files
    }, indent=2)
