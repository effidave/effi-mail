"""Storage utilities for email ingestion.

Handles file operations, seen ID tracking, and attachment saving.
"""

import json
from pathlib import Path
from typing import Set
import logging

logger = logging.getLogger(__name__)


def load_seen_ids(inbox_path: Path) -> Set[str]:
    """Load set of already-processed message IDs.
    
    Args:
        inbox_path: Path to _inbox folder
        
    Returns:
        Set of seen message IDs
    """
    seen_file = inbox_path / "_seen.json"
    if seen_file.exists():
        try:
            data = json.loads(seen_file.read_text(encoding="utf-8"))
            return set(data.get("seen_ids", []))
        except Exception as e:
            logger.warning(f"Failed to load seen IDs: {e}. Starting fresh.")
            return set()
    return set()


def save_seen_ids(inbox_path: Path, seen_ids: Set[str]) -> None:
    """Persist seen message IDs using atomic write.
    
    Args:
        inbox_path: Path to _inbox folder
        seen_ids: Set of message IDs to save
    """
    seen_file = inbox_path / "_seen.json"
    temp_file = seen_file.with_suffix(".tmp")
    
    try:
        data = {"seen_ids": sorted(list(seen_ids))}
        temp_file.write_text(json.dumps(data, indent=2), encoding="utf-8")
        temp_file.rename(seen_file)
        logger.debug(f"Saved {len(seen_ids)} seen IDs")
    except Exception as e:
        logger.error(f"Failed to save seen IDs to {seen_file}: {e}")
        raise
    finally:
        # Clean up temp file if it still exists (e.g., if rename failed)
        if temp_file.exists():
            try:
                temp_file.unlink()
            except Exception as cleanup_error:
                logger.warning(
                    f"Failed to remove temporary file {temp_file}: {cleanup_error}"
                )


def save_attachments(msg, attachments_dir: Path) -> list[dict]:
    """Save all attachments from an Outlook message.
    
    Args:
        msg: Outlook COM message object (win32com dispatch object)
        attachments_dir: Directory to save attachments
        
    Returns:
        List of attachment metadata dicts
    """
    saved = []
    
    if msg.Attachments.Count == 0:
        return saved
    
    attachments_dir.mkdir(parents=True, exist_ok=True)
    
    # Note: Outlook COM API uses 1-based indexing (not 0-based like Python)
    for i in range(1, msg.Attachments.Count + 1):
        att = msg.Attachments.Item(i)
        filename = att.FileName
        filepath = attachments_dir / filename
        
        # Handle duplicate filenames
        counter = 1
        while filepath.exists():
            stem = filepath.stem
            suffix = filepath.suffix
            filepath = attachments_dir / f"{stem}_{counter}{suffix}"
            counter += 1
        
        try:
            att.SaveAsFile(str(filepath))
            
            saved.append({
                "filename": filepath.name,
                "original_filename": filename,
                "local_path": f"./{attachments_dir.name}/{filepath.name}",
                "size_bytes": filepath.stat().st_size,
            })
            logger.debug(f"Saved attachment: {filename}")
        except Exception as e:
            logger.error(f"Failed to save attachment {filename}: {e}")
            continue
    
    return saved
