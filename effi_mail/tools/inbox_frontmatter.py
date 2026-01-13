"""Inbox frontmatter tools for effi-mail MCP server.

Provides tools to add/update YAML front matter in email markdown files
within the effi-work inbox folder.
"""

import re
from pathlib import Path
from typing import Optional


# Default inbox folder path
DEFAULT_INBOX_PATH = Path("C:/Users/DavidSant/effi-work/inbox")


def parse_yaml_frontmatter(content: str) -> tuple[dict, str]:
    """Parse YAML front matter from markdown content.
    
    Args:
        content: Full markdown file content
        
    Returns:
        Tuple of (frontmatter dict, body without frontmatter)
    """
    frontmatter = {}
    body = content
    
    # Check if file starts with YAML front matter
    if content.startswith("---"):
        # Find closing ---
        lines = content.split("\n")
        end_index = None
        for i, line in enumerate(lines[1:], start=1):
            if line.strip() == "---":
                end_index = i
                break
        
        if end_index:
            # Parse the YAML between the --- markers
            yaml_lines = lines[1:end_index]
            for line in yaml_lines:
                if ":" in line:
                    key, value = line.split(":", 1)
                    key = key.strip()
                    value = value.strip()
                    # Handle boolean values
                    if value.lower() == "true":
                        value = True
                    elif value.lower() == "false":
                        value = False
                    # Handle quoted strings
                    elif value.startswith('"') and value.endswith('"'):
                        value = value[1:-1]
                    elif value.startswith("'") and value.endswith("'"):
                        value = value[1:-1]
                    frontmatter[key] = value
            
            # Body is everything after the closing ---
            body = "\n".join(lines[end_index + 1:])
    
    return frontmatter, body


def format_yaml_frontmatter(frontmatter: dict) -> str:
    """Format a dict as YAML front matter.
    
    Args:
        frontmatter: Dict of key-value pairs
        
    Returns:
        YAML front matter string with --- delimiters
    """
    if not frontmatter:
        return ""
    
    lines = ["---"]
    for key, value in frontmatter.items():
        if isinstance(value, bool):
            lines.append(f"{key}: {str(value).lower()}")
        elif value is None:
            # Skip None values
            continue
        elif isinstance(value, str):
            # Quote strings with special characters
            if any(c in value for c in [":", "#", "[", "]", "{", "}", ",", "<", ">"]):
                lines.append(f'{key}: "{value}"')
            else:
                lines.append(f"{key}: {value}")
        else:
            lines.append(f"{key}: {value}")
    lines.append("---")
    
    return "\n".join(lines)


def find_email_file_by_id(
    email_id: str,
    inbox_path: Path = DEFAULT_INBOX_PATH
) -> Optional[Path]:
    """Find an email markdown file by its Outlook EntryID.
    
    Args:
        email_id: The Outlook EntryID to search for (long hex string)
        inbox_path: Base path to search in
        
    Returns:
        Path to the matching file, or None if not found
    """
    if not inbox_path.exists():
        return None
    
    # Search all .md files in inbox and subfolders
    for filepath in inbox_path.rglob("*.md"):
        try:
            content = filepath.read_text(encoding="utf-8")
            # Look for the Email ID line (Outlook EntryID)
            if f"**Email ID:** {email_id}" in content:
                return filepath
        except Exception:
            continue
    
    return None


def add_email_frontmatter(
    email_id: str,
    client: Optional[str] = None,
    matter: Optional[str] = None,
    filed: bool = False
) -> dict:
    """Add or update YAML front matter on an email markdown file in the inbox.
    
    Searches the inbox folder and subfolders for a markdown file containing
    the specified Outlook EntryID, then adds/updates its YAML front matter
    with the provided metadata.
    
    Args:
        email_id: Outlook EntryID to search for (long hex string from **Email ID:** field)
        client: Optional client folder name for filing
        matter: Optional matter folder name for filing
        filed: Filing status - set True after email is copied to correspondence/
        
    Returns:
        Dict with:
            - success: bool
            - file_path: str (path to updated file, if found)
            - message: str (status message)
            - frontmatter: dict (current frontmatter after update)
    """
    # Find the email file
    email_file = find_email_file_by_id(email_id, DEFAULT_INBOX_PATH)
    
    if not email_file:
        return {
            "success": False,
            "file_path": None,
            "message": f"No email file found with Email ID: {email_id}",
            "frontmatter": {}
        }
    
    try:
        # Read current content
        content = email_file.read_text(encoding="utf-8")
        
        # Parse existing front matter
        existing_frontmatter, body = parse_yaml_frontmatter(content)
        
        # Build new front matter
        # Start with existing values, then overlay new ones
        new_frontmatter = dict(existing_frontmatter)
        
        # Always include email_id
        new_frontmatter["email_id"] = email_id
        
        # Update with provided values (only if not None)
        if client is not None:
            new_frontmatter["client"] = client
        if matter is not None:
            new_frontmatter["matter"] = matter
        
        # Filed is always set (default False)
        new_frontmatter["filed"] = filed
        
        # Format the new content
        frontmatter_str = format_yaml_frontmatter(new_frontmatter)
        
        # Ensure body starts with newline for clean formatting
        if body and not body.startswith("\n"):
            body = "\n" + body
        
        new_content = frontmatter_str + body
        
        # Write back
        email_file.write_text(new_content, encoding="utf-8")
        
        return {
            "success": True,
            "file_path": str(email_file),
            "message": f"Updated front matter for: {email_file.name}",
            "frontmatter": new_frontmatter
        }
        
    except Exception as e:
        return {
            "success": False,
            "file_path": str(email_file) if email_file else None,
            "message": f"Error updating file: {str(e)}",
            "frontmatter": {}
        }
