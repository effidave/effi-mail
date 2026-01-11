"""Workspace filing tools for effi-mail MCP server.

Provides atomic email filing to workspace folders, preventing context overflow
by returning only filenames instead of email content.

Includes thread-aware filing that splits threads into individual emails,
with deduplication and edited-quote detection.
"""

import hashlib
import json
import re
from datetime import datetime
from pathlib import Path
from typing import Optional, List, Tuple
from html import unescape

from effi_mail.helpers import outlook


def fix_mojibake(text: str) -> str:
    """Fix UTF-8 mojibake from Outlook (UTF-8 bytes decoded as Windows-1252).
    
    Outlook COM can return text where UTF-8 bytes were incorrectly decoded
    as Windows-1252, causing "smart quotes" and other Unicode characters
    to appear as garbled sequences like "â€œ" instead of proper quotes.
    
    Args:
        text: Text potentially containing mojibake
        
    Returns:
        Text with mojibake sequences replaced by correct Unicode characters
    """
    if not text:
        return text
    
    # First try the encode/decode trick: if the text looks like mojibake,
    # re-encode as windows-1252 and decode as UTF-8
    # Common mojibake signatures: sequences starting with \xe2 interpreted as Windows-1252
    # become characters like â (0xe2), € (0x80), etc.
    try:
        # Check for common mojibake patterns: â followed by special characters
        # These are UTF-8 multi-byte sequences misread as Windows-1252
        if '\xe2\x80' in text.encode('windows-1252', errors='ignore').decode('windows-1252', errors='ignore'):
            # This text might have mojibake - try to fix it
            fixed = text.encode('windows-1252', errors='surrogateescape').decode('utf-8', errors='replace')
            # Only use the fixed version if it reduced garbled sequences
            if '\ufffd' not in fixed and len(fixed) >= len(text) * 0.5:
                return fixed
    except (UnicodeDecodeError, UnicodeEncodeError, LookupError):
        pass
    
    # Try direct detection of mojibake patterns in the string
    # When UTF-8 bytes are decoded as Windows-1252:
    # U+201C (") = E2 80 9C -> â € œ (but € is 0x80 which is € in cp1252)
    # The actual characters we see depend on the Windows-1252 mapping
    
    # Common patterns when UTF-8 is misread as Windows-1252:
    mojibake_map = {
        # The key is what we see when UTF-8 bytes E2 80 XX are read as Windows-1252
        'â\x80\x9c': '\u201c',  # left double quote
        'â\x80\x9d': '\u201d',  # right double quote  
        'â\x80\x98': '\u2018',  # left single quote
        'â\x80\x99': '\u2019',  # right single quote / apostrophe
        'â\x80\x93': '\u2013',  # en-dash
        'â\x80\x94': '\u2014',  # em-dash
        'â\x80\xa2': '\u2022',  # bullet
        'â\x80\xa6': '\u2026',  # ellipsis
        'Â\xa0': '\u00a0',      # non-breaking space
        'Â ': ' ',               # NBSP that became "Â " 
    }
    
    for bad, good in mojibake_map.items():
        if bad in text:
            text = text.replace(bad, good)
    
    return text


def parse_sender_name(sender: str) -> str:
    """Extract display name from sender string and convert to filename-safe format.
    
    Args:
        sender: Sender string in formats like:
            - "Katie Brownridge <katie.brownridge@biorelate.com>"
            - "David Sant </o=ExchangeLabs/ou=Exchange Administrative Group...>"
            - "john.smith@example.com"
    
    Returns:
        Lowercase hyphenated name suitable for filenames (e.g., "katie-brownridge")
    """
    if not sender:
        return "unknown"
    
    # Extract text before < if present
    if "<" in sender:
        name_part = sender.split("<")[0].strip()
    else:
        name_part = sender.strip()
    
    # If empty or looks like an email, parse from email address
    if not name_part or "@" in name_part:
        # Extract email part
        if "<" in sender and ">" in sender:
            email_match = re.search(r"<([^>]+)>", sender)
            email = email_match.group(1) if email_match else sender
        else:
            email = sender
        
        # Skip Exchange addresses - extract name from start if possible
        if "/o=ExchangeLabs" in email or email.startswith("/"):
            # Try to find any reasonable name before the Exchange path
            pre_exchange = sender.split("/o=")[0].strip()
            if pre_exchange and pre_exchange not in ["<", ""]:
                name_part = pre_exchange.strip("<").strip()
            else:
                return "unknown"
        else:
            # Get local part of email (before @)
            local_part = email.split("@")[0] if "@" in email else email
            # Replace dots and underscores with spaces
            name_part = local_part.replace(".", " ").replace("_", " ")
    
    # Convert to filename-safe format
    name = name_part.lower()
    # Replace spaces with hyphens
    name = name.replace(" ", "-")
    # Remove any non-alphanumeric characters except hyphens
    name = re.sub(r"[^a-z0-9-]", "", name)
    # Collapse multiple hyphens
    name = re.sub(r"-+", "-", name)
    # Strip leading/trailing hyphens
    name = name.strip("-")
    
    return name if name else "unknown"


def slugify_subject(subject: str) -> str:
    """Convert email subject to filename-safe topic slug.
    
    Args:
        subject: Email subject line
    
    Returns:
        Lowercase hyphenated slug suitable for filenames (max 50 chars)
    """
    if not subject:
        return "no-subject"
    
    text = subject
    
    # Remove common prefixes (case-insensitive)
    prefixes = [r"re:\s*", r"fwd:\s*", r"fw:\s*"]
    for prefix in prefixes:
        text = re.sub(prefix, "", text, flags=re.IGNORECASE)
    
    # Remove reference numbers like [HJ-xxx-xxx-xxx] or [xxxxx-xxxxxxxxxx-xxx]
    text = re.sub(r"\[[A-Z]*-?\d+[-\d]*\]", "", text)
    
    # Strip whitespace
    text = text.strip()
    
    # Lowercase
    text = text.lower()
    
    # Replace slashes with hyphens (e.g., "GDPR/Data" -> "gdpr-data")
    text = text.replace("/", "-")
    
    # Replace spaces with hyphens
    text = text.replace(" ", "-")
    
    # Remove any non-alphanumeric characters except hyphens
    text = re.sub(r"[^a-z0-9-]", "", text)
    
    # Collapse multiple hyphens
    text = re.sub(r"-+", "-", text)
    
    # Strip leading/trailing hyphens
    text = text.strip("-")
    
    # Truncate to 50 characters at word boundary if possible
    if len(text) > 50:
        # Try to cut at a hyphen (word boundary)
        truncated = text[:50]
        last_hyphen = truncated.rfind("-")
        if last_hyphen > 30:  # Only use word boundary if reasonable
            text = truncated[:last_hyphen]
        else:
            text = truncated
    
    return text if text else "no-subject"


def html_to_plain_text(html: str) -> str:
    """Convert HTML to plain text.
    
    Args:
        html: HTML content
        
    Returns:
        Plain text with tags stripped and entities decoded
    """
    if not html:
        return ""
    
    text = html
    
    # Fix UTF-8 mojibake from Outlook before processing
    text = fix_mojibake(text)
    
    # Replace common block elements with newlines
    text = re.sub(r"<br\s*/?>", "\n", text, flags=re.IGNORECASE)
    text = re.sub(r"</p>", "\n\n", text, flags=re.IGNORECASE)
    text = re.sub(r"</div>", "\n", text, flags=re.IGNORECASE)
    text = re.sub(r"</tr>", "\n", text, flags=re.IGNORECASE)
    text = re.sub(r"</li>", "\n", text, flags=re.IGNORECASE)
    
    # Strip all remaining HTML tags
    text = re.sub(r"<[^>]+>", "", text)
    
    # Decode HTML entities
    text = unescape(text)
    
    # Replace em-dashes with regular hyphens (house style)
    text = text.replace("—", "-")
    text = text.replace("–", "-")
    
    # Normalize whitespace
    text = re.sub(r" +", " ", text)
    text = re.sub(r"\n{3,}", "\n\n", text)
    
    return text.strip()


# =============================================================================
# Thread Parsing and Deduplication Helpers
# =============================================================================

# Common reply/forward separator patterns (single-line)
REPLY_SEPARATORS = [
    # Outlook-style separators
    r"_{5,}",  # _____ (5+ underscores)
    r"-{5,}",  # ----- (5+ hyphens)
    # Gmail-style
    r"On\s+.+\s+wrote:",  # "On Mon, Jan 9, 2026, John wrote:"
    # Original message markers
    r"-+\s*Original Message\s*-+",
    r"-+\s*Forwarded Message\s*-+",
    r"-+\s*Forwarded by\s+.+\s+-+",
]

# Multi-line header block pattern (From/Sent/To/Subject on separate lines)
HEADER_BLOCK_START = re.compile(r"^From:\s+.+$", re.IGNORECASE)


def extract_new_content_only(body: str) -> str:
    """Extract only the new content from an email, stripping quoted replies.
    
    Args:
        body: Full email body text
        
    Returns:
        Just the new content before any quoted/forwarded content
    """
    if not body:
        return ""
    
    lines = body.split("\n")
    new_content_lines = []
    
    for i, line in enumerate(lines):
        # Check if this line starts a quoted section
        is_separator = False
        
        # Check for quote markers (lines starting with >)
        if line.strip().startswith(">"):
            is_separator = True
        
        # Check for multi-line header block (From: followed by Sent:/Date:, To:, Subject:)
        if not is_separator and HEADER_BLOCK_START.match(line.strip()):
            # Look ahead for Sent/Date, To, Subject pattern
            remaining = lines[i:i+5]  # Check next few lines
            remaining_text = "\n".join(remaining)
            if (re.search(r"Sent:|Date:", remaining_text, re.IGNORECASE) and
                re.search(r"To:", remaining_text, re.IGNORECASE) and
                re.search(r"Subject:", remaining_text, re.IGNORECASE)):
                is_separator = True
        
        # Check for single-line separator patterns
        if not is_separator:
            for pattern in REPLY_SEPARATORS:
                if re.search(pattern, line, re.IGNORECASE):
                    # Special case: don't treat signature separators as reply separators
                    # Signature separator is exactly "-- " or "--" at start of line
                    if line.strip() == "--" or line.strip() == "-- ":
                        continue
                    is_separator = True
                    break
        
        if is_separator:
            # Stop here - everything after is quoted content
            break
        
        new_content_lines.append(line)
    
    # Join and clean up
    result = "\n".join(new_content_lines)
    
    # Remove trailing whitespace but preserve internal formatting
    result = result.rstrip()
    
    return result


def compute_body_hash(body: str) -> str:
    """Compute a normalized hash of email body for comparison.
    
    Normalizes whitespace and removes signatures before hashing
    to make comparison more robust.
    
    Args:
        body: Email body text
        
    Returns:
        SHA256 hash of normalized body (first 16 chars)
    """
    if not body:
        return ""
    
    text = body
    
    # Normalize whitespace
    text = re.sub(r"\s+", " ", text)
    text = text.strip().lower()
    
    # Remove common signature patterns
    # Signature typically starts with "-- " on its own line
    sig_match = re.search(r"\s--\s", text)
    if sig_match:
        text = text[:sig_match.start()]
    
    # Hash it
    return hashlib.sha256(text.encode("utf-8")).hexdigest()[:16]


def find_existing_email_file(
    folder: Path, 
    internet_message_id: Optional[str] = None,
    sender_slug: Optional[str] = None,
    timestamp: Optional[str] = None
) -> Optional[Path]:
    """Check if an email is already filed in the workspace.
    
    Uses a multi-tier matching strategy:
    1. Primary: Match by Internet Message ID in file content
    2. Secondary: Match by filename pattern (timestamp + sender)
    
    Args:
        folder: Folder to search in
        internet_message_id: The email's unique Message-ID header
        sender_slug: Slugified sender name (e.g., "katie-brownridge")
        timestamp: Timestamp string (e.g., "2026-01-09-1409")
        
    Returns:
        Path to existing file if found, None otherwise
    """
    if not folder.exists():
        return None
    
    # Primary: Search for Internet Message ID in file contents
    if internet_message_id:
        for filepath in folder.glob("*.md"):
            try:
                content = filepath.read_text(encoding="utf-8")
                if f"**Internet Message ID:** {internet_message_id}" in content:
                    return filepath
            except Exception:
                continue
    
    # Secondary: Match by filename pattern
    if timestamp and sender_slug:
        pattern = f"{timestamp}__{sender_slug}____*.md"
        matches = list(folder.glob(pattern))
        if matches:
            return matches[0]
    
    return None


def get_thread_emails_for_filing(
    email_id: str,
    include_sent: bool = True
) -> Tuple[List[dict], Optional[str], Optional[str]]:
    """Retrieve all emails in a thread for filing, sorted oldest-first.
    
    Args:
        email_id: Any email ID in the thread
        include_sent: Include Sent Items folder
        
    Returns:
        Tuple of (list of email dicts with full bodies, conversation_id, error)
    """
    try:
        # Get source email to find conversation info
        source_email = outlook.get_email_full(email_id)
        if not source_email:
            return [], None, f"Email not found: {email_id}"
        
        conversation_id = source_email.get("conversation_id")
        conversation_topic = source_email.get("conversation_topic")
        
        if not conversation_id or not conversation_topic:
            # Not part of a thread, return just this email
            return [source_email], None, None
        
        # Get all emails in the conversation
        emails = outlook.get_emails_by_conversation_id(
            conversation_id=conversation_id,
            conversation_topic=conversation_topic,
            include_sent=include_sent,
            include_dms=False,
            limit=100
        )
        
        if not emails:
            # Fallback to just the source email
            return [source_email], conversation_id, None
        
        # Sort chronologically (oldest first) - critical for edit detection
        emails.sort(key=lambda e: e.received_time)
        
        # Fetch full email data for each
        full_emails = []
        for email in emails:
            full_email = outlook.get_email_full(email.id)
            if full_email:
                full_emails.append(full_email)
        
        return full_emails, conversation_id, None
        
    except Exception as e:
        return [], None, str(e)


def detect_quote_modification(
    original_body: str,
    later_email_body: str,
    original_sender: str
) -> bool:
    """Detect if an original email's content was modified when quoted.
    
    Looks for the original content in the later email's quoted section
    and checks if it differs significantly.
    
    Args:
        original_body: The original email's body
        later_email_body: Body of a later email that might quote the original
        original_sender: Sender of the original email (for finding quote attribution)
        
    Returns:
        True if the quote appears to be modified, False otherwise
    """
    if not original_body or not later_email_body:
        return False
    
    # Get just the new content from original (what would be quoted)
    original_new = extract_new_content_only(original_body)
    if not original_new or len(original_new) < 50:
        # Too short to meaningfully compare
        return False
    
    # Normalize for comparison
    original_normalized = re.sub(r"\s+", " ", original_new.lower()).strip()
    
    # Look for the quoted section in the later email
    # This is the content AFTER the separator
    later_lines = later_email_body.split("\n")
    in_quote = False
    quoted_content = []
    
    for line in later_lines:
        if in_quote:
            # Strip quote markers
            clean_line = re.sub(r"^>+\s*", "", line)
            quoted_content.append(clean_line)
        else:
            # Check if we're entering a quote
            for pattern in REPLY_SEPARATORS:
                if re.search(pattern, line, re.IGNORECASE):
                    in_quote = True
                    break
    
    if not quoted_content:
        return False
    
    quoted_text = "\n".join(quoted_content)
    quoted_normalized = re.sub(r"\s+", " ", quoted_text.lower()).strip()
    
    # Check if original content appears in the quote (with some fuzzy matching)
    # We use a substring check - if >80% of original appears, it's not modified
    original_words = set(original_normalized.split())
    quoted_words = set(quoted_normalized.split())
    
    if not original_words:
        return False
    
    overlap = len(original_words & quoted_words) / len(original_words)
    
    # If less than 70% overlap, consider it modified
    return overlap < 0.7


def format_email_markdown(email: dict) -> str:
    """Format email data as markdown.
    
    Args:
        email: Dict with email fields:
            - subject
            - sender_name, sender_email (or sender)
            - received_time (ISO timestamp)
            - body (plain text preferred)
            - html_body (fallback if body empty)
            - recipients_to (list)
            - recipients_cc (list, may be empty)
            - id (EntryID)
            - internet_message_id (optional)
            - attachments (list of {name, size})
    
    Returns:
        Formatted markdown string
    """
    subject = fix_mojibake(email.get("subject", "(No Subject)"))
    
    # Parse received time
    received = email.get("received_time", "")
    if received:
        try:
            if isinstance(received, str):
                # Handle ISO format with timezone
                dt = datetime.fromisoformat(received.replace("Z", "+00:00"))
            else:
                dt = received
            date_formatted = dt.strftime("%Y-%m-%d %H:%M")
        except (ValueError, AttributeError):
            date_formatted = str(received)
    else:
        date_formatted = "Unknown"
    
    # Format sender
    sender_name = email.get("sender_name", "")
    sender_email = email.get("sender_email", "")
    if not sender_name and not sender_email:
        # Try combined sender field
        sender = email.get("sender", "Unknown")
        if "<" in sender:
            parts = sender.split("<")
            sender_name = parts[0].strip()
            sender_email = parts[1].rstrip(">").strip()
        else:
            sender_email = sender
    
    if sender_name:
        from_line = f"{sender_name} ({sender_email})"
    else:
        from_line = sender_email
    
    # Recipients
    recipients_to = email.get("recipients_to", [])
    if isinstance(recipients_to, list):
        to_line = ", ".join(recipients_to) if recipients_to else "Unknown"
    else:
        to_line = str(recipients_to)
    
    recipients_cc = email.get("recipients_cc", [])
    if isinstance(recipients_cc, list):
        cc_line = ", ".join(recipients_cc) if recipients_cc else None
    else:
        cc_line = str(recipients_cc) if recipients_cc else None
    
    # Email IDs
    email_id = email.get("id", "")
    internet_message_id = email.get("internet_message_id", "")
    
    # Body - prefer plain text, fall back to HTML converted
    body = email.get("body", "")
    if not body or body.strip() == "":
        html_body = email.get("html_body", "")
        body = html_to_plain_text(html_body)
    else:
        # Fix mojibake in plain text body too
        body = fix_mojibake(body)
    
    # Replace em-dashes (house style)
    body = body.replace("—", "-").replace("–", "-")
    subject = subject.replace("—", "-").replace("–", "-")
    
    # Build markdown
    lines = [
        f"# Email: {subject}",
        "",
        f"**Date:** {date_formatted}",
        f"**From:** {from_line}",
        f"**To:** {to_line}",
    ]
    
    if cc_line:
        lines.append(f"**CC:** {cc_line}")
    
    lines.extend([
        f"**Subject:** {subject}",
        f"**Email ID:** {email_id}",
    ])
    
    if internet_message_id:
        lines.append(f"**Internet Message ID:** {internet_message_id}")
    
    lines.extend([
        "",
        "---",
        "",
        body,
        "",
        "---",
    ])
    
    # Attachments section (only if there are attachments)
    attachments = email.get("attachments", [])
    if attachments:
        lines.extend([
            "",
            "## Attachments",
            "",
        ])
        for att in attachments:
            name = att.get("name", att.get("filename", "unknown"))
            size = att.get("size", 0)
            # Format size
            if size >= 1024 * 1024:
                size_str = f"{size / (1024 * 1024):.1f} MB"
            elif size >= 1024:
                size_str = f"{size / 1024:.1f} KB"
            else:
                size_str = f"{size} bytes"
            lines.append(f"- {name} ({size_str})")
    
    return "\n".join(lines)


def generate_email_filename(email: dict, topic_slug: Optional[str] = None) -> str:
    """Generate filename for an email.
    
    Format: YYYY-MM-DD-HHMM_email__sender-name____topic.md
    
    Args:
        email: Email dict with received_time, sender_name/sender_email/sender
        topic_slug: Optional topic override. If None, generated from subject.
    
    Returns:
        Filename string
    """
    # Parse received time
    received = email.get("received_time", "")
    if received:
        try:
            if isinstance(received, str):
                dt = datetime.fromisoformat(received.replace("Z", "+00:00"))
            else:
                dt = received
            timestamp = dt.strftime("%Y-%m-%d-%H%M")
        except (ValueError, AttributeError):
            timestamp = datetime.now().strftime("%Y-%m-%d-%H%M")
    else:
        timestamp = datetime.now().strftime("%Y-%m-%d-%H%M")
    
    # Get sender name
    sender_name = email.get("sender_name", "")
    sender_email = email.get("sender_email", "")
    if not sender_name and not sender_email:
        sender = email.get("sender", "")
    else:
        sender = f"{sender_name} <{sender_email}>"
    
    sender_slug = parse_sender_name(sender)
    
    # Get topic
    if topic_slug:
        topic = topic_slug
    else:
        subject = email.get("subject", "")
        topic = slugify_subject(subject)
    
    return f"{timestamp}__{sender_slug}____{topic}.md"


def get_unique_filepath(folder: Path, filename: str) -> Path:
    """Get unique filepath, appending -2, -3 etc. if file exists.
    
    Args:
        folder: Target folder path
        filename: Desired filename
        
    Returns:
        Path that doesn't exist yet
    """
    filepath = folder / filename
    if not filepath.exists():
        return filepath
    
    # Split filename into base and extension
    base = filename.rsplit(".", 1)[0]
    ext = ".md"
    
    counter = 2
    while True:
        new_filename = f"{base}-{counter}{ext}"
        filepath = folder / new_filename
        if not filepath.exists():
            return filepath
        counter += 1
        if counter > 100:  # Safety limit
            raise ValueError(f"Too many duplicate files for {filename}")


def file_email_to_workspace(
    email_id: str,
    destination_folder: str,
    topic_slug: Optional[str] = None,
) -> str:
    """Atomically fetch and file an email to a workspace folder.
    
    Fetches the email, formats it as markdown, and saves to the destination
    folder. Returns only the filename - the agent never sees the email body,
    preventing context overflow when processing multiple emails.
    
    Args:
        email_id: Outlook EntryID of the email to file
        destination_folder: Full filesystem path to target folder
            (e.g., "C:/Users/DavidSant/effi-work/clients/Biorelate Limited/projects/GDPR Compliance/correspondence")
        topic_slug: Optional topic for filename. If omitted, auto-generated from subject.
    
    Returns:
        JSON with success status and filename/path, or error message.
        
    Filename format: YYYY-MM-DD-HHMM_email__sender-name____topic.md
    """
    if not email_id:
        return json.dumps({
            "success": False,
            "error": "email_id parameter is required"
        })
    
    if not destination_folder:
        return json.dumps({
            "success": False,
            "error": "destination_folder parameter is required"
        })
    
    try:
        # Fetch email
        email = outlook.get_email_full(email_id)
        if not email:
            return json.dumps({
                "success": False,
                "error": f"Email not found: {email_id}"
            })
        
        # Create destination folder if it doesn't exist
        folder_path = Path(destination_folder)
        folder_path.mkdir(parents=True, exist_ok=True)
        
        # Generate filename
        filename = generate_email_filename(email, topic_slug)
        
        # Get unique filepath
        filepath = get_unique_filepath(folder_path, filename)
        
        # Format email as markdown
        markdown_content = format_email_markdown(email)
        
        # Write file
        filepath.write_text(markdown_content, encoding="utf-8")
        
        return json.dumps({
            "success": True,
            "filename": filepath.name,
            "path": str(filepath).replace("\\", "/")
        })
        
    except Exception as e:
        return json.dumps({
            "success": False,
            "error": str(e)
        })


def format_email_markdown_new_content_only(email: dict) -> str:
    """Format email as markdown with only the NEW content (quotes stripped).
    
    Same as format_email_markdown but strips quoted replies from the body.
    
    Args:
        email: Email dict with standard fields
        
    Returns:
        Formatted markdown with only new content
    """
    # Get the body
    body = email.get("body", "")
    if not body or body.strip() == "":
        html_body = email.get("html_body", "")
        body = html_to_plain_text(html_body)
    
    # Extract only new content
    new_content = extract_new_content_only(body)
    
    # Create a modified email dict with stripped body
    email_copy = email.copy()
    email_copy["body"] = new_content
    email_copy["html_body"] = ""  # Don't fall back to HTML
    
    return format_email_markdown(email_copy)


def file_thread_to_workspace(
    email_id: str,
    destination_folder: str,
    topic_slug: Optional[str] = None,
    include_sent: bool = True,
    strip_quotes: bool = True,
) -> str:
    """File all emails in a thread to workspace, each as a separate file.
    
    Retrieves all emails in the conversation, processes them oldest-first,
    and files each as a separate markdown file. Handles:
    - Deduplication: Skips emails already filed to the folder
    - Quote stripping: Removes quoted content so each file has only new content
    - Edit detection: If quoted content was modified, preserves thread structure
    
    Args:
        email_id: Any email ID in the thread (we'll find all related emails)
        destination_folder: Full filesystem path to target folder
        topic_slug: Optional topic for filenames. If omitted, auto-generated from subject.
        include_sent: Include Sent Items folder when finding thread emails (default: True)
        strip_quotes: Strip quoted replies from each email (default: True)
    
    Returns:
        JSON with success status, filed/skipped files, and thread info
    """
    if not email_id:
        return json.dumps({
            "success": False,
            "error": "email_id parameter is required"
        })
    
    if not destination_folder:
        return json.dumps({
            "success": False,
            "error": "destination_folder parameter is required"
        })
    
    try:
        # Create destination folder if it doesn't exist
        folder_path = Path(destination_folder)
        folder_path.mkdir(parents=True, exist_ok=True)
        
        # Get all thread emails (oldest first)
        emails, conversation_id, error = get_thread_emails_for_filing(
            email_id, include_sent=include_sent
        )
        
        if error:
            return json.dumps({
                "success": False,
                "error": error
            })
        
        if not emails:
            return json.dumps({
                "success": False,
                "error": "No emails found in thread"
            })
        
        # Determine topic slug from first email if not provided
        if not topic_slug:
            topic_slug = slugify_subject(emails[0].get("subject", ""))
        
        # Track results
        filed = []
        skipped = []
        edit_detected = None
        
        # Process emails oldest-first
        for i, email in enumerate(emails):
            email_msg_id = email.get("internet_message_id", "")
            sender_name = email.get("sender_name", "")
            sender_email = email.get("sender_email", "")
            if not sender_name and not sender_email:
                sender = email.get("sender", "")
            else:
                sender = f"{sender_name} <{sender_email}>"
            sender_slug = parse_sender_name(sender)
            
            # Generate timestamp for filename matching
            received = email.get("received_time", "")
            if received:
                try:
                    if isinstance(received, str):
                        dt = datetime.fromisoformat(received.replace("Z", "+00:00"))
                    else:
                        dt = received
                    timestamp = dt.strftime("%Y-%m-%d-%H%M")
                except (ValueError, AttributeError):
                    timestamp = None
            else:
                timestamp = None
            
            # Check if already filed
            existing = find_existing_email_file(
                folder_path,
                internet_message_id=email_msg_id,
                sender_slug=sender_slug,
                timestamp=timestamp
            )
            
            if existing:
                skipped.append({
                    "filename": existing.name,
                    "reason": "already_exists",
                    "internet_message_id": email_msg_id
                })
                continue
            
            # Check for quote modification in later emails (if we have more emails)
            quote_modified = False
            if i < len(emails) - 1 and strip_quotes:
                email_body = email.get("body", "")
                for later_email in emails[i + 1:]:
                    if detect_quote_modification(
                        email_body,
                        later_email.get("body", ""),
                        sender
                    ):
                        quote_modified = True
                        edit_detected = {
                            "at_email_index": i,
                            "internet_message_id": email_msg_id,
                            "reason": "Quoted content was edited in a later email"
                        }
                        break
            
            # Generate filename
            filename = generate_email_filename(email, topic_slug)
            
            # Get unique filepath
            filepath = get_unique_filepath(folder_path, filename)
            
            # Format email - strip quotes unless modification detected
            if strip_quotes and not quote_modified:
                markdown_content = format_email_markdown_new_content_only(email)
            else:
                markdown_content = format_email_markdown(email)
            
            # Write file
            filepath.write_text(markdown_content, encoding="utf-8")
            
            filed.append({
                "filename": filepath.name,
                "path": str(filepath).replace("\\", "/"),
                "status": "new",
                "internet_message_id": email_msg_id,
                "quotes_preserved": quote_modified
            })
        
        # Build response
        result = {
            "success": True,
            "filed": filed,
            "skipped": skipped,
            "thread_info": {
                "conversation_id": conversation_id,
                "total_emails": len(emails),
                "filed_count": len(filed),
                "skipped_count": len(skipped)
            }
        }
        
        if edit_detected:
            result["edit_detected"] = edit_detected
        
        return json.dumps(result, indent=2)
        
    except Exception as e:
        return json.dumps({
            "success": False,
            "error": str(e)
        })
