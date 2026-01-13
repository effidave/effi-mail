"""Main email ingestion logic.

Polls Outlook inbox and saves emails as markdown with YAML frontmatter.
"""

import logging
from pathlib import Path
from datetime import datetime
from typing import List, Optional
import win32com.client
import frontmatter

from effi_mail.ingestion.storage import load_seen_ids, save_seen_ids, save_attachments
from effi_mail.ingestion.thread_parser import extract_new_content

logger = logging.getLogger(__name__)

# Outlook folder constants
OL_FOLDER_INBOX = 6
OL_FOLDER_SENT = 5
OL_FOLDER_DRAFTS = 16


def get_outlook_folder(folder: str = "Inbox"):
    """Connect to Outlook and return specified folder.
    
    Args:
        folder: Folder name - "Inbox", "Sent Items", or custom folder path
                e.g. "Projects/Active" for nested folders
    
    Returns:
        Outlook folder object
    """
    outlook = win32com.client.Dispatch("Outlook.Application")
    namespace = outlook.GetNamespace("MAPI")
    
    # Handle common folder names
    folder_lower = folder.lower()
    if folder_lower == "inbox":
        return namespace.GetDefaultFolder(OL_FOLDER_INBOX)
    elif folder_lower == "sent items" or folder_lower == "sent":
        return namespace.GetDefaultFolder(OL_FOLDER_SENT)
    elif folder_lower == "drafts":
        return namespace.GetDefaultFolder(OL_FOLDER_DRAFTS)
    
    # Navigate to custom folder path (e.g. "Projects/Active")
    inbox = namespace.GetDefaultFolder(OL_FOLDER_INBOX)
    root = inbox.Parent  # Gets the mailbox root
    
    current = root
    for part in folder.split("/"):
        current = current.Folders[part]
    
    return current


def build_email_markdown(
    message_id: str,
    received: datetime,
    from_address: str,
    to_addresses: List[str],
    cc_addresses: List[str],
    subject: str,
    new_content: str,
    quoted_content: str = "",
    thread_id: Optional[str] = None,
    in_reply_to: Optional[str] = None,
    attachments: Optional[List[dict]] = None,
) -> str:
    """Build markdown string with YAML frontmatter.
    
    Args:
        message_id: Outlook EntryID
        received: DateTime when email was received
        from_address: Sender email address
        to_addresses: List of recipient emails
        cc_addresses: List of CC recipient emails
        subject: Email subject
        new_content: New content from this email
        quoted_content: Quoted/threaded content
        thread_id: Conversation ID (optional)
        in_reply_to: In-Reply-To header (optional)
        attachments: List of attachment metadata (optional)
        
    Returns:
        Markdown string with YAML frontmatter
    """
    metadata = {
        "message_id": message_id,
        "received": received.isoformat(),
        "from_address": from_address,
        "to_addresses": to_addresses,
        "cc_addresses": cc_addresses,
        "subject": subject,
        "state": "received",
    }
    
    if thread_id:
        metadata["thread_id"] = thread_id
    if in_reply_to:
        metadata["in_reply_to"] = in_reply_to
    if attachments:
        metadata["attachments"] = attachments
    
    # Build body content
    body_parts = ["# New Content\n", new_content]
    
    if quoted_content:
        body_parts.append("\n\n<details>")
        body_parts.append("<summary>Previous messages in thread</summary>\n")
        body_parts.append(quoted_content)
        body_parts.append("\n</details>")
    
    body = "\n".join(body_parts)
    
    post = frontmatter.Post(body, **metadata)
    return frontmatter.dumps(post)


def save_email(msg, inbox_path: Path) -> Path:
    """Save a single Outlook message to the inbox folder.
    
    Args:
        msg: Outlook COM message object
        inbox_path: Path to _inbox folder
    
    Returns:
        Path to saved markdown file
    """
    # Extract identifiers
    message_id = msg.EntryID
    received = msg.ReceivedTime  # This is a pywintypes datetime
    
    # Convert COM datetime to Python datetime
    received_dt = datetime(
        received.year, received.month, received.day,
        received.hour, received.minute, received.second
    )
    
    # Create filename using entry ID slug as specified
    timestamp = received_dt.strftime("%Y-%m-%d-%H%M%S")
    # Use last 16 chars of EntryID (more unique than first 12)
    slug = message_id[-16:].replace("/", "_").replace("+", "_")
    base_name = f"{timestamp}_{slug}"
    
    md_path = inbox_path / f"{base_name}.md"
    attachments_dir = inbox_path / f"{base_name}_attachments"
    
    # Save attachments first
    attachments = save_attachments(msg, attachments_dir)
    
    # Extract addresses
    from_address = msg.SenderEmailAddress
    to_addresses = [r.Address for r in msg.Recipients if r.Type == 1]  # 1 = To
    cc_addresses = [r.Address for r in msg.Recipients if r.Type == 2]  # 2 = CC
    
    # Extract thread info
    thread_id = getattr(msg, "ConversationID", None)
    
    # Extract body and handle threading
    body = msg.Body or ""
    new_content, quoted = extract_new_content(body)
    
    # Build and save markdown
    markdown = build_email_markdown(
        message_id=message_id,
        received=received_dt,
        from_address=from_address,
        to_addresses=to_addresses,
        cc_addresses=cc_addresses,
        subject=msg.Subject or "(No subject)",
        new_content=new_content,
        quoted_content=quoted,
        thread_id=thread_id,
        attachments=attachments,
    )
    
    md_path.write_text(markdown, encoding="utf-8")
    
    return md_path


def ingest_new_emails(
    inbox_path: Path,
    folder: str = "Inbox",
    limit: int = 50
) -> List[Path]:
    """Poll Outlook folder and save new emails locally.
    
    Args:
        inbox_path: Path to local _inbox folder
        folder: Outlook folder to poll - "Inbox", "Sent Items", or custom path
        limit: Maximum emails to process per run
    
    Returns:
        List of paths to newly saved email files
    """
    inbox_path.mkdir(parents=True, exist_ok=True)
    
    # Load already-seen message IDs
    seen_ids = load_seen_ids(inbox_path)
    logger.info(f"Loaded {len(seen_ids)} previously seen message IDs")
    
    # Connect to Outlook
    logger.info(f"Connecting to Outlook folder: {folder}")
    outlook_folder = get_outlook_folder(folder)
    
    saved_paths = []
    processed = 0
    
    # Iterate through folder items (most recent first)
    items = outlook_folder.Items
    items.Sort("[ReceivedTime]", True)  # Descending
    
    logger.info(f"Processing up to {limit} new emails...")
    
    for msg in items:
        if processed >= limit:
            break
        
        try:
            message_id = msg.EntryID
            
            if message_id in seen_ids:
                continue
            
            # Save immediately
            path = save_email(msg, inbox_path)
            saved_paths.append(path)
            
            # Mark as seen
            seen_ids.add(message_id)
            
            logger.info(f"Saved: {msg.Subject}")
            processed += 1
            
        except Exception as e:
            logger.error(f"Failed to save email: {e}")
            continue
    
    # Persist seen IDs
    save_seen_ids(inbox_path, seen_ids)
    
    logger.info(f"Ingestion complete: {len(saved_paths)} new emails saved")
    return saved_paths
