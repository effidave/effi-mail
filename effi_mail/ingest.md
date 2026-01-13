# Email Ingestion System - Build Instructions

Instructions for building Phase 1 of the email processing system: Email Ingestion.

**Reference Documents:**
- [Final_Solution.md](Final_Solution.md) - Architecture and design decisions
- [System_formatted.md](System_formatted.md) - Detailed implementation discussion

---

## Objective

Build a system that:
1. Polls Outlook inbox via COM
2. Saves emails immediately to `_inbox/` as Markdown with YAML frontmatter
3. Deduplicates using message_id
4. Extracts and saves attachments
5. Handles email threads appropriately

**Key Principle:** Capture emails immediately on ingestion, before any LLM processing. COM is unreliable - store locally ASAP.

---

## Dependencies

Install these packages:

```bash
pip install pywin32           # Outlook COM automation
pip install python-frontmatter # Markdown with YAML frontmatter
pip install quotequail        # Email thread extraction (optional - can use regex)
pip install pydantic          # Data models
pip install pytest            # Testing
```

---

## Folder Structure

```
_inbox/                             # Unprocessed emails land here first
  {YYYY-MM-DD-HHMMSS}_{entry_id_slug}.md
  {YYYY-MM-DD-HHMMSS}_{entry_id_slug}_attachments/
    attachment1.docx
    attachment2.pdf
  _seen.json                       # Track processed message IDs

---

## Data Models

Define Pydantic models for email data:

```python
from pydantic import BaseModel
from datetime import datetime
from typing import Optional

class Attachment(BaseModel):
    filename: str
    local_path: str
    size_bytes: int
    content_type: Optional[str] = None

class EmailMetadata(BaseModel):
    message_id: str
    received: datetime
    from_address: str
    to_addresses: list[str]
    cc_addresses: list[str] = []
    subject: str
    thread_id: Optional[str] = None
    in_reply_to: Optional[str] = None
    attachments: list[Attachment] = []
    state: str = "received"
```

---

## Email Storage Format

Save emails as Markdown with YAML frontmatter:

```markdown
---
message_id: "000000000243D14276FFF7478FD384CE8FBFD6B00700D5E87B7F419C4B46ACB9A64273A5643400000000010C0000D5E87B7F419C4B46ACB9A64273A564340004FA4A6A560000"
received: 2025-01-11T14:30:22Z
from_address: john.smith@acmecorp.com
to_addresses:
  - you@yourfirm.com
cc_addresses: []
subject: "RE: Contract review"
thread_id: "AAQkAGE0M2I5N2M2LTNmNTQtNDFhMi1hMDU2LWM2MWE0NzY3YzBiNAAQABCDEF"
in_reply_to: null
attachments:
  - filename: contract_draft.docx
    local_path: ./2025-01-11-143022_4FA4A6A560000_attachments/contract_draft.docx
    size_bytes: 45032
state: received
---

# New Content

Please review clause 4.2 and confirm it's acceptable. We need to 
close by Friday if possible.

Also, can you draft the settlement statement based on the attached?

Thanks,
John

<details>
<summary>Previous messages in thread</summary>

> On 10 Jan, you wrote:
> Please find attached the initial draft for your review.

</details>
```

---

## Implementation Steps

### Step 1: Outlook COM Connection

```python
import win32com.client
from pathlib import Path

# Outlook folder constants
OL_FOLDER_INBOX = 6
OL_FOLDER_SENT = 5
OL_FOLDER_DRAFTS = 16

def get_outlook_folder(folder: str = "Inbox"):
    """
    Connect to Outlook and return specified folder.
    
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
```

### Step 2: Track Seen Messages

Maintain a simple JSON file of seen message IDs:

```python
import json
from pathlib import Path

def load_seen_ids(inbox_path: Path) -> set:
    """Load set of already-processed message IDs."""
    seen_file = inbox_path / "_seen.json"
    if seen_file.exists():
        data = json.loads(seen_file.read_text(encoding="utf-8"))
        return set(data.get("seen_ids", []))
    return set()

def save_seen_ids(inbox_path: Path, seen_ids: set):
    """Persist seen message IDs. Use atomic write."""
    seen_file = inbox_path / "_seen.json"
    temp_file = seen_file.with_suffix(".tmp")
    
    data = {"seen_ids": list(seen_ids)}
    temp_file.write_text(json.dumps(data, indent=2), encoding="utf-8")
    temp_file.rename(seen_file)
```

### Step 3: Thread Content Extraction

Extract new content from threaded emails:

```python
import re

def extract_new_content(body: str) -> tuple[str, str]:
    """
    Extract new content from email, separating from quoted thread.
    
    Returns: (new_content, quoted_remainder)
    """
    # Common patterns indicating quoted content
    patterns = [
        r'\n\s*On .+wrote:\s*\n',              # "On 10 Jan, John wrote:"
        r'\n\s*From:.+\nSent:.+\nTo:.+\n',     # Outlook-style header block
        r'\n\s*-{3,}\s*Original Message',      # "--- Original Message ---"
        r'\n_{5,}\n',                          # Underscore separator
    ]
    
    earliest_match = len(body)
    for pattern in patterns:
        match = re.search(pattern, body, re.IGNORECASE | re.MULTILINE)
        if match and match.start() < earliest_match:
            earliest_match = match.start()
    
    new_content = body[:earliest_match].strip()
    quoted = body[earliest_match:].strip() if earliest_match < len(body) else ""
    
    return new_content, quoted
```

**Alternative:** Use `quotequail` library:

```python
import quotequail

def extract_new_content_quotequail(body: str) -> tuple[str, str]:
    result = quotequail.unwrap(body)
    if result:
        return result.get("text", body), result.get("quoted", "")
    return body, ""
```

### Step 4: Save Attachments

```python
from pathlib import Path

def save_attachments(msg, attachments_dir: Path) -> list[dict]:
    """Save all attachments from an Outlook message."""
    saved = []
    
    if msg.Attachments.Count == 0:
        return saved
    
    attachments_dir.mkdir(parents=True, exist_ok=True)
    
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
        
        att.SaveAsFile(str(filepath))
        
        saved.append({
            "filename": filename,
            "local_path": f"./{attachments_dir.name}/{filepath.name}",
            "size_bytes": filepath.stat().st_size,
        })
    
    return saved
```

### Step 5: Build Markdown Email

```python
import frontmatter
from datetime import datetime

def build_email_markdown(
    message_id: str,
    received: datetime,
    from_address: str,
    to_addresses: list[str],
    cc_addresses: list[str],
    subject: str,
    new_content: str,
    quoted_content: str = "",
    thread_id: str = None,
    in_reply_to: str = None,
    attachments: list[dict] = None,
) -> str:
    """Build markdown string with YAML frontmatter."""
    
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
```

### Step 6: Save Single Email

```python
from pathlib import Path
from datetime import datetime

def save_email(msg, inbox_path: Path) -> Path:
    """
    Save a single Outlook message to the inbox folder.
    
    Returns: Path to saved markdown file
    """
    # Extract identifiers
    message_id = msg.EntryID
    received = msg.ReceivedTime  # This is a pywintypes datetime
    
    # Convert COM datetime to Python datetime
    received_dt = datetime(
        received.year, received.month, received.day,
        received.hour, received.minute, received.second
    )
    
    # Create filename
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
    # in_reply_to is harder to get from COM - may need to parse headers
    
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
```

### Step 7: Main Ingestion Loop

```python
from pathlib import Path
import logging

logger = logging.getLogger(__name__)

def ingest_new_emails(
    inbox_path: Path,
    folder: str = "Inbox",
    limit: int = 50
) -> list[Path]:
    """
    Poll Outlook folder and save new emails locally.
    
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
    
    # Connect to Outlook
    outlook_folder = get_outlook_folder(folder)
    
    saved_paths = []
    processed = 0
    
    # Iterate through folder items (most recent first)
    items = outlook_folder.Items
    items.Sort("[ReceivedTime]", True)  # Descending
    
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
    
    return saved_paths
```

### Step 8: Entry Point

```python
import argparse
from pathlib import Path

def main():
    """Run email ingestion."""
    parser = argparse.ArgumentParser(description="Ingest emails from Outlook")
    parser.add_argument(
        "--folder", "-f",
        default="Inbox",
        help="Outlook folder to poll (default: Inbox). Examples: 'Sent Items', 'Projects/Active'"
    )
    parser.add_argument(
        "--limit", "-l",
        type=int,
        default=50,
        help="Maximum emails to process (default: 50)"
    )
    args = parser.parse_args()
    
    workspace = Path(r"C:\Users\DavidSant\effi-work")
    inbox_path = workspace / "_inbox"
    
    saved = ingest_new_emails(inbox_path, folder=args.folder, limit=args.limit)
    print(f"Ingested {len(saved)} new emails from '{args.folder}'")

if __name__ == "__main__":
    main()
```

**Usage examples:**

```bash
# Default - poll Inbox
python -m email_ingestion.ingest

# Poll Sent Items
python -m email_ingestion.ingest --folder "Sent Items"

# Poll a custom folder
python -m email_ingestion.ingest -f "Clients/Active"

# Limit to 10 emails
python -m email_ingestion.ingest -f Inbox -l 10
```

---

## Testing Strategy

### Unit Tests (Pure Functions)

Test without Outlook or file system:

```python
# tests/unit/test_thread_extraction.py

def test_extract_outlook_style_quote():
    body = """Thanks, approved.

From: You
Sent: 10 January 2025
To: John
Subject: RE: Contract

Please review."""
    
    new, quoted = extract_new_content(body)
    
    assert new == "Thanks, approved."
    assert "Please review" in quoted

def test_extract_on_date_wrote():
    body = """Looks good.

On 10 Jan 2025, John wrote:
> Initial draft attached."""
    
    new, quoted = extract_new_content(body)
    
    assert new == "Looks good."
    assert "Initial draft" in quoted

def test_no_quoted_content():
    body = "Simple email with no thread."
    
    new, quoted = extract_new_content(body)
    
    assert new == body
    assert quoted == ""
```

### Integration Tests (File System)

Test with temporary directories:

```python
# tests/integration/test_email_save.py
import tempfile
from pathlib import Path
import frontmatter

def test_build_and_parse_email_markdown(tmp_path):
    markdown = build_email_markdown(
        message_id="ABC123",
        received=datetime(2025, 1, 11, 14, 30, 0),
        from_address="john@example.com",
        to_addresses=["you@firm.com"],
        cc_addresses=[],
        subject="Test email",
        new_content="Hello world",
        quoted_content="",
    )
    
    # Save and reload
    path = tmp_path / "test.md"
    path.write_text(markdown, encoding="utf-8")
    
    loaded = frontmatter.load(path)
    
    assert loaded["message_id"] == "ABC123"
    assert loaded["from_address"] == "john@example.com"
    assert "Hello world" in loaded.content

def test_seen_ids_persistence(tmp_path):
    seen = {"id1", "id2", "id3"}
    
    save_seen_ids(tmp_path, seen)
    loaded = load_seen_ids(tmp_path)
    
    assert loaded == seen
```

### COM Mocking Tests

Test ingestion logic with mocked Outlook:

```python
# tests/unit/test_ingestion.py
from unittest.mock import MagicMock, patch

def test_skips_already_seen_emails(tmp_path):
    # Pre-populate seen IDs
    save_seen_ids(tmp_path, {"existing_id"})
    
    # Mock Outlook
    mock_msg = MagicMock()
    mock_msg.EntryID = "existing_id"
    mock_msg.Subject = "Old email"
    
    mock_inbox = MagicMock()
    mock_inbox.Items = [mock_msg]
    
    with patch("mymodule.get_outlook_inbox", return_value=mock_inbox):
        saved = ingest_new_emails(tmp_path, limit=10)
    
    assert len(saved) == 0

def test_saves_new_email(tmp_path):
    mock_msg = MagicMock()
    mock_msg.EntryID = "new_id_123"
    mock_msg.Subject = "New email"
    mock_msg.Body = "Hello"
    mock_msg.SenderEmailAddress = "sender@example.com"
    mock_msg.Recipients = []
    mock_msg.Attachments.Count = 0
    mock_msg.ReceivedTime = MagicMock(
        year=2025, month=1, day=11,
        hour=14, minute=30, second=0
    )
    
    mock_inbox = MagicMock()
    mock_inbox.Items = [mock_msg]
    mock_inbox.Items.Sort = MagicMock()
    
    with patch("mymodule.get_outlook_inbox", return_value=mock_inbox):
        saved = ingest_new_emails(tmp_path, limit=10)
    
    assert len(saved) == 1
    assert saved[0].exists()
```

---

## Error Handling

Handle common COM failures gracefully:

```python
import pythoncom
import pywintypes

def ingest_new_emails_safe(
    inbox_path: Path,
    folder: str = "Inbox",
    limit: int = 50
) -> list[Path]:
    """Ingestion with COM error handling."""
    
    # Initialize COM for this thread
    pythoncom.CoInitialize()
    
    try:
        return ingest_new_emails(inbox_path, folder=folder, limit=limit)
    
    except pywintypes.com_error as e:
        logger.error(f"Outlook COM error: {e}")
        return []
    
    finally:
        pythoncom.CoUninitialize()
```

---

## File Structure

Organize the code as follows:

```
src/
  email_ingestion/
    __init__.py
    models.py          # Pydantic models
    outlook.py         # COM connection and message iteration
    thread_parser.py   # Thread content extraction
    storage.py         # Markdown save/load, seen IDs
    ingest.py          # Main ingestion loop
    
tests/
  unit/
    test_thread_extraction.py
    test_markdown_builder.py
  integration/
    test_file_storage.py
    test_seen_ids.py
  fixtures/
    sample_emails/
```

---

## Acceptance Criteria

The ingestion system is complete when:

1. ✅ Connects to Outlook via COM without errors
2. ✅ Saves new emails as Markdown with YAML frontmatter
3. ✅ Skips already-seen emails (deduplication works)
4. ✅ Extracts new content from threaded emails
5. ✅ Saves attachments to companion folder
6. ✅ Handles COM errors gracefully (doesn't crash)
7. ✅ All unit tests pass
8. ✅ Integration tests pass with temp directories
9. ✅ Can be run repeatedly without duplicating emails

---

## Next Phase

Once ingestion is working, proceed to Phase 2: Classification and Routing.

The processing pipeline will:
1. Scan `_inbox/` for unprocessed emails
2. Classify each email (client, internal, personal, marketing, spam)
3. Match client emails to client/matter folders
4. Move processed emails to `{client}/{matter}/correspondence/`
