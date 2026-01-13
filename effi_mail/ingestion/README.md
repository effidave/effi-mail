# Email Ingestion Module

The email ingestion module provides a deterministic script to fetch emails from Outlook and convert them to markdown files with YAML frontmatter before triaging.

## Overview

The ingestion system:
1. Polls Outlook inbox (or specified folder) via COM
2. Saves emails immediately to `_inbox/` as markdown with YAML frontmatter
3. Deduplicates using message IDs tracked in `_seen.json`
4. Extracts and saves attachments to companion folders
5. Separates new content from quoted thread content

**Key Principle:** Capture emails immediately on ingestion, before any LLM processing. COM is unreliable - store locally ASAP.

## Installation

Install this module as part of the main project using the standard installation method for the repository (e.g., via `pip install .` or `pip install -e .` for development from the project root). This will automatically install all required dependencies declared in `pyproject.toml`, including `python-frontmatter`.

The script requires:
- Python 3.10+
- Windows OS (for Outlook COM access)
- `pywin32` package (Windows only, installed automatically)
- `python-frontmatter` (installed automatically via project dependencies)

## Usage

### Basic Usage

Run the ingestion script from the repository root:

```bash
python scripts/ingest_emails.py
```

This will:
- Connect to your Outlook Inbox
- Process up to 50 new emails
- Save them to `_inbox/` in the current directory

### Command-Line Options

```bash
# Specify Outlook folder
python scripts/ingest_emails.py --folder "Sent Items"

# Set processing limit
python scripts/ingest_emails.py --limit 100

# Specify custom inbox path
python scripts/ingest_emails.py --inbox-path "C:/Users/YourName/effi-work/_inbox"

# Enable verbose logging
python scripts/ingest_emails.py --verbose

# Combine options
python scripts/ingest_emails.py -f "Projects/Active" -l 20 -v
```

### Options

| Option | Short | Default | Description |
|--------|-------|---------|-------------|
| `--folder` | `-f` | `Inbox` | Outlook folder to poll (e.g., "Sent Items", "Projects/Active") |
| `--limit` | `-l` | `50` | Maximum emails to process per run |
| `--inbox-path` | | `_inbox` | Path to _inbox directory |
| `--verbose` | `-v` | | Enable debug logging |

## Output Structure

### File Naming

Emails are saved with the following filename format:

```
YYYY-MM-DD-HHMMSS_{entry_id_slug}.md
```

Example: `2026-01-13-143022_4FA4A6A560000.md`

### Attachments

Attachments are saved in companion directories:

```
_inbox/
  2026-01-13-143022_4FA4A6A560000.md
  2026-01-13-143022_4FA4A6A560000_attachments/
    document.pdf
    spreadsheet.xlsx
```

### Email Format

Each email is saved as markdown with YAML frontmatter:

```markdown
---
message_id: "ABC123..."
received: 2026-01-13T14:30:22
from_address: john.smith@example.com
to_addresses:
  - you@yourfirm.com
cc_addresses: []
subject: "Contract review"
thread_id: "THREAD456"
attachments:
  - filename: contract.pdf
    local_path: ./2026-01-13-143022_4FA4A6A560000_attachments/contract.pdf
    size_bytes: 45032
state: received
---

# New Content

Please review the attached contract and let me know your thoughts.

Thanks,
John

<details>
<summary>Previous messages in thread</summary>

> On 12 Jan, you wrote:
> I've drafted the initial version. Please have a look.

</details>
```

### Deduplication Tracking

The `_seen.json` file tracks processed message IDs:

```json
{
  "seen_ids": [
    "ABC123...",
    "DEF456...",
    "GHI789..."
  ]
}
```

## Architecture

### Module Structure

```
effi_mail/ingestion/
  __init__.py           # Module exports
  ingest.py             # Main ingestion logic and Outlook connection
  storage.py            # File operations and seen ID tracking
  thread_parser.py      # Thread content extraction
```

### Key Functions

#### `ingest_new_emails(inbox_path, folder, limit)`

Main entry point that:
- Connects to Outlook
- Loads seen message IDs
- Iterates through emails (newest first)
- Skips already-seen emails
- Saves new emails to disk
- Updates seen IDs

#### `save_email(msg, inbox_path)`

Saves a single Outlook message:
- Extracts metadata (sender, recipients, subject, etc.)
- Saves attachments
- Separates new content from quoted content
- Writes markdown file with YAML frontmatter

#### `extract_new_content(body)`

Separates email content into new vs quoted:
- Detects common reply/forward patterns
- Extracts only the new content
- Returns tuple of (new_content, quoted_content)

## Thread Content Extraction

The system automatically detects and separates quoted content using these patterns:

- **Outlook-style**: `From: ... Sent: ... To: ...` header blocks
- **Gmail-style**: `On ... wrote:` lines
- **Separator lines**: `_____` or `-----` (5+ characters)
- **Quote markers**: Lines starting with `>`

New content appears in the main body, quoted content is collapsed in a `<details>` section.

## Running Repeatedly

The ingestion script can be run repeatedly without duplicating emails:

```bash
# First run - ingests all new emails
python scripts/ingest_emails.py

# Second run - only processes emails received since last run
python scripts/ingest_emails.py
```

The `_seen.json` file ensures emails are never processed twice.

## Error Handling

The script handles common failures gracefully:

- **COM errors**: Logs error and continues with next email
- **Attachment failures**: Logs error but still saves the email
- **Missing metadata**: Uses sensible defaults (e.g., "(No subject)")

All errors are logged to help with debugging.

## Integration with Triage

After ingestion, emails in `_inbox/` are ready for triage:

1. **Ingestion phase** (this module): Fetch and convert emails to markdown
2. **Triage phase** (separate tools): Process emails, categorize, route to appropriate folders

The triage tools read from `_inbox/` but that aspect is handled separately.

## Testing

Run tests with:

```bash
pytest tests/test_ingestion.py -v
```

Tests cover:
- Seen ID persistence
- Thread content extraction
- Markdown formatting (Windows only)

Note: Some tests are skipped on non-Windows platforms due to COM dependencies.

## Limitations

- **Windows only**: Requires Windows and Outlook COM interface
- **Single account**: Connects to the default Outlook profile
- **Text-only**: Does not preserve rich formatting (HTML is converted to plain text)
- **No inline images**: Inline images are saved as attachments

## Troubleshooting

### "ModuleNotFoundError: No module named 'win32com'"

Install pywin32:
```bash
pip install pywin32
```

### "No such folder: Projects/Active"

Check that the folder exists in Outlook and the path is correct. Use forward slashes for nested folders.

### "Failed to save attachment"

Check disk space and file permissions. The error is logged but ingestion continues.

### Duplicate files with "-2" suffix

This happens if an email with the same timestamp and entry ID slug already exists. This is expected behavior to prevent overwrites.
