# effi-mail Developer Guide

This guide covers non-obvious implementation details that developers need to know when working with the effi-mail codebase.

## Architecture Overview

effi-mail uses a **database-free architecture** built on **FastMCP** for tool registration. All state is stored directly in Outlook (via categories) or in simple JSON files. This eliminates sync issues and ensures the source of truth is always Outlook itself.

```
┌─────────────────┐     ┌──────────────────────────────────┐     ┌─────────────────────┐
│   MCP Client    │────▶│          effi_mail/              │────▶│  outlook_client.py  │
│  (e.g. Claude)  │     │  ┌────────────────────────────┐  │     │     (COM/MAPI)      │
└─────────────────┘     │  │  main.py (FastMCP server)  │  │     └──────────┬──────────┘
                        │  └─────────────┬──────────────┘  │                │
                        │                │                 │                ▼
                        │  ┌─────────────▼──────────────┐  │     ┌─────────────────────┐
                        │  │        tools/              │  │     │  Microsoft Outlook  │
                        │  │  ├── email_retrieval.py   │  │     │  - Email storage    │
                        │  │  ├── triage.py            │  │     │  - Triage categories│
                        │  │  ├── domain_categories.py │  │     └─────────────────────┘
                        │  │  ├── client_search.py     │  │
                        │  │  └── dms.py               │  │
                        │  └────────────────────────────┘  │
                        └────────────────┬─────────────────┘
                                         │
                                         ▼
                                ┌──────────────────┐
                                │domain_categories │
                                │     .json        │
                                │(Domain → Client/ │
                                │ Marketing/etc)   │
                                └──────────────────┘
```

### Key Design Decisions

1. **FastMCP Framework**: Uses [FastMCP](https://github.com/jlowin/fastmcp) for clean, decorator-based tool registration instead of manual `list_tools()`/`call_tool()` dispatch.

2. **No SQLite Database**: Previously, emails were synced to a local database. This created sync issues and stale data. Now all queries go directly to Outlook.

3. **Triage via Outlook Categories**: Triage status (`processed`, `deferred`, `archived`) is stored as Outlook categories on the emails themselves. This persists across devices and survives Outlook reinstalls.

4. **Domain Categories in JSON**: Domain categorization (Client, Marketing, Personal, Internal, Spam) is stored in `domain_categories.json` for simplicity and easy editing.

### Package Structure

```
effi_mail/
├── __init__.py          # Package entry, exports mcp, main, run_server
├── main.py              # FastMCP server setup, tool registration
├── config.py            # Transport configuration (stdio/http/sse)
├── helpers.py           # Shared outlook client, truncate_text, format_email_summary
└── tools/
    ├── __init__.py
    ├── email_retrieval.py   # get_pending_emails, get_inbox_emails_by_domain, get_email_details, get_email_content
    ├── triage.py            # triage_email, batch_triage, batch_archive_domain
    ├── domain_categories.py # categorize_domain, get_domain_summary, get_uncategorized_domains
    ├── client_search.py     # search_emails_by_client, search_outlook_direct
    └── dms.py               # save_attachment
```

## Triage System

Triage status is stored directly on emails using Outlook categories.

### Triage Categories

```python
TRIAGE_CATEGORY_PREFIX = "effi:"
TRIAGE_CATEGORIES = {
    "processed": "effi:processed",
    "deferred": "effi:deferred", 
    "archived": "effi:archived"
}
```

### Triage Methods (outlook_client.py)

| Method | Description |
|--------|-------------|
| `set_triage_status(email_id, status)` | Set triage status on single email |
| `get_triage_status(email_id)` | Get current triage status (or `None` if pending) |
| `clear_triage_status(email_id)` | Remove triage category, making email pending again |
| `batch_set_triage_status(email_ids, status)` | Set status on multiple emails |
| `get_pending_emails(days)` | Get emails without any triage category |
| `get_pending_emails_from_domain(domain, days)` | Get pending emails from specific domain |

### How Categories Work

Outlook categories are stored as a semicolon-separated string:
```python
# Email with multiple categories
email.Categories = "Project Alpha; effi:processed; Important"
```

When setting triage status:
1. Parse existing categories
2. Remove any existing `effi:*` category
3. Add the new triage category
4. Preserve all non-Effi categories

### Filtering by Triage Status

To find pending emails, we use DASL queries that exclude triaged emails:
```python
# Exclude emails with any effi: category
query = '@SQL=NOT("urn:schemas:httpmail:categories" LIKE \'%effi:%\')'
```

## Domain Categories

Domain categorization is stored in `domain_categories.json`:

```json
{
  "acme.com": "Client",
  "mailchimp.com": "Marketing",
  "ourcompany.com": "Internal",
  "friend@gmail.com": "Personal"
}
```

### Valid Categories
- `Client` - Email from/to client domains
- `Marketing` - Newsletters, promotions, automated emails
- `Internal` - Your organization's domains
- `Personal` - Personal contacts
- `Spam` - Unwanted emails to ignore

### Domain Category Methods (domain_categories.py)

| Method | Description |
|--------|-------------|
| `get_domain_category(domain)` | Get category or `"Uncategorized"` |
| `set_domain_category(domain, category)` | Set category for domain |
| `get_all_categories()` | Get full domain→category mapping |
| `get_domains_by_category(category)` | Get all domains with given category |

## Outlook COM Quirks

### DASL vs Jet Query Syntax

Outlook's `Items.Restrict()` method accepts two different query syntaxes that **cannot be mixed**:

#### Jet Syntax (Property-based)
Used for standard Outlook properties like dates:
```python
query = "[ReceivedTime] >= '12/01/2025 00:00 AM'"
items.Restrict(query)
```

#### DASL Syntax (Schema-based)
Used for email-specific fields. **Must have `@SQL=` prefix and quoted property names**:
```python
# ✅ CORRECT
query = '@SQL="urn:schemas:httpmail:fromemail" LIKE \'%@client.com\''

# ❌ WRONG - missing @SQL= prefix
query = 'urn:schemas:httpmail:fromemail LIKE \'%@client.com\''

# ❌ WRONG - unquoted property name
query = '@SQL=urn:schemas:httpmail:fromemail LIKE \'%@client.com\''
```

#### Chaining Filters
Since Jet and DASL can't be mixed, we chain `.Restrict()` calls:
```python
# First filter by date (Jet)
date_filtered = items.Restrict("[ReceivedTime] >= '12/01/2025'")

# Then filter by email domain (DASL)  
final = date_filtered.Restrict('@SQL="urn:schemas:httpmail:fromemail" LIKE \'%@client.com\'')
```

### Exchange vs SMTP Addresses

Internal Exchange emails have `SenderEmailType = 'EX'` and use X500 Distinguished Names:
```
/O=EXCHANGELABS/OU=EXCHANGE ADMINISTRATIVE GROUP/CN=RECIPIENTS/CN=...
```

External emails have `SenderEmailType = 'SMTP'` with normal addresses:
```
user@example.com
```

The `_message_to_email()` method handles this by checking `SenderEmailType` and using `Sender.GetExchangeUser().PrimarySmtpAddress` for Exchange users.

### MAPI Property Access

Some properties require the `PropertyAccessor` interface with MAPI property tags:

```python
# Internet Message ID (permanent identifier)
PR_INTERNET_MESSAGE_ID = "http://schemas.microsoft.com/mapi/proptag/0x1035001F"
message_id = message.PropertyAccessor.GetProperty(PR_INTERNET_MESSAGE_ID)

# Content ID for inline images
PR_CONTENT_ID = "http://schemas.microsoft.com/mapi/proptag/0x3712001F"
```

## Email Identifiers

### EntryID (Volatile)
- Outlook's native identifier
- **Changes when email is moved between folders**
- Format: Long hex string (150+ chars)
- Use for: Immediate operations within a session

### internet_message_id (Permanent)
- RFC 2822 Message-ID header
- **Persists across folder moves and Outlook reinstalls**
- Format: `<unique-id@domain.com>`
- Use for: Long-term references, cross-session lookups

The `get_email_by_id` tool auto-detects the format:
- Contains `<` and `@` → internet_message_id
- Otherwise → EntryID

### Lazy Backfill Strategy

Existing emails won't have `internet_message_id` populated. It's extracted and stored when:
1. Email is synced via `sync_emails_by_client` or `sync_email_by_id`
2. Email is retrieved via `get_email_by_id` with an EntryID

## Client Identification

### Domain Mapping

Clients are identified by email domains registered in `effi-work`:
```json
{
  "client_id": "acme",
  "domains": ["acme.com", "acme.co.uk"],
  "contact_emails": ["john.personal@gmail.com"]
}
```

### recipient_domains Field

Computed at sync time for efficient outbound email queries:
```python
# For email sent to: alice@acme.com, bob@partner.org
# Stored as: "acme.com,partner.org"
email.recipient_domains = "acme.com,partner.org"
```

This allows finding emails **sent to** a client without parsing recipient fields on every query.

## Tool Categories

Tools are organized into 5 modules under `effi_mail/tools/`. There are 17 tools total.

### Email Retrieval (`tools/email_retrieval.py`)
| Tool | Description |
|------|-------------|
| `get_pending_emails` | Get emails awaiting triage, grouped by domain |
| `get_inbox_emails_by_domain` | Get emails from a specific sender domain (Inbox only) |
| `get_email_details` | Get email metadata by EntryID or internet_message_id |
| `get_email_content` | Get full email body and attachments |

### Triage (`tools/triage.py`)
| Tool | Description |
|------|-------------|
| `triage_email` | Mark single email as processed/deferred/archived |
| `batch_triage` | Mark multiple emails with same status |
| `batch_archive_domain` | Archive all pending emails from a domain |

### Domain Categories (`tools/domain_categories.py`)
| Tool | Description |
|------|-------------|
| `categorize_domain` | Set category (Client/Marketing/Internal/Personal/Spam) for a domain |
| `get_domain_summary` | Get all domains with their categories |
| `get_uncategorized_domains` | List domains not yet categorized |

### Client Search (`tools/client_search.py`)
| Tool | Description |
|------|-------------|
| `search_emails_by_client` | Query Outlook for client correspondence using effi-clients data |
| `search_outlook_direct` | Ad-hoc Outlook queries with flexible filters |

### DMS (`tools/dms.py`)
| Tool | Description |
|------|-------------|
| `save_attachment` | Save email attachment to local filesystem |

## Testing

### Running Tests
```bash
# All tests
python -m pytest tests/ -v

# Triage category tests
python -m pytest tests/test_triage_categories.py -v

# MCP server tests
python -m pytest tests/test_mcp_server.py -v

# DMS tools tests
python -m pytest tests/test_dms_tools.py -v

# With coverage
python -m pytest tests/ --cov=effi_mail --cov-report=html
```

### Test Files

| File | Coverage |
|------|----------|
| `test_mcp_server.py` | Core MCP tool handlers, error handling, integration workflows |
| `test_triage_categories.py` | Triage status via Outlook categories, domain categories via JSON |
| `test_dms_tools.py` | Attachment saving to filesystem |

### Mocking Outlook COM

Since tools are now distributed across multiple modules, we use a `patch_outlook()` context manager that patches the `outlook` client in all tool modules simultaneously:

```python
from contextlib import ExitStack
from unittest.mock import patch, Mock

def patch_outlook(mock_client):
    """Patch outlook in all tool modules."""
    stack = ExitStack()
    modules_to_patch = [
        'effi_mail.helpers',
        'effi_mail.tools.email_retrieval',
        'effi_mail.tools.triage',
        'effi_mail.tools.domain_categories',
        'effi_mail.tools.client_search',
        'effi_mail.tools.dms',
    ]
    for module in modules_to_patch:
        stack.enter_context(patch(f'{module}.outlook', mock_client))
    return stack
```

Usage in tests:
```python
@pytest.fixture
def mock_outlook():
    mock = Mock()
    mock.set_triage_status = Mock(return_value=True)
    mock.get_triage_status = Mock(return_value=None)
    mock.batch_set_triage_status = Mock(return_value={"success": 3, "failed": 0})
    mock.get_pending_emails = Mock(return_value={
        "total": 5,
        "domains": [{"domain": "example.com", "count": 5, "emails": []}]
    })
    return mock

async def test_triage_email(mock_outlook):
    with patch_outlook(mock_outlook):
        result = await call_tool("triage_email", {
            "email_id": "test-id",
            "status": "processed"
        })
        assert json.loads(result[0].text)["success"] is True
```

### Domain Categories Testing

Tests use a temporary JSON file:
```python
@pytest.fixture
def temp_categories_file(tmp_path, monkeypatch):
    temp_file = tmp_path / "test_categories.json"
    temp_file.write_text('{}')
    monkeypatch.setattr('domain_categories.CATEGORIES_FILE', str(temp_file))
    return temp_file
```

## Common Issues

### "Condition is not valid" Error
DASL query syntax is wrong. Check:
1. `@SQL=` prefix present
2. Property name in double quotes
3. String values in single quotes
4. No mixing of Jet and DASL in one query

### Emails Not Filtering Correctly
The DASL query might be silently failing. Outlook falls back to returning all items. Wrap in try/except and log the query for debugging.

### Exchange Address Resolution Fails
The sender might not be in the GAL (Global Address List). Handle `None` from `GetExchangeUser()`:
```python
try:
    exchange_user = sender.GetExchangeUser()
    if exchange_user:
        return exchange_user.PrimarySmtpAddress
except:
    pass
return sender.Address  # Fallback
```

### internet_message_id is None
Not all emails have this header (e.g., calendar invites, meeting requests). Treat as optional.

## Data Storage

### Outlook Categories (Triage Status)

Triage status is stored as Outlook categories directly on emails:

| Status | Category Name |
|--------|---------------|
| Processed | `effi:processed` |
| Deferred | `effi:deferred` |
| Archived | `effi:archived` |
| Pending | (no effi: category) |

Categories persist across:
- Outlook restarts
- Device sync (if using Exchange/O365)
- Folder moves

### domain_categories.json

Simple JSON file mapping domains to categories:

```json
{
  "acme.com": "Client",
  "bigcorp.com": "Client", 
  "mailchimp.com": "Marketing",
  "linkedin.com": "Marketing",
  "ourcompany.com": "Internal"
}
```

### Why No Database?

The previous SQLite-based architecture had issues:

1. **Sync Drift**: Database could become stale if emails were moved/deleted in Outlook
2. **Duplicate State**: Triage status existed in both DB and needed sync
3. **Complexity**: Required sync tools and reconciliation logic

The new architecture:
- **Single Source of Truth**: Outlook is always authoritative
- **No Sync Required**: Queries go directly to Outlook
- **Simpler Code**: No database.py, no sync tools, no migrations

## Environment Setup

### Prerequisites
- Python 3.10+
- Microsoft Outlook (desktop, not web)
- Windows (COM requires it)

### Installation
```bash
cd effi-mail
python -m venv .venv
.venv\Scripts\Activate.ps1
pip install -e .
```

### Running the MCP Server

The server uses FastMCP and can be run in multiple ways:

#### Via Entry Point (Recommended)
```bash
# Stdio transport (default, for MCP clients)
effi-mail

# HTTP transport for web clients
effi-mail --transport streamable-http --port 8000

# SSE transport
effi-mail --transport sse --port 8000
```

#### VS Code Configuration
Configure in `.vscode/mcp.json`:
```json
{
  "servers": {
    "effi-mail": {
      "type": "stdio",
      "command": "effi-mail"
    }
  }
}
```

#### For Manual Testing
```python
import asyncio
from mcp_server import call_tool

result = asyncio.run(call_tool('search_emails_by_client', {
    'client_id': 'acme',
    'days': 30
}))
print(result[0].text)
```

#### Inspecting Available Tools
```python
from effi_mail import mcp

# List all registered tools
for tool in mcp._tool_manager._tools.values():
    print(f"{tool.name}: {tool.description[:60]}...")
```

## Future Enhancements

See [TODO.md](TODO.md) for planned improvements:
- Subfolder support for Outlook queries
- Smart triage suggestions based on domain category
- Bulk domain categorization from email patterns
- Triage status sync verification
