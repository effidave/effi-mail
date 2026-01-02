# effi-mail Developer Guide

This guide covers non-obvious implementation details that developers need to know when working with the effi-mail codebase.

## Architecture Overview

```
┌─────────────────┐     ┌──────────────────┐     ┌─────────────────┐
│   MCP Client    │────▶│   mcp_server.py  │────▶│   database.py   │
│  (e.g. Claude)  │     │   (Tool Layer)   │     │   (SQLite)      │
└─────────────────┘     └────────┬─────────┘     └─────────────────┘
                                 │
                                 ▼
                        ┌──────────────────┐
                        │ outlook_client.py│
                        │   (COM/MAPI)     │
                        └──────────────────┘
                                 │
                                 ▼
                        ┌──────────────────┐
                        │  Microsoft       │
                        │  Outlook         │
                        └──────────────────┘
```

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

### Database Tools (Fast, Synced Data)
- `search_emails_by_client` - Query local database
- `get_emails` - General email retrieval

### Direct Outlook Tools (Slower, Live Data)
- `search_outlook_by_client` - Query Outlook directly
- `search_outlook_direct` - Ad-hoc Outlook queries
- `get_email_by_id` - Retrieve single email with full body

### Sync Tools
- `sync_emails_by_client` - Sync client emails to database
- `sync_email_by_id` - Sync single email

## Testing

### Running Tests
```bash
# All tests
python -m pytest tests/ -v

# Client-centric search tests only
python -m pytest tests/test_client_centric_search.py -v

# With coverage
python -m pytest tests/ --cov=. --cov-report=html
```

### Mocking Outlook COM

Tests mock the COM interface to avoid requiring Outlook:

```python
@pytest.fixture
def mock_outlook_client(self):
    with patch('outlook_client.win32com.client'):
        from outlook_client import OutlookClient
        client = OutlookClient(db=Mock())
        client._outlook = Mock()
        client._namespace = Mock()
        return client
```

For chained `.Restrict()` calls:
```python
mock_messages.Restrict.return_value = mock_date_filtered
mock_date_filtered.Restrict.return_value = mock_dasl_filtered
mock_dasl_filtered.__iter__ = Mock(return_value=iter([]))
```

### Test Database

Tests use an in-memory SQLite database:
```python
@pytest.fixture
def test_db(tmp_path):
    db_path = tmp_path / "test.db"
    db = Database(str(db_path))
    yield db
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

## Database Schema

Key tables and fields:

```sql
CREATE TABLE emails (
    id TEXT PRIMARY KEY,           -- Outlook EntryID
    internet_message_id TEXT,      -- RFC 2822 Message-ID (nullable)
    subject TEXT,
    sender_email TEXT,
    sender_domain TEXT,
    recipients_to TEXT,            -- JSON array
    recipients_cc TEXT,            -- JSON array  
    recipient_domains TEXT,        -- Comma-separated domains
    received_time TEXT,            -- ISO format
    direction TEXT,                -- 'inbound' or 'outbound'
    triage_status TEXT,            -- 'pending', 'triaged', 'archived'
    client_id TEXT,                -- Foreign key to effi-work client
    ...
);

CREATE INDEX idx_emails_sender_domain ON emails(sender_domain);
CREATE INDEX idx_emails_client_id ON emails(client_id);
CREATE INDEX idx_emails_internet_message_id ON emails(internet_message_id);
```

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
The server is configured in `.vscode/mcp.json` and runs automatically when VS Code connects.

For manual testing:
```python
import asyncio
from mcp_server import call_tool

result = asyncio.run(call_tool('search_outlook_by_client', {
    'client_id': 'acme',
    'days': 30
}))
print(result[0].text)
```

## Future Enhancements

See [TODO.md](TODO.md) for planned improvements:
- FTS5 full-text search on email bodies
- Subfolder support for Outlook queries
- Batch backfill of internet_message_id
