# effi-mail

MCP server for Outlook email management with triage workflow, built with [FastMCP](https://github.com/jlowin/fastmcp).

## Features

- **Direct Outlook Access**: Fetch emails from Outlook via Windows COM
- **Triage via Categories**: Triage status stored as Outlook categories (effi:processed, effi:deferred, effi:archived)
- **Domain Categorization**: Classify sender domains as Client, Internal, Marketing, Personal, or Spam (stored in domain_categories.json)
- **Client Search**: Search emails by client using domains from effi-clients MCP server
- **Batch Operations**: Archive marketing emails, process multiple emails at once
- **DMS Integration**: Read-only access to emails filed in DMSforLegal

## Architecture

This server operates **without a local database**:
- **Triage status** is stored directly on emails as Outlook categories
- **Domain categories** are stored in `domain_categories.json`
- **Client information** is retrieved from the effi-clients MCP server
- All email queries go directly to Outlook COM

## Installation

```bash
pip install -e .
```

Requires Python 3.10+ and Windows (for Outlook COM access).

## MCP Configuration

Add to your MCP settings:

```json
{
  "mcpServers": {
    "effi-mail": {
      "command": "effi-mail"
    }
  }
}
```

Or run directly:

```json
{
  "mcpServers": {
    "effi-mail": {
      "command": "python",
      "args": ["-m", "effi_mail.main"]
    }
  }
}
```

### Transport Options

The server supports multiple transports via environment variables:

| Variable | Default | Options |
|----------|---------|---------|
| `MCP_TRANSPORT` | `stdio` | `stdio`, `streamable-http`, `sse` |
| `MCP_HOST` | `0.0.0.0` | Any valid host |
| `MCP_PORT` | `8000` | Any valid port |

## Tools (17 total)

### Email Retrieval
- `get_pending_emails` - Get emails pending triage (no effi: category), grouped by domain
- `get_inbox_emails_by_domain` - Get emails from a specific domain (Inbox only)
- `get_email_content` - Get full email body by ID
- `get_email_by_id` - Get email details by EntryID or internet_message_id

### Triage
- `triage_email` - Set triage status (adds effi:* category to email)
- `batch_triage` - Triage multiple emails at once
- `batch_archive_domain` - Archive all pending emails from a domain

### Domain Categorization
- `get_uncategorized_domains` - List domains that haven't been categorized yet
- `categorize_domain` - Set category for a domain (saves to domain_categories.json)
- `get_domain_summary` - Summary of domains by category

### Client Search
- `search_emails_by_client` - Search Outlook for client correspondence (uses effi-clients)
- `search_outlook_by_client` - Alias for search_emails_by_client
- `search_outlook_direct` - Search Outlook with flexible filters

### DMS (DMSforLegal)
Read-only access to emails filed in the DMSforLegal Outlook store.

- `list_dms_clients` - List all client folders in DMSforLegal
- `list_dms_matters` - List matter folders for a client
- `get_dms_emails` - Get emails filed under a client/matter
- `search_dms` - Search across DMS with client/matter/subject/date filters

**DMS folder structure**: `\\DMSforLegal\_My Matters\{Client}\{Matter}\Emails`

## Workflow

1. Use `get_pending_emails` to see un-triaged emails grouped by domain
2. Use `get_uncategorized_domains` to categorize new sender domains
3. Archive marketing emails in bulk with `batch_archive_domain`
4. Mark individual emails as processed/deferred using `triage_email`
5. Search client emails using `search_emails_by_client`
6. Access filed emails via DMS tools (`list_dms_clients`, `get_dms_emails`)

## Project Structure

```
effi_mail/
├── __init__.py           # Package exports
├── main.py               # FastMCP server setup & tool registration
├── config.py             # Transport configuration
├── helpers.py            # Shared utilities (outlook client, formatters)
└── tools/
    ├── email_retrieval.py    # Email fetching tools
    ├── triage.py             # Triage status tools
    ├── domain_categories.py  # Domain categorization tools
    ├── client_search.py      # Client search tools
    └── dms.py                # DMSforLegal tools
```

## Development

Run tests:

```bash
pytest tests/ -v
```

## License

MIT
