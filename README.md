# effi-mail

MCP server for Outlook email management with triage workflow, built with [FastMCP](https://github.com/jlowin/fastmcp).

## Features

- **Email Ingestion**: Deterministic script to fetch and convert emails to markdown before triaging
- **Direct Outlook Access**: Fetch emails from Outlook via Windows COM
- **Triage via Categories**: Triage status stored as Outlook categories (effi:action, effi:waiting, effi:processed, effi:archived)
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

## Auto-File for Large Results

To reduce memory usage in long chat sessions, tools that return potentially large payloads automatically save results to a cache file when the count exceeds a threshold.

### Behaviour

| Result Count | Action |
|--------------|--------|
| ≤ 20 items | Return inline (as before) |
| > 20 items | Save to `~/.effi/cache/{prefix}_{timestamp}.json`, return 5-item preview + `full_data_file` path |

### Override Parameters

All search/retrieval tools support these optional parameters:

| Parameter | Type | Default | Description |
|-----------|------|---------|-------------|
| `output_file` | string | `""` | Agent-specified path to save results (takes priority) |
| `force_inline` | bool | `False` | Return full payload inline regardless of size |
| `auto_file_threshold` | int | `20` | Item count above which to auto-file |

### Example Response (auto-filed)

```json
{
  "count": 87,
  "limit_applied": 100,
  "results_truncated": false,
  "preview": [...5 items...],
  "full_data_file": "C:/Users/david/.effi/cache/pending_emails_20260107_143022.json",
  "auto_filed": true,
  "auto_file_note": "Results (87) exceeded threshold (20). Full data saved to file."
}
```

### Affected Tools

- `get_pending_emails`
- `get_inbox_emails_by_domain`
- `get_sent_emails_by_domain`
- `search_inbox_by_subject`
- `get_emails_by_client`
- `search_outlook_direct`
- `scan_for_commitments`
- `get_uncategorized_domains`
- `get_email_thread`
- `get_dms_emails`
- `get_dms_admin_emails`
- `search_dms`

### Cache File Structure

Auto-filed results are stored with metadata and tracking flags:

```json
{
  "metadata": {
    "created": "2026-01-07T14:30:22",
    "source_tool": "get_pending_emails",
    "total_items": 87,
    "retrieved_count": 20,
    "processed_count": 12
  },
  "items": [
    {"id": "xxx", "subject": "...", "_retrieved": true, "_processed": false},
    {"id": "yyy", "subject": "...", "_retrieved": false, "_processed": false}
  ]
}
```

### Cache Tools

| Tool | Purpose |
|------|---------|
| `read_cache_file` | Paginate through cached results; auto-marks items as retrieved |
| `mark_cache_processed` | Mark items as processed after taking action |
| `get_cache_status` | Check progress (counts, percentages) |
| `reset_cache_flags` | Reset retrieved/processed flags to reprocess |
| `list_cache_files` | List recent cache files with status |

### Cache Workflow Example

```
1. get_pending_emails() → 87 items auto-filed

2. read_cache_file(path, limit=15) → next 15 unretrieved items
   File: retrieved=15, processed=0
   
3. Agent archives 12 emails
   mark_cache_processed(path, ids=["a","b","c"...])
   File: retrieved=15, processed=12

4. read_cache_file(path, limit=15) → next 15 unretrieved
   File: retrieved=30, processed=12

5. Agent interrupted, resumes later:
   get_cache_status(path) → "87 total, 30 retrieved, 12 processed"
   read_cache_file(path, unprocessed_only=true) → 18 retrieved-but-not-processed
   
6. Need to start over:
   reset_cache_flags(path) → all flags reset to false
```

## Workflow

### Email Ingestion (Pre-Triage)

Before triaging, use the ingestion script to fetch and convert emails:

```bash
# After installation, use the entry point command:
ingest-emails --folder Inbox --limit 50

# Or run directly during development:
python scripts/ingest_emails.py --folder Inbox --limit 50
```

This saves emails to `_inbox/` as markdown with YAML frontmatter. See [Ingestion README](effi_mail/ingestion/README.md) for details.

### Email Triage

1. Use `get_pending_emails` to see un-triaged emails grouped by domain
2. Use `get_uncategorized_domains` to categorize new sender domains
3. Archive marketing emails in bulk with `batch_archive_domain`
4. Mark individual emails as action/waiting/processed using `triage_email`
5. Search client emails using `search_emails_by_client`
6. Access filed emails via DMS tools (`list_dms_clients`, `get_dms_emails`)

## Project Structure

```
effi_mail/
├── __init__.py           # Package exports
├── main.py               # FastMCP server setup & tool registration
├── config.py             # Transport configuration
├── helpers.py            # Shared utilities (outlook client, formatters)
├── ingestion/            # Email ingestion module
│   ├── ingest.py         # Main ingestion logic
│   ├── storage.py        # File operations and seen ID tracking
│   ├── thread_parser.py  # Thread content extraction
│   └── README.md         # Ingestion documentation
└── tools/
    ├── email_retrieval.py    # Email fetching tools
    ├── triage.py             # Triage status tools
    ├── domain_categories.py  # Domain categorization tools
    ├── client_search.py      # Client search tools
    └── dms.py                # DMSforLegal tools

scripts/
└── ingest_emails.py      # CLI script for email ingestion
```

## Development

Run tests:

```bash
pytest tests/ -v
```

## License

MIT
