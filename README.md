# effi-mail

MCP server for Outlook email management with triage workflow.

## Features

- **Email Sync**: Fetch emails from Outlook via Windows COM
- **SQLite Storage**: Persistent storage of email metadata and triage status
- **Domain Categorization**: Classify sender domains as Client, Internal, Marketing, or Personal
- **Triage Workflow**: Mark emails as pending, processed, deferred, or archived
- **Client/Matter Association**: Link emails to clients and matters for legal workflow
- **Batch Operations**: Archive marketing emails, process multiple emails at once

## Installation

```bash
pip install -e .
```

## MCP Configuration

Add to your MCP settings:

```json
{
  "mcpServers": {
    "effi-mail": {
      "command": "python",
      "args": ["path/to/mcp_server.py"]
    }
  }
}
```

## Tools

### Email Retrieval
- `sync_emails` - Sync emails from Outlook to database
- `get_pending_emails` - Get emails pending triage
- `get_emails_by_domain` - Get emails from a specific domain
- `get_email_content` - Get full email body

### Triage
- `triage_email` - Set triage status for one email
- `batch_triage` - Triage multiple emails
- `batch_archive_domain` - Archive all emails from a domain

### Domain Categorization
- `get_uncategorized_domains` - List uncategorized domains
- `categorize_domain` - Set category for a domain
- `get_domain_summary` - Summary of domains by category

### Client/Matter Management
- `create_client` - Create client record
- `create_matter` - Create matter for client
- `list_clients` - List all clients and matters

### Statistics
- `get_triage_stats` - Email counts by status and category

## Workflow

1. Call `sync_emails` to fetch recent emails
2. Use `get_uncategorized_domains` to categorize new domains
3. Use `get_pending_emails` with category filters to triage
4. Archive marketing emails in bulk with `batch_archive_domain`
5. Process client emails individually, linking to matters
