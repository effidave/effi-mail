# effi-mail

MCP server for Outlook email management with triage workflow.

## Features

- **Direct Outlook Access**: Fetch emails from Outlook via Windows COM
- **Triage via Categories**: Triage status stored as Outlook categories (Effi:Processed, Effi:Deferred, Effi:Archived)
- **Domain Categorization**: Classify sender domains as Client, Internal, Marketing, Personal, or Spam (stored in domain_categories.json)
- **Client Search**: Search emails by client using domains from effi-clients MCP server
- **Batch Operations**: Archive marketing emails, process multiple emails at once

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
- `get_pending_emails` - Get emails pending triage (no Effi: category)
- `get_emails_by_domain` - Get emails from a specific domain
- `get_email_content` - Get full email body
- `get_email_by_id` - Get email by EntryID

### Triage
- `triage_email` - Set triage status (adds Effi:* category to email)
- `batch_triage` - Triage multiple emails
- `batch_archive_domain` - Archive all pending emails from a domain

### Domain Categorization
- `get_uncategorized_domains` - List uncategorized domains from pending emails
- `categorize_domain` - Set category for a domain (saves to domain_categories.json)
- `get_domain_summary` - Summary of domains by category

### Client Search
- `search_emails_by_client` - Search Outlook for client correspondence
- `search_outlook_by_client` - Alias for search_emails_by_client
- `search_outlook_direct` - Search Outlook with flexible filters

## Workflow

1. Use `get_pending_emails` to see un-triaged emails grouped by domain
2. Use `get_uncategorized_domains` to categorize new domains
3. Archive marketing emails in bulk with `batch_archive_domain`
4. Mark individual emails as processed/deferred using `triage_email`
5. Search client emails using `search_emails_by_client`
