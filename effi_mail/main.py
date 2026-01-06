"""FastMCP server for Outlook email management - effi-mail."""

from fastmcp import FastMCP

from effi_mail.config import get_transport_config
from effi_mail.tools import (
    # Email retrieval
    get_pending_emails,
    get_inbox_emails_by_domain,
    get_email_by_id,
    download_attachment,
    # Triage
    triage_email,
    batch_triage,
    batch_archive_domain,
    # Domain categories
    get_uncategorized_domains,
    categorize_domain,
    get_domain_summary,
    # Client search
    get_emails_by_client,
    search_outlook_direct,
    # DMS
    list_dms_clients,
    list_dms_matters,
    get_dms_emails,
    search_dms,
    file_email_to_dms,
    batch_file_emails_to_dms,
)


# Create FastMCP server
mcp = FastMCP("effi-mail")

# Register email retrieval tools
mcp.tool()(get_pending_emails)
mcp.tool()(get_inbox_emails_by_domain)
mcp.tool()(get_email_by_id)
mcp.tool()(download_attachment)

# Register triage tools
mcp.tool()(triage_email)
mcp.tool()(batch_triage)
mcp.tool()(batch_archive_domain)

# Register domain categorization tools
mcp.tool()(get_uncategorized_domains)
mcp.tool()(categorize_domain)
mcp.tool()(get_domain_summary)

# Register client search tools
mcp.tool()(get_emails_by_client)
mcp.tool()(search_outlook_direct)

# Register DMS tools
mcp.tool()(list_dms_clients)
mcp.tool()(list_dms_matters)
mcp.tool()(get_dms_emails)
mcp.tool()(search_dms)
mcp.tool()(file_email_to_dms)
mcp.tool()(batch_file_emails_to_dms)


def run_server():
    """Run the MCP server with configured transport."""
    config = get_transport_config()
    
    if config['transport'] == 'stdio':
        mcp.run(transport='stdio')
    elif config['transport'] == 'streamable-http':
        mcp.run(transport='streamable-http', host=config['host'], port=config['port'])
    elif config['transport'] == 'sse':
        mcp.run(transport='sse', host=config['host'], port=config['port'])
    else:
        mcp.run(transport='stdio')


def main():
    """Entry point for effi-mail MCP server."""
    run_server()


if __name__ == "__main__":
    main()
