"""FastMCP server for Outlook email management - effi-mail."""

from fastmcp import FastMCP

from effi_mail.config import get_transport_config
from effi_mail.tools import (
    # Email retrieval
    get_pending_emails,
    get_inbox_emails_by_domain,
    get_sent_emails_by_domain,
    get_email_by_id,
    download_attachment,
    search_inbox_by_subject,
    # Triage
    triage_email,
    batch_triage,
    batch_archive_domain,
    archive_email,
    batch_archive_emails,
    list_subfolders,
    # Domain categories
    get_uncategorized_domains,
    categorize_domain,
    get_domain_summary,
    # Client search
    get_emails_by_client,
    search_outlook_direct,
    # Commitment scanning
    scan_for_commitments,
    mark_scanned,
    batch_mark_scanned,
    # DMS
    list_dms_clients,
    list_dms_matters,
    get_dms_emails,
    get_dms_admin_emails,
    search_dms,
    file_email_to_dms,
    file_admin_email_to_dms,
    batch_file_emails_to_dms,
    # Workspace filing
    file_email_to_workspace,
    file_thread_to_workspace,
    # Thread tracking
    get_email_thread,
    get_thread_locations,
    # Cache operations
    read_cache_file,
    mark_cache_processed,
    get_cache_status,
    reset_cache_flags,
    list_cache_files,
    # Inbox frontmatter
    add_email_frontmatter,
)


# Create FastMCP server
mcp = FastMCP("effi-mail")

# Register email retrieval tools
mcp.tool()(get_pending_emails)
mcp.tool()(get_inbox_emails_by_domain)
mcp.tool()(get_sent_emails_by_domain)
mcp.tool()(get_email_by_id)
mcp.tool()(download_attachment)
mcp.tool()(search_inbox_by_subject)

# Register triage tools
mcp.tool()(triage_email)
mcp.tool()(batch_triage)
mcp.tool()(batch_archive_domain)
mcp.tool()(archive_email)
mcp.tool()(batch_archive_emails)
mcp.tool()(list_subfolders)

# Register domain categorization tools
mcp.tool()(get_uncategorized_domains)
mcp.tool()(categorize_domain)
mcp.tool()(get_domain_summary)

# Register client search tools
mcp.tool()(get_emails_by_client)
mcp.tool()(search_outlook_direct)

# Register commitment scanning tools
mcp.tool()(scan_for_commitments)
mcp.tool()(mark_scanned)
mcp.tool()(batch_mark_scanned)

# Register DMS tools
mcp.tool()(list_dms_clients)
mcp.tool()(list_dms_matters)
mcp.tool()(get_dms_emails)
mcp.tool()(get_dms_admin_emails)
mcp.tool()(search_dms)
mcp.tool()(file_email_to_dms)
mcp.tool()(file_admin_email_to_dms)
mcp.tool()(batch_file_emails_to_dms)

# Register workspace filing tools
mcp.tool()(file_email_to_workspace)
mcp.tool()(file_thread_to_workspace)

# Register thread tracking tools
mcp.tool()(get_email_thread)
mcp.tool()(get_thread_locations)

# Register cache tools
mcp.tool()(read_cache_file)
mcp.tool()(mark_cache_processed)
mcp.tool()(get_cache_status)
mcp.tool()(reset_cache_flags)
mcp.tool()(list_cache_files)

# Register inbox frontmatter tools
mcp.tool()(add_email_frontmatter)


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
