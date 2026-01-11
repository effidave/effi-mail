"""Tool functions for effi-mail MCP server.

All tools are importable from this module for registration in main.py.
"""

from effi_mail.tools.email_retrieval import (
    get_pending_emails,
    get_inbox_emails_by_domain,
    get_sent_emails_by_domain,
    get_email_by_id,
    download_attachment,
    search_inbox_by_subject,
)
from effi_mail.tools.triage import (
    triage_email,
    batch_triage,
    batch_archive_domain,
    archive_email,
    batch_archive_emails,
    list_subfolders,
)
from effi_mail.tools.domain_categories import (
    get_uncategorized_domains,
    categorize_domain,
    get_domain_summary,
)
from effi_mail.tools.client_search import (
    get_emails_by_client,
    search_outlook_direct,
    scan_for_commitments,
    mark_scanned,
    batch_mark_scanned,
)
from effi_mail.tools.dms import (
    list_dms_clients,
    list_dms_matters,
    get_dms_emails,
    get_dms_admin_emails,
    search_dms,
    file_email_to_dms,
    file_admin_email_to_dms,
    batch_file_emails_to_dms,
)
from effi_mail.tools.workspace_filing import (
    file_email_to_workspace,
    file_thread_to_workspace,
)
from effi_mail.tools.thread import (
    get_email_thread,
    get_thread_locations,
)
from effi_mail.tools.cache import (
    read_cache_file,
    mark_cache_processed,
    get_cache_status,
    reset_cache_flags,
    list_cache_files,
)

__all__ = [
    # Email retrieval
    "get_pending_emails",
    "get_inbox_emails_by_domain",
    "get_sent_emails_by_domain",
    "get_email_by_id",
    "download_attachment",
    "search_inbox_by_subject",
    # Triage
    "triage_email",
    "batch_triage",
    "batch_archive_domain",
    "archive_email",
    "batch_archive_emails",
    "list_subfolders",
    # Domain categories
    "get_uncategorized_domains",
    "categorize_domain",
    "get_domain_summary",
    # Client search
    "get_emails_by_client",
    "search_outlook_direct",
    # Commitment scanning
    "scan_for_commitments",
    "mark_scanned",
    "batch_mark_scanned",
    # DMS
    "list_dms_clients",
    "list_dms_matters",
    "get_dms_emails",
    "get_dms_admin_emails",
    "search_dms",
    "file_email_to_dms",
    "file_admin_email_to_dms",
    "batch_file_emails_to_dms",
    # Workspace filing
    "file_email_to_workspace",
    "file_thread_to_workspace",
    # Thread tracking
    "get_email_thread",
    "get_thread_locations",
    # Cache operations
    "read_cache_file",
    "mark_cache_processed",
    "get_cache_status",
    "reset_cache_flags",
    "list_cache_files",
]
