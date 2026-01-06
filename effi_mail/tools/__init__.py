"""Tool functions for effi-mail MCP server.

All tools are importable from this module for registration in main.py.
"""

from effi_mail.tools.email_retrieval import (
    get_pending_emails,
    get_inbox_emails_by_domain,
    get_email_by_id,
    download_attachment,
)
from effi_mail.tools.triage import (
    triage_email,
    batch_triage,
    batch_archive_domain,
)
from effi_mail.tools.domain_categories import (
    get_uncategorized_domains,
    categorize_domain,
    get_domain_summary,
)
from effi_mail.tools.client_search import (
    get_emails_by_client,
    search_outlook_direct,
)
from effi_mail.tools.dms import (
    list_dms_clients,
    list_dms_matters,
    get_dms_emails,
    search_dms,
    file_email_to_dms,
    batch_file_emails_to_dms,
)

__all__ = [
    # Email retrieval
    "get_pending_emails",
    "get_inbox_emails_by_domain",
    "get_email_by_id",
    "download_attachment",
    # Triage
    "triage_email",
    "batch_triage",
    "batch_archive_domain",
    # Domain categories
    "get_uncategorized_domains",
    "categorize_domain",
    "get_domain_summary",
    # Client search
    "get_emails_by_client",
    "search_outlook_direct",
    # DMS
    "list_dms_clients",
    "list_dms_matters",
    "get_dms_emails",
    "search_dms",
    "file_email_to_dms",
    "batch_file_emails_to_dms",
]
