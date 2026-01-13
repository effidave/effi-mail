"""Email ingestion module for effi-mail.

Provides deterministic script to fetch and convert emails to markdown
before triaging them.
"""

from effi_mail.ingestion.ingest import ingest_new_emails

__all__ = [
    "ingest_new_emails",
]
