"""Outlook client package - split by functionality.

This package provides specialized client classes for different Outlook operations:

- BaseOutlookClient: Connection management and shared utilities
- DMSClient: DMS (DMSforLegal) read/write operations  
- TriageClient: Triage status via Outlook categories
- RetrievalClient: Email fetching, body retrieval, attachments
- SearchClient: DASL query building and flexible search
- FoldersClient: Folder navigation, moving, archiving

Each client manages its own COM connection. For a long-running MCP server,
create singleton instances in helpers.py.
"""

from outlook_client.base import BaseOutlookClient
from outlook_client.dms import DMSClient
from outlook_client.triage import TriageClient
from outlook_client.retrieval import RetrievalClient
from outlook_client.search import SearchClient
from outlook_client.folders import FoldersClient

__all__ = [
    "BaseOutlookClient",
    "DMSClient", 
    "TriageClient",
    "RetrievalClient",
    "SearchClient",
    "FoldersClient",
]
