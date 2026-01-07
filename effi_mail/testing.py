"""Test utilities for effi-mail.

Provides list_tools() and call_tool() functions for testing
that simulate the MCP server behavior without running the full server.
"""

import json
from typing import Any, List
from mcp.types import Tool, TextContent

from effi_mail.tools import (
    get_pending_emails as _get_pending_emails,
    get_inbox_emails_by_domain as _get_inbox_emails_by_domain,
    get_email_by_id as _get_email_by_id,
    triage_email as _triage_email,
    batch_triage as _batch_triage,
    batch_archive_domain as _batch_archive_domain,
    get_uncategorized_domains as _get_uncategorized_domains,
    categorize_domain as _categorize_domain,
    get_domain_summary as _get_domain_summary,
    get_emails_by_client as _get_emails_by_client,
    search_outlook_direct as _search_outlook_direct,
    scan_for_commitments as _scan_for_commitments,
    mark_scanned as _mark_scanned,
    batch_mark_scanned as _batch_mark_scanned,
    list_dms_clients as _list_dms_clients,
    list_dms_matters as _list_dms_matters,
    get_dms_emails as _get_dms_emails,
    search_dms as _search_dms,
    file_email_to_dms as _file_email_to_dms,
    batch_file_emails_to_dms as _batch_file_emails_to_dms,
)


async def list_tools() -> List[Tool]:
    """List available MCP tools.
    
    For testing - returns tool definitions without running the MCP server.
    """
    return [
        # Email retrieval tools
        Tool(
            name="get_pending_emails",
            description="Get emails pending triage (no effi: category), grouped by domain.",
            inputSchema={
                "type": "object",
                "properties": {
                    "days": {"type": "integer", "default": 30},
                    "limit": {"type": "integer", "default": 100},
                    "category_filter": {
                        "type": "string",
                        "enum": ["Client", "Internal", "Marketing", "Personal", "Uncategorized"]
                    }
                }
            }
        ),
        Tool(
            name="get_inbox_emails_by_domain",
            description="Get emails from a specific sender domain (Inbox only).",
            inputSchema={
                "type": "object",
                "properties": {
                    "domain": {"type": "string"},
                    "limit": {"type": "integer", "default": 20}
                },
                "required": ["domain"]
            }
        ),
        Tool(
            name="get_email_by_id",
            description="Get full email details by ID.",
            inputSchema={
                "type": "object",
                "properties": {
                    "email_id": {"type": "string"},
                    "include_body": {"type": "boolean", "default": True},
                    "include_attachments": {"type": "boolean", "default": True}
                },
                "required": ["email_id"]
            }
        ),
        
        # Triage tools
        Tool(
            name="triage_email",
            description="Assign triage status to an email using Outlook categories.",
            inputSchema={
                "type": "object",
                "properties": {
                    "email_id": {"type": "string"},
                    "status": {"type": "string", "enum": ["action", "waiting", "processed", "archived"]}
                },
                "required": ["email_id", "status"]
            }
        ),
        Tool(
            name="batch_triage",
            description="Triage multiple emails at once with the same status.",
            inputSchema={
                "type": "object",
                "properties": {
                    "email_ids": {"type": "array", "items": {"type": "string"}},
                    "status": {"type": "string", "enum": ["action", "waiting", "processed", "archived"]}
                },
                "required": ["email_ids", "status"]
            }
        ),
        Tool(
            name="batch_archive_domain",
            description="Archive all pending emails from a specific domain.",
            inputSchema={
                "type": "object",
                "properties": {
                    "domain": {"type": "string"},
                    "days": {"type": "integer", "default": 30}
                },
                "required": ["domain"]
            }
        ),
        
        # Domain categorization tools
        Tool(
            name="get_uncategorized_domains",
            description="Get list of domains that haven't been categorized yet.",
            inputSchema={
                "type": "object",
                "properties": {
                    "days": {"type": "integer", "default": 30},
                    "limit": {"type": "integer", "default": 20}
                }
            }
        ),
        Tool(
            name="categorize_domain",
            description="Set the category for a sender domain.",
            inputSchema={
                "type": "object",
                "properties": {
                    "domain": {"type": "string"},
                    "category": {"type": "string", "enum": ["Client", "Internal", "Marketing", "Personal", "Spam"]}
                },
                "required": ["domain", "category"]
            }
        ),
        Tool(
            name="get_domain_summary",
            description="Get summary of all domains grouped by category.",
            inputSchema={"type": "object", "properties": {}}
        ),
        
        # Client search tools
        Tool(
            name="get_emails_by_client",
            description="Get all email correspondence with a client from Outlook.",
            inputSchema={
                "type": "object",
                "properties": {
                    "client_id": {"type": "string"},
                    "days": {"type": "integer", "default": 30},
                    "date_from": {"type": "string"},
                    "date_to": {"type": "string"},
                    "limit": {"type": "integer", "default": 100}
                },
                "required": ["client_id"]
            }
        ),
        Tool(
            name="search_outlook_direct",
            description="Query Outlook directly with flexible filters.",
            inputSchema={
                "type": "object",
                "properties": {
                    "sender_domain": {"type": "string"},
                    "sender_email": {"type": "string"},
                    "recipient_domain": {"type": "string"},
                    "recipient_email": {"type": "string"},
                    "subject_contains": {"type": "string"},
                    "body_contains": {"type": "string"},
                    "date_from": {"type": "string"},
                    "date_to": {"type": "string"},
                    "days": {"type": "integer", "default": 30},
                    "folder": {"type": "string", "default": "Inbox"},
                    "limit": {"type": "integer", "default": 50}
                }
            }
        ),
        
        # Commitment scanning tools
        Tool(
            name="scan_for_commitments",
            description="Scan sent emails for commitment detection. Returns unscanned emails with full body.",
            inputSchema={
                "type": "object",
                "properties": {
                    "days": {"type": "integer", "default": 14},
                    "limit": {"type": "integer", "default": 100}
                }
            }
        ),
        Tool(
            name="mark_scanned",
            description="Mark an email as scanned for commitments.",
            inputSchema={
                "type": "object",
                "properties": {
                    "email_id": {"type": "string"}
                },
                "required": ["email_id"]
            }
        ),
        Tool(
            name="batch_mark_scanned",
            description="Mark multiple emails as scanned for commitments.",
            inputSchema={
                "type": "object",
                "properties": {
                    "email_ids": {"type": "array", "items": {"type": "string"}}
                },
                "required": ["email_ids"]
            }
        ),
        
        # DMS tools
        Tool(
            name="list_dms_clients",
            description="List all client folders in DMSforLegal.",
            inputSchema={"type": "object", "properties": {}}
        ),
        Tool(
            name="list_dms_matters",
            description="List all matter folders for a client in DMSforLegal.",
            inputSchema={
                "type": "object",
                "properties": {"client": {"type": "string"}},
                "required": ["client"]
            }
        ),
        Tool(
            name="get_dms_emails",
            description="Get emails filed under a specific client/matter in DMSforLegal.",
            inputSchema={
                "type": "object",
                "properties": {
                    "client": {"type": "string"},
                    "matter": {"type": "string"},
                    "limit": {"type": "integer", "default": 50}
                },
                "required": ["client", "matter"]
            }
        ),
        Tool(
            name="search_dms",
            description="Search emails across DMSforLegal with filters.",
            inputSchema={
                "type": "object",
                "properties": {
                    "client": {"type": "string"},
                    "matter": {"type": "string"},
                    "subject_contains": {"type": "string"},
                    "date_from": {"type": "string"},
                    "date_to": {"type": "string"},
                    "limit": {"type": "integer", "default": 50}
                }
            }
        ),
        Tool(
            name="file_email_to_dms",
            description="File an email to a DMS client/matter folder.",
            inputSchema={
                "type": "object",
                "properties": {
                    "email_id": {"type": "string"},
                    "client": {"type": "string"},
                    "matter": {"type": "string"}
                },
                "required": ["email_id", "client", "matter"]
            }
        ),
        Tool(
            name="batch_file_emails_to_dms",
            description="File multiple emails to a DMS client/matter folder.",
            inputSchema={
                "type": "object",
                "properties": {
                    "email_ids": {"type": "array", "items": {"type": "string"}},
                    "client": {"type": "string"},
                    "matter": {"type": "string"}
                },
                "required": ["email_ids", "client", "matter"]
            }
        ),
    ]


async def call_tool(name: str, arguments: dict[str, Any]) -> list[TextContent]:
    """Handle tool calls for testing.
    
    Dispatches to the actual tool implementations.
    """
    try:
        # Email retrieval tools
        if name == "get_pending_emails":
            result = _get_pending_emails(
                days=arguments.get("days", 30),
                limit=arguments.get("limit", 100),
                category_filter=arguments.get("category_filter")
            )
            return [TextContent(type="text", text=result)]
        
        elif name == "get_inbox_emails_by_domain":
            result = _get_inbox_emails_by_domain(
                domain=arguments["domain"],
                limit=arguments.get("limit", 20)
            )
            return [TextContent(type="text", text=result)]
        
        elif name == "get_email_by_id":
            result = _get_email_by_id(
                email_id=arguments["email_id"],
                include_body=arguments.get("include_body", True),
                include_attachments=arguments.get("include_attachments", True),
                max_body_length=arguments.get("max_body_length")
            )
            return [TextContent(type="text", text=result)]
        
        # Triage tools
        elif name == "triage_email":
            result = _triage_email(
                email_id=arguments["email_id"],
                status=arguments["status"]
            )
            return [TextContent(type="text", text=result)]
        
        elif name == "batch_triage":
            result = _batch_triage(
                email_ids=arguments["email_ids"],
                status=arguments["status"]
            )
            return [TextContent(type="text", text=result)]
        
        elif name == "batch_archive_domain":
            result = _batch_archive_domain(
                domain=arguments["domain"],
                days=arguments.get("days", 30)
            )
            return [TextContent(type="text", text=result)]
        
        # Domain categorization tools
        elif name == "get_uncategorized_domains":
            result = _get_uncategorized_domains(
                days=arguments.get("days", 30),
                limit=arguments.get("limit", 20)
            )
            return [TextContent(type="text", text=result)]
        
        elif name == "categorize_domain":
            result = _categorize_domain(
                domain=arguments["domain"],
                category=arguments["category"]
            )
            return [TextContent(type="text", text=result)]
        
        elif name == "get_domain_summary":
            result = _get_domain_summary()
            return [TextContent(type="text", text=result)]
        
        # Client search tools (async)
        elif name == "get_emails_by_client":
            result = await _get_emails_by_client(
                client_id=arguments["client_id"],
                days=arguments.get("days", 30),
                date_from=arguments.get("date_from"),
                date_to=arguments.get("date_to"),
                limit=arguments.get("limit", 100)
            )
            return [TextContent(type="text", text=result)]
        
        elif name == "search_outlook_direct":
            result = _search_outlook_direct(
                sender_domain=arguments.get("sender_domain"),
                sender_email=arguments.get("sender_email"),
                recipient_domain=arguments.get("recipient_domain"),
                recipient_email=arguments.get("recipient_email"),
                subject_contains=arguments.get("subject_contains"),
                body_contains=arguments.get("body_contains"),
                date_from=arguments.get("date_from"),
                date_to=arguments.get("date_to"),
                days=arguments.get("days", 30),
                folder=arguments.get("folder", "Inbox"),
                limit=arguments.get("limit", 50)
            )
            return [TextContent(type="text", text=result)]
        
        # Commitment scanning tools
        elif name == "scan_for_commitments":
            result = _scan_for_commitments(
                days=arguments.get("days", 14),
                limit=arguments.get("limit", 100)
            )
            return [TextContent(type="text", text=result)]
        
        elif name == "mark_scanned":
            result = _mark_scanned(
                email_id=arguments["email_id"]
            )
            return [TextContent(type="text", text=result)]
        
        elif name == "batch_mark_scanned":
            result = _batch_mark_scanned(
                email_ids=arguments["email_ids"]
            )
            return [TextContent(type="text", text=result)]
        
        # DMS tools
        elif name == "list_dms_clients":
            result = _list_dms_clients()
            return [TextContent(type="text", text=result)]
        
        elif name == "list_dms_matters":
            result = _list_dms_matters(
                client=arguments.get("client")
            )
            return [TextContent(type="text", text=result)]
        
        elif name == "get_dms_emails":
            result = _get_dms_emails(
                client=arguments.get("client"),
                matter=arguments.get("matter"),
                limit=arguments.get("limit", 50)
            )
            return [TextContent(type="text", text=result)]
        
        elif name == "search_dms":
            result = _search_dms(
                client=arguments.get("client"),
                matter=arguments.get("matter"),
                subject_contains=arguments.get("subject_contains"),
                date_from=arguments.get("date_from"),
                date_to=arguments.get("date_to"),
                limit=arguments.get("limit", 50)
            )
            return [TextContent(type="text", text=result)]
        
        elif name == "file_email_to_dms":
            result = _file_email_to_dms(
                email_id=arguments.get("email_id"),
                client=arguments.get("client"),
                matter=arguments.get("matter")
            )
            return [TextContent(type="text", text=result)]
        
        elif name == "batch_file_emails_to_dms":
            result = _batch_file_emails_to_dms(
                email_ids=arguments.get("email_ids", []),
                client=arguments.get("client"),
                matter=arguments.get("matter")
            )
            return [TextContent(type="text", text=result)]
        
        else:
            return [TextContent(type="text", text=json.dumps({"error": f"Unknown tool: {name}"}))]
    
    except Exception as e:
        return [TextContent(type="text", text=json.dumps({"error": str(e)}))]
