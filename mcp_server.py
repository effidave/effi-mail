"""MCP Server for Outlook email management - effi-mail."""

import json
import asyncio
from datetime import datetime
from typing import Any, Optional
from mcp.server import Server
from mcp.server.stdio import stdio_server
from mcp.types import (
    Tool,
    TextContent,
    CallToolResult,
)

from outlook_client import OutlookClient
from models import EmailCategory
from effi_work_client import get_client_identifiers_from_effi_work
from domain_categories import (
    get_domain_category,
    set_domain_category,
    get_all_domain_categories,
    get_domains_by_category,
    get_uncategorized_domains,
)


# Initialize components
outlook = OutlookClient()

# Create MCP server
server = Server("effi-mail")


def truncate_text(text: str, max_length: int = 500) -> str:
    """Truncate text to max length with indicator."""
    if len(text) <= max_length:
        return text
    return text[:max_length] + f"... [{len(text) - max_length} more chars]"


def format_email_summary(email, include_preview: bool = False, include_recipients: bool = False) -> dict:
    """Format email for MCP response."""
    result = {
        "id": email.id,
        "subject": email.subject,
        "sender": f"{email.sender_name} <{email.sender_email}>",
        "domain": email.domain,
        "received": email.received_time.isoformat(),
        "has_attachments": email.has_attachments,
        "direction": email.direction,
    }
    # Get triage status from Outlook categories
    triage = outlook.get_triage_status(email.id)
    if triage:
        result["triage_status"] = triage
    if include_preview:
        result["preview"] = truncate_text(email.body_preview, 200)
    if include_recipients:
        result["recipients_to"] = email.recipients_to
        result["recipients_cc"] = email.recipients_cc
    return result


@server.list_tools()
async def list_tools() -> list[Tool]:
    """List available MCP tools."""
    return [
        # Email retrieval tools
        Tool(
            name="get_pending_emails",
            description="Get emails pending triage (no Effi: category), grouped by domain. Queries Outlook directly.",
            inputSchema={
                "type": "object",
                "properties": {
                    "days": {
                        "type": "integer",
                        "description": "Days to look back (default: 30)",
                        "default": 30
                    },
                    "limit": {
                        "type": "integer",
                        "description": "Maximum emails to return (default: 100)",
                        "default": 100
                    },
                    "category_filter": {
                        "type": "string",
                        "description": "Filter by domain category: Client, Internal, Marketing, Personal, Uncategorized",
                        "enum": ["Client", "Internal", "Marketing", "Personal", "Uncategorized"]
                    }
                }
            }
        ),
        Tool(
            name="get_emails_by_domain",
            description="Get emails from a specific sender domain.",
            inputSchema={
                "type": "object",
                "properties": {
                    "domain": {
                        "type": "string",
                        "description": "Domain name to filter by (e.g., 'gmail.com')"
                    },
                    "limit": {
                        "type": "integer",
                        "description": "Maximum emails to return",
                        "default": 20
                    }
                },
                "required": ["domain"]
            }
        ),
        Tool(
            name="get_email_content",
            description="Get full content of a specific email by ID. Use for reading email body.",
            inputSchema={
                "type": "object",
                "properties": {
                    "email_id": {
                        "type": "string",
                        "description": "Email EntryID"
                    },
                    "max_length": {
                        "type": "integer",
                        "description": "Maximum body length (default: 5000)",
                        "default": 5000
                    }
                },
                "required": ["email_id"]
            }
        ),
        
        # Triage tools
        Tool(
            name="triage_email",
            description="Assign triage status to an email using Outlook categories. Status is stored in the email itself.",
            inputSchema={
                "type": "object",
                "properties": {
                    "email_id": {
                        "type": "string",
                        "description": "Email EntryID"
                    },
                    "status": {
                        "type": "string",
                        "description": "Triage status (stored as Effi:Processed, Effi:Deferred, or Effi:Archived category)",
                        "enum": ["processed", "deferred", "archived"]
                    }
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
                    "email_ids": {
                        "type": "array",
                        "items": {"type": "string"},
                        "description": "List of email EntryIDs"
                    },
                    "status": {
                        "type": "string",
                        "description": "Triage status to apply",
                        "enum": ["processed", "deferred", "archived"]
                    }
                },
                "required": ["email_ids", "status"]
            }
        ),
        Tool(
            name="batch_archive_domain",
            description="Archive all pending emails from a specific domain (useful for marketing). Gets pending emails from Outlook and archives them.",
            inputSchema={
                "type": "object",
                "properties": {
                    "domain": {
                        "type": "string",
                        "description": "Domain to archive all emails from"
                    },
                    "days": {
                        "type": "integer",
                        "description": "Days to look back (default: 30)",
                        "default": 30
                    }
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
                    "limit": {
                        "type": "integer",
                        "description": "Maximum domains to return",
                        "default": 20
                    }
                }
            }
        ),
        Tool(
            name="categorize_domain",
            description="Set the category for a sender domain. Saves to domain_categories.json.",
            inputSchema={
                "type": "object",
                "properties": {
                    "domain": {
                        "type": "string",
                        "description": "Domain name"
                    },
                    "category": {
                        "type": "string",
                        "description": "Category to assign",
                        "enum": ["Client", "Internal", "Marketing", "Personal", "Spam"]
                    }
                },
                "required": ["domain", "category"]
            }
        ),
        Tool(
            name="get_domain_summary",
            description="Get summary of all domains grouped by category from domain_categories.json.",
            inputSchema={
                "type": "object",
                "properties": {}
            }
        ),
        
        # Client-centric email search tools
        Tool(
            name="search_emails_by_client",
            description="Search Outlook for all correspondence with a client. Searches by client domains and contact emails from effi-clients.",
            inputSchema={
                "type": "object",
                "properties": {
                    "client_id": {
                        "type": "string",
                        "description": "Client identifier (looked up from effi-clients)"
                    },
                    "days": {
                        "type": "integer",
                        "description": "Days to look back (default: 30)",
                        "default": 30
                    },
                    "date_from": {
                        "type": "string",
                        "description": "Start date (YYYY-MM-DD), overrides days"
                    },
                    "date_to": {
                        "type": "string",
                        "description": "End date (YYYY-MM-DD)"
                    },
                    "limit": {
                        "type": "integer",
                        "description": "Maximum results (default: 100)",
                        "default": 100
                    }
                },
                "required": ["client_id"]
            }
        ),
        Tool(
            name="search_outlook_by_client",
            description="Search Outlook for client correspondence. Gets domains from effi-clients. Alias for search_emails_by_client.",
            inputSchema={
                "type": "object",
                "properties": {
                    "client_id": {
                        "type": "string",
                        "description": "Client identifier (looked up from effi-clients)"
                    },
                    "days": {
                        "type": "integer",
                        "description": "Days to look back (default: 30)",
                        "default": 30
                    },
                    "date_from": {
                        "type": "string",
                        "description": "Start date (YYYY-MM-DD)"
                    },
                    "date_to": {
                        "type": "string",
                        "description": "End date (YYYY-MM-DD)"
                    },
                    "limit": {
                        "type": "integer",
                        "description": "Maximum results (default: 100)",
                        "default": 100
                    }
                },
                "required": ["client_id"]
            }
        ),
        Tool(
            name="search_outlook_direct",
            description="Query Outlook directly with flexible filters. For ad-hoc historical searches.",
            inputSchema={
                "type": "object",
                "properties": {
                    "sender_domain": {
                        "type": "string",
                        "description": "Filter by sender's domain"
                    },
                    "sender_email": {
                        "type": "string",
                        "description": "Filter by exact sender email"
                    },
                    "recipient_domain": {
                        "type": "string",
                        "description": "Filter by recipient domain"
                    },
                    "recipient_email": {
                        "type": "string",
                        "description": "Filter by exact recipient email"
                    },
                    "subject_contains": {
                        "type": "string",
                        "description": "Subject contains text"
                    },
                    "body_contains": {
                        "type": "string",
                        "description": "Body contains text"
                    },
                    "date_from": {
                        "type": "string",
                        "description": "Start date (YYYY-MM-DD)"
                    },
                    "date_to": {
                        "type": "string",
                        "description": "End date (YYYY-MM-DD)"
                    },
                    "days": {
                        "type": "integer",
                        "description": "Days to look back (default: 30)",
                        "default": 30
                    },
                    "folder": {
                        "type": "string",
                        "description": "Outlook folder: 'Inbox' or 'Sent Items'",
                        "default": "Inbox"
                    },
                    "limit": {
                        "type": "integer",
                        "description": "Maximum results (default: 50)",
                        "default": 50
                    }
                }
            }
        ),
        Tool(
            name="get_email_by_id",
            description="Get full email details by ID. Accepts EntryID or internet_message_id (auto-detected).",
            inputSchema={
                "type": "object",
                "properties": {
                    "email_id": {
                        "type": "string",
                        "description": "Email ID - either Outlook EntryID or internet_message_id (format: <...@...>)"
                    },
                    "include_body": {
                        "type": "boolean",
                        "description": "Include full email body",
                        "default": True
                    },
                    "include_attachments": {
                        "type": "boolean",
                        "description": "Include attachment metadata",
                        "default": True
                    }
                },
                "required": ["email_id"]
            }
        ),
    ]


@server.call_tool()
async def call_tool(name: str, arguments: dict[str, Any]) -> CallToolResult:
    """Handle tool calls."""
    
    try:
        if name == "get_pending_emails":
            days = arguments.get("days", 30)
            limit = arguments.get("limit", 100)
            category_filter = arguments.get("category_filter")
            
            # Query Outlook directly for pending emails (no Effi: category)
            result = outlook.get_pending_emails(days=days, limit=limit, group_by_domain=True)
            
            # Filter by domain category if specified
            if category_filter:
                filtered_domains = []
                for domain_data in result.get("domains", []):
                    domain_name = domain_data["domain"]
                    domain_cat = get_domain_category(domain_name)
                    if domain_cat == category_filter:
                        domain_data["category"] = domain_cat
                        filtered_domains.append(domain_data)
                
                # Format emails for response
                for domain_data in filtered_domains:
                    domain_data["emails"] = [
                        format_email_summary(e, include_preview=True) 
                        for e in domain_data["emails"]
                    ]
                
                total = sum(len(d["emails"]) for d in filtered_domains)
                return CallToolResult(
                    content=[TextContent(
                        type="text",
                        text=json.dumps({
                            "total_pending": total,
                            "domains": filtered_domains
                        }, indent=2)
                    )]
                )
            
            # No filter - add category info and format
            for domain_data in result.get("domains", []):
                domain_name = domain_data["domain"]
                domain_data["category"] = get_domain_category(domain_name)
                domain_data["emails"] = [
                    format_email_summary(e, include_preview=True) 
                    for e in domain_data["emails"]
                ]
            
            return CallToolResult(
                content=[TextContent(
                    type="text",
                    text=json.dumps({
                        "total_pending": result["total"],
                        "domains": result.get("domains", [])
                    }, indent=2)
                )]
            )
        
        elif name == "get_emails_by_domain":
            domain = arguments["domain"]
            limit = arguments.get("limit", 20)
            # Search Outlook directly for emails from this domain
            emails = outlook.search_outlook(sender_domain=domain, limit=limit)
            return CallToolResult(
                content=[TextContent(
                    type="text",
                    text=json.dumps({
                        "domain": domain,
                        "category": get_domain_category(domain),
                        "count": len(emails),
                        "emails": [format_email_summary(e, include_preview=True) for e in emails]
                    }, indent=2)
                )]
            )
        
        elif name == "get_email_content":
            email_id = arguments["email_id"]
            max_length = arguments.get("max_length", 5000)
            
            # Get email directly from Outlook
            full_email = outlook.get_email_full(email_id)
            
            if full_email:
                return CallToolResult(
                    content=[TextContent(
                        type="text",
                        text=json.dumps({
                            "subject": full_email.get("subject"),
                            "from": full_email.get("sender"),
                            "received": full_email.get("received"),
                            "attachments": full_email.get("attachments", []),
                            "body": truncate_text(full_email.get("body", ""), max_length)
                        }, indent=2)
                    )]
                )
            return CallToolResult(
                content=[TextContent(type="text", text=json.dumps({"error": "Email not found"}))]
            )
        
        elif name == "triage_email":
            email_id = arguments["email_id"]
            status = arguments["status"]
            
            success = outlook.set_triage_status(email_id, status)
            if success:
                return CallToolResult(
                    content=[TextContent(
                        type="text",
                        text=json.dumps({"success": True, "email_id": email_id, "status": status})
                    )]
                )
            return CallToolResult(
                content=[TextContent(type="text", text=json.dumps({"error": f"Failed to set triage status on {email_id}"}))]
            )
        
        elif name == "batch_triage":
            email_ids = arguments["email_ids"]
            status = arguments["status"]
            
            results = outlook.batch_set_triage_status(email_ids, status)
            
            return CallToolResult(
                content=[TextContent(
                    type="text",
                    text=json.dumps({
                        "success": results["failed"] == 0,
                        "triaged": results["success"],
                        "failed": results["failed"],
                        "status": status
                    })
                )]
            )
        
        elif name == "batch_archive_domain":
            domain = arguments["domain"]
            days = arguments.get("days", 30)
            
            # Get pending emails from this domain
            pending_emails = outlook.get_pending_emails_from_domain(domain, days=days)
            
            # Archive them all
            archived = 0
            for email in pending_emails:
                if outlook.set_triage_status(email.id, "archived"):
                    archived += 1
            
            return CallToolResult(
                content=[TextContent(
                    type="text",
                    text=json.dumps({
                        "success": True,
                        "domain": domain,
                        "archived_count": archived
                    })
                )]
            )
        
        elif name == "get_uncategorized_domains":
            limit = arguments.get("limit", 20)
            # Get pending emails from Outlook to find unique domains
            result = outlook.get_pending_emails(days=30, limit=500, group_by_domain=True)
            
            uncategorized = []
            for domain_data in result.get("domains", []):
                domain_name = domain_data["domain"]
                category = get_domain_category(domain_name)
                if category == "Uncategorized":
                    uncategorized.append({
                        "name": domain_name,
                        "email_count": domain_data["count"],
                        "sample_subjects": [e.subject for e in domain_data["emails"][:3]]
                    })
                    if len(uncategorized) >= limit:
                        break
            
            return CallToolResult(
                content=[TextContent(
                    type="text",
                    text=json.dumps({
                        "count": len(uncategorized),
                        "domains": uncategorized
                    }, indent=2)
                )]
            )
        
        elif name == "categorize_domain":
            domain = arguments["domain"]
            category = arguments["category"]
            set_domain_category(domain, category)
            return CallToolResult(
                content=[TextContent(
                    type="text",
                    text=json.dumps({"success": True, "domain": domain, "category": category})
                )]
            )
        
        elif name == "get_domain_summary":
            all_categories = get_all_domain_categories()
            
            # Group by category
            result = {}
            for domain, category in all_categories.items():
                if category not in result:
                    result[category] = {"count": 0, "domains": []}
                result[category]["count"] += 1
                if len(result[category]["domains"]) < 10:
                    result[category]["domains"].append(domain)
            
            return CallToolResult(
                content=[TextContent(type="text", text=json.dumps(result, indent=2))]
            )
        
        # Client-centric email search tools
        elif name == "search_emails_by_client":
            client_id = arguments["client_id"]
            days = arguments.get("days", 30)
            date_from_str = arguments.get("date_from")
            date_to_str = arguments.get("date_to")
            limit = arguments.get("limit", 100)
            
            # Parse dates
            date_from = datetime.strptime(date_from_str, "%Y-%m-%d") if date_from_str else None
            date_to = datetime.strptime(date_to_str, "%Y-%m-%d") if date_to_str else None
            
            # Get client identifiers from effi-clients (fresh data)
            identifiers = await get_client_identifiers_from_effi_work(client_id)
            if not identifiers.get("domains"):
                return CallToolResult(
                    content=[TextContent(
                        type="text",
                        text=json.dumps({"error": f"Client not found: {client_id}", "source": identifiers.get("source")})
                    )]
                )
            
            # Search Outlook directly
            emails = outlook.search_outlook_by_identifiers(
                domains=identifiers["domains"],
                contact_emails=identifiers.get("contact_emails", []),
                days=days,
                date_from=date_from,
                date_to=date_to,
                limit=limit,
            )
            
            return CallToolResult(
                content=[TextContent(
                    type="text",
                    text=json.dumps({
                        "client_id": client_id,
                        "identifiers": identifiers,
                        "count": len(emails),
                        "emails": [format_email_summary(e, include_preview=True, include_recipients=True) for e in emails]
                    }, indent=2)
                )]
            )
        
        elif name == "search_outlook_by_client":
            # This now does the same as search_emails_by_client (both use Outlook directly)
            client_id = arguments["client_id"]
            days = arguments.get("days", 30)
            date_from_str = arguments.get("date_from")
            date_to_str = arguments.get("date_to")
            limit = arguments.get("limit", 100)
            
            # Parse dates
            date_from = datetime.strptime(date_from_str, "%Y-%m-%d") if date_from_str else None
            date_to = datetime.strptime(date_to_str, "%Y-%m-%d") if date_to_str else None
            
            # Get client identifiers from effi-clients (fresh data)
            identifiers = await get_client_identifiers_from_effi_work(client_id)
            if not identifiers.get("domains"):
                return CallToolResult(
                    content=[TextContent(
                        type="text",
                        text=json.dumps({"error": f"Client not found: {client_id}", "source": identifiers.get("source")})
                    )]
                )
            
            # Search Outlook
            emails = outlook.search_outlook_by_identifiers(
                domains=identifiers["domains"],
                contact_emails=identifiers.get("contact_emails", []),
                days=days,
                date_from=date_from,
                date_to=date_to,
                limit=limit,
            )
            
            return CallToolResult(
                content=[TextContent(
                    type="text",
                    text=json.dumps({
                        "client_id": client_id,
                        "identifiers": identifiers,
                        "count": len(emails),
                        "emails": [format_email_summary(e, include_preview=True, include_recipients=True) for e in emails]
                    }, indent=2)
                )]
            )
        
        elif name == "search_outlook_direct":
            sender_domain = arguments.get("sender_domain")
            sender_email = arguments.get("sender_email")
            recipient_domain = arguments.get("recipient_domain")
            recipient_email = arguments.get("recipient_email")
            subject_contains = arguments.get("subject_contains")
            body_contains = arguments.get("body_contains")
            date_from_str = arguments.get("date_from")
            date_to_str = arguments.get("date_to")
            days = arguments.get("days", 30)
            folder = arguments.get("folder", "Inbox")
            limit = arguments.get("limit", 50)
            
            # Parse dates
            date_from = datetime.strptime(date_from_str, "%Y-%m-%d") if date_from_str else None
            date_to = datetime.strptime(date_to_str, "%Y-%m-%d") if date_to_str else None
            
            emails = outlook.search_outlook(
                sender_domain=sender_domain,
                sender_email=sender_email,
                recipient_domain=recipient_domain,
                recipient_email=recipient_email,
                subject_contains=subject_contains,
                body_contains=body_contains,
                date_from=date_from,
                date_to=date_to,
                days=days,
                folder=folder,
                limit=limit,
            )
            
            return CallToolResult(
                content=[TextContent(
                    type="text",
                    text=json.dumps({
                        "folder": folder,
                        "count": len(emails),
                        "emails": [format_email_summary(e, include_preview=True, include_recipients=True) for e in emails]
                    }, indent=2)
                )]
            )
        
        elif name == "get_email_by_id":
            email_id = arguments["email_id"]
            include_body = arguments.get("include_body", True)
            include_attachments = arguments.get("include_attachments", True)
            
            # Get email directly from Outlook
            full_email = outlook.get_email_full(email_id)
            
            if full_email:
                result = full_email
                if not include_body:
                    result.pop("body", None)
                    result.pop("html_body", None)
                if not include_attachments:
                    result.pop("attachments", None)
                return CallToolResult(
                    content=[TextContent(type="text", text=json.dumps(result, indent=2))]
                )
            else:
                return CallToolResult(
                    content=[TextContent(
                        type="text",
                        text=json.dumps({"error": f"Email not found: {email_id}"})
                    )]
                )
        
        else:
            return CallToolResult(
                content=[TextContent(type="text", text=json.dumps({"error": f"Unknown tool: {name}"}))]
            )
    
    except Exception as e:
        return CallToolResult(
            content=[TextContent(type="text", text=json.dumps({"error": str(e)}))]
        )


async def main():
    """Run the MCP server."""
    async with stdio_server() as (read_stream, write_stream):
        await server.run(read_stream, write_stream, server.create_initialization_options())


if __name__ == "__main__":
    asyncio.run(main())
