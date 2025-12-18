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

from database import Database
from outlook_client import OutlookClient
from models import EmailCategory, TriageStatus


# Initialize components
db = Database()
outlook = OutlookClient(db)

# Create MCP server
server = Server("effi-mail")


def truncate_text(text: str, max_length: int = 500) -> str:
    """Truncate text to max length with indicator."""
    if len(text) <= max_length:
        return text
    return text[:max_length] + f"... [{len(text) - max_length} more chars]"


def format_email_summary(email, include_preview: bool = False) -> dict:
    """Format email for MCP response."""
    result = {
        "id": email.id,
        "subject": email.subject,
        "sender": f"{email.sender_name} <{email.sender_email}>",
        "domain": email.domain,
        "received": email.received_time.isoformat(),
        "triage_status": email.triage_status.value,
        "has_attachments": email.has_attachments,
    }
    if include_preview:
        result["preview"] = truncate_text(email.body_preview, 200)
    if email.client_id:
        result["client_id"] = email.client_id
    if email.matter_id:
        result["matter_id"] = email.matter_id
    return result


@server.list_tools()
async def list_tools() -> list[Tool]:
    """List available MCP tools."""
    return [
        # Email retrieval tools
        Tool(
            name="sync_emails",
            description="Sync emails from Outlook to local database. Call this first to populate data.",
            inputSchema={
                "type": "object",
                "properties": {
                    "days": {
                        "type": "integer",
                        "description": "Number of days to sync (default: 7)",
                        "default": 7
                    },
                    "exclude_unfocused": {
                        "type": "boolean",
                        "description": "Exclude emails with 'Unfocused' category",
                        "default": True
                    }
                }
            }
        ),
        Tool(
            name="get_pending_emails",
            description="Get emails pending triage, grouped by domain category.",
            inputSchema={
                "type": "object",
                "properties": {
                    "limit": {
                        "type": "integer",
                        "description": "Maximum emails to return (default: 50)",
                        "default": 50
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
            description="Assign triage status to an email. Mark as processed, deferred, or archived.",
            inputSchema={
                "type": "object",
                "properties": {
                    "email_id": {
                        "type": "string",
                        "description": "Email EntryID"
                    },
                    "status": {
                        "type": "string",
                        "description": "Triage status",
                        "enum": ["processed", "deferred", "archived"]
                    },
                    "client_id": {
                        "type": "string",
                        "description": "Optional client ID to associate"
                    },
                    "matter_id": {
                        "type": "string",
                        "description": "Optional matter ID to associate"
                    },
                    "notes": {
                        "type": "string",
                        "description": "Optional notes about the email"
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
                    },
                    "client_id": {
                        "type": "string",
                        "description": "Optional client ID for all emails"
                    }
                },
                "required": ["email_ids", "status"]
            }
        ),
        Tool(
            name="batch_archive_domain",
            description="Archive all pending emails from a specific domain (useful for marketing).",
            inputSchema={
                "type": "object",
                "properties": {
                    "domain": {
                        "type": "string",
                        "description": "Domain to archive all emails from"
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
            description="Set the category for a sender domain.",
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
                        "enum": ["Client", "Internal", "Marketing", "Personal"]
                    }
                },
                "required": ["domain", "category"]
            }
        ),
        Tool(
            name="get_domain_summary",
            description="Get summary of all domains grouped by category.",
            inputSchema={
                "type": "object",
                "properties": {}
            }
        ),
        
        # Client/Matter tools
        Tool(
            name="create_client",
            description="Create a new client record.",
            inputSchema={
                "type": "object",
                "properties": {
                    "id": {
                        "type": "string",
                        "description": "Client ID (e.g., 'acme-corp')"
                    },
                    "name": {
                        "type": "string",
                        "description": "Client display name"
                    },
                    "domains": {
                        "type": "array",
                        "items": {"type": "string"},
                        "description": "Email domains associated with this client"
                    },
                    "folder_path": {
                        "type": "string",
                        "description": "Path in effi-work repository"
                    }
                },
                "required": ["id", "name"]
            }
        ),
        Tool(
            name="create_matter",
            description="Create a new matter for a client.",
            inputSchema={
                "type": "object",
                "properties": {
                    "id": {
                        "type": "string",
                        "description": "Matter ID"
                    },
                    "client_id": {
                        "type": "string",
                        "description": "Client ID this matter belongs to"
                    },
                    "name": {
                        "type": "string",
                        "description": "Matter name"
                    },
                    "description": {
                        "type": "string",
                        "description": "Matter description"
                    },
                    "folder_path": {
                        "type": "string",
                        "description": "Path in effi-work repository"
                    }
                },
                "required": ["id", "client_id", "name"]
            }
        ),
        Tool(
            name="list_clients",
            description="List all clients and their matters.",
            inputSchema={
                "type": "object",
                "properties": {}
            }
        ),
        
        # Statistics
        Tool(
            name="get_triage_stats",
            description="Get statistics about email triage status.",
            inputSchema={
                "type": "object",
                "properties": {}
            }
        ),
    ]


@server.call_tool()
async def call_tool(name: str, arguments: dict[str, Any]) -> CallToolResult:
    """Handle tool calls."""
    
    try:
        if name == "sync_emails":
            days = arguments.get("days", 7)
            exclude = ["Unfocused"] if arguments.get("exclude_unfocused", True) else []
            stats = outlook.sync_emails_to_db(days=days, exclude_categories=exclude)
            return CallToolResult(
                content=[TextContent(
                    type="text",
                    text=json.dumps({
                        "success": True,
                        "message": f"Synced emails from last {days} days",
                        "stats": stats
                    }, indent=2)
                )]
            )
        
        elif name == "get_pending_emails":
            limit = arguments.get("limit", 50)
            category_filter = arguments.get("category_filter")
            
            emails = db.get_emails_by_status(TriageStatus.PENDING, limit=limit)
            
            # Filter by domain category if specified
            if category_filter:
                target_category = EmailCategory(category_filter)
                filtered_emails = []
                for email in emails:
                    domain = db.get_domain(email.domain)
                    if domain and domain.category == target_category:
                        filtered_emails.append(email)
                emails = filtered_emails
            
            # Group by domain
            by_domain = {}
            for email in emails:
                if email.domain not in by_domain:
                    domain = db.get_domain(email.domain)
                    by_domain[email.domain] = {
                        "category": domain.category.value if domain else "Uncategorized",
                        "emails": []
                    }
                by_domain[email.domain]["emails"].append(format_email_summary(email, include_preview=True))
            
            return CallToolResult(
                content=[TextContent(
                    type="text",
                    text=json.dumps({
                        "total_pending": len(emails),
                        "domains": by_domain
                    }, indent=2)
                )]
            )
        
        elif name == "get_emails_by_domain":
            domain = arguments["domain"]
            limit = arguments.get("limit", 20)
            emails = db.get_emails_by_domain(domain, limit=limit)
            return CallToolResult(
                content=[TextContent(
                    type="text",
                    text=json.dumps({
                        "domain": domain,
                        "count": len(emails),
                        "emails": [format_email_summary(e, include_preview=True) for e in emails]
                    }, indent=2)
                )]
            )
        
        elif name == "get_email_content":
            email_id = arguments["email_id"]
            max_length = arguments.get("max_length", 5000)
            
            email = db.get_email(email_id)
            body = outlook.get_email_body(email_id, max_length=max_length)
            
            if email:
                return CallToolResult(
                    content=[TextContent(
                        type="text",
                        text=json.dumps({
                            "subject": email.subject,
                            "from": f"{email.sender_name} <{email.sender_email}>",
                            "received": email.received_time.isoformat(),
                            "attachments": email.attachment_names,
                            "body": body
                        }, indent=2)
                    )]
                )
            return CallToolResult(
                content=[TextContent(type="text", text=json.dumps({"error": "Email not found"}))]
            )
        
        elif name == "triage_email":
            email_id = arguments["email_id"]
            status = TriageStatus(arguments["status"])
            client_id = arguments.get("client_id")
            matter_id = arguments.get("matter_id")
            notes = arguments.get("notes", "")
            
            db.update_email_triage(email_id, status, client_id, matter_id, notes)
            return CallToolResult(
                content=[TextContent(
                    type="text",
                    text=json.dumps({"success": True, "email_id": email_id, "status": status.value})
                )]
            )
        
        elif name == "batch_triage":
            email_ids = arguments["email_ids"]
            status = TriageStatus(arguments["status"])
            client_id = arguments.get("client_id")
            
            for email_id in email_ids:
                db.update_email_triage(email_id, status, client_id)
            
            return CallToolResult(
                content=[TextContent(
                    type="text",
                    text=json.dumps({
                        "success": True,
                        "count": len(email_ids),
                        "status": status.value
                    })
                )]
            )
        
        elif name == "batch_archive_domain":
            domain = arguments["domain"]
            emails = db.get_emails_by_domain(domain)
            pending = [e for e in emails if e.triage_status == TriageStatus.PENDING]
            
            for email in pending:
                db.update_email_triage(email.id, TriageStatus.ARCHIVED)
            
            return CallToolResult(
                content=[TextContent(
                    type="text",
                    text=json.dumps({
                        "success": True,
                        "domain": domain,
                        "archived_count": len(pending)
                    })
                )]
            )
        
        elif name == "get_uncategorized_domains":
            limit = arguments.get("limit", 20)
            domains = db.get_uncategorized_domains()[:limit]
            return CallToolResult(
                content=[TextContent(
                    type="text",
                    text=json.dumps({
                        "count": len(domains),
                        "domains": [
                            {
                                "name": d.name,
                                "email_count": d.email_count,
                                "sample_senders": d.sample_senders
                            }
                            for d in domains
                        ]
                    }, indent=2)
                )]
            )
        
        elif name == "categorize_domain":
            domain = arguments["domain"]
            category = EmailCategory(arguments["category"])
            db.update_domain_category(domain, category)
            return CallToolResult(
                content=[TextContent(
                    type="text",
                    text=json.dumps({"success": True, "domain": domain, "category": category.value})
                )]
            )
        
        elif name == "get_domain_summary":
            result = {}
            for category in EmailCategory:
                domains = db.get_domains_by_category(category)
                result[category.value] = {
                    "count": len(domains),
                    "domains": [{"name": d.name, "emails": d.email_count} for d in domains[:10]]
                }
            return CallToolResult(
                content=[TextContent(type="text", text=json.dumps(result, indent=2))]
            )
        
        elif name == "create_client":
            from models import Client
            client = Client(
                id=arguments["id"],
                name=arguments["name"],
                domains=arguments.get("domains", []),
                folder_path=arguments.get("folder_path")
            )
            db.upsert_client(client)
            return CallToolResult(
                content=[TextContent(
                    type="text",
                    text=json.dumps({"success": True, "client_id": client.id, "name": client.name})
                )]
            )
        
        elif name == "create_matter":
            from models import Matter
            matter = Matter(
                id=arguments["id"],
                client_id=arguments["client_id"],
                name=arguments["name"],
                description=arguments.get("description", ""),
                folder_path=arguments.get("folder_path")
            )
            db.upsert_matter(matter)
            return CallToolResult(
                content=[TextContent(
                    type="text",
                    text=json.dumps({"success": True, "matter_id": matter.id, "client_id": matter.client_id})
                )]
            )
        
        elif name == "list_clients":
            clients = db.get_all_clients()
            result = []
            for client in clients:
                matters = db.get_matters_for_client(client.id)
                result.append({
                    "id": client.id,
                    "name": client.name,
                    "domains": client.domains,
                    "folder_path": client.folder_path,
                    "matters": [{"id": m.id, "name": m.name, "active": m.active} for m in matters]
                })
            return CallToolResult(
                content=[TextContent(type="text", text=json.dumps({"clients": result}, indent=2))]
            )
        
        elif name == "get_triage_stats":
            triage_stats = db.get_triage_stats()
            domain_stats = db.get_domain_stats()
            return CallToolResult(
                content=[TextContent(
                    type="text",
                    text=json.dumps({
                        "by_triage_status": triage_stats,
                        "by_domain_category": domain_stats
                    }, indent=2)
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
