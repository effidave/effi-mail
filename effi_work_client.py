"""Client for accessing client data via effi-core MCP server.

This module uses the MCP client to call effi-core tools, providing
a clean contract boundary between effi-mail and effi-core.
The MCP tool signatures serve as the stable interface.

Note: Client tools were moved from effi-work to effi-core MCP server.
"""

import os
import json
import asyncio
from typing import Dict, Any, Optional, List
from contextlib import asynccontextmanager

from mcp import ClientSession
from mcp.client.stdio import stdio_client, StdioServerParameters


# Configuration for effi-core server
EFFI_CLIENTS_PATH = os.environ.get("EFFI_CLIENTS_PATH", r"C:\Users\DavidSant\effi-core")
EFFI_CLIENTS_PYTHON = os.environ.get("EFFI_CLIENTS_PYTHON", r"C:\Users\DavidSant\effi-core\.venv\Scripts\python.exe")


@asynccontextmanager
async def get_effi_core_session():
    """Create a session with the effi-core MCP server.
    
    Yields:
        ClientSession connected to effi-core
    """
    server_params = StdioServerParameters(
        command=EFFI_CLIENTS_PYTHON,
        args=["-m", "mcp_server.main"],
        cwd=EFFI_CLIENTS_PATH,
    )
    
    async with stdio_client(server_params) as (read_stream, write_stream):
        async with ClientSession(read_stream, write_stream) as session:
            await session.initialize()
            yield session


async def get_client_identifiers_from_effi_work(client_id: str) -> Dict[str, Any]:
    """Get client identifiers (domains, contact emails) from effi-core via MCP.
    
    Note: Function name kept for backwards compatibility, but now uses effi-core.
    
    Args:
        client_id: Client identifier (case-insensitive)
        
    Returns:
        Dict with:
            - client_id: Resolved client ID (or None if not found)
            - domains: List of email domains
            - contact_emails: List of registered contact emails
            - source: 'effi-work' or error description
    """
    try:
        async with get_effi_core_session() as session:
            # Call effi-core' get_client_by_id tool
            result = await session.call_tool(
                "get_client_by_id",
                {"client_id": client_id}
            )
            
            # Parse the response
            if result.content and len(result.content) > 0:
                data = json.loads(result.content[0].text)
                
                if data.get("error"):
                    return {
                        "client_id": None,
                        "domains": [],
                        "contact_emails": [],
                        "source": "not-found"
                    }
                
                # effi-core returns: {folder, context: {domain or domains, key_contacts, ...}, ...}
                context = data.get("context", {})
                
                # Handle both 'domain' (singular string) and 'domains' (list)
                domains = context.get("domains", [])
                if not domains and context.get("domain"):
                    domains = [context.get("domain")]
                
                # Extract contact emails from key_contacts if available
                contact_emails = context.get("contact_emails", [])
                if not contact_emails:
                    # Try to extract from key_contacts (if they have email field)
                    for contact in context.get("key_contacts", []):
                        if isinstance(contact, dict) and contact.get("email"):
                            contact_emails.append(contact["email"])
                
                return {
                    "client_id": data.get("folder", client_id).lower(),
                    "domains": domains,
                    "contact_emails": contact_emails,
                    "source": "effi-core"
                }
            
            return {
                "client_id": None,
                "domains": [],
                "contact_emails": [],
                "source": "effi-core-empty-response"
            }
            
    except FileNotFoundError:
        return {
            "client_id": None,
            "domains": [],
            "contact_emails": [],
            "source": "effi-core-not-found"
        }
    except Exception as e:
        return {
            "client_id": None,
            "domains": [],
            "contact_emails": [],
            "source": f"effi-core-error: {str(e)}"
        }


async def get_all_clients_from_effi_work() -> List[Dict[str, Any]]:
    """Get all clients from effi-core via MCP.
    
    Note: Function name kept for backwards compatibility, but now uses effi-core.
    
    Returns:
        List of client dicts with client_id, name, domains, contact_emails
    """
    try:
        async with get_effi_core_session() as session:
            result = await session.call_tool("get_all_clients", {})
            
            if result.content and len(result.content) > 0:
                data = json.loads(result.content[0].text)
                return data.get("clients", [])
            
            return []
            
    except Exception:
        return []


async def find_client_by_email_domain(domain: str) -> Optional[Dict[str, Any]]:
    """Find which client an email domain belongs to via effi-core MCP.
    
    Args:
        domain: Email domain to look up
        
    Returns:
        Client dict if found, None otherwise
    """
    try:
        async with get_effi_core_session() as session:
            result = await session.call_tool(
                "find_client_by_email",
                {"domain": domain}
            )
            
            if result.content and len(result.content) > 0:
                data = json.loads(result.content[0].text)
                if data.get("error"):
                    return None
                return data
            
            return None
            
    except Exception:
        return None


# Synchronous wrappers for use in non-async contexts
def get_client_identifiers_sync(client_id: str) -> Dict[str, Any]:
    """Synchronous wrapper for get_client_identifiers_from_effi_work."""
    return asyncio.run(get_client_identifiers_from_effi_work(client_id))


def get_all_clients_sync() -> List[Dict[str, Any]]:
    """Synchronous wrapper for get_all_clients_from_effi_work."""
    return asyncio.run(get_all_clients_from_effi_work())
