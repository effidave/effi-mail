"""effi-mail MCP server package.

Provides email management tools via FastMCP.
"""

from effi_mail.main import mcp, main, run_server
from effi_mail.testing import list_tools, call_tool

__all__ = ["mcp", "main", "run_server", "list_tools", "call_tool"]
