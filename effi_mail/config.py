"""Configuration for effi-mail MCP server."""

import os


def get_transport_config() -> dict:
    """Get transport configuration from environment."""
    return {
        'transport': os.getenv('MCP_TRANSPORT', 'stdio'),
        'host': os.getenv('MCP_HOST', '0.0.0.0'),
        'port': int(os.getenv('MCP_PORT', '8000')),
    }
