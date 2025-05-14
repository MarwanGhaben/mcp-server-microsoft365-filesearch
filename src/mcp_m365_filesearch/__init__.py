# from mcp_m365_filesearch.server import server
import os
import sys

def main():
    """Initialize and run the MCP server."""

    required_keys = ["CLIENT_ID", "CLIENT_SECRET", "TENANT_ID"]

    # Check for required environment variables
    for key in required_keys:
        if key not in os.environ:
            print(
                f"Error: ${key} environment variable is required",
                file=sys.stderr,
            )
            sys.exit(1)

    print("Starting Microsoft 365 Search MCP server...", file=sys.stderr)

    server.run(transport="stdio")

__all__ = ["main", "server"]
