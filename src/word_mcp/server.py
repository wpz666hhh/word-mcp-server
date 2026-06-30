"""Word MCP Server — entry point.

Registers all tool modules and starts the stdio MCP server.
"""

import logging

from mcp.server.fastmcp import FastMCP

from .tools.lifecycle import register_lifecycle_tools
from .tools.operations import register_operation_tools
from .tools.vba import register_vba_tool

logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s [word-mcp] %(levelname)s: %(message)s",
)
logger = logging.getLogger("word-mcp")

mcp = FastMCP("word-mcp")

# Register all three tool layers
register_lifecycle_tools(mcp)
register_operation_tools(mcp)
register_vba_tool(mcp)


def main():
    """Entry point for `python -m word_mcp` or `word-mcp` console script."""
    logger.info("Starting Word MCP Server...")
    mcp.run(transport="stdio")
    logger.info("Word MCP Server stopped.")


if __name__ == "__main__":
    main()
