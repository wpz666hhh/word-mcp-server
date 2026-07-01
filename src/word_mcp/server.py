"""Word MCP Server — entry point.

Registers all tool modules and starts the stdio MCP server.
"""

import logging

from mcp.server.fastmcp import FastMCP

from .tools.lifecycle import register_lifecycle_tools
from .tools.text import (
    word_insert_text, word_format_text, word_format_paragraph,
    word_get_content, word_get_selection, word_select,
    word_apply_style, word_auto_numbering, word_find_replace,
    word_format_text_by_find, word_delete_target, word_replace_target
)
from .tools.tables import word_insert_table, word_format_table
from .tools.layout import (
    word_insert_image, word_insert_page_break,
    word_set_page_setup, word_set_header_footer
)
from .tools.structure import word_get_document_structure, word_read_table
from .tools.macro import word_execute_macro
from .tools.self_optimize import word_record_lesson

logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s [word-mcp] %(levelname)s: %(message)s",
)
logger = logging.getLogger("word-mcp")

mcp = FastMCP("word-mcp")

from functools import wraps
from .com_manager import release_word

def com_cleanup_decorator(func):
    """Decorator to ensure COM references are released after each tool execution."""
    @wraps(func)
    def wrapper(*args, **kwargs):
        try:
            return func(*args, **kwargs)
        finally:
            release_word()
    return wrapper

# Register tool layers (VBA module removed — requires unsafe trust setting)
register_lifecycle_tools(mcp)

# Register L2 atomic operation tools
for tool_func in [
    word_insert_text, word_format_text, word_format_paragraph,
    word_get_content, word_get_selection, word_select,
    word_apply_style, word_auto_numbering, word_find_replace,
    word_format_text_by_find, word_delete_target, word_replace_target,
    word_insert_table, word_format_table,
    word_insert_image, word_insert_page_break,
    word_set_page_setup, word_set_header_footer,
    word_get_document_structure, word_read_table,
    word_execute_macro, word_record_lesson
]:
    mcp.tool()(com_cleanup_decorator(tool_func))


def main():
    """Entry point for `python -m word_mcp` or `word-mcp` console script."""
    logger.info("Starting Word MCP Server...")
    mcp.run(transport="stdio")
    logger.info("Word MCP Server stopped.")


if __name__ == "__main__":
    main()
