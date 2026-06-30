"""Shared helper functions for Word MCP tools."""

import os
from pathlib import Path


def normalize_path(file_path: str) -> str:
    """Convert a path to absolute, resolving ~ and environment variables.

    Args:
        file_path: A path string, may be relative, with ~, or with %VAR%.

    Returns:
        Absolute path string with forward slashes.
    """
    expanded = os.path.expandvars(os.path.expanduser(file_path))
    absolute = str(Path(expanded).resolve())
    return absolute


def ensure_dir(file_path: str) -> None:
    """Create parent directories for a file path if they don't exist.

    Args:
        file_path: Path to a file whose parent dirs should exist.
    """
    parent = Path(file_path).parent
    parent.mkdir(parents=True, exist_ok=True)


def format_error(operation: str, error: Exception) -> str:
    """Format an exception into a user-friendly error message.

    Args:
        operation: What was being attempted (e.g. "打开文档").
        error: The caught exception.

    Returns:
        Formatted error string.
    """
    return f"{operation}失败: {error}"
