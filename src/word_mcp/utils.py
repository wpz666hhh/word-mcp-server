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


import json
from pathlib import Path

# Load error hints dynamically
_HINTS_FILE = Path(__file__).parent / "com_errors.json"
_COM_ERROR_HINTS: dict[str, str] = {}
if _HINTS_FILE.exists():
    try:
        with open(_HINTS_FILE, "r", encoding="utf-8") as f:
            _COM_ERROR_HINTS = json.load(f)
    except Exception:
        pass


def format_error(operation: str, error: Exception) -> str:
    """Format an exception into a user-friendly error message, preserving original details for debugging.

    Args:
        operation: What was being attempted (e.g. "打开文档").
        error: The caught exception.

    Returns:
        Formatted error string.
    """
    err_str = str(error)
    details = f" (原始错误: {err_str})"

    # Try to match known COM error codes
    for code, hint in _COM_ERROR_HINTS.items():
        if code in err_str:
            return f"{operation}失败: {hint}{details}"

    # Suppress low-level COM noise in user-facing messages but keep original details
    for noise in ["Command failed", "(-2147352567, '发生意外。'"]:
        if noise in err_str:
            # Extract just the human-readable part after the COM noise
            parts = err_str.split(", ")
            for p in parts:
                if "wdmain" not in p and "0x" not in p and len(p) > 10 and ("，" in p or "的" in p):
                    clean_p = p.strip("'\"")
                    return f"{operation}失败: {clean_p}{details}"
            return f"{operation}失败: 请检查 Word 窗口状态后重试{details}"

    return f"{operation}失败: {err_str}"
