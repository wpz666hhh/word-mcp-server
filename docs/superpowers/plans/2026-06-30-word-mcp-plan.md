# Word MCP Server Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** Build an MCP server that lets Claude Agent control Microsoft Word via COM automation, with real-time visible window.

**Architecture:** Three-layer hybrid — L1 document lifecycle (6 tools), L2 high-frequency atomic operations (15 tools), L3 VBA executor (1 tool). Python + pywin32 + FastMCP, stdio transport, lazy COM initialization with auto-reconnect.

**Tech Stack:** Python ≥3.10, pywin32 ≥306, mcp ≥1.0.0, Windows-only (COM dependency)

---

## File Map

| File | Responsibility |
|------|---------------|
| `pyproject.toml` | Project metadata, dependencies, entry point |
| `src/word_mcp/__init__.py` | Package init, version |
| `src/word_mcp/server.py` | FastMCP server entry, registers all tool modules |
| `src/word_mcp/com_manager.py` | COM lifecycle: Dispatch/GetObject, reconnect, visibility |
| `src/word_mcp/utils.py` | Shared helpers: path normalization, error formatting |
| `src/word_mcp/tools/__init__.py` | Tools package init |
| `src/word_mcp/tools/lifecycle.py` | L1: 6 document lifecycle tools |
| `src/word_mcp/tools/operations.py` | L2: 15 atomic operation tools |
| `src/word_mcp/tools/vba.py` | L3: VBA executor tool |
| `tests/__init__.py` | Test package init |
| `tests/test_com_manager.py` | COM manager unit tests |
| `tests/test_lifecycle.py` | L1 tool tests |
| `tests/test_operations.py` | L2 tool tests |
| `tests/test_vba.py` | L3 tool test |
| `README.md` | Usage and installation docs |

---

### Task 1: Project Scaffold

**Files:**
- Create: `pyproject.toml`
- Create: `src/word_mcp/__init__.py`
- Create: `src/word_mcp/tools/__init__.py`
- Create: `tests/__init__.py`

- [ ] **Step 1: Write `pyproject.toml`**

```toml
[build-system]
requires = ["setuptools>=68.0", "wheel"]
build-backend = "setuptools.backends._legacy:_Backend"

[project]
name = "word-mcp-server"
version = "0.1.0"
description = "MCP server for Microsoft Word COM automation"
requires-python = ">=3.10"
dependencies = [
    "pywin32>=306",
    "mcp>=1.0.0",
]

[project.scripts]
word-mcp = "word_mcp.server:main"

[tool.setuptools.packages.find]
where = ["src"]
```

- [ ] **Step 2: Write `src/word_mcp/__init__.py`**

```python
"""Word MCP Server — Microsoft Word COM automation for Claude Agent."""

__version__ = "0.1.0"
```

- [ ] **Step 3: Write `src/word_mcp/tools/__init__.py`**

```python
"""MCP tool modules for Word manipulation."""
```

- [ ] **Step 4: Write `tests/__init__.py`**

```python
"""Tests for word-mcp-server."""
```

- [ ] **Step 5: Install in dev mode and verify**

```bash
cd C:/Users/19773/Desktop/word-mcp-server
pip install -e .
```

Expected: install succeeds, no errors.

- [ ] **Step 6: Commit**

```bash
git add pyproject.toml src/word_mcp/__init__.py src/word_mcp/tools/__init__.py tests/__init__.py
git commit -m "chore: scaffold project structure with pyproject.toml"
```

---

### Task 2: COM Connection Manager

**Files:**
- Create: `src/word_mcp/com_manager.py`
- Create: `tests/test_com_manager.py`

- [ ] **Step 1: Write failing test for `get_word_app`**

Create `tests/test_com_manager.py`:

```python
"""Tests for COM connection manager."""
import sys
import pytest


class TestGetWordApp:
    """Tests for get_word_app()."""

    def test_creates_word_instance(self):
        """get_word_app() should return a Word Application object."""
        # Skip if not on Windows
        if sys.platform != "win32":
            pytest.skip("COM tests only run on Windows")

        from word_mcp.com_manager import get_word_app

        app = get_word_app(visible=False)
        assert app is not None
        # Word.Application.Name returns "Microsoft Word"
        assert "Word" in app.Name

    def test_returns_cached_instance(self):
        """Second call should return same instance without re-creating."""
        if sys.platform != "win32":
            pytest.skip("COM tests only run on Windows")

        from word_mcp.com_manager import get_word_app

        app1 = get_word_app(visible=False)
        app2 = get_word_app()
        assert app1 is app2

    def test_visible_flag(self):
        """visible=False should hide Word, visible=True should show it."""
        if sys.platform != "win32":
            pytest.skip("COM tests only run on Windows")

        from word_mcp.com_manager import get_word_app

        app = get_word_app(visible=True)
        assert app.Visible is True
```

- [ ] **Step 2: Run tests, verify they fail**

```bash
python -m pytest tests/test_com_manager.py -v
```

Expected: FAIL (import error — `com_manager` module doesn't exist).

- [ ] **Step 3: Write `src/word_mcp/com_manager.py`**

```python
"""COM connection lifecycle management for Microsoft Word.

Maintains a singleton Word.Application COM object. On first call it either
attaches to an existing Word process (GetObject) or creates a new one
(Dispatch). Subsequent calls return the cached instance. If the connection
is lost (user closed Word), the next call transparently reconnects.
"""

import logging
from typing import Optional

import pythoncom
import win32com.client

logger = logging.getLogger(__name__)

_word_app: Optional[object] = None
_coinit_done: bool = False


def get_word_app(visible: bool = True) -> object:
    """Get or create the singleton Word.Application COM object.

    Args:
        visible: Whether the Word window should be visible to the user.
                 Defaults to True so the user can watch agent operations.

    Returns:
        win32com.client.CDispatch for Word.Application.

    Raises:
        RuntimeError: If Microsoft Word is not installed or cannot be started.
    """
    global _word_app, _coinit_done

    # If we already have a reference, test if it's still alive
    if _word_app is not None:
        try:
            _ = _word_app.Name
            # Ensure visibility setting matches request
            if visible:
                _word_app.Visible = True
            return _word_app
        except Exception:
            logger.info("Word COM connection lost, reconnecting...")
            _word_app = None

    # Initialize COM for this thread
    if not _coinit_done:
        pythoncom.CoInitialize()
        _coinit_done = True

    # Try attaching to an existing Word process first
    try:
        _word_app = win32com.client.GetObject(Class="Word.Application")
        logger.info("Attached to existing Word process")
    except Exception:
        try:
            _word_app = win32com.client.Dispatch("Word.Application")
            logger.info("Created new Word process")
        except Exception as e:
            raise RuntimeError(
                "无法启动 Microsoft Word。请确认 Word 已安装。\n"
                f"原始错误: {e}"
            ) from e

    # Configure the Word instance
    _word_app.Visible = visible
    try:
        _word_app.DisplayAlerts = 0  # wdAlertsNone
    except Exception:
        pass  # Non-critical

    return _word_app


def release_word():
    """Release the COM reference. Does NOT close the Word application."""
    global _word_app, _coinit_done
    _word_app = None
    if _coinit_done:
        try:
            pythoncom.CoUninitialize()
        except Exception:
            pass
        _coinit_done = False
```

- [ ] **Step 4: Run tests, verify they pass**

```bash
python -m pytest tests/test_com_manager.py -v
```

Expected: 3 tests PASS (Word window flashes briefly).

- [ ] **Step 5: Commit**

```bash
git add src/word_mcp/com_manager.py tests/test_com_manager.py
git commit -m "feat: add COM connection manager with lazy init and reconnect"
```

---

### Task 3: MCP Server Skeleton

**Files:**
- Create: `src/word_mcp/server.py`

- [ ] **Step 1: Write `src/word_mcp/server.py`**

```python
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
```

- [ ] **Step 2: Write stub tool modules so server imports**

Create `src/word_mcp/tools/lifecycle.py`:

```python
"""L1: Document lifecycle tools (stub — implemented in Task 4)."""

from mcp.server.fastmcp import FastMCP


def register_lifecycle_tools(mcp: FastMCP):
    """Register document lifecycle tools. Stub: no tools yet."""
    pass
```

Create `src/word_mcp/tools/operations.py`:

```python
"""L2: High-frequency atomic operation tools (stub — implemented in Task 5-7)."""

from mcp.server.fastmcp import FastMCP


def register_operation_tools(mcp: FastMCP):
    """Register atomic operation tools. Stub: no tools yet."""
    pass
```

Create `src/word_mcp/tools/vba.py`:

```python
"""L3: VBA executor tool (stub — implemented in Task 8)."""

from mcp.server.fastmcp import FastMCP


def register_vba_tool(mcp: FastMCP):
    """Register VBA executor tool. Stub: no tools yet."""
    pass
```

- [ ] **Step 3: Verify server starts (will fail on tool call, but imports clean)**

```bash
python -c "from word_mcp.server import mcp; print('Server module OK:', mcp.name)"
```

Expected: `Server module OK: word-mcp`

- [ ] **Step 4: Commit**

```bash
git add src/word_mcp/server.py src/word_mcp/tools/lifecycle.py src/word_mcp/tools/operations.py src/word_mcp/tools/vba.py
git commit -m "feat: add MCP server skeleton with tool module stubs"
```

---

### Task 4: L1 — Document Lifecycle Tools

**Files:**
- Modify: `src/word_mcp/tools/lifecycle.py`
- Create: `src/word_mcp/utils.py`
- Create: `tests/test_lifecycle.py`

- [ ] **Step 1: Write shared utilities**

Create `src/word_mcp/utils.py`:

```python
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
```

- [ ] **Step 2: Write L1 tool tests**

Create `tests/test_lifecycle.py`:

```python
"""Tests for L1 document lifecycle tools."""
import os
import sys
import tempfile

import pytest


DOC_LIFECYCLE_AVAILABLE = sys.platform == "win32"


@pytest.mark.skipif(not DOC_LIFECYCLE_AVAILABLE, reason="COM tests only on Windows")
class TestDocumentLifecycle:
    """End-to-end tests for document lifecycle tools."""

    @pytest.fixture(autouse=True)
    def setup(self):
        """Ensure a Word instance exists and a temp file path is available."""
        from word_mcp.com_manager import get_word_app
        self.app = get_word_app(visible=False)
        self.tmpdir = tempfile.mkdtemp()
        yield
        # Cleanup: close any open documents
        try:
            while self.app.Documents.Count > 0:
                self.app.Documents(1).Close(SaveChanges=False)
        except Exception:
            pass

    def _run(self, coro):
        """Helper: run an async tool function synchronously."""
        import asyncio
        return asyncio.run(coro)

    def test_create_and_save_document(self):
        """word_create + word_save: create a doc and save it."""
        from word_mcp.tools.lifecycle import word_create, word_save

        result1 = self._run(word_create())
        assert "已创建" in result1

        doc_path = os.path.join(self.tmpdir, "test_create.docx")
        result2 = self._run(word_save(file_path=doc_path))
        assert "已保存" in result2
        assert os.path.exists(doc_path)

    def test_open_existing_document(self):
        """word_open: open a .docx file."""
        from word_mcp.tools.lifecycle import word_open

        doc_path = os.path.join(self.tmpdir, "test_open.docx")
        # Create a document first
        doc = self.app.Documents.Add()
        doc.SaveAs(doc_path)
        doc.Close()

        result = self._run(word_open(file_path=doc_path))
        assert "已打开" in result

    def test_export_pdf(self):
        """word_save_as_pdf: export document to PDF."""
        from word_mcp.tools.lifecycle import word_save_as_pdf

        doc = self.app.Documents.Add()
        doc.Content.Text = "Hello PDF"

        pdf_path = os.path.join(self.tmpdir, "test_export.pdf")
        result = self._run(word_save_as_pdf(output_path=pdf_path))
        assert "已导出" in result
        assert os.path.exists(pdf_path)

    def test_close_document(self):
        """word_close: close active document."""
        from word_mcp.tools.lifecycle import word_close

        self.app.Documents.Add()
        assert self.app.Documents.Count == 1

        result = self._run(word_close(save_before_close=False))
        assert "已关闭" in result
        assert self.app.Documents.Count == 0

    def test_get_active_document_info(self):
        """word_get_active_document: return doc info."""
        from word_mcp.tools.lifecycle import word_get_active_document

        doc = self.app.Documents.Add()
        doc.Content.Text = "Test content " * 100

        result = self._run(word_get_active_document())
        assert "当前文档" in result
        assert "页数" in result
```

- [ ] **Step 3: Run tests, verify they fail**

```bash
python -m pytest tests/test_lifecycle.py -v
```

Expected: FAIL (tools not defined in lifecycle.py).

- [ ] **Step 4: Implement `src/word_mcp/tools/lifecycle.py`**

```python
"""L1: Document lifecycle tools — open, create, save, export, close, info."""

import os

from mcp.server.fastmcp import FastMCP

from ..com_manager import get_word_app
from ..utils import ensure_dir, format_error


def register_lifecycle_tools(mcp: FastMCP):
    """Register all L1 document lifecycle tools on the MCP server."""

    @mcp.tool()
    async def word_open(file_path: str) -> str:
        """打开已有 Word 文档。

        Args:
            file_path: 文档的绝对路径或相对路径 (.docx, .doc)
        """
        try:
            app = get_word_app()
            path = os.path.abspath(os.path.expandvars(os.path.expanduser(file_path)))
            doc = app.Documents.Open(path)
            pages = doc.Content.ComputeStatistics(2)  # wdStatisticPages
            return f"已打开文档: {doc.Name}（共 {pages} 页）"
        except Exception as e:
            return format_error("打开文档", e)

    @mcp.tool()
    async def word_create(template_path: str | None = None) -> str:
        """新建 Word 文档，可指定模板。

        Args:
            template_path: 可选，模板文件路径 (.dotx, .dot)
        """
        try:
            app = get_word_app()
            if template_path:
                path = os.path.abspath(os.path.expandvars(os.path.expanduser(template_path)))
                doc = app.Documents.Add(Template=path)
            else:
                doc = app.Documents.Add()
            return f"已创建新文档: {doc.Name}"
        except Exception as e:
            return format_error("创建文档", e)

    @mcp.tool()
    async def word_save(file_path: str | None = None) -> str:
        """保存当前活动文档。

        Args:
            file_path: 可选，另存为目标路径。不提供则保存到原路径。
        """
        try:
            app = get_word_app()
            doc = app.ActiveDocument
            if file_path:
                path = os.path.abspath(os.path.expandvars(os.path.expanduser(file_path)))
                ensure_dir(path)
                doc.SaveAs(path)
            else:
                doc.Save()
            return f"已保存: {doc.FullName}"
        except Exception as e:
            return format_error("保存文档", e)

    @mcp.tool()
    async def word_save_as_pdf(
        output_path: str,
        page_range: str | None = None,
    ) -> str:
        """将当前文档导出为 PDF。

        Args:
            output_path: PDF 文件的输出路径。
            page_range: 可选，页码范围，如 "1-3"。不指定则导出全部。
        """
        try:
            app = get_word_app()
            doc = app.ActiveDocument
            path = os.path.abspath(os.path.expandvars(os.path.expanduser(output_path)))
            ensure_dir(path)

            if page_range:
                start, end = page_range.split("-")
                # ExportAsFixedFormat with wdExportFromTo = 3
                doc.ExportAsFixedFormat(
                    path, 17,  # wdExportFormatPDF = 17
                    Range=3,   # wdExportFromTo
                    From=int(start.strip()),
                    To=int(end.strip()),
                )
            else:
                doc.ExportAsFixedFormat(path, 17)  # wdExportFormatPDF = 17
            return f"已导出 PDF: {path}"
        except Exception as e:
            return format_error("导出 PDF", e)

    @mcp.tool()
    async def word_close(save_before_close: bool = True) -> str:
        """关闭当前活动文档（不关闭 Word 进程）。

        Args:
            save_before_close: 关闭前是否保存，默认 True。
        """
        try:
            app = get_word_app()
            if app.Documents.Count == 0:
                return "没有打开的文档"
            doc = app.ActiveDocument
            name = doc.Name
            doc.Close(SaveChanges=save_before_close)
            return f"已关闭文档: {name}"
        except Exception as e:
            return format_error("关闭文档", e)

    @mcp.tool()
    async def word_get_active_document() -> str:
        """获取当前活动文档的信息（路径、页数、字数等）。"""
        try:
            app = get_word_app()
            if app.Documents.Count == 0:
                return "没有打开的文档"
            doc = app.ActiveDocument
            pages = doc.Content.ComputeStatistics(2)   # wdStatisticPages
            words = doc.Content.ComputeStatistics(0)   # wdStatisticWords
            lines = doc.Content.ComputeStatistics(1)   # wdStatisticLines
            paragraphs = doc.Content.ComputeStatistics(4)  # wdStatisticParagraphs
            return (
                f"当前文档: {doc.FullName}\n"
                f"页数: {pages}\n"
                f"字数: {words}\n"
                f"行数: {lines}\n"
                f"段落数: {paragraphs}"
            )
        except Exception as e:
            return format_error("获取文档信息", e)
```

- [ ] **Step 5: Run tests, verify they pass**

```bash
python -m pytest tests/test_lifecycle.py -v
```

Expected: 5 tests PASS.

- [ ] **Step 6: Commit**

```bash
git add src/word_mcp/utils.py src/word_mcp/tools/lifecycle.py tests/test_lifecycle.py
git commit -m "feat: implement L1 document lifecycle tools (6 tools)"
```

---

### Task 5: L2 — Text & Format Tools (Batch A)

**Files:**
- Modify: `src/word_mcp/tools/operations.py`
- Create: `tests/test_operations.py`

- [ ] **Step 1: Write tests for text/format tools**

Create `tests/test_operations.py`:

```python
"""Tests for L2 atomic operation tools."""
import sys
import tempfile

import pytest


OPS_AVAILABLE = sys.platform == "win32"


@pytest.mark.skipif(not OPS_AVAILABLE, reason="COM tests only on Windows")
class TestTextOperations:
    """Tests for text insertion, formatting, and reading."""

    @pytest.fixture(autouse=True)
    def setup(self):
        import asyncio
        from word_mcp.com_manager import get_word_app
        from word_mcp.tools.lifecycle import word_create

        self.app = get_word_app(visible=False)
        asyncio.run(word_create())
        self.doc = self.app.ActiveDocument
        self.loop = asyncio.new_event_loop()
        yield
        self.loop.close()
        try:
            self.doc.Close(SaveChanges=False)
        except Exception:
            pass

    def _run(self, coro):
        return self.loop.run_until_complete(coro)

    def test_insert_text_at_selection(self):
        """word_insert_text: insert at cursor position."""
        from word_mcp.tools.operations import word_insert_text

        result = self._run(word_insert_text(text="Hello World"))
        assert "已插入" in result
        assert "Hello World" in self.doc.Content.Text

    def test_insert_text_at_start(self):
        """word_insert_text: insert at document start."""
        from word_mcp.tools.operations import word_insert_text

        self.doc.Content.Text = "Existing"
        result = self._run(word_insert_text(text="Start ", position="start"))
        assert "已插入" in result
        assert self.doc.Content.Text.startswith("Start ")

    def test_format_text_bold_and_font(self):
        """word_format_text: make text bold with custom font."""
        from word_mcp.tools.operations import (
            word_insert_text,
            word_format_text,
        )

        self._run(word_insert_text(text="Bold Text"))
        result = self._run(word_format_text(
            range="selection",
            bold=True,
            font_name="Arial",
            font_size=16,
            color="FF0000",
        ))
        assert "已设置" in result

    def test_format_paragraph_alignment(self):
        """word_format_paragraph: set center alignment."""
        from word_mcp.tools.operations import (
            word_insert_text,
            word_format_paragraph,
        )

        self._run(word_insert_text(text="Centered paragraph"))
        result = self._run(word_format_paragraph(
            range="selection",
            alignment="center",
        ))
        assert "已设置" in result

    def test_get_content(self):
        """word_get_content: read document text."""
        from word_mcp.tools.operations import (
            word_insert_text,
            word_get_content,
        )

        self._run(word_insert_text(text="Read this back"))
        result = self._run(word_get_content())
        assert "Read this back" in result

    def test_get_selection(self):
        """word_get_selection: get user-selected text."""
        from word_mcp.tools.operations import word_get_selection

        # Without user selection, returns info about cursor position
        result = self._run(word_get_selection())
        assert isinstance(result, str)

    def test_select_range(self):
        """word_select: select a range of text."""
        from word_mcp.tools.operations import (
            word_insert_text,
            word_select,
        )

        self._run(word_insert_text(text="Selectable text here"))
        result = self._run(word_select(range="start"))
        assert "已选中" in result or "选中" in result
```

- [ ] **Step 2: Run tests, verify they fail**

```bash
python -m pytest tests/test_operations.py -v
```

Expected: FAIL (tools not defined).

- [ ] **Step 3: Implement Batch A tools in `src/word_mcp/tools/operations.py`**

```python
"""L2: High-frequency atomic Word operation tools."""

from mcp.server.fastmcp import FastMCP

from ..com_manager import get_word_app
from ..utils import format_error

# Word enum constants
wdAlignCenter = 1
wdAlignLeft = 0
wdAlignRight = 2
wdAlignJustify = 3

ALIGNMENT_MAP = {
    "left": wdAlignLeft,
    "center": wdAlignCenter,
    "right": wdAlignRight,
    "justify": wdAlignJustify,
}


def _resolve_range(app, doc, range_spec: str):
    """Resolve a range specification string to a Word Range object.

    Args:
        app: Word Application object.
        doc: Word Document object.
        range_spec: "selection" | "start" | "end" | "all"

    Returns:
        A Word Range object.
    """
    if range_spec == "start":
        return doc.Range(0, 0)
    elif range_spec == "end":
        end = doc.Content.End - 1
        return doc.Range(end, end)
    elif range_spec == "all":
        return doc.Content
    else:  # "selection" or default
        return app.Selection.Range


def register_operation_tools(mcp: FastMCP):
    """Register all L2 atomic operation tools on the MCP server."""

    # === Batch A: Text & Format ===

    @mcp.tool()
    async def word_insert_text(
        text: str,
        position: str = "selection",
    ) -> str:
        """在文档中插入文字。

        Args:
            text: 要插入的文字内容。
            position: 插入位置 — "selection"(光标处), "start"(开头), "end"(末尾)
        """
        try:
            app = get_word_app()
            doc = app.ActiveDocument
            rng = _resolve_range(app, doc, position)
            rng.Text = text
            return f"已插入 {len(text)} 字"
        except Exception as e:
            return format_error("插入文字", e)

    @mcp.tool()
    async def word_format_text(
        range: str = "selection",
        font_name: str | None = None,
        font_size: float | None = None,
        bold: bool | None = None,
        italic: bool | None = None,
        underline: bool | None = None,
        color: str | None = None,
    ) -> str:
        """设置文字格式（字体、大小、加粗、颜色等）。

        Args:
            range: 应用范围 — "selection"(选中), "all"(全文)
            font_name: 字体名称，如 "微软雅黑"、"Arial"
            font_size: 字号（磅），如 12、16
            bold: 是否加粗
            italic: 是否斜体
            underline: 是否下划线
            color: 字体颜色，RGB十六进制如 "FF0000"（红色）
        """
        try:
            app = get_word_app()
            doc = app.ActiveDocument
            rng = _resolve_range(app, doc, range)

            changes = []
            font = rng.Font
            if font_name is not None:
                font.Name = font_name
                changes.append(f"字体={font_name}")
            if font_size is not None:
                font.Size = font_size
                changes.append(f"字号={font_size}pt")
            if bold is not None:
                font.Bold = bold
                changes.append(f"加粗={'是' if bold else '否'}")
            if italic is not None:
                font.Italic = italic
                changes.append(f"斜体={'是' if italic else '否'}")
            if underline is not None:
                font.Underline = 1 if underline else 0  # wdUnderlineSingle
                changes.append(f"下划线={'是' if underline else '否'}")
            if color is not None:
                # Convert hex RGB to Word color (long)
                font.Color = int(color, 16)
                changes.append(f"颜色=#{color}")

            if changes:
                return f"已设置: {', '.join(changes)}"
            return "未指定任何格式参数"
        except Exception as e:
            return format_error("设置文字格式", e)

    @mcp.tool()
    async def word_format_paragraph(
        range: str = "selection",
        alignment: str | None = None,
        line_spacing: float | None = None,
        first_line_indent: float | None = None,
        space_before: float | None = None,
        space_after: float | None = None,
    ) -> str:
        """设置段落格式（对齐、行距、缩进等）。

        Args:
            range: 应用范围 — "selection", "all"
            alignment: 对齐方式 — "left", "center", "right", "justify"
            line_spacing: 行距倍数，如 1.5、2.0
            first_line_indent: 首行缩进（磅），如 24 表示两个字符
            space_before: 段前间距（磅）
            space_after: 段后间距（磅）
        """
        try:
            app = get_word_app()
            doc = app.ActiveDocument
            rng = _resolve_range(app, doc, range)

            changes = []
            pf = rng.ParagraphFormat

            if alignment is not None:
                if alignment in ALIGNMENT_MAP:
                    pf.Alignment = ALIGNMENT_MAP[alignment]
                    changes.append(f"对齐={alignment}")

            if line_spacing is not None:
                pf.LineSpacingRule = 0  # wdLineSpaceMultiple
                pf.LineSpacing = line_spacing
                changes.append(f"行距={line_spacing}x")

            if first_line_indent is not None:
                pf.FirstLineIndent = first_line_indent
                changes.append(f"首行缩进={first_line_indent}pt")

            if space_before is not None:
                pf.SpaceBefore = space_before
                changes.append(f"段前={space_before}pt")

            if space_after is not None:
                pf.SpaceAfter = space_after
                changes.append(f"段后={space_after}pt")

            if changes:
                return f"已设置段落格式: {', '.join(changes)}"
            return "未指定任何段落格式参数"
        except Exception as e:
            return format_error("设置段落格式", e)

    @mcp.tool()
    async def word_get_content(
        range: str = "all",
        format: str = "text",
    ) -> str:
        """读取文档内容。

        Args:
            range: 读取范围 — "all"(全文), "selection"(选中)
            format: 返回格式 — "text"(纯文本), "html"(带格式)
        """
        try:
            app = get_word_app()
            doc = app.ActiveDocument

            if range == "selection":
                rng = app.Selection.Range
                label = "选中内容"
            else:
                rng = doc.Content
                label = "全文"

            if format == "html":
                content = rng.FormattedText.WordOpenXML if hasattr(rng, 'WordOpenXML') else rng.Text
            else:
                content = rng.Text

            if not content.strip():
                return f"[文档为空] ({label})"
            return content
        except Exception as e:
            return format_error("读取内容", e)

    @mcp.tool()
    async def word_get_selection() -> str:
        """获取当前用户在 Word 中选中的文字。"""
        try:
            app = get_word_app()
            sel = app.Selection
            text = sel.Text
            if not text.strip():
                return "当前未选中任何文字"
            return f"选中文字: {text}"
        except Exception as e:
            return format_error("获取选中内容", e)

    @mcp.tool()
    async def word_select(range: str) -> str:
        """选中文档中的指定区域（用户可在 Word 窗口中看到高亮）。

        Args:
            range: "start"(跳到开头), "end"(跳到末尾), "all"(全选)
        """
        try:
            app = get_word_app()
            doc = app.ActiveDocument
            rng = _resolve_range(app, doc, range)
            rng.Select()
            return f"已选中: {range}"
        except Exception as e:
            return format_error("选中区域", e)

    @mcp.tool()
    async def word_apply_style(
        range: str = "selection",
        style_name: str = "Normal",
    ) -> str:
        """对指定区域应用 Word 内置样式。

        Args:
            range: 应用范围
            style_name: 样式名，如 "Normal", "Heading 1", "Heading 2", "Title"
        """
        try:
            app = get_word_app()
            doc = app.ActiveDocument
            rng = _resolve_range(app, doc, range)

            # Try to get the style by name
            try:
                style = doc.Styles(style_name)
            except Exception:
                # List available styles
                available = []
                for i in range(1, min(doc.Styles.Count + 1, 30)):
                    try:
                        s = doc.Styles(i)
                        if s.NameLocal:
                            available.append(s.NameLocal)
                    except Exception:
                        pass
                return f"样式 '{style_name}' 不存在。可用样式: {', '.join(available[:15])}"

            rng.Style = style
            return f"已应用样式: {style_name}"
        except Exception as e:
            return format_error("应用样式", e)

    @mcp.tool()
    async def word_auto_numbering(
        range: str = "selection",
        type: str = "bullet",
    ) -> str:
        """为指定段落设置编号或项目符号。

        Args:
            range: 应用范围
            type: "bullet"(项目符号), "number"(数字编号)
        """
        try:
            app = get_word_app()
            doc = app.ActiveDocument
            rng = _resolve_range(app, doc, range)

            if type == "bullet":
                rng.ListFormat.ApplyBulletDefault()
            elif type == "number":
                rng.ListFormat.ApplyNumberDefault()
            else:
                return f"不支持的编号类型: {type}，请使用 'bullet' 或 'number'"

            label = "项目符号" if type == "bullet" else "数字编号"
            return f"已设置: {label}"
        except Exception as e:
            return format_error("设置编号", e)
```

- [ ] **Step 4: Run tests, verify they pass**

```bash
python -m pytest tests/test_operations.py -v
```

Expected: 7 tests PASS.

- [ ] **Step 5: Commit**

```bash
git add src/word_mcp/tools/operations.py tests/test_operations.py
git commit -m "feat: implement L2 text, format, and content tools (7 tools)"
```

---

### Task 6: L2 — Table & Image Tools (Batch B)

**Files:**
- Modify: `src/word_mcp/tools/operations.py`
- Modify: `tests/test_operations.py`

- [ ] **Step 1: Add tests for table and image tools**

Append to `tests/test_operations.py`:

```python
@pytest.mark.skipif(not OPS_AVAILABLE, reason="COM tests only on Windows")
class TestTableOperations:
    """Tests for table insertion and formatting."""

    @pytest.fixture(autouse=True)
    def setup(self):
        import asyncio
        from word_mcp.com_manager import get_word_app
        from word_mcp.tools.lifecycle import word_create

        self.app = get_word_app(visible=False)
        asyncio.run(word_create())
        self.doc = self.app.ActiveDocument
        self.loop = asyncio.new_event_loop()
        yield
        self.loop.close()
        try:
            self.doc.Close(SaveChanges=False)
        except Exception:
            pass

    def _run(self, coro):
        return self.loop.run_until_complete(coro)

    def test_insert_table(self):
        """word_insert_table: insert a table with data."""
        from word_mcp.tools.operations import word_insert_table

        data = [["Name", "Age"], ["Alice", "30"], ["Bob", "25"]]
        result = self._run(word_insert_table(rows=3, cols=2, data=data))
        assert "已插入" in result
        assert self.doc.Tables.Count == 1

    def test_format_table(self):
        """word_format_table: apply style to a table."""
        from word_mcp.tools.operations import (
            word_insert_table,
            word_format_table,
        )

        self._run(word_insert_table(rows=2, cols=2))
        result = self._run(word_format_table(
            table_index=1,
            header_row=True,
            auto_fit=True,
        ))
        assert "已设置" in result


@pytest.mark.skipif(not OPS_AVAILABLE, reason="COM tests only on Windows")
class TestImageAndBreakOperations:
    """Tests for image insertion and page breaks."""

    @pytest.fixture(autouse=True)
    def setup(self):
        import asyncio
        from word_mcp.com_manager import get_word_app
        from word_mcp.tools.lifecycle import word_create

        self.app = get_word_app(visible=False)
        asyncio.run(word_create())
        self.doc = self.app.ActiveDocument
        self.loop = asyncio.new_event_loop()
        self.tmpdir = tempfile.mkdtemp()
        yield
        self.loop.close()
        try:
            self.doc.Close(SaveChanges=False)
        except Exception:
            pass

    def _run(self, coro):
        return self.loop.run_until_complete(coro)

    def test_insert_page_break(self):
        """word_insert_page_break: insert a page break."""
        from word_mcp.tools.operations import word_insert_page_break

        result = self._run(word_insert_page_break())
        assert "已插入" in result

    def test_find_replace(self):
        """word_find_replace: find and replace text."""
        from word_mcp.tools.operations import (
            word_insert_text,
            word_find_replace,
        )

        self._run(word_insert_text(text="Hello Alice, hello Bob"))
        result = self._run(word_find_replace(
            find_text="Hello",
            replace_text="Hi",
        ))
        assert "替换" in result
        assert "Hi Alice" in self.doc.Content.Text
```

- [ ] **Step 2: Run tests, verify they fail**

```bash
python -m pytest tests/test_operations.py -v -k "Table or Image or Break or find"
```

Expected: FAIL (tools not implemented yet).

- [ ] **Step 3: Add Batch B tools to `register_operation_tools`**

Add these tools inside `register_operation_tools()` in `src/word_mcp/tools/operations.py`, before the function's `return`:

```python
    # === Batch B: Table, Image, Page Break, Find/Replace ===

    @mcp.tool()
    async def word_insert_table(
        rows: int,
        cols: int,
        data: list[list[str]] | None = None,
        position: str = "selection",
    ) -> str:
        """插入表格，可附带数据填充。

        Args:
            rows: 行数
            cols: 列数
            data: 可选，二维数组，用于填充表格内容。第一行作为表头。
            position: 插入位置
        """
        try:
            app = get_word_app()
            doc = app.ActiveDocument
            rng = _resolve_range(app, doc, position)

            table = doc.Tables.Add(rng, rows, cols)

            # Apply default borders
            table.Borders.Enable = True

            # Fill data if provided
            if data:
                for i, row_data in enumerate(data):
                    for j, cell_text in enumerate(row_data):
                        if i < rows and j < cols:
                            cell = table.Cell(i + 1, j + 1)
                            cell.Range.Text = str(cell_text)

            if data:
                return f"已插入表格: {rows}行×{cols}列（已填充 {len(data)} 行数据）"
            return f"已插入表格: {rows}行×{cols}列"
        except Exception as e:
            return format_error("插入表格", e)

    @mcp.tool()
    async def word_format_table(
        table_index: int = 1,
        style: str | None = None,
        header_row: bool = False,
        auto_fit: bool = False,
    ) -> str:
        """设置表格格式。

        Args:
            table_index: 表格序号（1=第一个表格）
            style: Word 表格样式名，如 "Grid Table 1 Light"
            header_row: 是否设置首行为表头（加粗 + 重复标题行）
            auto_fit: 是否自适应列宽
        """
        try:
            app = get_word_app()
            doc = app.ActiveDocument
            table = doc.Tables(table_index)

            changes = []

            if style is not None:
                try:
                    table.Style = style
                    changes.append(f"样式={style}")
                except Exception:
                    pass  # Style name may not exist; skip silently

            if header_row:
                # Format first row as bold
                for j in range(1, table.Columns.Count + 1):
                    table.Cell(1, j).Range.Font.Bold = True
                # Repeat header row across pages
                table.Rows(1).HeadingFormat = -1  # True in VBA
                changes.append("首行为表头")

            if auto_fit:
                table.AutoFitBehavior(2)  # wdAutoFitWindow
                changes.append("自适应列宽")

            if changes:
                return f"已设置表格格式: {', '.join(changes)}"
            return "未指定任何表格格式参数"
        except Exception as e:
            return format_error("设置表格格式", e)

    @mcp.tool()
    async def word_insert_image(
        image_path: str,
        position: str = "selection",
        width: float | None = None,
        height: float | None = None,
    ) -> str:
        """在文档中插入图片。

        Args:
            image_path: 图片文件路径（PNG, JPG, GIF, BMP 等）
            position: 插入位置
            width: 可选，图片宽度（磅），不指定则保持原始宽度
            height: 可选，图片高度（磅）
        """
        try:
            import os

            app = get_word_app()
            doc = app.ActiveDocument
            path = os.path.abspath(os.path.expandvars(os.path.expanduser(image_path)))

            if not os.path.exists(path):
                return f"图片文件不存在: {path}"

            rng = _resolve_range(app, doc, position)
            shape = doc.InlineShapes.AddPicture(
                path,
                LinkToFile=False,
                SaveWithDocument=True,
                Range=rng,
            )

            if width is not None:
                shape.Width = width
            if height is not None:
                shape.Height = height

            return f"已插入图片: {os.path.basename(path)}（{shape.Width:.0f}×{shape.Height:.0f}pt）"
        except Exception as e:
            return format_error("插入图片", e)

    @mcp.tool()
    async def word_insert_page_break(position: str = "selection") -> str:
        """在指定位置插入分页符。

        Args:
            position: 插入位置
        """
        try:
            app = get_word_app()
            doc = app.ActiveDocument
            rng = _resolve_range(app, doc, position)
            rng.InsertBreak(7)  # wdPageBreak = 7
            return "已插入分页符"
        except Exception as e:
            return format_error("插入分页符", e)

    @mcp.tool()
    async def word_find_replace(
        find_text: str,
        replace_text: str,
        match_case: bool = False,
    ) -> str:
        """在全文范围内查找并替换文字。

        Args:
            find_text: 要查找的文字
            replace_text: 替换后的文字
            match_case: 是否区分大小写
        """
        try:
            app = get_word_app()
            doc = app.ActiveDocument

            find = doc.Content.Find
            find.Text = find_text
            find.Replacement.Text = replace_text
            find.MatchCase = match_case
            find.Forward = True
            find.Wrap = 1  # wdFindContinue

            # wdReplaceAll = 2
            find.Execute(Replace=2)

            return f"已完成查找替换: '{find_text}' → '{replace_text}'"
        except Exception as e:
            return format_error("查找替换", e)
```

- [ ] **Step 4: Run tests, verify they pass**

```bash
python -m pytest tests/test_operations.py -v
```

Expected: 11 tests PASS.

- [ ] **Step 5: Commit**

```bash
git add src/word_mcp/tools/operations.py tests/test_operations.py
git commit -m "feat: implement L2 table, image, page break, and find/replace tools"
```

---

### Task 7: L2 — Page Setup & Header/Footer Tools (Batch C)

**Files:**
- Modify: `src/word_mcp/tools/operations.py`
- Modify: `tests/test_operations.py`

- [ ] **Step 1: Add tests for page setup and header/footer**

Append to `tests/test_operations.py`:

```python
@pytest.mark.skipif(not OPS_AVAILABLE, reason="COM tests only on Windows")
class TestPageSetupHeaderFooter:
    """Tests for page layout and header/footer."""

    @pytest.fixture(autouse=True)
    def setup(self):
        import asyncio
        from word_mcp.com_manager import get_word_app
        from word_mcp.tools.lifecycle import word_create

        self.app = get_word_app(visible=False)
        asyncio.run(word_create())
        self.doc = self.app.ActiveDocument
        self.loop = asyncio.new_event_loop()
        yield
        self.loop.close()
        try:
            self.doc.Close(SaveChanges=False)
        except Exception:
            pass

    def _run(self, coro):
        return self.loop.run_until_complete(coro)

    def test_set_page_setup_a4_landscape(self):
        """word_set_page_setup: set A4 landscape with custom margins."""
        from word_mcp.tools.operations import word_set_page_setup

        result = self._run(word_set_page_setup(
            orientation="landscape",
            page_size="A4",
            margins={"top": 72, "bottom": 72, "left": 90, "right": 90},
        ))
        assert "已设置" in result

    def test_set_header_footer(self):
        """word_set_header_footer: add header with page number."""
        from word_mcp.tools.operations import word_set_header_footer

        result = self._run(word_set_header_footer(
            type="header",
            text="Chapter 1: Introduction",
            include_pagenum=True,
        ))
        assert "已设置" in result
```

- [ ] **Step 2: Run tests, verify they fail**

```bash
python -m pytest tests/test_operations.py -v -k "PageSetup or HeaderFooter"
```

Expected: FAIL (tools not implemented).

- [ ] **Step 3: Add Batch C tools to `register_operation_tools`**

```python
    # === Batch C: Page Setup, Header/Footer ===

    @mcp.tool()
    async def word_set_page_setup(
        margins: dict | None = None,
        orientation: str | None = None,
        page_size: str | None = None,
    ) -> str:
        """设置页面布局（边距、方向、纸张大小）。

        Args:
            margins: 边距字典，单位磅（1英寸=72磅）。
                     例: {"top": 72, "bottom": 72, "left": 90, "right": 90}
            orientation: 纸张方向 — "portrait"(纵向), "landscape"(横向)
            page_size: 纸张尺寸 — "A4", "A3", "Letter", "Legal"
        """
        try:
            app = get_word_app()
            doc = app.ActiveDocument
            ps = doc.PageSetup
            changes = []

            if margins:
                if "top" in margins:
                    ps.TopMargin = margins["top"]
                    changes.append(f"上边距={margins['top']}pt")
                if "bottom" in margins:
                    ps.BottomMargin = margins["bottom"]
                    changes.append(f"下边距={margins['bottom']}pt")
                if "left" in margins:
                    ps.LeftMargin = margins["left"]
                    changes.append(f"左边距={margins['left']}pt")
                if "right" in margins:
                    ps.RightMargin = margins["right"]
                    changes.append(f"右边距={margins['right']}pt")

            if orientation:
                # wdOrientLandscape = 1, wdOrientPortrait = 0
                ps.Orientation = 1 if orientation == "landscape" else 0
                changes.append(f"方向={orientation}")

            if page_size:
                page_sizes = {
                    "A4": (595.3, 841.9),
                    "A3": (841.9, 1190.5),
                    "Letter": (612, 792),
                    "Legal": (612, 1008),
                }
                if page_size in page_sizes:
                    w, h = page_sizes[page_size]
                    ps.PageWidth = w
                    ps.PageHeight = h
                    changes.append(f"纸张={page_size}")

            if changes:
                return f"已设置页面: {', '.join(changes)}"
            return "未指定任何页面设置参数"
        except Exception as e:
            return format_error("设置页面", e)

    @mcp.tool()
    async def word_set_header_footer(
        type: str = "header",
        text: str = "",
        include_pagenum: bool = False,
    ) -> str:
        """设置页眉或页脚内容。

        Args:
            type: "header"(页眉) 或 "footer"(页脚)
            text: 页眉/页脚的文字内容
            include_pagenum: 是否添加页码
        """
        try:
            app = get_word_app()
            doc = app.ActiveDocument

            # Get the correct section and header/footer
            section = doc.Sections(1)
            if type == "header":
                hf = section.Headers(1)  # wdHeaderFooterPrimary
                label = "页眉"
            else:
                hf = section.Footers(1)  # wdHeaderFooterPrimary
                label = "页脚"

            rng = hf.Range

            if text:
                rng.Text = text

            if include_pagenum:
                # Insert page number field
                if text:
                    # Add separator then page number
                    rng.InsertAfter("  ")
                    rng.Collapse(0)  # wdCollapseEnd

                # Add "Page X of Y" format
                rng.InsertAfter("第 ")
                rng.Collapse(0)

                # PAGE field
                page_field = doc.Fields.Add(rng, 33)  # wdFieldPage = 33

                rng.InsertAfter(" 页 / 共 ")
                rng.Collapse(0)

                # NUMPAGES field
                num_field = doc.Fields.Add(rng, 26)  # wdFieldNumPages = 26

                rng.InsertAfter(" 页")

            parts = []
            if text:
                parts.append(f"内容='{text}'")
            if include_pagenum:
                parts.append("含页码")
            return f"已设置{label}: {', '.join(parts)}"
        except Exception as e:
            return format_error(f"设置{type}", e)
```

- [ ] **Step 4: Run tests, verify they pass**

```bash
python -m pytest tests/test_operations.py -v
```

Expected: 13 tests PASS.

- [ ] **Step 5: Commit**

```bash
git add src/word_mcp/tools/operations.py tests/test_operations.py
git commit -m "feat: implement L2 page setup and header/footer tools"
```

---

### Task 8: L3 — VBA Executor

**Files:**
- Modify: `src/word_mcp/tools/vba.py`
- Create: `tests/test_vba.py`

- [ ] **Step 1: Write VBA executor tests**

Create `tests/test_vba.py`:

```python
"""Tests for L3 VBA executor tool."""
import sys

import pytest


VBA_AVAILABLE = sys.platform == "win32"


@pytest.mark.skipif(not VBA_AVAILABLE, reason="COM tests only on Windows")
class TestVBAExecutor:
    """Tests for word_execute_vba."""

    @pytest.fixture(autouse=True)
    def setup(self):
        import asyncio
        from word_mcp.com_manager import get_word_app
        from word_mcp.tools.lifecycle import word_create

        self.app = get_word_app(visible=False)
        asyncio.run(word_create())
        self.doc = self.app.ActiveDocument
        self.loop = asyncio.new_event_loop()
        yield
        self.loop.close()
        try:
            self.doc.Close(SaveChanges=False)
        except Exception:
            pass

    def _run(self, coro):
        return self.loop.run_until_complete(coro)

    def test_insert_text_via_vba(self):
        """Execute VBA to insert text at cursor position."""
        from word_mcp.tools.vba import word_execute_vba

        code = 'Sub Main()\n  Selection.TypeText "VBA inserted text"\nEnd Sub'
        result = self._run(word_execute_vba(code=code))
        assert "已执行" in result or "VBA" in result

    def test_vba_type_mismatch_error(self):
        """VBA runtime errors should be caught and reported."""
        from word_mcp.tools.vba import word_execute_vba

        code = 'Sub Main()\n  Dim x As Integer\n  x = "not a number"\nEnd Sub'
        result = self._run(word_execute_vba(code=code))
        # Should not crash; should return an error message
        assert isinstance(result, str)
        assert len(result) > 0

    def test_vba_syntax_error(self):
        """VBA compile errors should be caught."""
        from word_mcp.tools.vba import word_execute_vba

        code = "Sub Main()\n  invalid syntax here !!!\nEnd Sub"
        result = self._run(word_execute_vba(code=code))
        assert isinstance(result, str)
        assert len(result) > 0
```

- [ ] **Step 2: Run tests, verify they fail**

```bash
python -m pytest tests/test_vba.py -v
```

Expected: FAIL (tool not implemented).

- [ ] **Step 3: Implement `src/word_mcp/tools/vba.py`**

```python
"""L3: VBA executor — the final fallback for any Word operation.

Injects VBA code into Word's VBE (Visual Basic Editor) as a temporary
module, runs the Main() subroutine, then removes the module.

Requires: Word Trust Center → "Trust access to the VBA project object model"
"""

import logging

from mcp.server.fastmcp import FastMCP

from ..com_manager import get_word_app
from ..utils import format_error

logger = logging.getLogger(__name__)

# COM enum constants
vbext_ct_StdModule = 1
wdLine = 5


def register_vba_tool(mcp: FastMCP):
    """Register the VBA executor tool on the MCP server."""

    @mcp.tool()
    async def word_execute_vba(code: str) -> str:
        """在 Word 进程中执行 VBA 代码，覆盖所有 L2 未覆盖的操作。

        VBA 代码必须包含一个名为 Main 的 Sub 过程。
        例如:
            Sub Main()
                Selection.TypeText "Hello from VBA"
            End Sub

        Args:
            code: 完整的 VBA 代码字符串，必须包含 Sub Main()。

        Notes:
            - 代码在临时模块中运行，执行后自动删除
            - 不修改 Normal.dotm 或创建持久化宏
            - 需要 Word 启用「信任对 VBA 项目对象模型的访问」
              (文件 → 选项 → 信任中心 → 宏设置)
        """
        vb_comp = None
        app = get_word_app()

        # Validate: must contain Sub Main()
        if "Sub Main(" not in code and "Sub Main (" not in code:
            return (
                "错误: VBA 代码必须包含 Sub Main() 过程。\n"
                "示例:\n"
                "Sub Main()\n"
                "    Selection.TypeText \"Hello\"\n"
                "End Sub"
            )

        try:
            # Access the VBA project of the active document
            try:
                vb_project = app.ActiveDocument.VBProject
            except Exception:
                return (
                    "错误: 无法访问 VBA 项目对象模型。\n"
                    "请在 Word 中执行以下操作:\n"
                    "1. 文件 → 选项 → 信任中心 → 信任中心设置\n"
                    "2. 宏设置 → 勾选「信任对 VBA 项目对象模型的访问」\n"
                    "3. 重新启动 Word"
                )

            # Create a temporary module and inject the code
            vb_comp = vb_project.VBComponents.Add(vbext_ct_StdModule)
            vb_comp.CodeModule.AddFromString(code)

            # Run Sub Main()
            app.Run("Main")

            return "VBA 代码已执行成功"

        except Exception as e:
            error_msg = str(e)

            # Extract VBA error details if available
            if hasattr(e, 'args') and len(e.args) > 0:
                error_msg = str(e.args[0])

            return f"VBA 执行错误: {error_msg}"

        finally:
            # Clean up: remove the temporary module
            if vb_comp is not None:
                try:
                    vb_project.VBComponents.Remove(vb_comp)
                except Exception as cleanup_err:
                    logger.warning(f"Failed to clean up VBA module: {cleanup_err}")
```

- [ ] **Step 4: Run tests, verify they pass**

```bash
python -m pytest tests/test_vba.py -v
```

Expected: 3 tests PASS (or 1 PASS + 2 SKIP if VBA access is disabled — the error message should still be informative).

- [ ] **Step 5: Commit**

```bash
git add src/word_mcp/tools/vba.py tests/test_vba.py
git commit -m "feat: implement L3 VBA executor tool"
```

---

### Task 9: End-to-End Integration Verification

**Files:**
- Modify: `tests/test_lifecycle.py` (add E2E scenario)

- [ ] **Step 1: Add end-to-end scenario test**

Append to `tests/test_lifecycle.py`:

```python
@pytest.mark.skipif(not DOC_LIFECYCLE_AVAILABLE, reason="COM tests only on Windows")
class TestEndToEnd:
    """End-to-end scenario: create a complete document."""

    @pytest.fixture(autouse=True)
    def setup(self):
        import asyncio
        from word_mcp.com_manager import get_word_app

        self.app = get_word_app(visible=False)
        self.loop = asyncio.new_event_loop()
        self.tmpdir = tempfile.mkdtemp()
        yield
        self.loop.close()
        try:
            while self.app.Documents.Count > 0:
                self.app.Documents(1).Close(SaveChanges=False)
        except Exception:
            pass

    def _run(self, coro):
        return self.loop.run_until_complete(coro)

    def test_create_resume_document(self):
        """Simulate: agent creates a resume with table + formatted text."""
        from word_mcp.tools.lifecycle import (
            word_create,
            word_save,
            word_get_active_document,
        )
        from word_mcp.tools.operations import (
            word_insert_text,
            word_format_text,
            word_insert_table,
            word_format_table,
            word_insert_page_break,
        )

        # 1. Create new document
        self._run(word_create())
        assert self.app.Documents.Count == 1

        # 2. Insert title
        self._run(word_insert_text(text="张三的简历\n", position="start"))
        self._run(word_format_text(range="all", bold=True, font_size=22))

        # 3. Insert personal info table
        data = [
            ["姓名", "张三"],
            ["电话", "138-0000-0000"],
            ["邮箱", "zhangsan@example.com"],
        ]
        self._run(word_insert_table(rows=3, cols=2, data=data))
        self._run(word_format_table(table_index=1, header_row=True))

        # 4. Insert work experience section
        self._run(word_insert_text(text="\n工作经历\n"))
        self._run(word_insert_text(
            text="2020-2024  某科技有限公司  高级工程师\n"
        ))

        # 5. Insert page break
        self._run(word_insert_page_break())

        # 6. Save
        doc_path = os.path.join(self.tmpdir, "简历.docx")
        result = self._run(word_save(file_path=doc_path))
        assert "已保存" in result
        assert os.path.exists(doc_path)

        # 7. Verify document has content
        info = self._run(word_get_active_document())
        assert "页数" in info
```

- [ ] **Step 2: Run the E2E test**

```bash
python -m pytest tests/test_lifecycle.py::TestEndToEnd -v
```

Expected: 1 test PASS. Document created successfully via tool chain.

- [ ] **Step 3: Run full test suite**

```bash
python -m pytest tests/ -v
```

Expected: All tests PASS or SKIP (non-Windows).

- [ ] **Step 4: Commit**

```bash
git add tests/test_lifecycle.py
git commit -m "test: add end-to-end resume creation scenario"
```

---

### Task 10: README Documentation

**Files:**
- Create: `README.md`

- [ ] **Step 1: Write `README.md`**

```markdown
# Word MCP Server

MCP server that gives Claude Agent full control over Microsoft Word via COM automation. The Word window stays visible so you can watch the agent work in real time.

## Requirements

- Windows 10 or later
- Microsoft Word 2016 or later (desktop version)
- Python 3.10+

## Installation

```bash
cd word-mcp-server
pip install -e .
```

## Claude Code Configuration

Add to `settings.local.json`:

```json
{
  "mcpServers": {
    "word": {
      "command": "python",
      "args": ["-m", "word_mcp"],
      "type": "local"
    }
  }
}
```

Restart Claude Code.

## Enabling VBA Support (L3)

For the `word_execute_vba` tool to work:

1. Open Word
2. File → Options → Trust Center → Trust Center Settings
3. Macro Settings → Check "Trust access to the VBA project object model"
4. Restart Word

## Available Tools

### L1 — Document Lifecycle (6 tools)

| Tool | Description |
|------|-------------|
| `word_open` | Open a .docx/.doc file |
| `word_create` | Create a new document |
| `word_save` | Save the active document |
| `word_save_as_pdf` | Export to PDF |
| `word_close` | Close the active document |
| `word_get_active_document` | Get document info |

### L2 — Atomic Operations (15 tools)

| Tool | Description |
|------|-------------|
| `word_insert_text` | Insert text at cursor/start/end |
| `word_format_text` | Set font, size, bold, italic, color |
| `word_format_paragraph` | Set alignment, spacing, indent |
| `word_insert_table` | Insert a table with data |
| `word_format_table` | Style a table |
| `word_insert_image` | Insert an image file |
| `word_insert_page_break` | Insert a page break |
| `word_set_page_setup` | Set margins, orientation, paper size |
| `word_set_header_footer` | Set header/footer with page numbers |
| `word_find_replace` | Find and replace text |
| `word_get_content` | Read document text |
| `word_get_selection` | Get user-selected text |
| `word_select` | Select a range (visible highlight) |
| `word_apply_style` | Apply Word built-in style |
| `word_auto_numbering` | Bullet or numbered list |

### L3 — VBA Executor (1 tool)

| Tool | Description |
|------|-------------|
| `word_execute_vba` | Execute arbitrary VBA code |

## Example: Ask Claude to Create a Resume

```
User: 帮我创建一份简历，包含个人信息表和工作经历

Agent will:
1. word_create()              — new document
2. word_insert_text("张三的简历") — title
3. word_format_text(bold=true) — format title
4. word_insert_table(rows, cols, data) — info table
5. word_format_table(header_row=true) — style table
6. word_insert_text("工作经历") — section heading
7. word_apply_style("Heading 2") — heading style
8. word_insert_text("2020-2024 ...") — experience
9. word_save("简历.docx") — save

You watch it happen in the Word window.
```
```

- [ ] **Step 2: Commit**

```bash
git add README.md
git commit -m "docs: add README with tool reference and usage examples"
```

---

### Task 11: Final Verification & MCP Config

**Files:**
- Modify: `.claude/settings.local.json` (in Desktop)

- [ ] **Step 1: Run full test suite one final time**

```bash
cd C:/Users/19773/Desktop/word-mcp-server
python -m pytest tests/ -v --tb=short
```

Expected: all tests PASS or SKIP cleanly.

- [ ] **Step 2: Add MCP server configuration**

Read the current `.claude/settings.local.json` and add the `mcpServers.word` entry:

```json
{
  "permissions": { ... },
  "mcpServers": {
    "word": {
      "command": "python",
      "args": ["-m", "word_mcp"],
      "type": "local"
    }
  }
}
```

- [ ] **Step 3: Verify MCP server starts standalone**

```bash
cd C:/Users/19773/Desktop/word-mcp-server
python -m word_mcp --help 2>&1 || python -c "from word_mcp.server import mcp; print('Server ready:', mcp.name)"
```

Expected: `Server ready: word-mcp`

- [ ] **Step 4: Final commit**

```bash
git add -A
git commit -m "chore: finalize configuration and verify integration"
```
