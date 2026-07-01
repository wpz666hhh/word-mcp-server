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

        import pythoncom

        from word_mcp.com_manager import get_word_app

        # Ensure COM is initialized
        try:
            pythoncom.CoInitialize()
        except Exception:
            pass

        self.app = get_word_app(visible=False)
        # Create a new document for testing
        self.doc = self.app.Documents.Add()
        self.loop = asyncio.new_event_loop()
        yield
        self.loop.close()
        try:
            self.doc.Close(SaveChanges=False)
        except Exception:
            pass
        # Also close any other stray docs
        try:
            while self.app.Documents.Count > 0:
                self.app.Documents(1).Close(SaveChanges=False)
        except Exception:
            pass

    def _run(self, coro):
        return self.loop.run_until_complete(coro)

    def test_insert_text_at_selection(self):
        """word_insert_text: insert at cursor position."""
        from word_mcp.tools.text import word_insert_text

        result = word_insert_text(text="Hello World")
        assert "已插入" in result
        assert "Hello World" in self.doc.Content.Text

    def test_insert_text_at_start(self):
        """word_insert_text: insert at document start."""
        from word_mcp.tools.text import word_insert_text

        self.doc.Content.Text = "Existing"
        result = word_insert_text(text="Start ", position="start")
        assert "已插入" in result
        assert self.doc.Content.Text.startswith("Start ")

    def test_format_text_bold_and_font(self):
        """word_format_text: make text bold with custom font."""
        from word_mcp.tools.text import word_insert_text, word_format_text

        word_insert_text(text="Bold Text")
        result = word_format_text(
            range_spec="selection",
            bold=True,
            font_name="Arial",
            font_size=16,
            color="FF0000",
        )
        assert "已设置" in result

    def test_format_paragraph_alignment(self):
        """word_format_paragraph: set center alignment."""
        from word_mcp.tools.text import word_insert_text, word_format_paragraph

        word_insert_text(text="Centered paragraph")
        result = word_format_paragraph(range_spec="selection", alignment="center")
        assert "已设置" in result

    def test_get_content(self):
        """word_get_content: read document text."""
        from word_mcp.tools.text import word_insert_text, word_get_content

        word_insert_text(text="Read this back")
        result = word_get_content()
        assert "Read this back" in result

    def test_get_selection(self):
        """word_get_selection: get user-selected text."""
        from word_mcp.tools.text import word_get_selection

        result = word_get_selection()
        assert isinstance(result, str)

    def test_select_range(self):
        """word_select: select a range of text."""
        from word_mcp.tools.text import word_insert_text, word_select

        word_insert_text(text="Selectable text here")
        result = word_select(range_spec="start")
        assert "已选中" in result or "选中" in result

    def test_apply_style(self):
        """word_apply_style: apply a Word style."""
        from word_mcp.tools.text import word_insert_text, word_apply_style

        word_insert_text(text="A heading")
        result = word_apply_style(range_spec="selection", style_name="Heading 1")
        assert "已应用" in result or "样式" in result

    def test_auto_numbering_bullet(self):
        """word_auto_numbering: apply bullet list."""
        from word_mcp.tools.text import word_insert_text, word_auto_numbering

        word_insert_text(text="Bullet item")
        result = word_auto_numbering(range_spec="selection", type="bullet")
        assert "已设置" in result


@pytest.mark.skipif(not OPS_AVAILABLE, reason="COM tests only on Windows")
class TestTableOperations:
    """Tests for table insertion and formatting."""

    @pytest.fixture(autouse=True)
    def setup(self):
        import asyncio
        from word_mcp.com_manager import get_word_app
        from word_mcp.tools.lifecycle import word_create

        self.app = get_word_app(visible=False)
        word_create()
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
        from word_mcp.tools.tables import word_insert_table

        data = [["Name", "Age"], ["Alice", "30"], ["Bob", "25"]]
        result = word_insert_table(rows=3, cols=2, data=data)
        assert "已插入" in result
        assert self.doc.Tables.Count == 1

    def test_format_table(self):
        """word_format_table: apply style to a table."""
        from word_mcp.tools.tables import word_insert_table, word_format_table

        word_insert_table(rows=2, cols=2)
        result = word_format_table(
            table_index=1,
            header_row=True,
            auto_fit=True,
        )
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
        word_create()
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
        from word_mcp.tools.layout import word_insert_page_break

        word_insert_page_break()
        assert self.doc.Content.Text.count("\x0c") >= 1  # Form feed char

    def test_find_replace(self):
        """word_find_replace: find and replace text."""
        from word_mcp.tools.text import word_insert_text, word_find_replace

        word_insert_text(text="Hello Alice, hello Bob")
        result = word_find_replace(
            find_text="Hello",
            replace_text="Hi",
        )
        assert "已完成" in result
        assert "Hi Alice" in self.doc.Content.Text


@pytest.mark.skipif(not OPS_AVAILABLE, reason="COM tests only on Windows")
class TestPageSetupHeaderFooter:
    """Tests for page layout and header/footer."""

    @pytest.fixture(autouse=True)
    def setup(self):
        import asyncio
        from word_mcp.com_manager import get_word_app
        from word_mcp.tools.lifecycle import word_create

        self.app = get_word_app(visible=False)
        word_create()
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
        from word_mcp.tools.layout import word_set_page_setup

        result = word_set_page_setup(
            orientation="landscape",
            page_size="A4",
            margins={"top": 72, "bottom": 72, "left": 90, "right": 90},
        )
        assert "已设置" in result

    def test_set_header_footer(self):
        """word_set_header_footer: add header with page number."""
        from word_mcp.tools.layout import word_set_header_footer

        result = word_set_header_footer(
            type="header",
            text="Chapter 1: Introduction",
            include_pagenum=True,
        )
        assert "已设置" in result
