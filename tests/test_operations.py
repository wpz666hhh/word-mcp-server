"""Tests for L2 atomic operation tools."""
import sys

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
        from word_mcp.tools.operations import word_insert_text, word_format_text

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
        from word_mcp.tools.operations import word_insert_text, word_format_paragraph

        self._run(word_insert_text(text="Centered paragraph"))
        result = self._run(word_format_paragraph(range="selection", alignment="center"))
        assert "已设置" in result

    def test_get_content(self):
        """word_get_content: read document text."""
        from word_mcp.tools.operations import word_insert_text, word_get_content

        self._run(word_insert_text(text="Read this back"))
        result = self._run(word_get_content())
        assert "Read this back" in result

    def test_get_selection(self):
        """word_get_selection: get user-selected text."""
        from word_mcp.tools.operations import word_get_selection

        result = self._run(word_get_selection())
        assert isinstance(result, str)

    def test_select_range(self):
        """word_select: select a range of text."""
        from word_mcp.tools.operations import word_insert_text, word_select

        self._run(word_insert_text(text="Selectable text here"))
        result = self._run(word_select(range="start"))
        assert "已选中" in result or "选中" in result

    def test_apply_style(self):
        """word_apply_style: apply a Word style."""
        from word_mcp.tools.operations import word_insert_text, word_apply_style

        self._run(word_insert_text(text="A heading"))
        result = self._run(word_apply_style(range="selection", style_name="Heading 1"))
        assert "已应用" in result or "样式" in result

    def test_auto_numbering_bullet(self):
        """word_auto_numbering: apply bullet list."""
        from word_mcp.tools.operations import word_insert_text, word_auto_numbering

        self._run(word_insert_text(text="Bullet item"))
        result = self._run(word_auto_numbering(range="selection", type="bullet"))
        assert "已设置" in result
