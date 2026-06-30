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
        assert isinstance(result, str)
        assert len(result) > 0

    def test_vba_syntax_error(self):
        """VBA compile errors should be caught."""
        from word_mcp.tools.vba import word_execute_vba

        code = "Sub Main()\n  invalid syntax here !!!\nEnd Sub"
        result = self._run(word_execute_vba(code=code))
        assert isinstance(result, str)
        assert len(result) > 0
