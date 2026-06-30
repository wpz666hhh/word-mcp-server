"""Tests for COM connection manager."""
import sys
import pytest


class TestGetWordApp:
    """Tests for get_word_app()."""

    def test_creates_word_instance(self):
        """get_word_app() should return a Word Application object."""
        if sys.platform != "win32":
            pytest.skip("COM tests only run on Windows")

        from word_mcp.com_manager import get_word_app

        app = get_word_app(visible=False)
        assert app is not None
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
