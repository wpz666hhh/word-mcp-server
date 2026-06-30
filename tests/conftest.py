"""Pytest fixtures for word-mcp-server tests."""
import pytest
from word_mcp import com_manager


@pytest.fixture(autouse=True)
def reset_com_manager():
    """Reset COM manager singleton state before and after each test.

    This ensures tests are independently runnable regardless of ordering
    and that Word instances don't leak after interrupted tests.
    """
    _cleanup_com_manager()
    yield
    _cleanup_com_manager()


def _cleanup_com_manager():
    """Release any held COM resources without killing the Word process."""
    import pythoncom

    if com_manager._word_app is not None:
        try:
            # Close all open documents without saving (gentler than Quit)
            while com_manager._word_app.Documents.Count > 0:
                com_manager._word_app.Documents(1).Close(SaveChanges=False)
        except Exception:
            pass
        com_manager._word_app = None

    if com_manager._coinit_done:
        try:
            pythoncom.CoUninitialize()
        except Exception:
            pass
        com_manager._coinit_done = False
