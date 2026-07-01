"""COM connection lifecycle management for Microsoft Word.

Maintains a singleton Word.Application COM object. On first call it either
attaches to an existing Word process (GetObject) or creates a new one
(Dispatch). Subsequent calls return the cached instance. If the connection
is lost (user closed Word), the next call transparently reconnects.
"""

import atexit
import logging
from typing import Optional

import pythoncom
import win32com.client

logger = logging.getLogger(__name__)

_word_app: Optional[object] = None
_coinit_done: bool = False
_created_by_us: bool = False


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
    global _word_app, _coinit_done, _created_by_us

    # If we already have a reference, test if it's still alive
    if _word_app is not None:
        try:
            _ = _word_app.Name
            
            # 如果缓存的是我们自己创建的空实例（无文档），且检测到用户在外部打开了文档，
            # 我们主动释放并退出这个空实例，以便后续重新 GetObject 连接到用户的活动 Word 实例。
            if _created_by_us:
                try:
                    if _word_app.Documents.Count == 0:
                        import win32gui
                        hwnd = win32gui.FindWindow('OPUSAPP', None)
                        if hwnd:
                            title = win32gui.GetWindowText(hwnd)
                            # 如果窗口标题包含 " - Word" 且不是默认的空窗口（如 "Microsoft Word" 或 "Word"），
                            # 说明用户有打开的文档。而我们自己的实例 Documents.Count == 0，说明这不是我们的实例。
                            if title and " - Word" in title and not title.startswith("Microsoft Word") and title != "Word":
                                logger.info(f"Detected user Word window '{title}' but cached instance is empty. Re-attaching...")
                                try:
                                    _word_app.Quit(0)
                                except Exception:
                                    pass
                                _word_app = None
                                _created_by_us = False
                except Exception as ex:
                    logger.debug(f"Error checking active window: {ex}")

            if _word_app is not None:
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
        _created_by_us = False
        logger.info("Attached to existing Word process")
    except Exception:
        try:
            _word_app = win32com.client.Dispatch("Word.Application")
            _created_by_us = True
            logger.info("Created new Word process")
        except Exception as e:
            raise RuntimeError(
                "Failed to start Microsoft Word. Please verify Word is installed.\n"
                f"Original error: {e}"
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

@atexit.register
def _cleanup_word():
    """Clean up Word process on exit if we created it."""
    global _word_app, _created_by_us
    if _created_by_us and _word_app is not None:
        try:
            # We created it, so quit to avoid zombie WINWORD.EXE
            # Pass SaveChanges=False (wdDoNotSaveChanges=0)
            _word_app.Quit(0)
            logger.info("Cleaned up Word process on exit")
        except Exception:
            pass
    release_word()
