# CLAUDE.md

This file provides guidance to Claude Code (claude.ai/code) when working with code in this repository.

## Project Overview

Word MCP Server — an MCP server that controls Microsoft Word via pywin32 COM automation. Built with FastMCP (`mcp>=1.0.0`). All tools are synchronous functions (FastMCP wraps them transparently) and return Chinese-language status messages. Windows-only; requires Microsoft Word 2016+ and Python 3.10+.

## Build / Test / Run

```bash
# Install in editable mode
pip install -e .

# Run the server directly (stdio transport)
python -m word_mcp

# Run all tests (COM tests auto-skip on non-Windows)
pytest

# Run a single test file
pytest tests/test_operations.py -v

# Run a single test class
pytest tests/test_lifecycle.py::TestDocumentLifecycle -v

# Run with stdout live (debug hanging tests)
pytest -s
```

MCP config for `settings.local.json` / `.mcp.json`:
```json
{"mcpServers": {"word": {"command": "python", "args": ["-m", "word_mcp"], "type": "local"}}}
```
The repo ships with `.mcp.json` for auto-discovery in Claude Code.

## Architecture

### Three-layer structure split across 8 modules

All tools are registered in `server.py` — lifecycle tools via `register_lifecycle_tools(mcp)`, all others via a loop over `mcp.tool()`.

```
src/word_mcp/
├── server.py          — Entry: creates FastMCP, registers 29 tool functions, runs stdio
├── com_manager.py     — Singleton Word.Application via win32com (GetObject → Dispatch fallback)
├── utils.py           — Path helpers (normalize_path, ensure_dir), COM error-code → Chinese mapping
├── models.py          — TypedDict definitions (ParagraphLocator, HeadingLocator, BookmarkLocator, SearchLocator, Margins)
├── locator.py         — resolve_range(): maps locator dicts/strings to Word Range objects
└── tools/
    ├── __init__.py    — empty
    ├── lifecycle.py   — L1: 8 tools (open/create/save/save_as_pdf/close/get_active_doc/list/activate)
    ├── text.py        — L2: 12 tools (insert_text, format_text, format_paragraph, get_content, get_selection,
    │                    select, apply_style, auto_numbering, find_replace, format_text_by_find,
    │                    delete_target, replace_target)
    ├── tables.py      — L2: 2 tools (insert_table, format_table)
    ├── layout.py      — L2: 4 tools (insert_image, insert_page_break, set_page_setup, set_header_footer)
    ├── structure.py   — L2: 2 tools (get_document_structure, read_table)
    └── macro.py       — L3: 1 tool (word_execute_macro — batch Python execution via exec())
```

### L3: `macro.py` — batch execution for complex operations

`word_execute_macro` runs arbitrary Python code directly against the live Word COM objects (`app` and `doc` are injected as local variables). This is the escape hatch for operations too expensive to orchestrate one tool call at a time:

```python
word_execute_macro(python_code="""
for i in range(1, doc.Tables(1).Rows.Count + 1):
    doc.Tables(1).Rows(i).Height = 30
""")
```

Key details:
- Uses `exec()` with `ScreenUpdating=False` during execution, restoring it in a `finally` block
- Captures any `print()` output via `contextlib.redirect_stdout`
- Must NOT open new documents — only operate on the active `doc`
- Prefer Range over Selection for performance

### `com_manager.py` — singleton `_word_app`

- First call: `GetObject("Word.Application")` to attach to running Word → fallback `Dispatch("Word.Application")` to launch new
- Subsequent calls return cached reference with keepalive check (`_word_app.Name`)
- Lost connections trigger transparent reconnect
- `atexit` handler quits Word only if `_created_by_us` (no zombie processes)
- `release_word()` clears reference without killing Word (for test cleanup)
- Word started with `Visible=True` and `DisplayAlerts=0`

### `locator.py` + `models.py` — position resolution system

All L2 tools accept a `LocatorDef` parameter (`Union[...]` of TypedDicts + str):

| Value | Effect |
|-------|--------|
| `"selection"` (default) | Current cursor/selection |
| `"start"` | Document beginning |
| `"end"` | Document end |
| `"all"` | Whole document |
| `{"type": "paragraph", "index": 5}` | Nth paragraph (1-based) |
| `{"type": "heading", "text": "工作经历"}` | First heading containing text |
| `{"type": "search", "text": "关键词", "instance": 1}` | Nth occurrence via Find |
| `{"type": "bookmark", "name": "MyMark"}` | Named bookmark location |

### Color handling

RGB hex strings (e.g. `"FF0000"` or `"#FF0000"`) → Word's `R + G*256 + B*65536` format. Used consistently across `format_text`, `format_table`, `find_replace`, and `format_text_by_find`.

### Font-aware find/replace

`word_find_replace` and `word_format_text_by_find` use Word's `Find.Font` / `Replacement.Font` for filtering text by existing font properties and changing those properties during replacement.

## Key patterns

- **No async** — all tool functions are plain `def`, not `async def`. FastMCP's `mcp.tool()` wraps them.
- **Error handling** — every tool wraps in `try/except` and calls `format_error(operation, e)`, which maps known COM HRESULTs to Chinese messages using `_COM_ERROR_HINTS` dict. Falls through to the raw error string if no hint matches.
- **Registration** — `lifecycle.py` tools registered via `register_lifecycle_tools(mcp)`; all others (including `word_execute_macro`) registered via `mcp.tool()(tool_func)` in a loop in `server.py`.
- **Tests** — `conftest.py` auto-resets COM singleton before/after each test via `reset_com_manager()`. All tests `@pytest.mark.skipif(sys.platform != "win32")`. Test classes create their own `asyncio.new_event_loop()` (needed for FastMCP's async test setup) but call synchronous tool functions directly (not via `run_until_complete`). The `_run()` helper in test classes is unused and can be removed.
- **Test files** — 6 test files in `tests/`: `conftest.py`, `test_com_manager.py`, `test_lifecycle.py`, `test_operations.py`, `test_structure_locator.py`. Each test class has an `autouse` fixture that creates a new document and cleans it up (`Close(SaveChanges=False)`) after each test.
- **VBA removed** — requires Word trust center toggle; all functionality covered by L1+L2+L3 tools.

## Known issues

- `src/word_mcp/__init__.py` version string (`0.1.0`) is outdated — `pyproject.toml` has `0.2.0`. Keep them in sync.
- `operations.py` was split into `text.py` / `tables.py` / `layout.py` / `structure.py` — the old file was deleted, no import stubs remain.
- Test classes create `asyncio.new_event_loop()` and expose a `_run()` helper that wraps `loop.run_until_complete()`, but all tools are synchronous and tests call them directly — the loop object and `_run()` are vestigial.
