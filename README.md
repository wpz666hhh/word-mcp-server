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
1. word_create()                — new document
2. word_insert_text("张三的简历") — title
3. word_format_text(bold=true)  — format title
4. word_insert_table(rows, cols, data) — info table
5. word_format_table(header_row=true) — style table
6. word_insert_text("工作经历") — section heading
7. word_apply_style("Heading 2") — heading style
8. word_save("简历.docx") — save

You watch it happen in the Word window.
```
