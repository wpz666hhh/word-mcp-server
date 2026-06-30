# Word MCP Server

[**中文**](README.md) | [**English**](#)

---

**Claude Agent controlling Microsoft Word via COM automation — real-time and visible.**

## Overview

Word MCP Server is a Python-based MCP server that connects to Microsoft Word via pywin32 COM automation. Your agent can open documents, edit text, format tables, insert images, set headers & footers — all while keeping the Word window visible so you can watch it work.

## Requirements

- Windows 10 or later
- Microsoft Word 2016 or later (desktop version)
- Python 3.10+

## Installation

Choose one of the following methods.

### Method 1: pip install from GitHub (recommended)

No need to clone — install directly:

```bash
pip install git+https://github.com/wpz666hhh/word-mcp-server.git
```

Then configure Claude Code (see below).

### Method 2: uv install (faster)

If you use [uv](https://docs.astral.sh/uv/):

```bash
uv pip install git+https://github.com/wpz666hhh/word-mcp-server.git
```

### Method 3: Clone + install

```bash
git clone https://github.com/wpz666hhh/word-mcp-server.git
cd word-mcp-server
pip install -e .
```

Editable mode — useful if you want to modify the code.

### Method 4: One-liner install script

Run this in PowerShell to install & configure automatically:

```powershell
iex "& { $(curl -fsL https://raw.githubusercontent.com/wpz666hhh/word-mcp-server/main/scripts/install.ps1) }"
```

> The script will:
> 1. Check your Python environment
> 2. Install via pip from GitHub
> 3. Write the MCP entry into Claude Code's `settings.local.json`
> 4. Prompt you to restart Claude Code when done

## Claude Code Configuration

Add to your `settings.local.json` (`C:\Users\<your-username>\.claude\` or `<project>\.claude\`):

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

Restart Claude Code to activate.

> **Tip:** This repo includes a `.mcp.json` file at the root. If you cloned the repo, Claude Code may auto-discover this MCP server. See [MCP auto-discovery docs](https://docs.anthropic.com/en/docs/agents/claude-code/mcp#auto-discovery).

## Available Tools

### L1 — Document Lifecycle

| Tool | Description |
|------|-------------|
| `word_open` | Open a .docx/.doc file |
| `word_create` | Create a new document |
| `word_save` | Save or Save As |
| `word_save_as_pdf` | Export to PDF |
| `word_close` | Close the active document |
| `word_get_active_document` | Get document info (pages, words, etc.) |
| `word_list_documents` | List all open documents |
| `word_activate_document` | Switch to a specific document |

### L2 — Atomic Operations

| Tool | Description |
|------|-------------|
| `word_insert_text` | Insert text |
| `word_format_text` | Set font, size, color, etc. |
| `word_format_paragraph` | Set alignment, spacing, indent |
| `word_insert_table` | Insert a table (with optional data) |
| `word_format_table` | Format table (style, font, header) |
| `word_insert_image` | Insert an image file |
| `word_insert_page_break` | Insert a page break |
| `word_set_page_setup` | Set margins, orientation, paper size |
| `word_set_header_footer` | Set header/footer (multi-section, page number style, alignment) |
| `word_find_replace` | Find and replace text |
| `word_get_content` | Read document text |
| `word_get_selection` | Get user-selected text |
| `word_select` | Select a range (visible highlight) |
| `word_apply_style` | Apply Word built-in style |
| `word_auto_numbering` | Bullet or numbered list |

## Example Usage

Tell Claude: *"Create a resume with a personal info table and work experience"* — the agent will orchestrate these tools:

```
 1. word_create()                         → new document
 2. word_insert_text("张三的简历")        → title
 3. word_format_text(bold=true, size=22)  → bold title
 4. word_insert_table(5, 2, [...])        → personal info table
 5. word_format_table(header_row=true)    → style table
 6. word_insert_text("工作经历")          → section heading
 7. word_apply_style("Heading 2")         → heading style
 8. word_save("简历.docx")                → save
```

The Word window stays visible throughout, so you can watch every change in real time.

## About VBA

The VBA executor was removed by default for security — it requires manually enabling "Trust access to the VBA project object model" in Word's Trust Center. Most use cases are covered by L1 + L2 tools. The VBA module can be restored if needed.

---

**Made with ❤️ for Windows + Claude Code**
