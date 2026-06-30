# Word MCP Server

[**中文**](#中文) | [**English**](#english)

---

<a id="中文"></a>

**通过 COM 自动化让 Claude Agent 操控 Microsoft Word，实时可见。**

## 简介

Word MCP Server 是一个基于 Python 的 MCP 服务器，通过 pywin32 COM 自动化连接 Microsoft Word。Agent 可以像人类一样操作 Word——打开文档、编辑文字、格式化表格、插入图片、设置页眉页脚——且 Word 窗口可见，用户实时看到每一步操作。

## 系统要求

- Windows 10 或更高
- Microsoft Word 2016 或更高（桌面版）
- Python 3.10+

## 安装

提供多种方式，任选其一。

### 方式一：pip 直接安装（推荐）

无需克隆仓库，一行命令安装：

```bash
pip install git+https://github.com/wpz666hhh/word-mcp-server.git
```

安装后配置 Claude Code（见下方配置说明）即可使用。

### 方式二：uv 安装（更快）

如果你使用 [uv](https://docs.astral.sh/uv/)：

```bash
uv pip install git+https://github.com/wpz666hhh/word-mcp-server.git
```

### 方式三：克隆后安装

```bash
git clone https://github.com/wpz666hhh/word-mcp-server.git
cd word-mcp-server
pip install -e .
```

开发模式，方便修改代码。

### 方式四：一键安装脚本

在 PowerShell 中运行以下命令自动完成安装 + 配置：

```powershell
# 安装 word-mcp-server 并自动写入 Claude Code 配置
iex "& { $(curl -fsL https://raw.githubusercontent.com/wpz666hhh/word-mcp-server/main/scripts/install.ps1) }"
```

> 安装脚本会自动：
> 1. 检测 Python 环境
> 2. 用 pip 从 GitHub 安装
> 3. 写入 Claude Code 的 `settings.local.json`
> 4. 完成后提示重启 Claude Code

## Claude Code 配置

安装后在 `settings.local.json`（位于 `C:\Users\<你的用户名>\.claude\` 或项目目录下的 `.claude\`）中添加：

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

重启 Claude Code 即可生效。

> **提示：** 本仓库自带 `.mcp.json` 配置文件，如果你克隆了仓库，Claude Code 可能能自动发现此 MCP 服务。见 [MCP 服务器自动发现](https://docs.anthropic.com/en/docs/agents/claude-code/mcp#auto-discovery)。

## 可用工具

### L1 — 文档生命周期

| 工具 | 说明 |
|------|------|
| `word_open` | 打开文档 |
| `word_create` | 新建文档 |
| `word_save` | 保存或另存 |
| `word_save_as_pdf` | 导出为 PDF |
| `word_close` | 关闭文档 |
| `word_get_active_document` | 获取文档信息（页数、字数等） |
| `word_list_documents` | 列出所有打开的文档 |
| `word_activate_document` | 切换到指定文档 |

### L2 — 原子操作

| 工具 | 说明 |
|------|------|
| `word_insert_text` | 插入文字 |
| `word_format_text` | 设置字体、大小、颜色等 |
| `word_format_paragraph` | 设置对齐、行距、缩进 |
| `word_insert_table` | 插入表格（可带数据） |
| `word_format_table` | 格式化表格（样式、字体、表头） |
| `word_insert_image` | 插入图片 |
| `word_insert_page_break` | 插入分页符 |
| `word_set_page_setup` | 设置页面布局（边距、方向、纸张） |
| `word_set_header_footer` | 设置页眉/页脚（多节、页码样式、对齐） |
| `word_find_replace` | 查找替换 |
| `word_get_content` | 读取文档内容 |
| `word_get_selection` | 获取用户选中文字 |
| `word_select` | 选中区域（可见高亮） |
| `word_apply_style` | 应用内置样式 |
| `word_auto_numbering` | 设置编号或项目符号 |

## 使用示例

告诉 Claude：“帮我创建一份简历，包含个人信息表和工作经历”——Agent 会自动组合以下工具完成：

```
 1. word_create()                         → 新建文档
 2. word_insert_text("张三的简历")        → 标题
 3. word_format_text(bold=true, size=22)  → 标题加粗
 4. word_insert_table(5, 2, [...])        → 个人信息表
 5. word_format_table(header_row=true)    → 表格样式
 6. word_insert_text("工作经历")          → 章节标题
 7. word_apply_style("Heading 2")         → 标题样式
 8. word_save("简历.docx")                → 保存
```

整个过程 Word 窗口可见，用户可实时看到每一步变化。

## 关于 VBA

VBA 执行器需要 Word 信任中心手动启用「信任对 VBA 项目对象模型的访问」，出于安全考虑已默认移除。绝大部分功能已通过 L1 + L2 工具覆盖，如需 VBA 可自行恢复对应模块。

---

<a id="english"></a>

# English

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
