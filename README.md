# Word MCP Server

**通过 COM 自动化让 Claude Agent 操控 Microsoft Word，实时可见。**

**An MCP server that empowers Claude Agent to control Microsoft Word via COM automation — real-time, visible, and interactive.**

---

## 简介 / Overview

> **中文**
> Word MCP Server 是一个基于 Python 的 MCP 服务器，通过 pywin32 COM 自动化连接 Microsoft Word。Agent 可以像人类一样操作 Word——打开文档、编辑文字、格式化表格、插入图片、设置页眉页脚——且 Word 窗口可见，用户实时看到每一步操作。
>
> **English**
> Word MCP Server is a Python-based MCP server that connects to Microsoft Word via pywin32 COM automation. Your agent can open documents, edit text, format tables, insert images, set headers & footers — all while keeping the Word window visible so you can watch it work.

## 系统要求 / Requirements

- Windows 10 或更高 / Windows 10 or later
- Microsoft Word 2016 或更高（桌面版）/ Microsoft Word 2016 or later (desktop)
- Python 3.10+

## 安装 / Installation

```bash
cd D:\myskill_mcp\mcp\word-mcp-server
pip install -e .
```

## Claude Code 配置 / Configuration

在 `settings.local.json` 中添加以下内容 / Add to your `settings.local.json`:

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

重启 Claude Code 即可生效 / Restart Claude Code to activate.

## 可用工具 / Available Tools

### L1 — 文档生命周期 / Document Lifecycle

| 工具 Tool | 说明 Description |
|-----------|------------------|
| `word_open` | 打开文档 / Open a .docx/.doc file |
| `word_create` | 新建文档 / Create a new document |
| `word_save` | 保存或另存 / Save or Save As |
| `word_save_as_pdf` | 导出为 PDF / Export to PDF |
| `word_close` | 关闭文档 / Close the active document |
| `word_get_active_document` | 获取文档信息 / Get document info (pages, words, etc.) |
| `word_list_documents` | 列出所有打开的文档 / List all open documents |
| `word_activate_document` | 切换到指定文档 / Activate a specific document |

### L2 — 原子操作 / Atomic Operations

| 工具 Tool | 说明 Description |
|-----------|------------------|
| `word_insert_text` | 插入文字 / Insert text |
| `word_format_text` | 设置字体、大小、颜色等 / Set font, size, color, etc. |
| `word_format_paragraph` | 设置对齐、行距、缩进 / Set alignment, spacing, indent |
| `word_insert_table` | 插入表格（可带数据） / Insert a table (with optional data) |
| `word_format_table` | 格式化表格（样式、字体、表头） / Format table (style, font, header) |
| `word_insert_image` | 插入图片 / Insert an image file |
| `word_insert_page_break` | 插入分页符 / Insert a page break |
| `word_set_page_setup` | 设置页面布局 / Set margins, orientation, paper size |
| `word_set_header_footer` | 设置页眉/页脚（多节、页码样式、对齐） / Set header/footer (multi-section, page number style, alignment) |
| `word_find_replace` | 查找替换 / Find and replace text |
| `word_get_content` | 读取文档内容 / Read document text |
| `word_get_selection` | 获取用户选中文字 / Get user-selected text |
| `word_select` | 选中区域（可见高亮） / Select a range (visible highlight) |
| `word_apply_style` | 应用内置样式 / Apply Word built-in style |
| `word_auto_numbering` | 设置编号或项目符号 / Bullet or numbered list |

## 使用示例 / Example

> **中文**
> 告诉 Claude："帮我创建一份简历，包含个人信息表和工作经历"——Agent 会自动组合以下工具完成：
>
> **English**
> Tell Claude: "Create a resume with a personal info table and work experience" — the agent will orchestrate these tools:

```
1. word_create()                        → 新建文档 / new document
2. word_insert_text("张三的简历")       → 标题 / title
3. word_format_text(bold=true, size=22) → 标题加粗 / bold title
4. word_insert_table(5, 2, [...])       → 个人信息表 / info table
5. word_format_table(header_row=true)   → 表格样式 / style table
6. word_insert_text("工作经历")         → 章节标题 / section heading
7. word_apply_style("Heading 2")        → 标题样式 / heading style
8. word_save("简历.docx")               → 保存 / save
```

## 关于 VBA / About VBA

> **中文**
> VBA 执行器需要 Word 信任中心手动启用「信任对 VBA 项目对象模型的访问」，出于安全考虑已默认移除。绝大部分功能已通过 L1 + L2 工具覆盖，如需 VBA 可参阅文档恢复。
>
> **English**
> The VBA executor was removed by default for security — it requires manually enabling "Trust access to the VBA project object model" in Word's Trust Center. Most use cases are covered by L1 + L2 tools. VBA can be restored if needed.

---

**Made with ❤️ for Windows + Claude Code**
