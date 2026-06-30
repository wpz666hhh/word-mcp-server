# Word MCP Server

[**中文**](#) | [**English**](README_EN.md)

---

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

> **提示：** 本仓库自带 `.mcp.json` 配置文件，克隆仓库后 Claude Code 会自动发现此 MCP 服务。详见 [Claude Code MCP 文档](https://docs.anthropic.com/en/docs/claude-code/mcp)。

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

告诉 Claude："帮我创建一份简历，包含个人信息表和工作经历"——Agent 会自动组合以下工具完成：

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

**Made with ❤️ for Windows + Claude Code**
