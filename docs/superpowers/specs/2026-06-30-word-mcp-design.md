# Word MCP Server — 设计规格说明书

**日期：** 2026-06-30  
**状态：** 待审核  
**类型：** MCP 服务器  

---

## 1. 目标

让 Claude Agent 具备通过 COM 自动化操控 Microsoft Word 的能力，覆盖 Word 绝大部分功能。用户可实时查看 Word 窗口中 agent 的操作过程。

## 2. 整体架构

```
┌──────────────────────────────────────────────┐
│                 Claude Agent                  │
│  (自然语言 → 智能组合 tool calls)              │
└──────────────┬───────────────────────────────┘
               │ MCP 协议 (stdio)
┌──────────────▼───────────────────────────────┐
│            Word MCP Server (Python)           │
│                                               │
│  ┌─────────┐  ┌──────────────┐  ┌─────────┐ │
│  │ L1 文档  │  │ L2 高频操作   │  │ L3 VBA  │ │
│  │ 生命周期 │  │ 原子工具      │  │ 执行器  │ │
│  │ (6个)    │  │ (15个)       │  │ (1个)   │ │
│  └─────────┘  └──────────────┘  └─────────┘ │
│                                               │
│  COM 封装层 (win32com.client)                  │
│  连接管理、错误处理、可见性控制                │
└──────────────┬───────────────────────────────┘
               │ COM Automation
┌──────────────▼───────────────────────────────┐
│          Microsoft Word 进程                   │
│     (窗口可见，用户实时看到操作过程)            │
└──────────────────────────────────────────────┘
```

**核心决策：**
- 采用**混合 MCP** 架构：文档级 tool + 高频原子 tool + VBA 兜底
- Python + pywin32 实现，COM 自动化最成熟的生态
- stdio 传输，无需网络端口
- Word 启动时强制 `Visible=True`
- 选择 MCP 而非 Skill：需要类型化工具定义 + JSON Schema 参数约束

## 3. 三层工具清单

### L1 — 文档生命周期（6 个 tool）

| Tool | 核心参数 | 说明 |
|------|------|------|
| `word_open` | `file_path: str` | 打开已有文档，Word 窗口可见 |
| `word_create` | `template_path: str?` | 新建文档，可选模板路径 |
| `word_save` | `file_path: str?` | 保存当前文档（可选另存路径） |
| `word_save_as_pdf` | `output_path: str`, `page_range: str?` | 导出为 PDF |
| `word_close` | `save_before_close: bool?` | 关闭活动文档 |
| `word_get_active_document` | — | 返回当前文档路径、页数等信息 |

### L2 — 高频原子操作（15 个 tool）

| Tool | 核心参数 | 说明 |
|------|------|------|
| `word_insert_text` | `text: str`, `position: str?` | 在指定位置或光标处插入文字 |
| `word_format_text` | `range`, `font_name?`, `font_size?`, `bold?`, `italic?`, `color?`, `underline?` | 设置文字格式 |
| `word_format_paragraph` | `range`, `alignment?`, `line_spacing?`, `first_line_indent?`, `space_before?`, `space_after?` | 设置段落格式 |
| `word_insert_table` | `rows: int`, `cols: int`, `data: str[][]?`, `position: str?` | 插入表格，可选填充数据 |
| `word_format_table` | `table_index: int`, `style?`, `header_row: bool?`, `auto_fit: bool?` | 设置表格样式 |
| `word_insert_image` | `image_path: str`, `position: str?`, `width?`, `height?` | 插入图片 |
| `word_insert_page_break` | `position: str?` | 插入分页符 |
| `word_set_page_setup` | `margins: object?`, `orientation: str?`, `page_size: str?` | 页面边距/方向/尺寸 |
| `word_set_header_footer` | `type: str`, `text: str`, `include_pagenum: bool?` | 设置页眉/页脚 |
| `word_find_replace` | `find_text: str`, `replace_text: str`, `match_case: bool?` | 全文查找替换 |
| `word_get_content` | `range: str?`, `format: str?` | 读取文档内容（纯文本/HTML） |
| `word_get_selection` | — | 获取当前用户选中的文本 |
| `word_select` | `range: str` | 选中指定区域（用户在窗口中看到高亮） |
| `word_apply_style` | `range: str`, `style_name: str` | 应用 Word 内置样式 |
| `word_auto_numbering` | `range: str`, `type: str` | 设置编号或项目符号 |

### L3 — VBA 执行器（1 个 tool）

| Tool | 参数 | 说明 |
|------|------|------|
| `word_execute_vba` | `code: str` | 在 Word 进程中执行 VBA 代码，返回结果或错误详情 |

VBA 执行器是终极兜底 — 任何 L2 未覆盖的操作，agent 均可生成 VBA 来完成。不持久化 VBA 代码，不修改 Normal.dotm。

## 4. COM 连接管理

### 4.1 进程生命周期

```
首次调用 tool
  │
  ▼
┌──────────────┐    存在
│ 检测运行中的  │──────────► GetObject() 获取已有进程
│ Word 进程？   │
└──────┬───────┘    不存在
       │
       ▼
  Dispatch("Word.Application")
  app.Visible = True         ← 强制可见
  app.DisplayAlerts = False  ← 抑制弹窗
  │
  ▼
  执行 tool 操作
  │
  ▼
  保持 COM 引用，不关闭 Word
  （复用同一进程，直到会话结束或用户手动关闭）
```

**关键规则：**
- Word 进程永不自动关闭。用户想关自己关
- `word_close` 关闭的是**文档**，不是 Word 进程
- 用户关闭 Word 后，下次 tool 调用自动 `Dispatch` 重新打开
- 延迟初始化：agent 启动时不开 Word，首次 tool 调用才打开

### 4.2 错误处理

| 场景 | 处理方式 |
|------|------|
| Word 未安装 | 启动时检测注册表，返回明确错误 + 安装指引（"请安装 Microsoft Word"） |
| 文档被其他进程锁定 | 返回可读错误，提示以只读模式打开 |
| VBA 执行错误 | 捕获 COM 异常，返回错误行号 + 消息原文 |
| COM 超时（RPC_E_CALL_REJECTED） | 重试 1 次（间隔 500ms），失败则释放连接重建 |
| 保存路径不存在 | 自动 `mkdir -p` 创建目录后保存 |

## 5. MCP 配置

### 5.1 配置方式

在 Claude Code 的 `settings.local.json` 中添加：

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

### 5.2 安装步骤（用户侧）

```bash
cd word-mcp-server
pip install -e .          # 安装依赖：pywin32, mcp[cli]
# 在 Claude Code 配置中添加上述 JSON
# 重启 Claude Code 生效
```

### 5.3 依赖

| 包 | 最低版本 | 用途 |
|----|------|------|
| `pywin32` | 306 | COM 自动化核心 |
| `mcp` | 1.x | MCP 协议 SDK，`@server.tool()` 装饰器 |
| `pydantic` | 2.x | （mcp 间接依赖）参数类型与 Schema 生成 |

## 6. 项目结构

```
word-mcp-server/
├── pyproject.toml
├── src/
│   └── word_mcp/
│       ├── __init__.py
│       ├── server.py          # MCP 入口，注册所有 tool
│       ├── com_manager.py     # COM 连接生命周期管理
│       ├── tools/
│       │   ├── __init__.py
│       │   ├── lifecycle.py   # L1: word_open/create/save/close/pdf/active
│       │   ├── operations.py  # L2: 15 个原子操作 tool
│       │   └── vba.py         # L3: word_execute_vba
│       └── utils.py           # 共享辅助函数
├── README.md
└── .gitignore
```

## 7. 开发阶段

| 阶段 | 内容 | 预估 | 验证标准 |
|------|------|------|------|
| Phase 1: 骨架 | `com_manager.py` + `server.py` 注册 1 个 tool（`word_get_active_document`） | 30 min | Claude Code 中调 tool，Word 窗口弹出 |
| Phase 2: L1 | 6 个文档生命周期 tool | 1 h | 创建→写→保存→导出 PDF→关闭，全链路通过 |
| Phase 3: L2 | 15 个原子操作 tool（分 3 批） | 2-3 h | 每批完成后用脚本验证 |
| Phase 4: L3 | VBA 执行器 | 30 min | 执行 `Selection.TypeText "Hello"` 等验证 |
| Phase 5: 端到端 | 自然语言场景测试 | 30 min | "创建一份简历文档" — agent 自主组合 tool 完成 |

## 8. 注意事项

- VBA 执行器不持久化：不创建宏模块，不修改 Normal.dotm
- 所有路径参数统一使用绝对路径，避免 Word 工作目录歧义
- COM 对象操作后及时 Release，避免内存泄漏
- 文件编码统一 UTF-8，支持中文路径和内容
- 这是 Windows 专属项目（依赖 COM），不支持 macOS/Linux
- Word 版本要求：Microsoft Word 2016 或更高（COM 接口兼容）

## 9. 典型使用场景

```
用户: "帮我创建一份工作简历，包含个人信息表格和两段工作经历"

Agent:
  1. word_create()                         → 新建文档
  2. word_insert_text("张三的简历", ...)    → 标题
  3. word_format_text(..., bold=true, size=22) → 标题加粗
  4. word_insert_table(5, 2, [...个人信息数据])  → 个人信息表
  5. word_format_table(1, header_row=true) → 表格样式
  6. word_insert_text("工作经历", ...)      → 章节标题
  7. word_apply_style(..., "Heading 2")    → 标题样式
  8. word_insert_text("2020-2024 ...")     → 工作经历内容
  9. word_insert_page_break()               → 分页
 10. word_save("C:/Users/.../简历.docx")    → 保存

整个过程 Word 窗口可见，用户实时看到变化
```
