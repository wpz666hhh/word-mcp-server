"""L2: High-frequency atomic Word operation tools."""

from mcp.server.fastmcp import FastMCP

from ..com_manager import get_word_app
from ..utils import format_error

# Word enumeration constants
wdAlignParagraphCenter = 1
wdAlignParagraphLeft = 0
wdAlignParagraphRight = 2
wdAlignParagraphJustify = 3

ALIGNMENT_MAP = {
    "left": wdAlignParagraphLeft,
    "center": wdAlignParagraphCenter,
    "right": wdAlignParagraphRight,
    "justify": wdAlignParagraphJustify,
}


def _resolve_range(app, doc, range_spec: str):
    """Resolve a range specification string to a Word Range object.

    Args:
        app: Word Application object.
        doc: Word Document object.
        range_spec: "selection" | "start" | "end" | "all"

    Returns:
        A Word Range object.
    """
    if range_spec == "start":
        return doc.Range(0, 0)
    elif range_spec == "end":
        end = doc.Content.End - 1
        return doc.Range(end, end)
    elif range_spec == "all":
        return doc.Content
    else:  # "selection" or default
        return app.Selection.Range


async def word_insert_text(
    text: str,
    position: str = "selection",
) -> str:
    """在文档中插入文字。

    Args:
        text: 要插入的文字内容。
        position: 插入位置 — "selection"(光标处), "start"(开头), "end"(末尾)
    """
    try:
        app = get_word_app()
        doc = app.ActiveDocument
        rng = _resolve_range(app, doc, position)
        rng.Text = text
        return f"已插入 {len(text)} 字"
    except Exception as e:
        return format_error("插入文字", e)


async def word_format_text(
    range: str = "selection",
    font_name: str | None = None,
    font_size: float | None = None,
    bold: bool | None = None,
    italic: bool | None = None,
    underline: bool | None = None,
    color: str | None = None,
) -> str:
    """设置文字格式（字体、大小、加粗、颜色等）。

    Args:
        range: 应用范围 — "selection"(选中), "all"(全文)
        font_name: 字体名称，如 "微软雅黑"、"Arial"
        font_size: 字号（磅），如 12、16
        bold: 是否加粗
        italic: 是否斜体
        underline: 是否下划线
        color: 字体颜色，RGB十六进制如 "FF0000"（红色）
    """
    try:
        app = get_word_app()
        doc = app.ActiveDocument
        rng = _resolve_range(app, doc, range)

        changes = []
        font = rng.Font
        if font_name is not None:
            font.Name = font_name
            changes.append(f"字体={font_name}")
        if font_size is not None:
            font.Size = font_size
            changes.append(f"字号={font_size}pt")
        if bold is not None:
            font.Bold = bold
            changes.append(f"加粗={'是' if bold else '否'}")
        if italic is not None:
            font.Italic = italic
            changes.append(f"斜体={'是' if italic else '否'}")
        if underline is not None:
            font.Underline = 1 if underline else 0
            changes.append(f"下划线={'是' if underline else '否'}")
        if color is not None:
            font.Color = int(color, 16)
            changes.append(f"颜色=#{color}")

        if changes:
            return f"已设置: {', '.join(changes)}"
        return "未指定任何格式参数"
    except Exception as e:
        return format_error("设置文字格式", e)


async def word_format_paragraph(
    range: str = "selection",
    alignment: str | None = None,
    line_spacing: float | None = None,
    first_line_indent: float | None = None,
    space_before: float | None = None,
    space_after: float | None = None,
) -> str:
    """设置段落格式（对齐、行距、缩进等）。

    Args:
        range: 应用范围 — "selection", "all"
        alignment: 对齐方式 — "left", "center", "right", "justify"
        line_spacing: 行距倍数，如 1.5、2.0
        first_line_indent: 首行缩进（磅），如 24 表示两个字符
        space_before: 段前间距（磅）
        space_after: 段后间距（磅）
    """
    try:
        app = get_word_app()
        doc = app.ActiveDocument
        rng = _resolve_range(app, doc, range)

        changes = []
        pf = rng.ParagraphFormat

        if alignment is not None:
            if alignment in ALIGNMENT_MAP:
                pf.Alignment = ALIGNMENT_MAP[alignment]
                changes.append(f"对齐={alignment}")

        if line_spacing is not None:
            pf.LineSpacingRule = 0  # wdLineSpaceMultiple
            pf.LineSpacing = line_spacing
            changes.append(f"行距={line_spacing}x")

        if first_line_indent is not None:
            pf.FirstLineIndent = first_line_indent
            changes.append(f"首行缩进={first_line_indent}pt")

        if space_before is not None:
            pf.SpaceBefore = space_before
            changes.append(f"段前={space_before}pt")

        if space_after is not None:
            pf.SpaceAfter = space_after
            changes.append(f"段后={space_after}pt")

        if changes:
            return f"已设置段落格式: {', '.join(changes)}"
        return "未指定任何段落格式参数"
    except Exception as e:
        return format_error("设置段落格式", e)


async def word_get_content(
    range: str = "all",
    format: str = "text",
) -> str:
    """读取文档内容。

    Args:
        range: 读取范围 — "all"(全文), "selection"(选中)
        format: 返回格式 — "text"(纯文本), "html"(带格式)
    """
    try:
        app = get_word_app()
        doc = app.ActiveDocument

        if range == "selection":
            rng = app.Selection.Range
            label = "选中内容"
        else:
            rng = doc.Content
            label = "全文"

        if format == "html":
            try:
                content = rng.FormattedText.WordOpenXML
            except Exception:
                content = rng.Text
        else:
            content = rng.Text

        if not content.strip():
            return f"[文档为空] ({label})"
        return content
    except Exception as e:
        return format_error("读取内容", e)


async def word_get_selection() -> str:
    """获取当前用户在 Word 中选中的文字。"""
    try:
        app = get_word_app()
        sel = app.Selection
        text = sel.Text
        if not text.strip():
            return "当前未选中任何文字"
        return f"选中文字: {text}"
    except Exception as e:
        return format_error("获取选中内容", e)


async def word_select(range: str) -> str:
    """选中文档中的指定区域（用户可在 Word 窗口中看到高亮）。

    Args:
        range: "start"(跳到开头), "end"(跳到末尾), "all"(全选)
    """
    try:
        app = get_word_app()
        doc = app.ActiveDocument
        rng = _resolve_range(app, doc, range)
        rng.Select()
        return f"已选中: {range}"
    except Exception as e:
        return format_error("选中区域", e)


async def word_apply_style(
    range: str = "selection",
    style_name: str = "Normal",
) -> str:
    """对指定区域应用 Word 内置样式。

    Args:
        range: 应用范围
        style_name: 样式名，如 "Normal", "Heading 1", "Heading 2", "Title"
    """
    try:
        app = get_word_app()
        doc = app.ActiveDocument
        rng = _resolve_range(app, doc, range)

        try:
            style = doc.Styles(style_name)
        except Exception:
            available = []
            for i in range(1, min(doc.Styles.Count + 1, 30)):
                try:
                    s = doc.Styles(i)
                    if s.NameLocal:
                        available.append(s.NameLocal)
                except Exception:
                    pass
            return f"样式 '{style_name}' 不存在。可用样式: {', '.join(available[:15])}"

        rng.Style = style
        return f"已应用样式: {style_name}"
    except Exception as e:
        return format_error("应用样式", e)


async def word_auto_numbering(
    range: str = "selection",
    type: str = "bullet",
) -> str:
    """为指定段落设置编号或项目符号。

    Args:
        range: 应用范围
        type: "bullet"(项目符号), "number"(数字编号)
    """
    try:
        app = get_word_app()
        doc = app.ActiveDocument
        rng = _resolve_range(app, doc, range)

        if type == "bullet":
            rng.ListFormat.ApplyBulletDefault()
        elif type == "number":
            rng.ListFormat.ApplyNumberDefault()
        else:
            return f"不支持的编号类型: {type}，请使用 'bullet' 或 'number'"

        label = "项目符号" if type == "bullet" else "数字编号"
        return f"已设置: {label}"
    except Exception as e:
        return format_error("设置编号", e)


def register_operation_tools(mcp: FastMCP):
    """Register all L2 atomic operation tools on the MCP server."""
    mcp.tool()(word_insert_text)
    mcp.tool()(word_format_text)
    mcp.tool()(word_format_paragraph)
    mcp.tool()(word_get_content)
    mcp.tool()(word_get_selection)
    mcp.tool()(word_select)
    mcp.tool()(word_apply_style)
    mcp.tool()(word_auto_numbering)
