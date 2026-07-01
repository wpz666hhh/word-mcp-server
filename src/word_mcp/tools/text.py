"""Text formatting and manipulation operations for Microsoft Word."""

from ..com_manager import get_word_app
from ..utils import format_error
from ..locator import resolve_range, LocatorDef

wdAlignParagraphLeft = 0
wdAlignParagraphCenter = 1
wdAlignParagraphRight = 2
wdAlignParagraphJustify = 3

ALIGNMENT_MAP = {
    "left": wdAlignParagraphLeft,
    "center": wdAlignParagraphCenter,
    "right": wdAlignParagraphRight,
    "justify": wdAlignParagraphJustify,
}

def word_insert_text(
    text: str,
    position: LocatorDef = "selection",
) -> str:
    """在文档中插入文字。

    Args:
        text: 要插入的文字内容。
        position: 插入位置 — "selection"(光标处), "start"(开头), "end"(末尾)
    """
    try:
        app = get_word_app()
        doc = app.ActiveDocument
        rng = resolve_range(app, doc, position)
        rng.Text = text
        return f"已插入 {len(text)} 字"
    except Exception as e:
        return format_error("插入文字", e)

def word_format_text(
    range_spec: LocatorDef = "selection",
    font_name: str | None = None,
    font_size: float | None = None,
    bold: bool | None = None,
    italic: bool | None = None,
    underline: bool | None = None,
    color: str | None = None,
    table_index: int | None = None,
) -> str:
    """设置文字格式（字体、大小、加粗、颜色等）。

    Args:
        range_spec: 应用范围 — "selection"(选中), "all"(全文)。
                    当指定 table_index 时此参数被忽略。
        font_name: 字体名称，如 "微软雅黑"、"Arial"
        font_size: 字号（磅），如 12、16
        bold: 是否加粗
        italic: 是否斜体
        underline: 是否下划线
        color: 字体颜色，RGB十六进制如 "FF0000"（红色）
        table_index: 可选，指定表格序号（1=第一个表格），仅格式化该表格文字。
    """
    try:
        app = get_word_app()
        doc = app.ActiveDocument

        if table_index is not None:
            table = doc.Tables(table_index)
            rng = table.Range
            label = f"表格{table_index}"
        else:
            rng = resolve_range(app, doc, range_spec)
            label = range_spec

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
            hex_str = color.lstrip("#")
            r = int(hex_str[0:2], 16)
            g = int(hex_str[2:4], 16)
            b = int(hex_str[4:6], 16)
            font.Color = r + g * 256 + b * 65536
            changes.append(f"颜色=#{color}")

        range_label = f"[{label}] " if table_index is not None else ""
        if changes:
            return f"已设置{range_label}: {', '.join(changes)}"
        return "未指定任何格式参数"
    except Exception as e:
        return format_error("设置文字格式", e)

def word_format_paragraph(
    range_spec: LocatorDef = "selection",
    alignment: str | None = None,
    line_spacing: float | None = None,
    first_line_indent: float | None = None,
    space_before: float | None = None,
    space_after: float | None = None,
) -> str:
    """设置段落格式（对齐、行距、缩进等）。

    Args:
        range_spec: 应用范围 — "selection", "all"
        alignment: 对齐方式 — "left", "center", "right", "justify"
        line_spacing: 行距倍数，如 1.5、2.0
        first_line_indent: 首行缩进（磅），如 24 表示两个字符
        space_before: 段前间距（磅）
        space_after: 段后间距（磅）
    """
    try:
        app = get_word_app()
        doc = app.ActiveDocument
        rng = resolve_range(app, doc, range_spec)

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

def word_get_content(
    range_spec: LocatorDef = "all",
    format: str = "text",
) -> str:
    """读取文档内容。

    Args:
        range_spec: 读取范围 — "all"(全文), "selection"(选中)
        format: 返回格式 — "text"(纯文本), "html"(带格式)
    """
    try:
        app = get_word_app()
        doc = app.ActiveDocument

        if range_spec == "selection":
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

def word_get_selection() -> str:
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

def word_select(range_spec: str) -> str:
    """选中文档中的指定区域（用户可在 Word 窗口中看到高亮）。

    Args:
        range_spec: "start"(跳到开头), "end"(跳到末尾), "all"(全选)
    """
    try:
        app = get_word_app()
        doc = app.ActiveDocument
        rng = resolve_range(app, doc, range_spec)
        rng.Select()
        return f"已选中: {range_spec}"
    except Exception as e:
        return format_error("选中区域", e)

def word_apply_style(
    range_spec: LocatorDef = "selection",
    style_name: str = "Normal",
) -> str:
    """对指定区域应用 Word 内置样式。

    Args:
        range_spec: 应用范围
        style_name: 样式名，如 "Normal", "Heading 1", "Heading 2", "Title"
    """
    try:
        app = get_word_app()
        doc = app.ActiveDocument
        rng = resolve_range(app, doc, range_spec)

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

def word_auto_numbering(
    range_spec: LocatorDef = "selection",
    type: str = "bullet",
) -> str:
    """为指定段落设置编号或项目符号。

    Args:
        range_spec: 应用范围
        type: "bullet"(项目符号), "number"(数字编号)
    """
    try:
        app = get_word_app()
        doc = app.ActiveDocument
        rng = resolve_range(app, doc, range_spec)

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

def word_find_replace(
    find_text: str,
    replace_text: str,
    match_case: bool = False,
    match_wildcards: bool = False,
    find_font_name: str | None = None,
    replace_font_name: str | None = None,
    find_font_size: float | None = None,
    replace_font_size: float | None = None,
    find_font_color: str | None = None,
    replace_font_color: str | None = None,
    find_bold: bool | None = None,
    replace_bold: bool | None = None,
) -> str:
    """在全文范围内查找并替换文字，支持按字体筛选和替换。"""
    try:
        app = get_word_app()
        doc = app.ActiveDocument

        find = doc.Content.Find
        find.ClearFormatting()
        find.Replacement.ClearFormatting()

        find.Text = find_text
        find.Replacement.Text = replace_text
        find.MatchCase = match_case
        find.MatchWildcards = match_wildcards
        find.Forward = True
        find.Wrap = 1  # wdFindContinue

        if find_font_name is not None or find_font_size is not None or \
           find_font_color is not None or find_bold is not None:
            find_font = find.Font
            if find_font_name is not None:
                find_font.Name = find_font_name
            if find_font_size is not None:
                find_font.Size = find_font_size
            if find_font_color is not None:
                hex_str = find_font_color.lstrip("#")
                r = int(hex_str[0:2], 16)
                g = int(hex_str[2:4], 16)
                b = int(hex_str[4:6], 16)
                find_font.Color = r + g * 256 + b * 65536
            if find_bold is not None:
                find_font.Bold = find_bold

        if replace_font_name is not None or replace_font_size is not None or \
           replace_font_color is not None or replace_bold is not None:
            repl_font = find.Replacement.Font
            if replace_font_name is not None:
                repl_font.Name = replace_font_name
            if replace_font_size is not None:
                repl_font.Size = replace_font_size
            if replace_font_color is not None:
                hex_str = replace_font_color.lstrip("#")
                r = int(hex_str[0:2], 16)
                g = int(hex_str[2:4], 16)
                b = int(hex_str[4:6], 16)
                repl_font.Color = r + g * 256 + b * 65536
            if replace_bold is not None:
                repl_font.Bold = replace_bold

        find.Execute(Replace=2)  # wdReplaceAll

        parts = [f"'{find_text}' → '{replace_text}'"]
        if find_font_name:
            parts.append(f"查找字体={find_font_name}")
        if replace_font_name:
            parts.append(f"替换字体={replace_font_name}")
        if match_wildcards:
            parts.append("通配符模式")
        return f"已完成查找替换: {', '.join(parts)}"
    except Exception as e:
        return format_error("查找替换", e)

def word_format_text_by_find(
    find_text: str,
    font_name: str | None = None,
    font_size: float | None = None,
    bold: bool | None = None,
    italic: bool | None = None,
    underline: bool | None = None,
    color: str | None = None,
    match_case: bool = False,
    match_wildcards: bool = False,
    find_font_name: str | None = None,
) -> str:
    """查找特定文字并修改其格式（查找+格式化的组合操作）。"""
    try:
        app = get_word_app()
        doc = app.ActiveDocument

        find = doc.Content.Find
        find.ClearFormatting()
        find.Replacement.ClearFormatting()

        find.Text = find_text
        find.Replacement.Text = find_text
        find.MatchCase = match_case
        find.MatchWildcards = match_wildcards
        find.Forward = True
        find.Wrap = 1

        if find_font_name is not None:
            find.Font.Name = find_font_name

        repl_font = find.Replacement.Font
        set_any = False
        changes = []
        if font_name is not None:
            repl_font.Name = font_name
            changes.append(f"字体={font_name}")
            set_any = True
        if font_size is not None:
            repl_font.Size = font_size
            changes.append(f"字号={font_size}")
            set_any = True
        if bold is not None:
            repl_font.Bold = bold
            changes.append(f"加粗={'是' if bold else '否'}")
            set_any = True
        if italic is not None:
            repl_font.Italic = italic
            changes.append(f"斜体={'是' if italic else '否'}")
            set_any = True
        if underline is not None:
            repl_font.Underline = 1 if underline else 0
            changes.append(f"下划线={'是' if underline else '否'}")
            set_any = True
        if color is not None:
            hex_str = color.lstrip("#")
            r = int(hex_str[0:2], 16)
            g = int(hex_str[2:4], 16)
            b = int(hex_str[4:6], 16)
            repl_font.Color = r + g * 256 + b * 65536
            changes.append(f"颜色=#{color}")
            set_any = True

        if not set_any:
            return "未指定任何格式参数"

        find.Execute(Replace=2)

        constraint = f"（限定字体={find_font_name}）" if find_font_name else ""
        return f"已修改 '{find_text}'{constraint} 的格式: {', '.join(changes)}"
    except Exception as e:
        return format_error("按查找格式化", e)

def word_delete_target(target: LocatorDef) -> str:
    """定向删除特定目标内容。
    
    Args:
        target: 定位器字典，如 {"type": "paragraph", "index": 5}
    """
    try:
        app = get_word_app()
        doc = app.ActiveDocument
        rng = resolve_range(app, doc, target)
        rng.Delete()
        return f"已成功删除目标内容"
    except Exception as e:
        return format_error("删除目标内容", e)

def word_replace_target(target: LocatorDef, new_text: str) -> str:
    """定向替换特定目标的内容。
    
    Args:
        target: 定位器字典，如 {"type": "paragraph", "index": 5} 或 {"type": "heading", "text": "工作经历"}
        new_text: 替换后的新文本
    """
    try:
        app = get_word_app()
        doc = app.ActiveDocument
        rng = resolve_range(app, doc, target)
        rng.Text = new_text
        return f"已成功将目标内容替换为 {len(new_text)} 个字符的新文本"
    except Exception as e:
        return format_error("替换目标内容", e)

