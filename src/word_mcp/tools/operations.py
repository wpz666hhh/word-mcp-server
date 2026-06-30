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
    range_spec: str = "selection",
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

        # Resolve target range
        if table_index is not None:
            table = doc.Tables(table_index)
            rng = table.Range
            label = f"表格{table_index}"
        else:
            rng = _resolve_range(app, doc, range_spec)
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
            # Word Font.Color uses R + G*256 + B*65536 format
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


async def word_format_paragraph(
    range_spec: str = "selection",
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
        rng = _resolve_range(app, doc, range_spec)

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
    range_spec: str = "all",
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


async def word_select(range_spec: str) -> str:
    """选中文档中的指定区域（用户可在 Word 窗口中看到高亮）。

    Args:
        range_spec: "start"(跳到开头), "end"(跳到末尾), "all"(全选)
    """
    try:
        app = get_word_app()
        doc = app.ActiveDocument
        rng = _resolve_range(app, doc, range_spec)
        rng.Select()
        return f"已选中: {range_spec}"
    except Exception as e:
        return format_error("选中区域", e)


async def word_apply_style(
    range_spec: str = "selection",
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
        rng = _resolve_range(app, doc, range_spec)

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
    range_spec: str = "selection",
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
        rng = _resolve_range(app, doc, range_spec)

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


# === Batch B: Table, Image, Page Break, Find/Replace ===


async def word_insert_table(
    rows: int,
    cols: int,
    data: list[list[str]] | None = None,
    position: str = "selection",
) -> str:
    """插入表格，可附带数据填充。

    Args:
        rows: 行数
        cols: 列数
        data: 可选，二维数组，用于填充表格内容。第一行作为表头。
        position: 插入位置
    """
    try:
        app = get_word_app()
        doc = app.ActiveDocument
        rng = _resolve_range(app, doc, position)

        table = doc.Tables.Add(rng, rows, cols)
        table.Borders.Enable = True

        if data:
            for i, row_data in enumerate(data):
                for j, cell_text in enumerate(row_data):
                    if i < rows and j < cols:
                        cell = table.Cell(i + 1, j + 1)
                        cell.Range.Text = str(cell_text)

        if data:
            return f"已插入表格: {rows}行×{cols}列（已填充 {len(data)} 行数据）"
        return f"已插入表格: {rows}行×{cols}列"
    except Exception as e:
        return format_error("插入表格", e)


ROW_HEIGHT_RULE_MAP = {
    "auto": 0,      # wdRowHeightAuto
    "at_least": 1,  # wdRowHeightAtLeast
    "exactly": 2,   # wdRowHeightExactly
}


async def word_format_table(
    table_index: int = 1,
    style: str | None = None,
    header_row: bool = False,
    auto_fit: bool = False,
    font_name: str | None = None,
    font_size: float | None = None,
    font_color: str | None = None,
    font_bold: bool | None = None,
    row_height: float | None = None,
    row_height_rule: str | None = None,
) -> str:
    """设置表格格式。

    Args:
        table_index: 表格序号（1=第一个表格）
        style: Word 表格样式名，如 "Grid Table 1 Light"
        header_row: 是否设置首行为表头（加粗 + 重复标题行）
        auto_fit: 是否自适应列宽
        font_name: 表格字体名称，如 "宋体"、"微软雅黑"
        font_size: 表格字号（磅），如 12（小四）、10.5（五号）
        font_color: 表格字体颜色，RGB十六进制如 "000000"（黑色）
        font_bold: 表格文字是否加粗
        row_height: 统一行高（磅），如 45。需配合 row_height_rule 使用。
        row_height_rule: 行高规则 — "auto"(自动), "at_least"(至少), "exactly"(精确固定)。
                         设为 "exactly" 可强制所有行等高。
    """
    try:
        app = get_word_app()
        doc = app.ActiveDocument
        table = doc.Tables(table_index)

        changes = []

        if style is not None:
            try:
                table.Style = style
                changes.append(f"样式={style}")
            except Exception:
                pass

        if header_row:
            for j in range(1, table.Columns.Count + 1):
                table.Cell(1, j).Range.Font.Bold = True
            table.Rows(1).HeadingFormat = -1
            changes.append("首行为表头")

        if auto_fit:
            table.AutoFitBehavior(2)  # wdAutoFitWindow
            changes.append("自适应列宽")

        # Font formatting for the entire table
        if any(x is not None for x in (font_name, font_size, font_color, font_bold)):
            tbl_range = table.Range
            tbl_font = tbl_range.Font

            if font_name is not None:
                tbl_font.Name = font_name
                changes.append(f"字体={font_name}")
            if font_size is not None:
                tbl_font.Size = font_size
                changes.append(f"字号={font_size}pt")
            if font_color is not None:
                hex_str = font_color.lstrip("#")
                r = int(hex_str[0:2], 16)
                g = int(hex_str[2:4], 16)
                b = int(hex_str[4:6], 16)
                tbl_font.Color = r + g * 256 + b * 65536
                changes.append(f"颜色=#{font_color}")
            if font_bold is not None:
                tbl_font.Bold = font_bold
                changes.append(f"加粗={'是' if font_bold else '否'}")

        # Row height
        rule_label = row_height_rule if row_height_rule else ""
        if row_height is not None or row_height_rule is not None:
            for i in range(1, table.Rows.Count + 1):
                row = table.Rows(i)
                if row_height_rule is not None and row_height_rule in ROW_HEIGHT_RULE_MAP:
                    row.HeightRule = ROW_HEIGHT_RULE_MAP[row_height_rule]
                if row_height is not None:
                    row.Height = row_height
            if row_height and row_height_rule:
                changes.append(f"行高={row_height}pt({rule_label})")
            elif row_height:
                changes.append(f"行高={row_height}pt")
            elif row_height_rule:
                changes.append(f"行高规则={rule_label}")

        if changes:
            return f"已设置表格格式: {', '.join(changes)}"
        return "未指定任何表格格式参数"
    except Exception as e:
        return format_error("设置表格格式", e)


async def word_insert_image(
    image_path: str,
    position: str = "selection",
    width: float | None = None,
    height: float | None = None,
) -> str:
    """在文档中插入图片。

    Args:
        image_path: 图片文件路径（PNG, JPG, GIF, BMP 等）
        position: 插入位置
        width: 可选，图片宽度（磅），不指定则保持原始宽度
        height: 可选，图片高度（磅）
    """
    try:
        import os

        app = get_word_app()
        doc = app.ActiveDocument
        path = os.path.abspath(os.path.expandvars(os.path.expanduser(image_path)))

        if not os.path.exists(path):
            return f"图片文件不存在: {path}"

        rng = _resolve_range(app, doc, position)
        shape = doc.InlineShapes.AddPicture(
            path,
            LinkToFile=False,
            SaveWithDocument=True,
            Range=rng,
        )

        if width is not None:
            shape.Width = width
        if height is not None:
            shape.Height = height

        return f"已插入图片: {os.path.basename(path)}（{shape.Width:.0f}×{shape.Height:.0f}pt）"
    except Exception as e:
        return format_error("插入图片", e)


async def word_insert_page_break(position: str = "selection") -> str:
    """在指定位置插入分页符。

    Args:
        position: 插入位置
    """
    try:
        app = get_word_app()
        doc = app.ActiveDocument
        rng = _resolve_range(app, doc, position)
        rng.InsertBreak(7)  # wdPageBreak = 7
        return "已插入分页符"
    except Exception as e:
        return format_error("插入分页符", e)


async def word_find_replace(
    find_text: str,
    replace_text: str,
    match_case: bool = False,
) -> str:
    """在全文范围内查找并替换文字。

    Args:
        find_text: 要查找的文字
        replace_text: 替换后的文字
        match_case: 是否区分大小写
    """
    try:
        app = get_word_app()
        doc = app.ActiveDocument

        find = doc.Content.Find
        find.Text = find_text
        find.Replacement.Text = replace_text
        find.MatchCase = match_case
        find.Forward = True
        find.Wrap = 1  # wdFindContinue

        find.Execute(Replace=2)  # wdReplaceAll

        return f"已完成查找替换: '{find_text}' → '{replace_text}'"
    except Exception as e:
        return format_error("查找替换", e)


async def word_set_page_setup(
    margins: dict | None = None,
    orientation: str | None = None,
    page_size: str | None = None,
) -> str:
    """设置页面布局（边距、方向、纸张大小）。

    Args:
        margins: 边距字典，单位磅（1英寸=72磅）。
                 例: {"top": 72, "bottom": 72, "left": 90, "right": 90}
        orientation: 纸张方向 — "portrait"(纵向), "landscape"(横向)
        page_size: 纸张尺寸 — "A4", "A3", "Letter", "Legal"
    """
    try:
        app = get_word_app()
        doc = app.ActiveDocument
        ps = doc.PageSetup
        changes = []

        if margins:
            if "top" in margins:
                ps.TopMargin = margins["top"]
                changes.append(f"上边距={margins['top']}pt")
            if "bottom" in margins:
                ps.BottomMargin = margins["bottom"]
                changes.append(f"下边距={margins['bottom']}pt")
            if "left" in margins:
                ps.LeftMargin = margins["left"]
                changes.append(f"左边距={margins['left']}pt")
            if "right" in margins:
                ps.RightMargin = margins["right"]
                changes.append(f"右边距={margins['right']}pt")

        if orientation:
            ps.Orientation = 1 if orientation == "landscape" else 0
            changes.append(f"方向={orientation}")

        if page_size:
            page_sizes = {
                "A4": (595.3, 841.9),
                "A3": (841.9, 1190.5),
                "Letter": (612, 792),
                "Legal": (612, 1008),
            }
            if page_size in page_sizes:
                w, h = page_sizes[page_size]
                ps.PageWidth = w
                ps.PageHeight = h
                changes.append(f"纸张={page_size}")

        if changes:
            return f"已设置页面: {', '.join(changes)}"
        return "未指定任何页面设置参数"
    except Exception as e:
        return format_error("设置页面", e)


async def word_set_header_footer(
    type: str = "header",
    text: str = "",
    include_pagenum: bool = False,
    page_num_style: str = "page_of_total",
    alignment: str = "center",
    font_name: str | None = None,
    font_size: float | None = None,
) -> str:
    """设置页眉或页脚内容。

    Args:
        type: "header"(页眉) 或 "footer"(页脚)
        text: 前缀文字，放在页码格式之前。例如 "— " 得到 "— 第 1 页 / 共 2 页"。
              留空则只显示 "第 1 页 / 共 2 页"。
              如需只添加居中页码，可设 text="" + include_pagenum=true。
        include_pagenum: 是否添加页码。
        page_num_style: 页码格式 — "page_of_total"(第X页/共Y页), "simple"(仅数字)。
        alignment: 对齐方式 — "left", "center"(默认), "right"。
        font_name: 页眉/页脚字体名称，如 "Times New Roman"、"宋体"。
        font_size: 页眉/页脚字号（磅），如 10.5、12。
    """
    try:
        app = get_word_app()
        doc = app.ActiveDocument

        label = "页眉" if type == "header" else "页脚"
        align_val = ALIGNMENT_MAP.get(alignment, wdAlignParagraphCenter)

        # Apply to ALL sections
        sections_processed = 0
        for sec_idx in range(1, doc.Sections.Count + 1):
            section = doc.Sections(sec_idx)

            if type == "header":
                hf = section.Headers(1)  # wdHeaderFooterPrimary
            else:
                hf = section.Footers(1)  # wdHeaderFooterPrimary

            # Always clear existing content first
            hf.Range.Delete()

            if not text and not include_pagenum:
                sections_processed += 1
                continue

            # Get a fresh range after clearing
            rng = hf.Range

            # 1) Write prefix text
            if text:
                rng.Text = text
                rng.Collapse(0)  # wdCollapseEnd

            # 2) Append page-number fields
            if include_pagenum:
                if page_num_style == "simple":
                    # Just the page number
                    doc.Fields.Add(rng, 33)  # wdFieldPage
                else:
                    # "第 X 页 / 共 Y 页"
                    rng.InsertAfter("第 ")
                    rng.Collapse(0)
                    doc.Fields.Add(rng, 33)  # wdFieldPage
                    rng.InsertAfter(" 页 / 共 ")
                    rng.Collapse(0)
                    doc.Fields.Add(rng, 26)  # wdFieldNumPages
                    rng.InsertAfter(" 页")

            # Set paragraph alignment
            hf.Range.ParagraphFormat.Alignment = align_val

            # Apply font formatting to the header/footer
            if font_name is not None or font_size is not None:
                hf_font = hf.Range.Font
                if font_name is not None:
                    hf_font.Name = font_name
                if font_size is not None:
                    hf_font.Size = font_size

            sections_processed += 1

        # Build result message
        if not text and not include_pagenum:
            plural = "s" if sections_processed > 1 else ""
            return f"已清空{label}（{sections_processed} 节）"

        parts = []
        if text:
            parts.append(f"内容='{text}'")
        if include_pagenum:
            style_label = "数字" if page_num_style == "simple" else "第X页/共Y页"
            parts.append(f"页码({style_label})")
        parts.append(f"对齐={alignment}")
        if font_name is not None:
            parts.append(f"字体={font_name}")
        if font_size is not None:
            parts.append(f"字号={font_size}pt")
        plural = "s" if sections_processed > 1 else ""
        return f"已设置{label}（{sections_processed} 节{plural}）: {', '.join(parts)}"
    except Exception as e:
        return format_error(f"设置{type}", e)


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

    # === Batch B: Table, Image, Page Break, Find/Replace ===
    mcp.tool()(word_insert_table)
    mcp.tool()(word_format_table)
    mcp.tool()(word_insert_image)
    mcp.tool()(word_insert_page_break)
    mcp.tool()(word_find_replace)

    # === Batch C: Page Setup, Header/Footer ===
    mcp.tool()(word_set_page_setup)
    mcp.tool()(word_set_header_footer)
