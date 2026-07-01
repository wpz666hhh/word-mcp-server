"""Table operations for Microsoft Word."""

from ..com_manager import get_word_app
from ..utils import format_error
from ..locator import resolve_range, LocatorDef

def word_insert_table(
    rows: int,
    cols: int,
    data: list[list[str]] | None = None,
    position: LocatorDef = "selection",
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
        rng = resolve_range(app, doc, position)

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

def word_format_table(
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
