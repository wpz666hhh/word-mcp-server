"""Page layout and header/footer operations for Microsoft Word."""

import os
from ..com_manager import get_word_app
from ..utils import format_error
from ..locator import resolve_range, LocatorDef
from ..models import Margins

def word_insert_image(
    image_path: str,
    position: LocatorDef = "selection",
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
        app = get_word_app()
        doc = app.ActiveDocument
        path = os.path.abspath(os.path.expandvars(os.path.expanduser(image_path)))

        if not os.path.exists(path):
            return f"图片文件不存在: {path}"

        rng = resolve_range(app, doc, position)
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

def word_insert_page_break(position: LocatorDef = "selection") -> str:
    """在指定位置插入分页符。

    Args:
        position: 插入位置
    """
    try:
        app = get_word_app()
        doc = app.ActiveDocument
        rng = resolve_range(app, doc, position)
        rng.InsertBreak(7)  # wdPageBreak = 7
        return "已插入分页符"
    except Exception as e:
        return format_error("插入分页符", e)

def word_set_page_setup(
    margins: Margins | None = None,
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

def word_set_header_footer(
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
        align_map = {"left": 0, "center": 1, "right": 2}
        align_val = align_map.get(alignment, 1)

        sections_processed = 0
        for sec_idx in range(1, doc.Sections.Count + 1):
            section = doc.Sections(sec_idx)

            if type == "header":
                hf = section.Headers(1)  # wdHeaderFooterPrimary
            else:
                hf = section.Footers(1)  # wdHeaderFooterPrimary

            hf.Range.Delete()

            if not text and not include_pagenum:
                sections_processed += 1
                continue

            rng = hf.Range

            if text:
                rng.Text = text
                rng.Collapse(0)  # wdCollapseEnd

            if include_pagenum:
                if page_num_style == "simple":
                    doc.Fields.Add(rng, 33)  # wdFieldPage
                else:
                    rng.InsertAfter("第 ")
                    rng.Collapse(0)
                    doc.Fields.Add(rng, 33)  # wdFieldPage
                    rng.InsertAfter(" 页 / 共 ")
                    rng.Collapse(0)
                    doc.Fields.Add(rng, 26)  # wdFieldNumPages
                    rng.InsertAfter(" 页")

            hf.Range.ParagraphFormat.Alignment = align_val

            if font_name is not None or font_size is not None:
                hf_font = hf.Range.Font
                if font_name is not None:
                    hf_font.Name = font_name
                if font_size is not None:
                    hf_font.Size = font_size

            sections_processed += 1

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
