"""L1: Document lifecycle tools — open, create, save, export, close, info."""

import os

from mcp.server.fastmcp import FastMCP

from ..com_manager import get_word_app
from ..utils import ensure_dir, format_error


def word_open(file_path: str) -> str:
    """打开已有 Word 文档。

    Args:
        file_path: 文档的绝对路径或相对路径 (.docx, .doc)
    """
    try:
        app = get_word_app()
        path = os.path.abspath(os.path.expandvars(os.path.expanduser(file_path)))
        norm_path = os.path.normcase(os.path.normpath(path))
        base_name = os.path.basename(norm_path).lower()

        # Check if already open — if so, just activate it
        for i in range(1, app.Documents.Count + 1):
            doc = app.Documents(i)
            try:
                doc_name = doc.Name.lower()
                doc_path = os.path.normcase(os.path.normpath(doc.FullName))
                
                # 1. 优先通过文件名（basename）比较，因为 Word 不允许同一实例中打开两个同名文件。
                # 这能完美解决 OneDrive 路径（https://...）与本地绝对路径不一致、或相对路径解析偏差的问题。
                # 2. 其次通过完整路径比对。
                if doc_name == base_name or doc_path == norm_path:
                    doc.Activate()
                    pages = doc.Content.ComputeStatistics(2)
                    return f"文档已打开，已激活: {doc.Name}（共 {pages} 页）"
            except Exception:
                pass  # unsaved docs may not have FullName

        doc = app.Documents.Open(path)
        pages = doc.Content.ComputeStatistics(2)  # wdStatisticPages
        return f"已打开文档: {doc.Name}（共 {pages} 页）"
    except Exception as e:
        return format_error("打开文档", e)


def word_create(template_path: str | None = None) -> str:
    """新建 Word 文档，可指定模板。

    Args:
        template_path: 可选，模板文件路径 (.dotx, .dot)
    """
    try:
        app = get_word_app()
        if template_path:
            path = os.path.abspath(os.path.expandvars(os.path.expanduser(template_path)))
            doc = app.Documents.Add(Template=path)
        else:
            doc = app.Documents.Add()
        return f"已创建新文档: {doc.Name}"
    except Exception as e:
        return format_error("创建文档", e)


def word_save(file_path: str | None = None) -> str:
    """保存当前活动文档。

    Args:
        file_path: 可选，另存为目标路径。不提供则保存到原路径。
    """
    try:
        app = get_word_app()
        doc = app.ActiveDocument
        if file_path:
            path = os.path.abspath(os.path.expandvars(os.path.expanduser(file_path)))
            ensure_dir(path)
            doc.SaveAs(path)
        else:
            doc.Save()
        return f"已保存: {doc.FullName}"
    except Exception as e:
        return format_error("保存文档", e)


def word_save_as_pdf(
    output_path: str,
    page_range: str | None = None,
) -> str:
    """将当前文档导出为 PDF。

    Args:
        output_path: PDF 文件的输出路径。
        page_range: 可选，页码范围，如 "1-3"。不指定则导出全部。
    """
    try:
        app = get_word_app()
        doc = app.ActiveDocument
        path = os.path.abspath(os.path.expandvars(os.path.expanduser(output_path)))
        ensure_dir(path)

        if page_range:
            start, end = page_range.split("-")
            doc.ExportAsFixedFormat(
                path, 17,  # wdExportFormatPDF = 17
                Range=3,   # wdExportFromTo
                From=int(start.strip()),
                To=int(end.strip()),
            )
        else:
            doc.ExportAsFixedFormat(path, 17)
        return f"已导出 PDF: {path}"
    except Exception as e:
        return format_error("导出 PDF", e)


def word_close(save_before_close: bool = True) -> str:
    """关闭当前活动文档（不关闭 Word 进程）。

    Args:
        save_before_close: 关闭前是否保存，默认 True。
    """
    try:
        app = get_word_app()
        if app.Documents.Count == 0:
            return "没有打开的文档"
        doc = app.ActiveDocument
        name = doc.Name
        doc.Close(SaveChanges=save_before_close)
        return f"已关闭文档: {name}"
    except Exception as e:
        return format_error("关闭文档", e)


def word_get_active_document() -> str:
    """获取当前活动文档的信息（路径、页数、字数等）。"""
    try:
        app = get_word_app()
        if app.Documents.Count == 0:
            return "没有打开的文档"
        doc = app.ActiveDocument
        pages = doc.Content.ComputeStatistics(2)   # wdStatisticPages
        words = doc.Content.ComputeStatistics(0)   # wdStatisticWords
        lines = doc.Content.ComputeStatistics(1)   # wdStatisticLines
        paragraphs = doc.Content.ComputeStatistics(4)  # wdStatisticParagraphs
        return (
            f"当前文档: {doc.FullName}\n"
            f"页数: {pages}\n"
            f"字数: {words}\n"
            f"行数: {lines}\n"
            f"段落数: {paragraphs}"
        )
    except Exception as e:
        return format_error("获取文档信息", e)


def word_list_documents() -> str:
    """列出 Word 中当前所有已打开的文档。

    Returns:
        每行一个文档，格式: "序号. 文档名 — 路径（当前活动文档标有 *）"
    """
    try:
        app = get_word_app()
        count = app.Documents.Count
        if count == 0:
            return "Word 中没有打开的文档"

        active_name = app.ActiveDocument.Name if app.ActiveDocument else ""
        lines = []
        for i in range(1, count + 1):
            doc = app.Documents(i)
            try:
                path = doc.FullName or "(未保存)"
            except Exception:
                path = "(未保存)"
            marker = " *" if doc.Name == active_name else ""
            lines.append(f"{i}. {doc.Name} — {path}{marker}")

        header = f"共 {count} 个文档（* 表示当前活动文档）:\n"
        return header + "\n".join(lines)
    except Exception as e:
        return format_error("列出文档", e)


def word_activate_document(index: int) -> str:
    """激活（切换到）指定的已打开文档。

    Args:
        index: 文档序号，与 word_list_documents 返回的序号一致（1-based）。
    """
    try:
        app = get_word_app()
        count = app.Documents.Count
        if count == 0:
            return "Word 中没有打开的文档"
        if index < 1 or index > count:
            names = [app.Documents(i).Name for i in range(1, count + 1)]
            return f"序号超出范围（1-{count}），当前文档: {', '.join(names)}"

        doc = app.Documents(index)
        doc.Activate()
        pages = doc.Content.ComputeStatistics(2)
        return f"已激活文档: {doc.Name}（共 {pages} 页）"
    except Exception as e:
        return format_error("激活文档", e)


def register_lifecycle_tools(mcp: FastMCP):
    """Register all L1 document lifecycle tools on the MCP server."""
    mcp.tool()(word_open)
    mcp.tool()(word_create)
    mcp.tool()(word_save)
    mcp.tool()(word_save_as_pdf)
    mcp.tool()(word_close)
    mcp.tool()(word_get_active_document)
    mcp.tool()(word_list_documents)
    mcp.tool()(word_activate_document)
