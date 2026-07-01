"""Document structure and advanced reading operations for Microsoft Word."""

import json
from ..com_manager import get_word_app
from ..utils import format_error

def word_get_document_structure(include_paragraphs: bool = False, max_paragraphs: int = 100) -> str:
    """获取文档结构（标题大纲，可选包含段落）。
    
    Args:
        include_paragraphs: 是否包含普通段落，默认False只返回标题。
        max_paragraphs: 最大读取段落数，防止超长文档卡死。
    """
    try:
        app = get_word_app()
        doc = app.ActiveDocument
        
        structure = []
        count = 0
        for i in range(1, doc.Paragraphs.Count + 1):
            if count >= max_paragraphs:
                break
                
            p = doc.Paragraphs(i)
            style_name = ""
            try:
                style_name = p.Style.NameLocal
            except Exception:
                pass
                
            is_heading = "Heading" in style_name or "标题" in style_name
            
            if is_heading or include_paragraphs:
                text = p.Range.Text.strip()
                if text:
                    structure.append({
                        "index": i,
                        "type": "heading" if is_heading else "paragraph",
                        "style": style_name,
                        "text": text[:50] + ("..." if len(text) > 50 else "")
                    })
                    if not is_heading:
                        count += 1
                        
        return json.dumps(structure, ensure_ascii=False, indent=2)
    except Exception as e:
        return format_error("获取文档结构", e)

def word_read_table(table_index: int = 1) -> str:
    """读取指定的表格内容并返回为二维 JSON 数组。
    
    Args:
        table_index: 表格序号（1 = 第一个表格）
    """
    try:
        app = get_word_app()
        doc = app.ActiveDocument
        
        if table_index < 1 or table_index > doc.Tables.Count:
            return f"表格序号 {table_index} 超出范围 (共 {doc.Tables.Count} 个表格)"
            
        table = doc.Tables(table_index)
        data = []
        
        # Word Tables are 1-indexed
        for r in range(1, table.Rows.Count + 1):
            row_data = []
            try:
                row = table.Rows(r)
                for c in range(1, row.Cells.Count + 1):
                    # Remove the cell marker (bell character + CR)
                    text = row.Cells(c).Range.Text.replace('\x07', '').replace('\r', '').strip()
                    row_data.append(text)
                data.append(row_data)
            except Exception:
                # Merged cells can cause exceptions when iterating
                data.append(["[Error reading row due to merged cells]"])
                
        return json.dumps(data, ensure_ascii=False, indent=2)
    except Exception as e:
        return format_error("读取表格", e)
