"""Document locator and range resolution for Microsoft Word."""
from typing import Union
from .models import ParagraphLocator, HeadingLocator, SearchLocator, BookmarkLocator

LocatorDef = Union[ParagraphLocator, HeadingLocator, SearchLocator, BookmarkLocator, dict, str]

def resolve_range(app, doc, range_spec: LocatorDef):
    """Resolve a range specification string or dict to a Word Range object.

    Args:
        app: Word Application object.
        doc: Word Document object.
        range_spec: String ("selection"|"start"|"end"|"all") or Dict locator.
            Dict examples:
            {"type": "paragraph", "index": 5}
            {"type": "heading", "text": "工作经历"}
            {"type": "search", "text": "特定文本", "instance": 1}

    Returns:
        A Word Range object.
    """
    if isinstance(range_spec, dict):
        loc_type = range_spec.get("type")
        if loc_type == "paragraph":
            idx = range_spec.get("index", 1)
            if 1 <= idx <= doc.Paragraphs.Count:
                return doc.Paragraphs(idx).Range
            else:
                raise ValueError(f"段落索引 {idx} 越界 (1-{doc.Paragraphs.Count})")
                
        elif loc_type == "heading":
            text_match = range_spec.get("text", "")
            for i in range(1, doc.Paragraphs.Count + 1):
                p = doc.Paragraphs(i)
                style_name = ""
                try:
                    style_name = p.Style.NameLocal
                except Exception:
                    pass
                if "Heading" in style_name or "标题" in style_name:
                    if text_match in p.Range.Text:
                        return p.Range
            raise ValueError(f"未找到包含 '{text_match}' 的标题")
            
        elif loc_type == "search":
            text_match = range_spec.get("text", "")
            instance = range_spec.get("instance", 1)
            
            rng = doc.Content
            count = 0
            while rng.Find.Execute(FindText=text_match, Forward=True, Wrap=0):
                count += 1
                if count == instance:
                    return rng
            raise ValueError(f"未找到搜索文本 '{text_match}' 的第 {instance} 次出现")
            
        else:
            raise ValueError(f"未知的定位器类型: {loc_type}")
            
    else:
        if range_spec == "start":
            return doc.Range(0, 0)
        elif range_spec == "end":
            end = doc.Content.End - 1
            return doc.Range(end, end)
        elif range_spec == "all":
            return doc.Content
        else:  # "selection" or default
            return app.Selection.Range
