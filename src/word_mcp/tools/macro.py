"""Batch macro execution operations for Microsoft Word."""

from ..com_manager import get_word_app
from ..utils import format_error

def word_execute_macro(python_code: str) -> str:
    """在后台批量执行基于 win32com 的 Python 脚本，以极大地提高处理复杂格式和内容的效率。
    
    这对于减少与 MCP 服务器之间的多次通信（Token 节省）极其有用。
    
    可用的局部变量:
    - app: Word Application 对象 (win32com)
    - doc: ActiveDocument 对象
    
    注意事项:
    - 确保使用 Range 而非 Selection，以提高性能。
    - 使用 print() 打印的信息将被捕获并作为结果的一部分返回。
    - 不要尝试打开新文档，只操作当前活动的 doc。
    
    Args:
        python_code: 要执行的 Python 脚本内容。
    """
    import io
    import sys
    import contextlib

    try:
        app = get_word_app()
        doc = app.ActiveDocument
        
        # 优化性能：可选地关闭界面刷新
        original_updating = app.ScreenUpdating
        app.ScreenUpdating = False
        
        stdout_capture = io.StringIO()
        
        try:
            # 准备执行环境
            local_vars = {"app": app, "doc": doc}
            
            with contextlib.redirect_stdout(stdout_capture):
                exec(python_code, {}, local_vars)
                
        finally:
            # 确保无论脚本是否报错，都恢复屏幕刷新
            app.ScreenUpdating = original_updating
            
        output = stdout_capture.getvalue()
        if not output.strip():
            return "脚本执行成功 (无输出)"
        return f"脚本执行成功，输出:\n{output}"
        
    except Exception as e:
        return format_error("执行批量脚本", e)
