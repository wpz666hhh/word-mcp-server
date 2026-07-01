"""Self-optimization and learning tools for Word MCP."""

import os
import json
from pathlib import Path

def word_record_lesson(
    title: str,
    content: str,
    new_error_code: str | None = None,
    new_error_hint: str | None = None
) -> str:
    """记录遇到的问题与经验到工作区规则中，并自动补充新的 COM 错误映射，以实现 MCP 的自我优化。
    
    Args:
        title: 经验教训的主题/标题，例如 "ListTemplates 级别添加限制"。
        content: 经验教训的具体描述（Markdown 格式），包括出错原因、正确写法等。
        new_error_code: 可选，本次遇到的 COM 错误码（如 "-2147352567"），如有。
        new_error_hint: 可选，该错误码对应的中文含义提示，如有。
    """
    try:
        # 1. 写入当前工作区的 .agents/AGENTS.md
        workspace_dir = Path(os.getcwd())
        agents_dir = workspace_dir / ".agents"
        agents_dir.mkdir(exist_ok=True)
        agents_file = agents_dir / "AGENTS.md"
        
        # 格式化经验条目
        lesson_entry = f"\n\n### {title}\n* **记录时间**: 自动同步\n{content}\n"
        
        if agents_file.exists():
            with open(agents_file, "r", encoding="utf-8") as f:
                existing_content = f.read()
        else:
            existing_content = "# Project Rules & COM Lessons\n\n这是一个用于 Word MCP 操作与 Python win32com 自动化的规则记录文档。\n"
            
        with open(agents_file, "w", encoding="utf-8") as f:
            f.write(existing_content + lesson_entry)
            
        status = ["已同步经验到当前工作区: .agents/AGENTS.md"]
        
        # 2. 如果提供了新的错误码，更新全局/共享的 com_errors.json
        if new_error_code and new_error_hint:
            src_dir = Path(__file__).parent.parent
            hints_file = src_dir / "com_errors.json"
            
            hints = {}
            if hints_file.exists():
                try:
                    with open(hints_file, "r", encoding="utf-8") as f:
                        hints = json.load(f)
                except Exception:
                    pass
            
            hints[str(new_error_code)] = new_error_hint
            
            with open(hints_file, "w", encoding="utf-8") as f:
                json.dump(hints, f, ensure_ascii=False, indent=4)
                
            status.append(f"已将新错误码写入 MCP 服务器: {new_error_code} -> '{new_error_hint}'")
            
        return "【自我优化完成】\n" + "\n".join(f"- {s}" for s in status)
    except Exception as e:
        return f"自我优化执行失败: {e}"
