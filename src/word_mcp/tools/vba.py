"""L3: VBA executor — the final fallback for any Word operation.

Injects VBA code into Word's VBE (Visual Basic Editor) as a temporary
module, runs the Main() subroutine, then removes the module.

Requires: Word Trust Center → "Trust access to the VBA project object model"
"""

import logging

from mcp.server.fastmcp import FastMCP

from ..com_manager import get_word_app
from ..utils import format_error

logger = logging.getLogger(__name__)

vbext_ct_StdModule = 1


async def word_execute_vba(code: str) -> str:
    """在 Word 进程中执行 VBA 代码，覆盖所有 L2 未覆盖的操作。

    VBA 代码必须包含一个名为 Main 的 Sub 过程。
    例如:
        Sub Main()
            Selection.TypeText "Hello from VBA"
        End Sub

    Args:
        code: 完整的 VBA 代码字符串，必须包含 Sub Main()。

    Notes:
        - 代码在临时模块中运行，执行后自动删除
        - 不修改 Normal.dotm 或创建持久化宏
        - 需要 Word 启用「信任对 VBA 项目对象模型的访问」
          (文件 → 选项 → 信任中心 → 宏设置)
    """
    vb_comp = None

    # Validate: must contain Sub Main()
    if "Sub Main(" not in code and "Sub Main (" not in code:
        return (
            "错误: VBA 代码必须包含 Sub Main() 过程。\n"
            "示例:\n"
            "Sub Main()\n"
            "    Selection.TypeText \"Hello\"\n"
            "End Sub"
        )

    app = get_word_app()

    try:
        try:
            vb_project = app.ActiveDocument.VBProject
        except Exception:
            return (
                "错误: 无法访问 VBA 项目对象模型。\n"
                "请在 Word 中执行以下操作:\n"
                "1. 文件 → 选项 → 信任中心 → 信任中心设置\n"
                "2. 宏设置 → 勾选「信任对 VBA 项目对象模型的访问」\n"
                "3. 重新启动 Word"
            )

        vb_comp = vb_project.VBComponents.Add(vbext_ct_StdModule)
        vb_comp.CodeModule.AddFromString(code)

        app.Run("Main")

        return "VBA 代码已执行成功"

    except Exception as e:
        return f"VBA 执行错误: {e}"

    finally:
        if vb_comp is not None:
            try:
                vb_project = app.ActiveDocument.VBProject
                vb_project.VBComponents.Remove(vb_comp)
            except Exception as cleanup_err:
                logger.warning(f"Failed to clean up VBA module: {cleanup_err}")


def register_vba_tool(mcp: FastMCP):
    """Register the VBA executor tool on the MCP server."""
    mcp.tool()(word_execute_vba)
