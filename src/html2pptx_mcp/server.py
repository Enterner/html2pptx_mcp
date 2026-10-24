"""MCP 协议层：只做适配，业务全部在 ``convert`` 里。

三个工具、两个资源、一个提示词：

- ``validate_html``：零副作用校验，返回机读报告；
- ``convert_html``：生成并写出 pptx，失败时以 ToolError 返回完整诊断；
- ``get_contract``：导出契约（JSON 或给模型看的精简规范）；
- ``pptx://contract`` / ``pptx://contract/llm``：资源形式暴露同一份契约；
- ``design_html_slide``：把契约规范与任务拼成提示词，减少客户端自己组织上下文。

``convert`` 不 import 任何 mcp 包，因此这套适配层可以整体替换而不影响转换核心。
"""

from __future__ import annotations

import json
from pathlib import Path
from typing import Any

from mcp.server.mcpserver import MCPServer
from mcp.server.mcpserver.exceptions import ToolError

from . import __version__
from .convert import CompileOptions, compile_html, load_contract, validate_html
from .convert.contract_docs import render_json, render_llm
from .convert.reverse import outline_pptx, reverse_pptx

INSTRUCTIONS = """\
工作流：
1) 需要规范时先读资源 pptx://contract/llm（或调用 get_contract(format="llm")）。
2) 每页写一份独立 HTML：画布固定 1280×720，元素必须 position:absolute 且写全 box。
3) 用 validate_html 自查，按 error 的 hint 逐条修复（不要忽略 hint 自己猜）。
4) 用 convert_html 生成 pptx，output_path 由你指定；base_dir 默认取输出目录。
5) 需要调整时改 HTML 后整页重新生成；本服务不做增量编辑。
注意：契约之外的写法会被拒绝；字体、图片路径与限额都由服务端校验。
字体默认子集内嵌（convert_html embed_fonts=false 可关）——收到
FONT_EMBED_FILE_MISSING 属正常：缺字体机器与原版走同一回退链。

参考 pptx 复用（逆转）：
1) 先 ppt_outline 拿整册结构与角色初判（标题/背景/徽章/图片…），确定哪些是格式。
2) 再 ppt_reverse 逆转：pages/elements 可选——全逆转、按页、按元素均可；
   产物落盘（HTML+assets），片段只在不大时内联回传，别整册读进上下文。
3) 复用逆转出的 HTML 改内容后走 convert_html；格式即与参考版一致。

确认节奏：默认「先 HTML 后 pptx」——validate 通过后把 HTML 路径（配 Edge 截图更佳）
交用户过目，等明确确认再调 convert_html；用户开口就要求直接出片、或处于连续
多轮迭代中时，跳过确认直接生成。
"""


# ---------------------------------------------------------------- 工具实现
def validate_html_tool(
    html: str,
    font_strategy: str = "auto",
    lossy_fixes: str = "apply",
) -> str:
    """校验受限 HTML，返回 JSON 报告；不读图片、不写文件。"""
    report = validate_html(
        html,
        options=CompileOptions(font_strategy=font_strategy, lossy_fixes=lossy_fixes),
    )
    return json.dumps(report.to_dict(), ensure_ascii=False)


def _require_absolute(path: str, what: str) -> None:
    """服务进程的工作目录与调用方不同：相对路径会静默落错位置，必须显式拒绝。"""
    if not Path(path).is_absolute():
        raise ToolError(f"{what} 必须为绝对路径，得到 {path!r}")


def convert_html_tool(
    html: str,
    output_path: str,
    base_dir: str | None = None,
    font_strategy: str = "auto",
    lossy_fixes: str = "apply",
    embed_fonts: bool = True,
) -> str:
    """生成 pptx 并写到 ``output_path``；失败时抛 ToolError（内容为完整报告 JSON）。"""
    _require_absolute(output_path, "output_path")
    if base_dir is not None:
        _require_absolute(base_dir, "base_dir")
    target = Path(output_path)
    directory = Path(base_dir) if base_dir else target.parent
    result = compile_html(
        html,
        options=CompileOptions(
            base_dir=directory,
            font_strategy=font_strategy,
            lossy_fixes=lossy_fixes,
            embed_fonts=embed_fonts,
        ),
    )
    if not result.ok:
        raise ToolError(json.dumps(result.report.to_dict(), ensure_ascii=False))
    saved = result.save(target)
    payload: dict[str, Any] = {
        "ok": True,
        "path": str(saved),
        "bytes": len(result.pptx or b""),
        "stats": result.report.stats,
        "warnings": [item.to_dict() for item in result.report.warnings],
        "fixes": [item.to_dict() for item in result.report.fixes],
    }
    return json.dumps(payload, ensure_ascii=False)


def get_contract_tool(format: str = "llm") -> str:
    """导出契约：``llm`` 为精简规范，``json`` 为完整机读契约。"""
    contract = load_contract()
    if format == "json":
        return render_json(contract)
    if format == "llm":
        return render_llm(contract)
    raise ToolError(f"format 仅支持 llm / json，得到 {format!r}")


def outline_tool(pptx_path: str) -> str:
    """整册盘点：主题色板/字体 + 每页形状清单与角色初判；选择性逆转的选片依据。"""
    _require_absolute(pptx_path, "pptx_path")
    try:
        return json.dumps(outline_pptx(pptx_path), ensure_ascii=False)
    except Exception as exc:
        raise ToolError(f"盘点失败：{exc}") from exc


def reverse_tool(
    pptx_path: str,
    out_dir: str,
    pages: list[int] | None = None,
    elements: dict[str, list[int]] | None = None,
    include_html: str = "auto",
    placeholder_text: bool = False,
) -> str:
    """逆转参考 pptx 为受限 HTML。

    pages/elements 都缺省 = 全逆转；两者给其一即为选择性逆转。
    include_html：auto（片段总量 ≤12KB 才内联）/ always / never。
    """
    if include_html not in ("auto", "always", "never"):
        raise ToolError(f"include_html 仅支持 auto/always/never，得到 {include_html!r}")
    _require_absolute(pptx_path, "pptx_path")
    _require_absolute(out_dir, "out_dir")
    element_map = {int(k): v for k, v in elements.items()} if elements else None
    try:
        result = reverse_pptx(
            pptx_path,
            out_dir,
            pages=pages,
            elements=element_map,
            placeholder_text=placeholder_text,
        )
    except Exception as exc:
        raise ToolError(f"逆转失败：{exc}") from exc
    total = sum(len(fragment) for fragment in result.html_by_page.values())
    inline = include_html == "always" or (include_html == "auto" and total <= 12000)
    payload: dict[str, Any] = {
        "ok": True,
        "html_files": [str(p) for p in result.html_paths],
        "assets_dir": str(Path(out_dir) / "assets"),
        "asset_count": len(result.asset_paths),
        "notes": [vars(n) for n in result.notes],
        "fragments_bytes": total,
        "html_by_page": result.html_by_page if inline else None,
        "hint": (
            "片段已内联（.stage 内层，补视口壳后即可 convert_html）"
            if inline
            else "片段过大未内联，完整文件已落盘；用文件工具按需读取"
        ),
    }
    return json.dumps(payload, ensure_ascii=False)


def contract_json() -> str:
    return render_json(load_contract())


def contract_llm() -> str:
    return render_llm(load_contract())


def design_prompt(topic: str = "", requirements: str = "") -> str:
    """把契约规范与本次任务拼成一段可直接使用的提示词。"""
    task = f"\n## 本次任务\n\n主题：{topic or '未指定'}\n"
    if requirements:
        task += f"要求：{requirements}\n"
    task += (
        "\n输出：一份完整 HTML（含画布根 div），不要解释、不要 Markdown 代码围栏之外的内容。\n"
        "写完后自检：每个元素是否都有 position:absolute 与完整 box；"
        "是否只用了契约允许的标签与属性。\n"
    )
    return contract_llm() + task


# ---------------------------------------------------------------- 服务器装配
def build_server() -> MCPServer:
    """构造 MCP 服务器。返回新实例，便于测试与嵌入。"""
    server = MCPServer(
        name="html2pptx-mcp",
        title="PPT MCP Server",
        description="把受限 HTML 编译成原生可编辑的 PowerPoint 文档。",
        instructions=INSTRUCTIONS,
        version=__version__,
    )
    server.add_tool(
        validate_html_tool,
        name="validate_html",
        description="校验受限 HTML 是否符合契约；返回 errors/warnings/fixes 的 JSON 报告。",
    )
    server.add_tool(
        convert_html_tool,
        name="convert_html",
        description="把受限 HTML 编译为原生可编辑 .pptx 并写入 output_path。",
    )
    server.add_tool(
        get_contract_tool,
        name="get_contract",
        description="导出契约：format=llm 为精简规范，format=json 为完整契约。",
    )
    server.add_tool(
        outline_tool,
        name="ppt_outline",
        description=(
            "盘点参考 pptx：主题色板/字体体系 + 每页形状清单与角色初判"
            "（标题/背景/徽章/图片…）；选择性逆转的选片依据。"
        ),
    )
    server.add_tool(
        reverse_tool,
        name="ppt_reverse",
        description=(
            "把参考 pptx 逆转为受限 HTML（契约子集）供复用。pages/elements 缺省=全逆转，"
            "给其一即选择性逆转；产物落盘，片段不大于 12KB 时内联回传。"
        ),
    )

    @server.resource(
        "pptx://contract",
        name="html-contract",
        description="完整的受限 HTML 契约（JSON）",
        mime_type="application/json",
    )
    def _contract_json() -> str:
        return contract_json()

    @server.resource(
        "pptx://contract/llm",
        name="html-contract-llm",
        description="给模型看的契约精简规范（Markdown）",
        mime_type="text/markdown",
    )
    def _contract_llm() -> str:
        return contract_llm()

    @server.prompt(
        name="design_html_slide",
        description="按契约规范生成一页幻灯片的 HTML。",
    )
    def _design_html_slide(topic: str = "", requirements: str = "") -> str:
        return design_prompt(topic, requirements)

    return server
