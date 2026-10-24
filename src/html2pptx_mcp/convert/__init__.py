"""受限 HTML → Slide IR → 原生 .pptx 的转换子系统。

模块一览（同层平铺，依赖只能向下）：

    api.py            对外编排：compile_html / validate_html / CompileOptions
    cli.py            命令行入口
    contract.py       契约：数据模型、加载、引用校验
    contract_docs.py  契约文档生成（JSON / 人读 / 模型用）
    ir.py             Slide IR：可 JSON 序列化的中间表示
    diagnostics.py    诊断模型：error / fix / warning
    values.py         单个 CSS 值的解析（纯函数，无 I/O）
    html.py           DOM + 样式 → IR，唯一的校验判定点
    html_table.py     简单表格解析
    html_svg.py       SVG 子集解析
    assets.py         资源定位、路径沙箱、解码
    emit.py           IR → python-pptx / DrawingML
    textfit.py        文本溢出估算
    data/             契约 JSON（contract / colors / fonts）

不变量：校验与生成共用同一条 html→IR 路径；任何「不按字面实现」都必须产出诊断，
不允许静默丢弃。本子系统不依赖 MCP 层，可被 CLI 与未来的 server 直接复用。
"""

from __future__ import annotations

from .api import CompileOptions, CompileResult, compile_html, validate_html
from .contract import Contract, load_contract
from .diagnostics import Report

__all__ = [
    "CompileOptions",
    "CompileResult",
    "Contract",
    "Report",
    "compile_html",
    "load_contract",
    "validate_html",
]
