"""html2pptx-mcp：受限 HTML 契约与原生 PPTX 生成。

包内分层：``html2pptx_mcp.convert`` 是 HTML→PPTX 的独立子系统，不依赖 MCP 层；
将来的 ``html2pptx_mcp.server`` / ``html2pptx_mcp.tools`` 只做协议适配，
业务逻辑下沉到 ``convert`` 与核心层。
"""

from __future__ import annotations

__version__ = "0.4.1"

__all__ = ["__version__"]
