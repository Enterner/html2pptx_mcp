"""``html2pptx-mcp`` 入口：默认 stdio，可选 streamable-http。

用法：
    html2pptx-mcp                          # stdio（Claude Desktop 等客户端）
    html2pptx-mcp --transport streamable-http --host 127.0.0.1 --port 8000
"""

from __future__ import annotations

import argparse
import sys
from collections.abc import Sequence

from . import __version__
from .server import build_server


def main(argv: Sequence[str] | None = None) -> int:
    parser = argparse.ArgumentParser(
        prog="html2pptx-mcp",
        description="受限 HTML → 原生可编辑 PowerPoint 的 MCP 服务",
    )
    parser.add_argument(
        "--transport",
        choices=("stdio", "streamable-http"),
        default="stdio",
        help="传输方式（默认 stdio）",
    )
    parser.add_argument("--host", default="127.0.0.1", help="HTTP 监听地址")
    parser.add_argument("--port", type=int, default=8000, help="HTTP 端口")
    parser.add_argument("--path", default="/mcp", help="streamable-http 路径")
    parser.add_argument("--version", action="version", version=__version__)
    args = parser.parse_args(argv)

    server = build_server()
    if args.transport == "stdio":
        server.run("stdio")
    else:
        server.run(
            "streamable-http",
            host=args.host,
            port=args.port,
            streamable_http_path=args.path,
        )
    return 0


if __name__ == "__main__":  # pragma: no cover - 入口
    sys.exit(main())
