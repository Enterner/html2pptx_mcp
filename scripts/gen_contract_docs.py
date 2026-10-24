#!/usr/bin/env python
"""生成或校验契约参考文档。

用法：
    python scripts/gen_contract_docs.py           # 重新生成 docs/contracts/*
    python scripts/gen_contract_docs.py --check   # CI：与数据不一致则失败
"""

from __future__ import annotations

import sys
from pathlib import Path

from html2pptx_mcp.convert.contract import load_contract
from html2pptx_mcp.convert.contract_docs import write_docs

ROOT = Path(__file__).resolve().parents[1]
OUT_DIR = ROOT / "docs" / "contracts"


def main(argv: list[str]) -> int:
    return write_docs(load_contract(), OUT_DIR, check="--check" in argv)


if __name__ == "__main__":
    sys.exit(main(sys.argv[1:]))
