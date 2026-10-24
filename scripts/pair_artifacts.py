#!/usr/bin/env python
"""把候选 HTML 与对应的 pptx 配对输出，便于人工逐页对比。

用法：
    python scripts/pair_artifacts.py --dir docs/temp/trial/llm-r5 \
        --out docs/temp/trial/pair/llm-r5 --deck

输出目录结构：
    <out>/01_cover.html   原样拷贝（浏览器打开）
    <out>/01_cover.pptx   单页编译产物（PowerPoint 打开）
    <out>/assets/         资源原样拷贝
    <out>/full_deck.pptx  全部页面按文件名顺序拼成的整册（--deck 时）
    <out>/README.md       本目录说明与每页状态
"""

from __future__ import annotations

import argparse
import shutil
import sys
from pathlib import Path

from html2pptx_mcp.convert import CompileOptions, compile_html


def main(argv: list[str] | None = None) -> int:
    parser = argparse.ArgumentParser(description="HTML/pptx 配对产物输出")
    parser.add_argument("--dir", required=True, type=Path, help="候选 HTML 目录")
    parser.add_argument("--out", required=True, type=Path, help="配对输出目录")
    parser.add_argument("--deck", action="store_true", help="额外生成整册 full_deck.pptx")
    parser.add_argument("--font-strategy", choices=("auto", "portable", "windows"), default="auto")
    args = parser.parse_args(argv)

    files = sorted(args.dir.glob("*.html"))
    if not files:
        print(f"{args.dir} 下没有 .html", file=sys.stderr)
        return 1
    args.out.mkdir(parents=True, exist_ok=True)
    for source in files:
        shutil.copy2(source, args.out / source.name)
    if (args.dir / "assets").exists():
        shutil.copytree(args.dir / "assets", args.out / "assets", dirs_exist_ok=True)

    options = CompileOptions(base_dir=args.dir, font_strategy=args.font_strategy)
    rows: list[tuple[str, bool, int, str]] = []
    for source in files:
        result = compile_html(source.read_text(encoding="utf-8"), options=options)
        if result.ok:
            target = args.out / f"{source.stem}.pptx"
            result.save(target)
            rows.append((source.name, True, len(result.pptx or b""), ""))
        else:
            codes = "、".join(item.code for item in result.report.errors)
            rows.append((source.name, False, 0, codes))

    if args.deck:
        pages = [source.read_text(encoding="utf-8") for source in files]
        deck = compile_html(pages, options=options)
        if deck.ok:
            deck.save(args.out / "full_deck.pptx")
            rows.append(("full_deck.pptx", True, len(deck.pptx or b""), ""))
        else:
            codes = "、".join(item.code for item in deck.report.errors)
            rows.append(("full_deck.pptx", False, 0, codes))

    lines = [
        "# 配对产物",
        "",
        f"来源：`{args.dir}`",
        "",
        "对比方式：同名 `.html` 用浏览器打开，`.pptx` 用 PowerPoint 打开，逐页对照。",
        "",
        "| 文件 | 状态 | 字节 | 备注 |",
        "|---|---|---|---|",
    ]
    lines += [
        f"| {name} | {'OK' if ok else 'FAIL'} | {size} | {note} |" for name, ok, size, note in rows
    ]
    (args.out / "README.md").write_text("\n".join(lines) + "\n", encoding="utf-8", newline="\n")

    failed = [row for row in rows if not row[1]]
    for name, ok, size, note in rows:
        print(f"{'OK ' if ok else 'FAIL'}  {name:<24} {size:>8}  {note}")
    return 0 if not failed else 2


if __name__ == "__main__":
    sys.exit(main())
