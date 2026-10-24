#!/usr/bin/env python
"""视觉对比：Edge 截图（HTML 基准）vs PowerPoint 渲染（pptx 产物）逐页 diff。

用法：
    python scripts/compare_visual.py --dir docs/temp/trial/single \
        --out docs/temp/trial/visual/single

每页输出（--out 下）：
    <stem>/edge.png / ppt.png / diff_mask.png / side_by_side.png

report.json 记录每页与汇总的像素差异（阈值 16/255 起计）。几何与样式一致性由
pytest 冒烟保证，本脚本回答的是「看起来像不像」——这一层只能靠独立渲染对照。
"""

from __future__ import annotations

import argparse
import json
import subprocess
import sys
import time
from dataclasses import dataclass
from pathlib import Path
from typing import Any

from PIL import Image, ImageChops

from html2pptx_mcp.convert import CompileOptions, compile_html

_ROOT = Path(__file__).resolve().parents[1]
_EDGE_CANDIDATES = (
    r"C:\Program Files (x86)\Microsoft\Edge\Application\msedge.exe",
    r"C:\Program Files\Microsoft\Edge\Application\msedge.exe",
)
_DIFF_THRESHOLD = 16


@dataclass(frozen=True, slots=True)
class PageMetrics:
    diff_ratio_gt16: float
    diff_ratio_gt64: float
    mean_abs_diff: float

    def to_dict(self) -> dict[str, Any]:
        return {
            "diff_ratio_gt16": self.diff_ratio_gt16,
            "diff_ratio_gt64": self.diff_ratio_gt64,
            "mean_abs_diff": self.mean_abs_diff,
        }


def _find_edge() -> Path:
    for candidate in _EDGE_CANDIDATES:
        if Path(candidate).exists():
            return Path(candidate)
    raise FileNotFoundError("未找到 Edge；请安装或调整 _EDGE_CANDIDATES")


def _edge_screenshot(edge: Path, html: Path, target: Path) -> None:
    target.parent.mkdir(parents=True, exist_ok=True)
    subprocess.run(
        [
            str(edge),
            "--headless",
            "--disable-gpu",
            f"--screenshot={target}",
            "--window-size=1280,720",
            "--hide-scrollbars",
            html.as_uri(),
        ],
        capture_output=True,
        timeout=60,
    )
    for _ in range(40):  # Edge 立即返回，文件落盘有延迟
        if target.is_file() and target.stat().st_size > 0:
            return
        time.sleep(0.1)
    raise FileNotFoundError(f"Edge 截图未生成：{target}")


def _pptx_to_png(pptx: Path, out_dir: Path) -> Path:
    out_dir.mkdir(parents=True, exist_ok=True)
    result = subprocess.run(
        [
            "powershell",
            "-NoProfile",
            "-ExecutionPolicy",
            "Bypass",
            "-File",
            str(_ROOT / "scripts" / "pptx_to_png.ps1"),
            "-Path",
            str(pptx),
            "-OutDir",
            str(out_dir),
        ],
        capture_output=True,
        text=True,
        timeout=180,
    )
    produced = sorted(out_dir.glob(f"{pptx.stem}_p*.png"))
    if result.returncode != 0 or not produced:
        # PowerPoint COM 是单实例：残留进程会让下一次 Open 失败，清掉后重试一次
        subprocess.run(["taskkill", "/F", "/IM", "POWERPNT.EXE"], capture_output=True)
        time.sleep(1.0)
        subprocess.run(
            [
                "powershell",
                "-NoProfile",
                "-ExecutionPolicy",
                "Bypass",
                "-File",
                str(_ROOT / "scripts" / "pptx_to_png.ps1"),
                "-Path",
                str(pptx),
                "-OutDir",
                str(out_dir),
            ],
            capture_output=True,
            text=True,
            timeout=180,
        )
        produced = sorted(out_dir.glob(f"{pptx.stem}_p*.png"))
    if not produced:
        raise FileNotFoundError(f"PowerPoint 导出失败：{pptx}（{result.stderr[-200:]}）")
    return produced[0]


def _compare(edge: Path, ppt: Path) -> tuple[PageMetrics, Image.Image, Image.Image, Image.Image]:
    edge_image = Image.open(edge).convert("RGB")
    ppt_image = Image.open(ppt).convert("RGB")
    if edge_image.size != ppt_image.size:
        ppt_image = ppt_image.resize(edge_image.size)
    diff = ImageChops.difference(edge_image, ppt_image).convert("L")
    histogram = diff.histogram()
    total = edge_image.size[0] * edge_image.size[1]
    metrics = PageMetrics(
        diff_ratio_gt16=round(sum(histogram[_DIFF_THRESHOLD:]) / total, 4),
        diff_ratio_gt64=round(sum(histogram[64:]) / total, 4),
        mean_abs_diff=round(sum(i * c for i, c in enumerate(histogram)) / total, 2),
    )
    mask = diff.point(lambda v: 255 if v > _DIFF_THRESHOLD else 0)
    side = Image.new("RGB", (edge_image.size[0], edge_image.size[1] * 2 + 8), "#111111")
    side.paste(edge_image, (0, 0))
    side.paste(ppt_image, (0, edge_image.size[1] + 8))
    return metrics, mask, side, ppt_image


def main(argv: list[str] | None = None) -> int:
    parser = argparse.ArgumentParser(description="HTML vs PPT 视觉对比")
    parser.add_argument("--dir", required=True, type=Path, help="候选 HTML 目录")
    parser.add_argument("--out", required=True, type=Path, help="对比产物输出目录")
    parser.add_argument("--base-dir", type=Path, help="资源根目录（默认同 --dir）")
    args = parser.parse_args(argv)
    args.out = args.out.resolve()
    if args.base_dir:
        args.base_dir = args.base_dir.resolve()

    files = sorted(p.resolve() for p in args.dir.glob("*.html"))
    if not files:
        print(f"{args.dir} 下没有 .html", file=sys.stderr)
        return 1
    edge = _find_edge()
    base = args.base_dir or args.dir
    options = CompileOptions(base_dir=base)
    pages: list[dict[str, Any]] = []
    for html in files:
        stem = html.stem
        page_out = args.out / stem
        page_out.mkdir(parents=True, exist_ok=True)
        result = compile_html(html.read_text(encoding="utf-8"), options=options)
        if not result.ok:
            codes = "、".join(item.code for item in result.report.errors)
            print(f"FAIL  {html.name}  编译错误：{codes}")
            pages.append({"file": html.name, "ok": False, "errors": codes})
            continue
        pptx_path = page_out / f"{stem}.pptx"
        result.save(pptx_path)
        edge_png = page_out / "edge.png"
        _edge_screenshot(edge, html, edge_png)
        ppt_png = _pptx_to_png(pptx_path, page_out)
        metrics, mask, side, _ = _compare(edge_png, ppt_png)
        mask.save(page_out / "diff_mask.png")
        side.save(page_out / "side_by_side.png")
        pages.append(
            {
                "file": html.name,
                "ok": True,
                "metrics": metrics.to_dict(),
                "edge": str(edge_png),
                "ppt": str(ppt_png),
            }
        )
        print(
            f"OK    {html.name:<28} "
            f"diff>{_DIFF_THRESHOLD}={metrics.diff_ratio_gt16:.2%}  "
            f"mean={metrics.mean_abs_diff}"
        )

    ok_pages = [page for page in pages if page.get("ok")]
    ratios = [page["metrics"]["diff_ratio_gt16"] for page in ok_pages]
    worst = max(ok_pages, key=lambda page: page["metrics"]["diff_ratio_gt16"]) if ok_pages else None
    summary: dict[str, Any] = {
        "pages": len(pages),
        "compiled": len(ok_pages),
        "avg_diff_ratio_gt16": round(sum(ratios) / len(ratios), 4) if ratios else None,
        "worst_page": worst["file"] if worst else None,
    }
    report = {"directory": str(args.dir), "pages": pages, "summary": summary}
    args.out.mkdir(parents=True, exist_ok=True)
    (args.out / "report.json").write_text(
        json.dumps(report, ensure_ascii=False, indent=2) + "\n", encoding="utf-8", newline="\n"
    )
    print(
        f"\n对比完成：{summary['compiled']}/{summary['pages']} 页编译成功，"
        f"平均差异 {summary['avg_diff_ratio_gt16']:.2%}，"
        f"最差页 {summary['worst_page']}"
    )
    print(f"报告：{args.out / 'report.json'}")
    return 0


if __name__ == "__main__":
    sys.exit(main())
