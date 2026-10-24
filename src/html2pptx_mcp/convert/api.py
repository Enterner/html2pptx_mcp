"""对外编排层：validate / compile。

约定两条，避免调用方写样板：

- **用户错误不抛异常**：契约违规、资源失败全部落进 ``report``；异常只用于编程错误
  （空输入、内部不变量被破坏）。
- **校验与生成同源**：``validate_html`` 与 ``compile_html`` 都走 ``html.lower``，
  区别只是后者继续执行 emit。

多页输入按文件顺序拼成一个 Document，元素 id 全局重排，保证跨页唯一。
"""

from __future__ import annotations

import sys
from collections.abc import Sequence
from dataclasses import dataclass, replace
from io import BytesIO
from pathlib import Path

from . import fonts, ir, textfit
from .contract import Contract, load_contract
from .diagnostics import Report
from .emit import emit
from .html import lower


@dataclass(frozen=True, slots=True)
class CompileOptions:
    """一次编译的全部外部输入；显式传入，便于测试与批量任务复用。"""

    base_dir: Path = Path(".")
    contract: Contract | None = None
    font_strategy: str = "auto"
    lossy_fixes: str = "apply"
    fail_on_warnings: bool = False
    metrics: textfit.Metrics | None = None
    embed_fonts: bool = True


@dataclass(frozen=True, slots=True)
class CompileResult:
    ok: bool
    report: Report
    document: ir.Document
    pptx: bytes | None

    def save(self, path: str | Path) -> Path:
        """原子写入 pptx；没有产物（存在 error / 严格模式拦下警告）时抛 ValueError。"""
        if self.pptx is None:
            raise ValueError(
                "编译未产出 pptx：先修 report.errors"
                + ("（或关闭 fail_on_warnings）" if not self.report.ok else "")
            )
        target = Path(path)
        target.parent.mkdir(parents=True, exist_ok=True)
        temporary = target.with_name(f".{target.name}.tmp")
        temporary.write_bytes(self.pptx)
        temporary.replace(target)
        return target


def compile_html(
    source: str | Sequence[str],
    *,
    options: CompileOptions | None = None,
) -> CompileResult:
    """把一份或多份 HTML 编译成 pptx。

    单份 = 单页；多份 = 按顺序拼成多页。元素 id 在最终文档内重排为 ``e0001`` 起。
    返回结果对象而不是抛异常：调用方检查 ``ok`` 与 ``report``。
    """
    options = options or CompileOptions()
    contract = options.contract or load_contract()
    report = Report(contract, lossy_fixes=options.lossy_fixes)
    sources = _as_sources(source)

    slides: list[ir.Slide] = []
    canvas: ir.Canvas | None = None
    for html in sources:
        lowered = lower(
            html,
            contract=contract,
            font_strategy=_resolve_font_strategy(options.font_strategy),
            lossy_fixes=options.lossy_fixes,
        )
        _merge(report, lowered.report)
        canvas = canvas or lowered.document.canvas
        slides.append(lowered.document.slides[0])

    document = ir.Document(canvas=canvas or _canvas(contract), slides=_renumber(tuple(slides)))
    report.stats["slides"] = len(document.slides)
    report.stats["elements"] = sum(len(slide.elements) for slide in document.slides)

    if not report.ok:
        return CompileResult(ok=False, report=report, document=document, pptx=None)

    font_strategy = _resolve_font_strategy(options.font_strategy)
    emitted = emit(
        document,
        contract=contract,
        base_dir=options.base_dir,
        font_strategy=font_strategy,
        metrics=options.metrics or fonts.system_metrics(),
        report=report,
    )
    if options.embed_fonts:
        from .fontembed import embed_used_fonts

        embed_used_fonts(
            document,
            emitted.presentation,
            report,
            default_font=contract.font_set(font_strategy).default,
        )
    ok = report.ok and not (options.fail_on_warnings and report.warnings)
    if not ok:
        return CompileResult(ok=False, report=report, document=document, pptx=None)

    buffer = BytesIO()
    emitted.presentation.save(buffer)
    return CompileResult(ok=True, report=report, document=document, pptx=buffer.getvalue())


def validate_html(
    source: str | Sequence[str],
    *,
    options: CompileOptions | None = None,
) -> Report:
    """只校验，不生成、不读图片；与 ``compile_html`` 共用遍历层。"""
    options = options or CompileOptions()
    contract = options.contract or load_contract()
    report = Report(contract, lossy_fixes=options.lossy_fixes)
    for html in _as_sources(source):
        lowered = lower(
            html,
            contract=contract,
            font_strategy=_resolve_font_strategy(options.font_strategy),
            lossy_fixes=options.lossy_fixes,
        )
        _merge(report, lowered.report)
    return report


def _resolve_font_strategy(strategy: str) -> str:
    """``auto`` 按运行平台选默认字体族：Windows 用微软雅黑（目标机器必有），
    其他平台用可自由分发的 Noto/思源。文档最终在什么机器上打开，字体策略就该按
    什么机器的字体现实来选。"""
    if strategy == "auto":
        return "windows" if sys.platform.startswith("win") else "portable"
    if strategy in ("portable", "windows"):
        return strategy
    raise ValueError(f"未知字体策略 {strategy!r}，可选：auto / portable / windows")


def _as_sources(source: str | Sequence[str]) -> list[str]:
    sources = [source] if isinstance(source, str) else list(source)
    if not sources:
        raise ValueError("至少需要一份 HTML")
    return sources


def _canvas(contract: Contract) -> ir.Canvas:
    return ir.Canvas(width_px=contract.canvas.width_px, height_px=contract.canvas.height_px)


def _merge(target: Report, source: Report) -> None:
    target.errors.extend(source.errors)
    target.warnings.extend(source.warnings)
    target.fixes.extend(source.fixes)
    for key, value in source.stats.items():
        target.stats[key] = target.stats.get(key, 0) + value


def _renumber(slides: tuple[ir.Slide, ...]) -> tuple[ir.Slide, ...]:
    """跨页统一编号，避免多份 HTML 各自从 e0001 开始时撞号。"""
    counter = 1
    renumbered: list[ir.Slide] = []
    for slide in slides:
        elements: list[ir.Node] = []
        for element in slide.elements:
            elements.append(replace(element, id=f"e{counter:04d}"))
            counter += 1
        renumbered.append(ir.Slide(background=slide.background, elements=tuple(elements)))
    return tuple(renumbered)
