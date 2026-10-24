"""文本溢出估算。

这是估算器，不是排版引擎：真实换行由 PowerPoint 决定，本模块只负责在生成后给出
「大概率放不下」的预警，命中后由使用者决定扩盒还是缩字号。

度量可注入：api 层默认接 ``fonts.system_metrics()``（hmtx 真实前进宽 + hhea 真实
行高，家族解析失败时逐字符回退 CJK≈1em / 拉丁≈0.55em 的启发式），因此无字体
环境下结果依然稳定；测试也可注入确定性度量。
"""

from __future__ import annotations

import re
from dataclasses import dataclass
from typing import Protocol

from . import ir

_CJK_RE = re.compile(r"[\u2e80-\u9fff\uf900-\ufaff\uff00-\uffef]")
_DEFAULT_LINE_MULTIPLE = 1.2  # 浏览器 normal 行高的近似
_OVERFLOW_THRESHOLD = 1.05


class Metrics(Protocol):
    def width_px(
        self, text: str, size_px: float, *, family: str | None = None, bold: bool = False
    ) -> float:
        """返回一段文本在给定字号下的像素宽度。"""

    def line_height_px(
        self, size_px: float, *, family: str | None = None, bold: bool = False
    ) -> float | None:
        """返回单倍行距的像素高度；None = 调用方用 1.2 倍默认。"""


@dataclass(frozen=True, slots=True)
class HeuristicMetrics:
    cjk_em: float = 1.0
    latin_em: float = 0.55

    def width_px(
        self, text: str, size_px: float, *, family: str | None = None, bold: bool = False
    ) -> float:
        del family, bold  # 启发式不区分字体
        return sum(
            (self.cjk_em if _CJK_RE.match(char) else self.latin_em) * size_px for char in text
        )

    def line_height_px(
        self, size_px: float, *, family: str | None = None, bold: bool = False
    ) -> float | None:
        del family, bold
        return None


@dataclass(frozen=True, slots=True)
class Overflow:
    element_id: str
    ratio: float
    box_px: tuple[float, float]
    estimated_text_px: float
    line: int
    path: str


def estimate_overflow(
    document: ir.Document,
    metrics: Metrics | None = None,
) -> list[Overflow]:
    """逐元素估算文本所需高度；超过盒内高 5% 即计入。"""
    active = metrics or HeuristicMetrics()
    results: list[Overflow] = []
    for slide in document.slides:
        for element in slide.elements:
            if not isinstance(element, (ir.TextBox, ir.Shape)) or not element.paragraphs:
                continue
            padding = element.padding
            inner_w = element.box.width - padding.left - padding.right
            inner_h = element.box.height - padding.top - padding.bottom
            if inner_w <= 0 or inner_h <= 0:
                continue
            # 居中惯用法（单段 line-height=盒高）是定位手段，不产生两倍文本高——
            # 与 emit 的 center_idiom 同口径，命中时按真实行距估算
            lh = element.paragraphs[0].line_height if len(element.paragraphs) == 1 else None
            idiom = (
                lh is not None
                and lh.px is not None
                and element.box.height > 0
                and abs(lh.px - element.box.height) <= max(2.0, element.box.height * 0.02)
            )
            total = 0.0
            for paragraph in element.paragraphs:
                if not any(run.text for run in paragraph.runs):
                    continue
                size = max((run.size_px or 16.0) for run in paragraph.runs)
                multiple = _line_multiple(paragraph, size, active, idiom=idiom)
                # 逐 run 度量：字号/家族/粗细混排时宽度才准；换行按 token 贪心
                tokens: list[str] = []
                widths: list[float] = []
                for run in paragraph.runs:
                    if not run.text:
                        continue
                    run_size = run.size_px or 16.0
                    for token in _tokenize(run.text):
                        tokens.append(token)
                        widths.append(
                            active.width_px(token, run_size, family=run.font, bold=bool(run.bold))
                        )
                total += _wrap_tokens(widths, inner_w) * size * multiple
            ratio = total / inner_h
            if ratio > _OVERFLOW_THRESHOLD:
                results.append(
                    Overflow(
                        element_id=element.id,
                        ratio=round(ratio, 2),
                        box_px=(round(element.box.width, 1), round(element.box.height, 1)),
                        estimated_text_px=round(total, 1),
                        line=element.source.line,
                        path=element.source.path,
                    )
                )
    return results


def _line_multiple(
    paragraph: ir.Paragraph, size_px: float, metrics: Metrics, *, idiom: bool = False
) -> float:
    line_height = paragraph.line_height
    if idiom:
        # 惯用法：真实单倍行距（雅黑≈1.32、等线≈1.04），缺字体回退 1.2
        real = metrics.line_height_px(size_px)
        return real / size_px if real else _DEFAULT_LINE_MULTIPLE
    if line_height is None:
        # 默认行距用字体真实行高（等线≈1.04、雅黑≈1.32，差异足以影响判断）
        real = metrics.line_height_px(size_px)
        return real / size_px if real else _DEFAULT_LINE_MULTIPLE
    if line_height.multiple is not None:
        return line_height.multiple
    if line_height.px is not None:
        return line_height.px / size_px
    return _DEFAULT_LINE_MULTIPLE


def _tokenize(text: str) -> list[str]:
    """CJK 逐字成 token，拉丁按词（含后随空格）成 token。"""
    tokens: list[str] = []
    buffer = ""
    for char in text:
        if _CJK_RE.match(char):
            if buffer:
                tokens.append(buffer)
                buffer = ""
            tokens.append(char)
        elif char == " ":
            if buffer:
                tokens.append(buffer + " ")
                buffer = ""
        else:
            buffer += char
    if buffer:
        tokens.append(buffer)
    return tokens


def _wrap_tokens(widths: list[float], max_width: float) -> int:
    """贪心换行：CJK 逐字可断、拉丁按词；累计宽超盒宽即换行。"""
    lines = 1
    current = 0.0
    for width in widths:
        if current > 0 and current + width > max_width:
            lines += 1
            current = width
        else:
            current += width
    return lines
