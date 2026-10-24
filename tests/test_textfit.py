"""textfit 单测：换行、行距口径与居中惯用法。

``real_metrics`` 夹具用一套确定性的假度量（等线行高 1.04、雅黑 1.32 的近似）
驱动真实代码路径；不依赖本机字体，CI 稳定。
"""

from __future__ import annotations

import dataclasses

import pytest

from html2pptx_mcp.convert import ir
from html2pptx_mcp.convert.textfit import HeuristicMetrics, estimate_overflow


@dataclasses.dataclass(frozen=True)
class FakeMetrics:
    """确定性度量：advance 可注入，行高固定 1.32 倍（雅黑口径）。"""

    advance_em: float = 1.0

    def width_px(self, text: str, size_px: float, *, family=None, bold=False) -> float:
        return len(text) * self.advance_em * size_px

    def line_height_px(self, size_px: float, *, family=None, bold=False) -> float:
        return size_px * 1.32


def _box(x=0.0, y=0.0, w=400.0, h=100.0) -> ir.Rect:
    return ir.Rect(x=x, y=y, width=w, height=h)


def _textbox(
    text: str,
    *,
    size=16.0,
    font=None,
    bold=False,
    box=None,
    line_height: ir.LineHeight | None = None,
    padding: ir.Padding | None = None,
) -> ir.TextBox:
    return ir.TextBox(
        id="e0001",
        box=box or _box(),
        source=ir.SourceRef(line=1, path="t.html"),
        paragraphs=(
            ir.Paragraph(
                runs=(ir.TextRun(text=text, size_px=size, font=font, bold=bold),),
                line_height=line_height,
            ),
        ),
        padding=padding or ir.Padding(),
    )


def _overflow_ratio(element: ir.TextBox, metrics=None) -> float:
    document = ir.Document(
        canvas=ir.Canvas(width_px=1280, height_px=720),
        slides=(ir.Slide(background=None, elements=(element,)),),
    )
    overflows = estimate_overflow(document, metrics)
    assert len(overflows) <= 1
    return overflows[0].ratio if overflows else 0.0


def test_single_line_no_overflow() -> None:
    element = _textbox("短文本")
    assert _overflow_ratio(element) == 0.0


def test_cjk_wraps_per_char() -> None:
    # HeuristicMetrics CJK=1em：40 字 × 16px = 640px，内宽 400 → 25 字/行 → 2 行
    # 高 = 2×16×1.2 = 38.4，盒高 30 → ratio 1.28
    element = _textbox("构" * 40, box=_box(h=30.0))
    assert _overflow_ratio(element, HeuristicMetrics()) == pytest.approx(38.4 / 30)


def test_latin_wraps_by_word() -> None:
    # 拉丁按词断："word " 宽 5×0.55×16 = 44px，内宽 400 → 9 词/行 → 30 词 4 行
    # 高 = 4×16×1.2 = 76.8，盒高 50 → ratio 1.536（estimate 四舍五入到 2 位）
    element = _textbox("word " * 30, box=_box(h=50.0))
    assert _overflow_ratio(element, HeuristicMetrics()) == pytest.approx(1.54)


def test_idiom_line_height_not_double_counted() -> None:
    """回归：line-height=盒高 的居中惯用法曾被当双倍行距，假预警 6 条。

    单段 + lh=盒高 → 真实行距（1.32×16=21.1）而非 2×16=32。
    """
    element = _textbox(
        "正文一行",
        box=ir.Rect(x=0, y=0, width=200, height=44),
        line_height=ir.LineHeight(px=44.0),
    )
    ratio = _overflow_ratio(element, FakeMetrics())
    # 单行 × 16 × 1.32 = 21.1，内高 44 → 不足半盒
    assert ratio < 0.6


def test_explicit_multiple_line_height_respected() -> None:
    element = _textbox(
        "文字",
        line_height=ir.LineHeight(multiple=2.0),
        box=ir.Rect(x=0, y=0, width=100, height=25),
    )
    ratio = _overflow_ratio(element, FakeMetrics(advance_em=1.0))
    # 一行 × 16 × 2 = 32，盒高 25 → 1.28
    assert ratio == pytest.approx(32 / 25)


def test_per_run_size_and_family_measured() -> None:
    """混排：run 字号/家族不同时逐 run 计宽，不取整段最大值一刀切。"""
    metrics = FakeMetrics(advance_em=1.0)
    element = ir.TextBox(
        id="e0001",
        box=_box(w=100.0, h=20.0),
        source=ir.SourceRef(line=1, path="t.html"),
        paragraphs=(
            ir.Paragraph(
                runs=(
                    ir.TextRun(text="大字", size_px=20.0, font="F", bold=True),
                    ir.TextRun(text="小字", size_px=10.0, font="F", bold=False),
                ),
            ),
        ),
    )
    ratio = _overflow_ratio(element, metrics)
    # 宽 = (2×20 + 2×10) × 1.0 = 60 → 一行；行高按最大字号 20×1.32 = 26.4
    assert ratio == pytest.approx(26.4 / 20)


def test_padding_shrinks_inner_box() -> None:
    element = _textbox(
        "字" * 20,
        box=ir.Rect(x=0, y=0, width=200, height=60),
        padding=ir.Padding(top=10.0, right=50.0, bottom=10.0, left=50.0),
    )
    # 内宽 100：每行 6 字 → 20 字 4 行 → 高 4×16×1.2 = 76.8 / 内高 40 = 1.92
    ratio = _overflow_ratio(element, HeuristicMetrics())
    assert ratio == pytest.approx(1.92)


def test_non_text_elements_skipped() -> None:
    document = ir.Document(
        canvas=ir.Canvas(width_px=1280, height_px=720),
        slides=(
            ir.Slide(
                background=None,
                elements=(
                    ir.Picture(
                        id="e0001",
                        box=_box(),
                        source=ir.SourceRef(line=1, path="t.html"),
                        src="a.png",
                    ),
                ),
            ),
        ),
    )
    assert estimate_overflow(document) == []


def test_real_metrics_protocol_shape() -> None:
    """真字体度量实现满足协议：家族未解析时回退启发式。"""
    from html2pptx_mcp.convert.fonts import FontFileMetrics

    metrics = FontFileMetrics()
    width = metrics.width_px("构建", 24.0, family="不存在的字体家族")
    assert width == pytest.approx(2 * 24.0)  # 回退 CJK=1em


def test_line_height_none_falls_back_to_default() -> None:
    metrics = HeuristicMetrics()
    assert metrics.line_height_px(16.0) is None


def test_dataclass_freeze() -> None:
    metrics = HeuristicMetrics()
    with pytest.raises(AttributeError):
        metrics.cjk_em = 2.0  # type: ignore[misc]
