"""Slide IR：HTML 与 PPTX 之间的唯一中间表示。

IR 是第二个契约。HTML 只是一个前端，未来的 native 工具与模板系统都产出同一种 IR，发射器
只认 IR。因此这里的类型必须：

- **可序列化**：``to_dict`` 输出固定舍入（4 位小数）与固定键序，能直接做 golden 快照；
- **无外部依赖**：不出现 python-pptx 对象、文件句柄、路径；
- **语义完整**：几何、填充、描边、阴影、文本都已是最终形态，发射器不再做业务判断。

坐标统一为画布绝对坐标（px 浮点），元素顺序即绘制顺序（z 序）。
"""

from __future__ import annotations

from dataclasses import asdict, dataclass
from typing import Any, TypeAlias

Point: TypeAlias = tuple[float, float]


class IRInvariantError(ValueError):
    """IR 自检失败：重复 id、负尺寸等，属于实现缺陷而不是用户输入问题。"""


# ---------------------------------------------------------------- 值对象
@dataclass(frozen=True, slots=True)
class Color:
    hex: str | None  # None 表示透明
    alpha: float = 1.0


@dataclass(frozen=True, slots=True)
class SolidFill:
    hex: str
    alpha: float = 1.0


@dataclass(frozen=True, slots=True)
class GradientStop:
    hex: str
    alpha: float
    position: float  # 0..1


@dataclass(frozen=True, slots=True)
class LinearGradient:
    angle_css_deg: float
    stops: tuple[GradientStop, ...]


Fill: TypeAlias = SolidFill | LinearGradient


@dataclass(frozen=True, slots=True)
class Stroke:
    width_px: float
    color: Color
    style: str = "solid"  # solid / dashed / dotted


@dataclass(frozen=True, slots=True)
class PolygonGeometry:
    """clip-path polygon：顶点为声明盒内的相对坐标（0..1），发射期换算 EMU。"""

    vertices: tuple[tuple[float, float], ...]


@dataclass(frozen=True, slots=True)
class FlexSpec:
    """flex 简写：只保留 grow/shrink；basis 仅 auto（尺寸来自声明盒或内容）。"""

    grow: float = 0.0
    shrink: float = 1.0


@dataclass(frozen=True, slots=True)
class Shadow:
    dx_px: float
    dy_px: float
    blur_px: float
    color: Color


@dataclass(frozen=True, slots=True)
class Radius:
    """圆角：px 与百分比二选一，百分比在这里不折算，等拿到盒子再解析。"""

    px: float | None = None
    percent: float | None = None


@dataclass(frozen=True, slots=True)
class Padding:
    top: float = 0.0
    right: float = 0.0
    bottom: float = 0.0
    left: float = 0.0


@dataclass(frozen=True, slots=True)
class LineHeight:
    multiple: float | None = None
    px: float | None = None


@dataclass(frozen=True, slots=True)
class RectGeometry:
    pass


@dataclass(frozen=True, slots=True)
class RoundRectGeometry:
    radius_px: float


@dataclass(frozen=True, slots=True)
class EllipseGeometry:
    pass


Geometry: TypeAlias = RectGeometry | RoundRectGeometry | EllipseGeometry | PolygonGeometry


# ---------------------------------------------------------------- 文本
@dataclass(frozen=True, slots=True)
class TextRun:
    text: str
    bold: bool | None = None
    italic: bool | None = None
    color: str | None = None
    size_px: float | None = None
    font: str | None = None


@dataclass(frozen=True, slots=True)
class Paragraph:
    runs: tuple[TextRun, ...]
    align: str | None = None
    line_height: LineHeight | None = None
    spacing_px: float | None = None  # letter-spacing，作用于段内所有 run
    indent_px: float | None = None  # 首行缩进（CSS text-indent 语义）


@dataclass(frozen=True, slots=True)
class TableTextStyle:
    size_px: float = 16.0
    color: str | None = None
    font: str | None = None
    header_bold: bool = True


@dataclass(frozen=True, slots=True)
class CellStyle:
    """单元格级样式；未声明的字段为 None，发射器回退到表格默认。"""

    fill: SolidFill | None = None
    color: str | None = None
    bold: bool | None = None
    italic: bool | None = None
    align: str | None = None
    border: Stroke | None = None
    padding: Padding | None = None


# ---------------------------------------------------------------- 几何与来源
@dataclass(frozen=True, slots=True)
class Rect:
    x: float
    y: float
    width: float
    height: float


@dataclass(frozen=True, slots=True)
class SourceRef:
    """指回 HTML 的位置：行号 + 结构路径，诊断同时带两者。"""

    line: int
    path: str


@dataclass(frozen=True, slots=True)
class Canvas:
    width_px: float
    height_px: float


# ---------------------------------------------------------------- 元素
@dataclass(frozen=True, slots=True)
class Element:
    id: str
    box: Rect
    source: SourceRef
    rotation_deg: float = 0.0


@dataclass(frozen=True, slots=True, kw_only=True)
class TextBox(Element):
    paragraphs: tuple[Paragraph, ...] = ()
    padding: Padding = Padding()
    word_wrap: bool = True
    vertical: bool = False  # writing-mode: vertical-rl
    valign: str | None = None  # flex align-items → 文本纵向锚点（top/middle/bottom）


@dataclass(frozen=True, slots=True, kw_only=True)
class Shape(Element):
    geometry: Geometry = RectGeometry()
    fill: Fill | None = None
    stroke: Stroke | None = None
    shadow: Shadow | None = None
    opacity: float = 1.0
    paragraphs: tuple[Paragraph, ...] = ()
    padding: Padding = Padding()
    word_wrap: bool = True
    vertical: bool = False
    valign: str | None = None


@dataclass(frozen=True, slots=True, kw_only=True)
class Picture(Element):
    """图片：IR 只记声明位置与 fit；字节、自然尺寸与实际落位由发射阶段解析。

    这样 validate 可以完全零 I/O，且校验与生成共享同一份 IR。
    """

    src: str
    fit: str = "contain"
    alt: str = ""


@dataclass(frozen=True, slots=True, kw_only=True)
class Table(Element):
    rows: tuple[tuple[str, ...], ...] = ()
    col_width_px: tuple[float, ...] = ()
    header: bool = False
    text_style: TableTextStyle = TableTextStyle()
    cells: tuple[tuple[CellStyle | None, ...], ...] = ()


@dataclass(frozen=True, slots=True, kw_only=True)
class Connector(Element):
    start: Point = (0.0, 0.0)
    end: Point = (0.0, 0.0)
    stroke: Stroke = Stroke(width_px=1.0, color=Color(hex="000000"))


Node: TypeAlias = TextBox | Shape | Picture | Table | Connector


@dataclass(frozen=True, slots=True)
class Slide:
    background: Fill | None = None
    elements: tuple[Node, ...] = ()


@dataclass(frozen=True, slots=True)
class Document:
    canvas: Canvas
    slides: tuple[Slide, ...] = ()


# ---------------------------------------------------------------- 序列化
_NODE_KINDS: dict[type, str] = {
    TextBox: "text",
    Shape: "shape",
    Picture: "picture",
    Table: "table",
    Connector: "connector",
}


def _round(value: Any) -> Any:
    if isinstance(value, float):
        return round(value, 4)
    if isinstance(value, dict):
        return {key: _round(item) for key, item in value.items()}
    if isinstance(value, (list, tuple)):
        return [_round(item) for item in value]
    return value


def to_dict(document: Document) -> dict[str, Any]:
    """IR → 可 JSON 化的字典。键序、浮点舍入固定，便于逐字节比对。"""
    return {
        "canvas": _round(asdict(document.canvas)),
        "slides": [
            {
                "background": _round(asdict(slide.background)) if slide.background else None,
                "elements": [
                    {"kind": _NODE_KINDS[type(element)], **_round(asdict(element))}
                    for element in slide.elements
                ],
            }
            for slide in document.slides
        ],
    }


def element_count(document: Document) -> int:
    return sum(len(slide.elements) for slide in document.slides)


def validate_document(document: Document) -> None:
    """IR 自检。发射器与 golden 测试前调用，把实现缺陷挡在用户可见之前。"""
    if document.canvas.width_px <= 0 or document.canvas.height_px <= 0:
        raise IRInvariantError("画布宽高必须为正数")
    seen: set[str] = set()
    for slide in document.slides:
        for element in slide.elements:
            if element.id in seen:
                raise IRInvariantError(f"元素 id 重复：{element.id}")
            seen.add(element.id)
            if element.box.width < 0 or element.box.height < 0:
                raise IRInvariantError(f"{element.id} 的盒子宽高为负")
            if element.source.line < 0:
                raise IRInvariantError(f"{element.id} 的来源行号非法")
