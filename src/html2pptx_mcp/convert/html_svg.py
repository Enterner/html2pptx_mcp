"""SVG 子集解析：rect / circle / line → 图元描述。

坐标在这里就换算成画布绝对值，遍历层拿到即可建节点。本模块不分配 id、不判限额，
也不理解 <svg> 自身的定位盒；它只回答「这个图元画成什么」。

透明度语义：形状保留 opacity 交给发射层；连接线没有独立透明度，直接乘进描边 alpha。
"""

from __future__ import annotations

from collections.abc import Callable
from dataclasses import dataclass

from bs4 import Tag

from . import ir
from .contract import Contract
from .diagnostics import Report
from .values import ValueContext, ValueRejected, parse_value


@dataclass(frozen=True, slots=True)
class SvgShape:
    line: int
    box: ir.Rect
    geometry: ir.Geometry
    fill: ir.Fill | None
    stroke: ir.Stroke | None
    opacity: float


@dataclass(frozen=True, slots=True)
class SvgConnector:
    line: int
    box: ir.Rect
    start: tuple[float, float]
    end: tuple[float, float]
    stroke: ir.Stroke


SvgItem = SvgShape | SvgConnector


def parse_svg_children(
    el: Tag,
    *,
    box: ir.Rect,
    report: Report,
    contract: Contract,
    font_strategy: str,
    path: str,
) -> list[SvgItem]:
    """解析 <svg> 的直接子元素；不支持的图元记录诊断后跳过。"""
    items: list[SvgItem] = []
    counts: dict[str, int] = {}
    for child in el.children:
        if not isinstance(child, Tag):
            continue
        name = child.name.lower()
        counts[name] = counts.get(name, 0) + 1
        child_path = f"{path}/{name}[{counts[name]}]"
        item = _parse_child(
            child,
            box=box,
            report=report,
            contract=contract,
            font_strategy=font_strategy,
            path=child_path,
        )
        if item is not None:
            items.append(item)
    return items


def _parse_child(
    el: Tag,
    *,
    box: ir.Rect,
    report: Report,
    contract: Contract,
    font_strategy: str,
    path: str,
) -> SvgItem | None:
    line = el.sourceline or 0
    name = el.name.lower()
    tag_spec = contract.tags.get(name)
    if tag_spec is None or name not in {"rect", "circle", "line"}:
        report.error(
            "SVG_UNSUPPORTED_ELEMENT",
            line=line,
            message=f"svg 内不支持 <{name}>",
            hint="svg 仅支持 <rect> / <circle> / <line>",
            path=path,
        )
        return None

    allowed = set(tag_spec.attributes)
    attrs: dict[str, str] = {}
    for key, value in el.attrs.items():
        if key not in allowed:
            report.error(
                "CONTRACT_UNKNOWN_ATTRIBUTE",
                line=line,
                message=f"<{name}> 属性 “{key}” 不在白名单",
                hint=f"允许：{'、'.join(sorted(allowed))}",
                path=path,
            )
            continue
        attrs[key] = str(value)

    def number(key: str, default: float = 0.0) -> float:
        raw = attrs.get(key)
        if raw is None:
            return default
        try:
            return float(raw)
        except ValueError:
            report.error(
                "SVG_INVALID_GEOMETRY",
                line=line,
                message=f"<{name}> 的 {key} 不是数字：“{raw}”",
                hint='SVG 属性使用无单位数值，例如 width="200"',
                path=path,
            )
            raise _InvalidGeometry from None

    try:
        opacity = _opacity(attrs, report, contract, font_strategy, line=line, path=path)
        fill = _fill(attrs, report, contract, font_strategy, line=line, path=path)
        stroke = _stroke(attrs, report, contract, font_strategy, number, line=line, path=path)
        return _geometry_item(
            name,
            attrs,
            box,
            number,
            fill,
            stroke,
            opacity,
            line,
            report=report,
            path=path,
        )
    except _InvalidGeometry:
        return None


class _InvalidGeometry(Exception):
    pass


def _opacity(
    attrs: dict[str, str],
    report: Report,
    contract: Contract,
    font_strategy: str,
    *,
    line: int,
    path: str,
) -> float:
    raw = attrs.get("opacity")
    if raw is None:
        return 1.0
    context = ValueContext(contract=contract, font_strategy=font_strategy)
    try:
        return float(parse_value(contract.properties["opacity"], raw, context).value)
    except ValueRejected as error:
        report.error(
            error.code,
            line=line,
            message=error.message,
            hint=error.hint,
            prop="opacity",
            path=path,
        )
        return 1.0


def _fill(
    attrs: dict[str, str],
    report: Report,
    contract: Contract,
    font_strategy: str,
    *,
    line: int,
    path: str,
) -> ir.Fill | None:
    if "fill" not in attrs:
        return ir.SolidFill(hex="000000")  # SVG 缺省填充是黑色
    if attrs["fill"].lower() == "none":
        return None
    context = ValueContext(contract=contract, font_strategy=font_strategy)
    try:
        parsed = parse_value(contract.properties["color"], attrs["fill"], context)
    except ValueRejected as error:
        report.error(
            error.code,
            line=line,
            message=f"fill 非法：{error.message}",
            hint=error.hint,
            prop="fill",
            path=path,
        )
        return None
    color = parsed.value
    for fix in parsed.fixes:
        report.fix(
            line=line,
            rule=fix.rule,
            prop="fill",
            value_from=fix.value_from,
            value_to=fix.value_to,
        )
    if color.hex is None:
        return None
    return ir.SolidFill(hex=color.hex, alpha=color.alpha)


def _stroke(
    attrs: dict[str, str],
    report: Report,
    contract: Contract,
    font_strategy: str,
    number: Callable[..., float],
    *,
    line: int,
    path: str,
) -> ir.Stroke | None:
    raw = attrs.get("stroke")
    if raw is None or raw.lower() == "none":
        return None
    context = ValueContext(contract=contract, font_strategy=font_strategy)
    try:
        parsed = parse_value(contract.properties["color"], raw, context)
    except ValueRejected as error:
        report.error(
            error.code,
            line=line,
            message=f"stroke 非法：{error.message}",
            hint=error.hint,
            prop="stroke",
            path=path,
        )
        raise _InvalidGeometry from None
    color = parsed.value
    for fix in parsed.fixes:
        report.fix(
            line=line,
            rule=fix.rule,
            prop="stroke",
            value_from=fix.value_from,
            value_to=fix.value_to,
        )
    if color.hex is None:
        return None
    width = number("stroke-width", 1.0)
    return ir.Stroke(width_px=width, color=ir.Color(hex=color.hex, alpha=color.alpha))


def _geometry_item(
    name: str,
    attrs: dict[str, str],
    box: ir.Rect,
    number: Callable[..., float],
    fill: ir.Fill | None,
    stroke: ir.Stroke | None,
    opacity: float,
    line: int,
    *,
    report: Report,
    path: str,
) -> SvgItem:
    if name == "rect":
        width, height = number("width"), number("height")
        if width <= 0 or height <= 0:
            report.error(
                "SVG_INVALID_GEOMETRY",
                line=line,
                message="<rect> 需要 width/height > 0",
                hint='示例：<rect x="0" y="0" width="200" height="80" fill="#4F46E5">',
                path=path,
            )
            raise _InvalidGeometry
        rect_box = ir.Rect(box.x + number("x"), box.y + number("y"), width, height)
        geometry: ir.Geometry = (
            _rect_geometry(number("rx"), rect_box) if "rx" in attrs else ir.RectGeometry()
        )
        return SvgShape(
            line=line,
            box=rect_box,
            geometry=geometry,
            fill=fill,
            stroke=stroke,
            opacity=opacity,
        )

    if name == "circle":
        radius = number("r")
        if radius <= 0:
            report.error(
                "SVG_INVALID_GEOMETRY",
                line=line,
                message="<circle> 需要 r > 0",
                hint='示例：<circle cx="12" cy="12" r="10" fill="#4F46E5">',
                path=path,
            )
            raise _InvalidGeometry
        circle_box = ir.Rect(
            box.x + number("cx") - radius,
            box.y + number("cy") - radius,
            2 * radius,
            2 * radius,
        )
        return SvgShape(
            line=line,
            box=circle_box,
            geometry=ir.EllipseGeometry(),
            fill=fill,
            stroke=stroke,
            opacity=opacity,
        )

    if stroke is None:
        report.error(
            "SVG_INVALID_GEOMETRY",
            line=line,
            message="<line> 需要 stroke 颜色",
            hint='示例：<line x1="0" y1="1" x2="1152" y2="1" stroke="#CBD5E1" stroke-width="2">',
            path=path,
        )
        raise _InvalidGeometry
    start = (box.x + number("x1"), box.y + number("y1"))
    end = (box.x + number("x2"), box.y + number("y2"))
    return SvgConnector(
        line=line,
        box=ir.Rect(
            min(start[0], end[0]),
            min(start[1], end[1]),
            abs(end[0] - start[0]),
            abs(end[1] - start[1]),
        ),
        start=start,
        end=end,
        stroke=ir.Stroke(
            width_px=stroke.width_px,
            color=ir.Color(hex=stroke.color.hex, alpha=stroke.color.alpha * opacity),
        ),
    )


def _rect_geometry(radius_px: float, box: ir.Rect) -> ir.Geometry:
    if radius_px <= 0:
        return ir.RectGeometry()
    if abs(box.width - box.height) < 0.01 and radius_px >= box.width / 2 - 0.01:
        return ir.EllipseGeometry()
    return ir.RoundRectGeometry(radius_px=min(radius_px, min(box.width, box.height) / 2))
