"""CSS 单值解析：把契约里登记的写法解析成 IR 值对象。

本模块是纯函数层：不读文件、不碰 DOM、不碰 python-pptx。每个解析器返回
``Parsed(value, fixes)``，或者抛 ``ValueRejected(code, message, hint)``——没有第三种
结局。自动修复（百分比转 px、rgb 转 hex、字体回落）在这里被记录成具体规则 id，
由上层写入报告，避免「谁改了什么」散落在各个调用点。

所有 hint 都必须是可照抄的改法，而不是「请检查语法」。
"""

from __future__ import annotations

import re
from collections.abc import Callable
from dataclasses import dataclass
from typing import Any, NoReturn

from . import ir
from .contract import Contract, PropertySpec


# ---------------------------------------------------------------- 结果与错误
@dataclass(frozen=True, slots=True)
class ValueFix:
    rule: str
    value_from: str
    value_to: str


@dataclass(frozen=True, slots=True)
class Parsed:
    value: Any
    fixes: tuple[ValueFix, ...] = ()


class ValueRejected(Exception):
    """单个值不合法：code + 人读消息 + 可照抄的 hint。"""

    def __init__(self, code: str, message: str, hint: str):
        super().__init__(message)
        self.code = code
        self.message = message
        self.hint = hint


@dataclass(frozen=True, slots=True)
class ValueContext:
    """解析一个值需要的全部外部信息；显式传入，避免解析器读全局状态。"""

    contract: Contract
    font_strategy: str = "portable"
    ref_width_px: float | None = None
    ref_height_px: float | None = None
    font_px: float = 16.0
    viewport_width_px: float | None = None
    viewport_height_px: float | None = None


Parser = Callable[[str, PropertySpec, ValueContext], Parsed]

INVALID_VALUE = "CONTRACT_UNSUPPORTED_VALUE"
INVALID_UNIT = "CONTRACT_UNSUPPORTED_UNIT"


# ---------------------------------------------------------------- 词法工具
def split_top(text: str, sep: str) -> list[str]:
    """按 sep 切分，但跳过括号与引号内部；CSS 值里的逗号/空格常藏在函数里。"""
    parts: list[str] = []
    current: list[str] = []
    depth = 0
    quote: str | None = None
    for char in text:
        if quote:
            current.append(char)
            if char == quote:
                quote = None
            continue
        if char in "\"'":
            quote = char
            current.append(char)
            continue
        if char == "(":
            depth += 1
        elif char == ")":
            depth = max(0, depth - 1)
        if char == sep and depth == 0:
            parts.append("".join(current))
            current = []
        else:
            current.append(char)
    parts.append("".join(current))
    return [part.strip() for part in parts if part.strip()]


def _tokens(text: str) -> list[str]:
    return split_top(text, " ")


def _fmt(number: float) -> str:
    return f"{number:g}"


def _reject(code: str, message: str, hint: str) -> NoReturn:
    raise ValueRejected(code, message, hint)


_HEX_RE = re.compile(r"^#([0-9a-f]{3}|[0-9a-f]{6})$", re.I)
_FUNC_RE = re.compile(r"^(rgba?|hsla?)\((.*)\)$", re.I)
_LENGTH_RE = re.compile(r"^(-?\d+(?:\.\d+)?)(px|em|pt|%|vw|vh)?$", re.I)
_GRADIENT_RE = re.compile(r"^linear-gradient\((.*)\)$", re.I | re.S)
_ANGLE_RE = re.compile(r"^(-?\d+(?:\.\d+)?)deg$", re.I)
_ROTATE_RE = re.compile(r"^rotate\((-?\d+(?:\.\d+)?)deg\)$", re.I)
_MULTIPLE_RE = re.compile(r"^\d+(?:\.\d+)?$")

_GRADIENT_DIRECTIONS = {
    "to top": 0.0,
    "to right": 90.0,
    "to bottom": 180.0,
    "to left": 270.0,
}


def _number(raw: str, prop: str, hint: str) -> float:
    try:
        return float(raw)
    except ValueError:
        _reject(INVALID_VALUE, f"{prop}: 需要数字，得到 “{raw}”", hint)
    raise AssertionError("不可达")


# ---------------------------------------------------------------- 颜色
def _parse_color(raw: str, ctx: ValueContext, *, prop: str) -> Parsed:
    value = raw.strip()
    lowered = value.lower()
    if lowered == "transparent":
        return Parsed(ir.Color(hex=None, alpha=0.0))

    match = _HEX_RE.match(value)
    if match:
        digits = match.group(1)
        if len(digits) == 3:
            digits = "".join(char * 2 for char in digits)
        return Parsed(ir.Color(hex=digits.upper()))

    match = _FUNC_RE.match(value)
    if match:
        function, body = match.group(1).lower(), match.group(2)
        if function.startswith("hsl"):
            _reject(
                INVALID_VALUE,
                f"{prop}: 不支持 {function}() 颜色",
                "改用 #RRGGBB 或 rgba(r,g,b,a)。示例：rgba(79,70,229,1)",
            )
        parts = split_top(body, ",")
        if len(parts) not in (3, 4):
            _reject(
                INVALID_VALUE,
                f"{prop}: {function}() 参数数量错误",
                "示例：rgba(15,23,42,0.08)",
            )
        try:
            red, green, blue = (int(float(part)) for part in parts[:3])
            alpha = float(parts[3]) if len(parts) == 4 else 1.0
        except ValueError:
            _reject(INVALID_VALUE, f"{prop}: {function}() 参数非法", "示例：rgba(15,23,42,0.08)")
        for channel in (red, green, blue):
            if not 0 <= channel <= 255:
                _reject(INVALID_VALUE, f"{prop}: RGB 分量越界", "分量取值 0–255")
        if not 0.0 <= alpha <= 1.0:
            _reject(INVALID_VALUE, f"{prop}: alpha 越界", "alpha 取值 0–1")
        return Parsed(
            ir.Color(hex=f"{red:02X}{green:02X}{blue:02X}", alpha=alpha),
            (ValueFix("fix.color.rgb-to-hex", value, "#" + f"{red:02X}{green:02X}{blue:02X}"),),
        )

    named = ctx.contract.colors.get(lowered)
    if named is not None:
        return Parsed(
            ir.Color(hex=named),
            (ValueFix("fix.color.named-to-hex", value, "#" + named),),
        )

    _reject(
        INVALID_VALUE,
        f"{prop}: 无法解析的颜色 “{value}”",
        "使用 #RRGGBB；rgb()/rgba()/具名色会被自动归一为 #RRGGBB。示例：color: #0F172A",
    )
    raise AssertionError("不可达")


# ---------------------------------------------------------------- 长度
def _percent_ref(spec: PropertySpec, ctx: ValueContext) -> float | None:
    ref = spec.constraints.get("percent_ref")
    if ref is None:
        return None
    text = str(ref)
    if text.endswith(".width"):
        return ctx.ref_width_px
    if text.endswith(".height"):
        return ctx.ref_height_px
    if text == "parent.font_size":
        return ctx.font_px
    return None


def _parse_length(
    raw: str,
    *,
    prop: str,
    ref_px: float | None,
    font_px: float,
    allow_negative: bool = False,
    min_px: float | None = None,
    max_px: float | None = None,
    viewport_w_px: float | None = None,
    viewport_h_px: float | None = None,
) -> Parsed:
    value = raw.strip()
    match = _LENGTH_RE.match(value)
    if not match:
        _reject(
            INVALID_UNIT,
            f"{prop}: 不支持的长度写法 “{raw}”",
            f"使用 px；%/em/pt/vw/vh 会自动换算；calc()/rem 不支持。示例：{prop}: 64px",
        )
    number = float(match.group(1))
    unit = (match.group(2) or "px").lower()
    fixes: list[ValueFix] = []
    if unit == "px":
        px = number
    elif unit == "%":
        if ref_px is None:
            _reject(
                INVALID_UNIT,
                f"{prop}: 百分比缺少参照（父盒尺寸未知）",
                "改写为 px，或确保父元素有显式 width/height。示例：width: 640px",
            )
        px = ref_px * number / 100.0
        fixes.append(ValueFix("fix.unit.percent-to-px", value, f"{_fmt(px)}px"))
    elif unit == "em":
        px = number * font_px
        fixes.append(ValueFix("fix.unit.em-to-px", value, f"{_fmt(px)}px"))
    elif unit == "pt":
        px = number * 96.0 / 72.0
        fixes.append(ValueFix("fix.unit.pt-to-px", value, f"{_fmt(px)}px"))
    elif unit in ("vw", "vh"):
        viewport = viewport_w_px if unit == "vw" else viewport_h_px
        if viewport is None:
            _reject(
                INVALID_UNIT,
                f"{prop}: {unit} 需要画布尺寸（此属性未开放视口参照）",
                "改用 px，或把视口换算后再写。示例：width: 1280px",
            )
        px = viewport * number / 100.0
        fixes.append(ValueFix(f"fix.unit.{unit}-to-px", value, f"{_fmt(px)}px"))
    else:  # 正则已收口，保留兜底防御
        _reject(INVALID_UNIT, f"{prop}: 不支持的单位 “{unit}”", "使用 px")

    if px < 0 and not allow_negative:
        _reject(INVALID_VALUE, f"{prop}: 不允许负值", f"取非负值。示例：{prop}: 0")
    if min_px is not None and px < min_px:
        _reject(
            INVALID_VALUE,
            f"{prop}: {_fmt(px)}px 小于下限 {_fmt(min_px)}px",
            f"取值不小于 {_fmt(min_px)}px",
        )
    if max_px is not None and px > max_px:
        _reject(
            INVALID_VALUE,
            f"{prop}: {_fmt(px)}px 超过上限 {_fmt(max_px)}px",
            f"取值不大于 {_fmt(max_px)}px",
        )
    return Parsed(px, tuple(fixes))


# ---------------------------------------------------------------- 各 kind
def _enum(raw: str, spec: PropertySpec, _ctx: ValueContext) -> Parsed:
    value = raw.strip().lower()
    allowed = spec.values or ()
    if value not in allowed:
        hint = f"允许取值：{'、'.join(allowed)}"
        if spec.note:
            hint = f"{hint}。{spec.note}"
        if spec.name == "position":
            hint = "绝对定位模型只支持 position:absolute（坐标相对最近的定位祖先）"
        if spec.name == "object-fit":
            hint = "contain：等比缩放并留白；cover：等比缩放并裁剪超出部分"
        _reject(INVALID_VALUE, f"{spec.name}: 不支持的取值 “{raw}”", hint)
    return Parsed(value)


def _length(raw: str, spec: PropertySpec, ctx: ValueContext) -> Parsed:
    min_px = spec.constraints.get("min_px")
    max_px = spec.constraints.get("max_px")
    return _parse_length(
        raw,
        prop=spec.name,
        ref_px=_percent_ref(spec, ctx),
        font_px=ctx.font_px,
        allow_negative=bool(spec.constraints.get("allow_negative", False)),
        min_px=float(min_px) if min_px is not None else None,
        max_px=float(max_px) if max_px is not None else None,
        viewport_w_px=ctx.viewport_width_px,
        viewport_h_px=ctx.viewport_height_px,
    )


def _color(raw: str, spec: PropertySpec, ctx: ValueContext) -> Parsed:
    return _parse_color(raw, ctx, prop=spec.name)


def _gradient(raw: str, prop: str, ctx: ValueContext) -> Parsed:
    match = _GRADIENT_RE.match(raw.strip())
    if not match:
        _reject(
            INVALID_VALUE,
            f"{prop}: 无法解析的渐变 “{raw}”",
            "示例：linear-gradient(135deg, #1E1B4B 0%, #4C1D95 100%)",
        )
    parts = split_top(match.group(1), ",")
    angle = 180.0  # CSS 缺省方向 to bottom
    if parts and (_ANGLE_RE.match(parts[0]) or parts[0].lower() in _GRADIENT_DIRECTIONS):
        head = parts.pop(0)
        angle_match = _ANGLE_RE.match(head)
        angle = (
            float(angle_match.group(1)) % 360.0
            if angle_match
            else _GRADIENT_DIRECTIONS[head.lower()]
        )
    if len(parts) != 2:
        _reject(
            INVALID_VALUE,
            f"{prop}: linear-gradient 需要 2 个色标，得到 {len(parts)} 个",
            "示例：linear-gradient(135deg, #1E1B4B 0%, #4C1D95 100%)",
        )
    stops: list[ir.GradientStop] = []
    for index, part in enumerate(parts):
        tokens = _tokens(part)
        color = _parse_color(tokens[0], ctx, prop=prop).value
        if color.hex is None:
            _reject(INVALID_VALUE, f"{prop}: 渐变色标不能为透明", "使用不透明颜色")
        position = 0.0 if index == 0 else 1.0
        if len(tokens) > 1:
            position_token = tokens[1]
            if not position_token.endswith("%"):
                _reject(
                    INVALID_VALUE,
                    f"{prop}: 色标位置仅支持百分比 “{position_token}”",
                    "示例：#4C1D95 100%",
                )
            try:
                position = float(position_token[:-1]) / 100.0
            except ValueError:
                _reject(
                    INVALID_VALUE,
                    f"{prop}: 色标位置非法 “{position_token}”",
                    "示例：#4C1D95 100%",
                )
        stops.append(
            ir.GradientStop(hex=color.hex, alpha=color.alpha, position=min(max(position, 0.0), 1.0))
        )
    return Parsed(ir.LinearGradient(angle_css_deg=angle, stops=tuple(stops)))


def _fill(raw: str, spec: PropertySpec, ctx: ValueContext) -> Parsed:
    text = raw.strip()
    if "linear-gradient(" in text.lower():
        if not text.lower().startswith("linear-gradient("):
            _reject(
                INVALID_VALUE,
                f"{spec.name}: background 仅支持纯色或 linear-gradient",
                "示例：background: linear-gradient(135deg, #1E1B4B 0%, #4C1D95 100%)",
            )
        return _gradient(text, spec.name, ctx)
    try:
        color = _parse_color(raw, ctx, prop=spec.name)
    except ValueRejected:
        raise ValueRejected(
            INVALID_VALUE,
            f"{spec.name}: 不支持的填充 “{text}”",
            "仅支持纯色或两色标 linear-gradient。示例：background: "
            "linear-gradient(135deg, #1E1B4B 0%, #4C1D95 100%)",
        ) from None
    if color.value.hex is None:
        return Parsed(None, color.fixes)
    return Parsed(ir.SolidFill(hex=color.value.hex, alpha=color.value.alpha), color.fixes)


def _border(raw: str, spec: PropertySpec, ctx: ValueContext) -> Parsed:
    value = raw.strip()
    if value.lower() in ("none", "0"):
        return Parsed(None)
    parts = _tokens(value)
    if len(parts) != 3:
        _reject(
            INVALID_VALUE,
            f"{spec.name}: border 需要 “宽度 样式 颜色” 三段，得到 “{value}”",
            "示例：border: 1px solid #E2E8F0",
        )
    if parts[1].lower() not in ("solid", "dashed", "dotted"):
        _reject(
            INVALID_VALUE,
            f"{spec.name}: border 样式仅支持 solid / dashed / dotted，得到 “{parts[1]}”",
            "示例：border: 1px solid #E2E8F0",
        )
    width = _parse_length(parts[0], prop="border-width", ref_px=None, font_px=ctx.font_px, min_px=0)
    color = _parse_color(parts[2], ctx, prop=spec.name)
    return Parsed(
        ir.Stroke(
            width_px=width.value,
            color=color.value,
            style=parts[1].lower(),
        ),
        width.fixes + color.fixes,
    )


def _radius(raw: str, spec: PropertySpec, ctx: ValueContext) -> Parsed:
    parts = _tokens(raw)
    if len(parts) != 1:
        _reject(
            INVALID_VALUE,
            f"{spec.name}: border-radius 仅支持单值（四角统一），得到 “{raw}”",
            "示例：border-radius: 12px 或 50%",
        )
    value = parts[0]
    if value.endswith("%"):
        try:
            percent = float(value[:-1]) / 100.0
        except ValueError:
            _reject(INVALID_VALUE, f"{spec.name}: 百分比非法 “{value}”", "示例：border-radius: 50%")
        if not 0.0 <= percent <= 1.0:
            _reject(INVALID_VALUE, f"{spec.name}: 百分比越界 “{value}”", "取值 0–100%")
        return Parsed(ir.Radius(percent=percent))
    length = _parse_length(value, prop=spec.name, ref_px=None, font_px=ctx.font_px, min_px=0)
    return Parsed(ir.Radius(px=length.value), length.fixes)


def _opacity(raw: str, spec: PropertySpec, _ctx: ValueContext) -> Parsed:
    number = _number(raw.strip(), spec.name, "示例：opacity: 0.35")
    if not 0.0 <= number <= 1.0:
        _reject(INVALID_VALUE, f"{spec.name}: 取值越界 “{raw}”", "取值范围 0–1")
    return Parsed(number)


def _transform(raw: str, spec: PropertySpec, _ctx: ValueContext) -> Parsed:
    match = _ROTATE_RE.match(raw.strip())
    if not match:
        _reject(
            INVALID_VALUE,
            f"{spec.name}: 仅支持 rotate(<N>deg)，得到 “{raw}”",
            "示例：transform: rotate(5deg)；translate/scale/skew 不支持",
        )
    return Parsed(float(match.group(1)) % 360.0)


def _shadow(raw: str, spec: PropertySpec, ctx: ValueContext) -> Parsed:
    value = raw.strip()
    if len(split_top(value, ",")) > 1:
        _reject(
            INVALID_VALUE,
            f"{spec.name}: 仅支持单层 box-shadow",
            "示例：0 4px 12px rgba(15,23,42,0.08)",
        )
    parts = _tokens(value)
    if parts and parts[0].lower() == "inset":
        _reject(
            INVALID_VALUE,
            f"{spec.name}: 不支持 inset 阴影",
            "使用外部阴影，示例：0 4px 12px rgba(0,0,0,0.15)",
        )
    if len(parts) == 5:
        _reject(INVALID_VALUE, f"{spec.name}: 不支持 spread", "示例：0 4px 12px rgba(0,0,0,0.15)")
    if len(parts) not in (3, 4):
        _reject(
            INVALID_VALUE,
            f"{spec.name}: 需要 “dx dy [blur] 颜色”，得到 “{value}”",
            "示例：box-shadow: 0 4px 12px rgba(0,0,0,0.15)",
        )
    dx = _parse_length(parts[0], prop=spec.name, ref_px=None, font_px=ctx.font_px)
    dy = _parse_length(parts[1], prop=spec.name, ref_px=None, font_px=ctx.font_px)
    blur = Parsed(0.0)
    color_token = parts[2]
    if len(parts) == 4:
        blur = _parse_length(parts[2], prop=spec.name, ref_px=None, font_px=ctx.font_px, min_px=0)
        color_token = parts[3]
    color = _parse_color(color_token, ctx, prop=spec.name)
    if color.value.hex is None:
        _reject(INVALID_VALUE, f"{spec.name}: 阴影颜色不能为透明", "示例：rgba(15,23,42,0.08)")
    return Parsed(
        ir.Shadow(
            dx_px=dx.value,
            dy_px=dy.value,
            blur_px=blur.value,
            color=color.value,
        ),
        dx.fixes + dy.fixes + blur.fixes + color.fixes,
    )


def _font(raw: str, spec: PropertySpec, ctx: ValueContext) -> Parsed:
    if not raw.strip():
        _reject(INVALID_VALUE, f"{spec.name}: 字体族为空", "示例：font-family: 'Segoe UI', Arial")
    tokens = [token.strip().strip("\"'") for token in split_top(raw, ",")]
    fonts = ctx.contract.font_set(ctx.font_strategy)
    resolved: str | None = None
    matched_index: int | None = None
    for index, token in enumerate(tokens):
        name = fonts.whitelist.get(token.lower())
        if name:
            resolved, matched_index = name, index
            break
    if resolved is None:
        return Parsed(
            fonts.default,
            (ValueFix("fix.font.whitelist", raw, fonts.default),),
        )
    if matched_index != 0:
        return Parsed(resolved, (ValueFix("fix.font.whitelist", raw, resolved),))
    return Parsed(resolved)


def _font_weight(raw: str, spec: PropertySpec, _ctx: ValueContext) -> Parsed:
    value = raw.strip().lower()
    if value == "bold":
        return Parsed(True)
    if value == "normal":
        return Parsed(False)
    if re.fullmatch(r"[0-9]+", value):
        weight = int(value)
        threshold = int(spec.constraints.get("bold_threshold", 600))
        if 100 <= weight <= 900:
            return Parsed(weight >= threshold)
    _reject(INVALID_VALUE, f"{spec.name}: 取值非法 “{raw}”", "支持 normal / bold / 100–900")


def _line_height(raw: str, spec: PropertySpec, ctx: ValueContext) -> Parsed:
    value = raw.strip()
    if _MULTIPLE_RE.match(value):
        multiple = float(value)
        _check_multiple(multiple, spec)
        return Parsed(ir.LineHeight(multiple=multiple))
    if value.endswith("%"):
        try:
            multiple = float(value[:-1]) / 100.0
        except ValueError:
            _reject(INVALID_VALUE, f"{spec.name}: 百分比非法 “{value}”", "示例：line-height: 150%")
        _check_multiple(multiple, spec)
        return Parsed(
            ir.LineHeight(multiple=multiple),
            (ValueFix("fix.unit.percent-to-multiple", value, _fmt(multiple)),),
        )
    length = _parse_length(value, prop=spec.name, ref_px=None, font_px=ctx.font_px, min_px=0)
    return Parsed(ir.LineHeight(px=length.value), length.fixes)


def _check_multiple(multiple: float, spec: PropertySpec) -> None:
    maximum = float(spec.constraints.get("max_multiple", 5))
    if multiple <= 0 or multiple > maximum:
        _reject(
            INVALID_VALUE,
            f"{spec.name}: 倍数越界 “{multiple:g}”",
            f"取值范围 0–{_fmt(maximum)}，示例：line-height: 1.5",
        )


def _padding(raw: str, spec: PropertySpec, ctx: ValueContext) -> Parsed:
    parts = _tokens(raw)
    if not 1 <= len(parts) <= 4:
        _reject(
            INVALID_VALUE, f"{spec.name}: 需要 1–4 个值，得到 “{raw}”", "示例：padding: 24px 32px"
        )
    values = [
        _parse_length(part, prop=spec.name, ref_px=None, font_px=ctx.font_px, min_px=0)
        for part in parts
    ]
    px = [item.value for item in values]
    top = px[0]
    right = px[1] if len(px) > 1 else top
    bottom = px[2] if len(px) > 2 else top
    left = px[3] if len(px) > 3 else right
    fixes = tuple(fix for item in values for fix in item.fixes)
    return Parsed(ir.Padding(top=top, right=right, bottom=bottom, left=left), fixes)


def _letter_spacing(raw: str, spec: PropertySpec, ctx: ValueContext) -> Parsed:
    """字距：支持任意 px（0 = 无字距，负值收紧）；其他单位仍拒绝。"""
    value = raw.strip().lower()
    if value == "normal":
        return Parsed(0.0)
    length = _parse_length(
        value, prop=spec.name, ref_px=None, font_px=ctx.font_px, allow_negative=True
    )
    return Parsed(length.value, length.fixes)


def _display(raw: str, spec: PropertySpec, _ctx: ValueContext) -> Parsed:
    value = raw.strip().lower()
    if value in (spec.values or ()):
        return Parsed(value)
    if value in ("flex", "inline-flex", "grid", "inline-grid"):
        _reject(
            INVALID_VALUE,
            f"{spec.name}:{value} 不支持",
            "没有流式布局；用 position:absolute + left/top/width/height 逐个定位子元素。"
            '示例：<div style="position:absolute;left:64px;top:64px">',
        )
    _reject(
        INVALID_VALUE,
        f"{spec.name}: 不支持的取值 “{raw}”",
        f"允许取值：{'、'.join(spec.values or ())}",
    )


def _margin(raw: str, spec: PropertySpec, ctx: ValueContext) -> Parsed:
    parts = _tokens(raw)
    if not 1 <= len(parts) <= 4:
        _reject(INVALID_VALUE, f"{spec.name}: 需要 1–4 个值，得到 “{raw}”", "示例：margin: 0")
    for part in parts:
        value = _parse_length(part, prop=spec.name, ref_px=None, font_px=ctx.font_px)
        if value.value != 0:
            _reject(
                INVALID_VALUE,
                f"{spec.name} 非 0：“{raw}”",
                "定位用 left/top；居中请直接写 left。示例：left: 590px（=(1280-100)/2）",
            )
    return Parsed(None, (ValueFix("fix.compat.margin-zero-dropped", raw, "(丢弃)"),))


def _writing_mode(raw: str, spec: PropertySpec, _ctx: ValueContext) -> Parsed:
    value = raw.strip().lower()
    if value == "vertical-rl":
        return Parsed(True)
    if value in ("horizontal-tb", "horizontal-lr"):
        return Parsed(None)  # 水平排版即默认，按无操作处理
    _reject(
        INVALID_VALUE,
        f"{spec.name}: 仅支持 vertical-rl（水平排版写默认即可），得到 “{raw}”",
        "示例：writing-mode: vertical-rl",
    )


def _text_indent(raw: str, spec: PropertySpec, ctx: ValueContext) -> Parsed:
    length = _parse_length(
        raw, prop=spec.name, ref_px=None, font_px=ctx.font_px, allow_negative=True
    )
    if length.value == 0:
        return Parsed(None)
    return Parsed(length.value, length.fixes)


_CLIP_PATH_RE = re.compile(r"^polygon\((.*)\)$", re.I | re.S)
_CLIP_UNIT_RE = re.compile(r"(-?\d+(?:\.\d+)?)(%|px)?", re.I)


def _clip_path(raw: str, spec: PropertySpec, _ctx: ValueContext) -> Parsed:
    """clip-path:polygon —— 结构校验；单位（%/px）到顶点分数的换算需要声明盒，
    由 resolve_clip_path 在拿到盒之后进行。"""
    match = _CLIP_PATH_RE.match(raw.strip())
    if not match:
        _reject(
            INVALID_VALUE,
            f"{spec.name}: 仅支持 polygon(...)，得到 “{raw}”",
            "示例：clip-path: polygon(0 0, 100% 0, 100% 100%, 0 100%)",
        )
    pairs: list[tuple[tuple[float, str], tuple[float, str]]] = []
    for pair in split_top(match.group(1), ","):
        tokens = _tokens(pair)
        if len(tokens) != 2:
            _reject(
                INVALID_VALUE,
                f"{spec.name}: 顶点需要 “x y” 两个值，得到 “{pair}”",
                "示例：polygon(0 0, 100% 0, 100% 100%, 0 100%)",
            )
        resolved: list[tuple[float, str]] = []
        for token in tokens:
            unit_match = _CLIP_UNIT_RE.fullmatch(token.strip())
            if not unit_match:
                _reject(
                    INVALID_UNIT,
                    f"{spec.name}: 顶点坐标仅支持 %/px，得到 “{token}”",
                    "示例：polygon(0 0, 100% 50%, 0 100%)",
                )
            resolved.append((float(unit_match.group(1)), (unit_match.group(2) or "px").lower()))
        pairs.append((resolved[0], resolved[1]))
    if len(pairs) < 3:
        _reject(
            INVALID_VALUE,
            f"{spec.name}: polygon 至少需要 3 个顶点，得到 {len(pairs)} 个",
            "示例：polygon(0 0, 100% 50%, 0 100%)",
        )
    return Parsed(tuple(pairs))


def resolve_clip_path(
    pairs: tuple[tuple[tuple[float, str], tuple[float, str]], ...], box_w: float, box_h: float
) -> tuple[tuple[float, float], ...]:
    """把 clip-path 顶点解析为声明盒内的相对分数（0..1）。"""
    out: list[tuple[float, float]] = []
    for (x_num, x_unit), (y_num, y_unit) in pairs:
        # 零尺寸盒上的 px 顶点没有可换算基准，退化为 0（% 不受影响）。
        fx = (x_num / 100.0) if x_unit == "%" else (x_num / box_w if box_w > 0 else 0.0)
        fy = (y_num / 100.0) if y_unit == "%" else (y_num / box_h if box_h > 0 else 0.0)
        out.append((fx, fy))
    return tuple(out)


def _flex(raw: str, spec: PropertySpec, _ctx: ValueContext) -> Parsed:
    parts = raw.strip().lower().split()
    if parts == ["none"]:
        return Parsed(ir.FlexSpec(0.0, 0.0))
    if parts == ["auto"]:
        return Parsed(ir.FlexSpec(1.0, 1.0))
    if not 1 <= len(parts) <= 3:
        _reject(
            INVALID_VALUE,
            f"{spec.name}: 需要 none / auto / <grow> [<shrink> [auto]]，得到 “{raw}”",
            "示例：flex: 0 0 auto（不伸缩）/ flex: 1（占满剩余空间）",
        )
    try:
        grow = float(parts[0])
        shrink = float(parts[1]) if len(parts) > 1 else 1.0
    except ValueError:
        _reject(
            INVALID_VALUE,
            f"{spec.name}: grow/shrink 必须是数字，得到 “{raw}”",
            "示例：flex: 1 1 auto",
        )
    if grow < 0 or shrink < 0:
        _reject(INVALID_VALUE, f"{spec.name}: grow/shrink 不能为负，得到 “{raw}”", "示例：flex: 1")
    if len(parts) == 3 and parts[2] != "auto":
        _reject(
            INVALID_VALUE,
            f"{spec.name}: flex-basis 仅支持 auto，得到 “{parts[2]}”",
            "尺寸写在 width/height；flex:1 表示由布局分配剩余空间",
        )
    return Parsed(ir.FlexSpec(grow, shrink))


def _transform_origin(raw: str, spec: PropertySpec, _ctx: ValueContext) -> Parsed:
    tokens = raw.strip().lower().split()
    if tokens in (["center"], ["center", "center"]):
        return Parsed(None, (ValueFix("fix.style.transform-origin-dropped", raw, "(丢弃)"),))
    _reject(
        INVALID_VALUE,
        f"{spec.name}: 仅支持 center / center center，得到 “{raw}”",
        "DrawingML 旋转枢轴固定为形状中心；需要其他枢轴请调整元素几何",
    )


# ---------------------------------------------------------------- 注册表
REGISTRY: dict[str, Parser] = {
    "enum": _enum,
    "length": _length,
    "color": _color,
    "fill": _fill,
    "border": _border,
    "radius": _radius,
    "opacity": _opacity,
    "transform": _transform,
    "shadow": _shadow,
    "font": _font,
    "font-weight": _font_weight,
    "line-height": _line_height,
    "padding": _padding,
    "letter-spacing": _letter_spacing,
    "writing-mode": _writing_mode,
    "text-indent": _text_indent,
    "clip-path": _clip_path,
    "flex": _flex,
    "transform-origin": _transform_origin,
    "display": _display,
    "margin": _margin,
}


def registry_kinds() -> frozenset[str]:
    return frozenset(REGISTRY)


def parse_value(prop: PropertySpec, raw: str, ctx: ValueContext) -> Parsed:
    """按契约声明的 kind 解析一个值。kind 没注册视为契约与实现漂移，直接抛错。"""
    parser = REGISTRY.get(prop.kind)
    if parser is None:
        raise LookupError(f"kind {prop.kind!r} 没有注册解析器（契约与实现漂移）")
    return parser(raw, prop, ctx)
