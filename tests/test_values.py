"""CSS 值解析器单测：按契约 spec 走 parse_value，覆盖合法/非法/归一化三路。

纯函数测试：无 I/O，ctx 用最小构造；这些用例是 16 类 kind 的行为快照，
改解析器前先看这里。
"""

from __future__ import annotations

import pytest

from html2pptx_mcp.convert.contract import load_contract
from html2pptx_mcp.convert.values import ValueContext, parse_value

CONTRACT = load_contract()


def ctx(**overrides) -> ValueContext:
    return ValueContext(contract=CONTRACT, **overrides)


def parse(prop: str, raw: str, **ctx_overrides):
    spec = CONTRACT.properties[prop]
    return parse_value(spec, raw, ctx(**ctx_overrides))


# ---- color ----
def test_color_hex() -> None:
    parsed = parse("color", "#074F92")
    assert parsed.value.hex == "074F92"
    assert parsed.value.alpha == 1.0
    assert parsed.fixes == ()


def test_color_rgb_normalized() -> None:
    parsed = parse("color", "rgb(7, 79, 146)")
    assert parsed.value.hex == "074F92"


def test_color_rgba_alpha_kept() -> None:
    parsed = parse("background-color", "rgba(0, 63, 136, 0.5)")
    assert parsed.value.hex == "003F88"
    assert parsed.value.alpha == pytest.approx(0.5)


def test_color_named() -> None:
    parsed = parse("color", "white")
    assert parsed.value.hex == "FFFFFF"


def test_color_invalid_rejected() -> None:
    from html2pptx_mcp.convert.values import ValueRejected

    with pytest.raises(ValueRejected) as exc:
        parse("color", "not-a-color")
    assert exc.value.code == "CONTRACT_UNSUPPORTED_VALUE"


# ---- length ----
def test_length_px() -> None:
    assert parse("left", "12px").value == 12.0


def test_length_pt_converted() -> None:
    assert parse("left", "16pt").value == pytest.approx(16 * 96 / 72)


def test_length_em_uses_context_font() -> None:
    assert parse("left", "2em", font_px=16.0).value == 32.0


def test_length_percent_uses_ref() -> None:
    assert parse("left", "50%", ref_width_px=400.0).value == 200.0


def test_length_negative_allowed_on_left() -> None:
    assert parse("left", "-8px").value == -8.0


def test_length_unknown_unit_rejected() -> None:
    from html2pptx_mcp.convert.values import ValueRejected

    with pytest.raises(ValueRejected) as exc:
        parse("left", "3rem")
    assert exc.value.code == "CONTRACT_UNSUPPORTED_UNIT"


# ---- enum ----
def test_enum_hit() -> None:
    assert parse("text-align", "center").value == "center"


def test_enum_miss_rejected() -> None:
    from html2pptx_mcp.convert.values import ValueRejected

    with pytest.raises(ValueRejected):
        parse("text-align", "middle")


# ---- border ----
def test_border_solid() -> None:
    parsed = parse("border", "1px solid #9DB8D9")
    stroke = parsed.value
    assert stroke.color.hex == "9DB8D9"
    assert stroke.style == "solid"


def test_border_dashed_style_kept() -> None:
    parsed = parse("border", "1px dashed #9DB8D9")
    assert parsed.value.style == "dashed"


# ---- 渐变与阴影 ----
def test_gradient_two_stops() -> None:
    parsed = parse("background", "linear-gradient(135deg, #29D3FF 0%, #3E7BFF 100%)")
    gradient = parsed.value
    assert len(gradient.stops) == 2
    assert gradient.stops[0].hex == "29D3FF"
    assert gradient.angle_css_deg == pytest.approx(135.0)


def test_gradient_three_stops_rejected() -> None:
    from html2pptx_mcp.convert.values import ValueRejected

    with pytest.raises(ValueRejected):
        parse(
            "background",
            "linear-gradient(90deg, #000 0%, #555 50%, #fff 100%)",
        )


def test_shadow_single_layer() -> None:
    parsed = parse("box-shadow", "0 4px 14px rgba(0, 0, 0, 0.35)")
    shadow = parsed.value
    assert shadow.blur_px == pytest.approx(14.0)
    assert shadow.color.alpha == pytest.approx(0.35)


# ---- 字体 ----
def test_font_whitelist_hit() -> None:
    parsed = parse("font-family", "微软雅黑, sans-serif", font_strategy="windows")
    assert parsed.value == "Microsoft YaHei"


def test_font_whitelist_miss_falls_back() -> None:
    # 白名单外的字体回退策略默认字体，并记 lossy fix
    parsed = parse("font-family", "Comic Sans MS")
    assert parsed.value


# ---- 其他 kind 抽样 ----
def test_letter_spacing_negative() -> None:
    assert parse("letter-spacing", "-0.5px").value == pytest.approx(-0.5)


def test_line_height_multiple() -> None:
    parsed = parse("line-height", "1.7")
    assert parsed.value.multiple == pytest.approx(1.7)


def test_writing_mode_vertical() -> None:
    parsed = parse("writing-mode", "vertical-rl")
    assert parsed.value is True


def test_transform_rotate_only() -> None:
    from html2pptx_mcp.convert.values import ValueRejected

    # 角度归一化到 0–360：-6deg → 354（PowerPoint rot 同口径）
    assert parse("transform", "rotate(-6deg)").value == pytest.approx(354.0)
    with pytest.raises(ValueRejected):
        parse("transform", "translateX(10px)")


def test_clip_path_polygon() -> None:
    parsed = parse("clip-path", "polygon(0 0, 0 100%, 100% 100%)")
    assert parsed.value is not None


def test_padding_multi_value() -> None:
    parsed = parse("padding", "4.8px 9.6px")
    padding = parsed.value
    assert padding.top == pytest.approx(4.8)
    assert padding.right == pytest.approx(9.6)


def test_opacity_range() -> None:
    from html2pptx_mcp.convert.values import ValueRejected

    assert parse("opacity", "0.5").value == pytest.approx(0.5)
    with pytest.raises(ValueRejected):
        parse("opacity", "1.5")


def test_display_block_parses_absorb_at_lowering() -> None:
    """display:block 解析为取值本身；丢弃修复记录在 lowering 层（见 test_html.py）。"""
    assert parse("display", "block").value == "block"
    assert parse("display", "none").value == "none"
