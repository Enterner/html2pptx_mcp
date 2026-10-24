"""字体解析单测：注册表展示名解析（纯函数）+ 真机解析（Windows 才跑）。

覆盖回归：bold 后缀挂在组内每个子家族上（「A Bold & B Bold」），只看串尾
会错拆成 regular——这是真实踩过的坑。
"""

from __future__ import annotations

import sys

import pytest

from html2pptx_mcp.convert.fonts import FontResolver, _parse_display_name

# ---- 展示名解析（纯函数，全平台可跑） ----


@pytest.mark.parametrize(
    ("display", "families", "bold"),
    [
        (
            "Microsoft YaHei & Microsoft YaHei UI (TrueType)",
            ["Microsoft YaHei", "Microsoft YaHei UI"],
            False,
        ),
        (
            "Microsoft YaHei Bold & Microsoft YaHei UI Bold (TrueType)",
            ["Microsoft YaHei", "Microsoft YaHei UI"],
            True,
        ),
        ("DengXian (TrueType)", ["DengXian"], False),
        ("DengXian Bold (TrueType)", ["DengXian"], True),
        ("DengXian Light (TrueType)", ["DengXian Light"], False),
        ("Arial Bold Italic (TrueType)", ["Arial"], True),
        ("华康宋体W12(P) (TrueType)", ["华康宋体W12(P)"], False),
    ],
)
def test_parse_display_name(display, families, bold) -> None:
    got_families, got_bold = _parse_display_name(display)
    assert got_families == families
    assert got_bold is bold


def test_parse_display_name_mixed_group() -> None:
    """组内一个带 Bold 一个不带：bold 取或（PowerPoint 组内风格一致，防御性）。"""
    families, bold = _parse_display_name("A & B Bold (TrueType)")
    assert families == ["A", "B"]
    assert bold is True


# ---- 真机解析（Windows 注册表 + 本机字体文件） ----
win32 = pytest.mark.skipif(sys.platform != "win32", reason="注册表与系统字体仅 Windows")


@win32
def test_resolve_yahei_regular_and_bold_distinct() -> None:
    """回归：bold 面必须命中 msyhbd.ttc，而非 regular 的 msyh.ttc。"""
    resolver = FontResolver()
    regular = resolver.resolve("Microsoft YaHei", bold=False)
    bold = resolver.resolve("Microsoft YaHei", bold=True)
    assert regular is not None and bold is not None
    assert regular.path != bold.path
    assert "msyhbd" in bold.path.name


@win32
def test_resolve_dengxian_via_chinese_alias() -> None:
    """回归：中文名「等线」靠 name 表别名命中 Deng.ttf（注册表只登记英文名）。"""
    face = FontResolver().resolve("等线", bold=False)
    assert face is not None and face.path.name == "Deng.ttf"


@win32
def test_resolve_missing_family_returns_none() -> None:
    assert FontResolver().resolve("华康宋体W12(P)", bold=False) is None
    assert FontResolver().resolve("绝不存在的字体", bold=False) is None


@win32
def test_face_metrics_shape() -> None:
    face = FontResolver().resolve("Microsoft YaHei", bold=False)
    assert face is not None
    width = face.width_px("构建桁架", 24.0)
    line_height = face.line_height_px(24.0)
    assert width > 0
    # 雅黑单倍行距 ≈ 1.32 倍字号，明显大于启发式默认 1.2
    assert line_height > 24.0 * 1.25
    assert face.fs_type() in (0, 2, 3, 8)  # 可嵌入的许可位


@win32
def test_font_file_metrics_fallback_on_unknown_family() -> None:
    from html2pptx_mcp.convert.fonts import FontFileMetrics

    metrics = FontFileMetrics()
    # 未知家族：逐字符回退启发式（CJK=1em），不抛错
    assert metrics.width_px("构建", 24.0, family="绝不存在的字体") == pytest.approx(48.0)
    assert metrics.line_height_px(24.0, family="绝不存在的字体") is None
