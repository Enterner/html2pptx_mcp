"""lowering 层单测：兼容声明（display/margin/box-sizing）的吸收与丢弃。"""

from __future__ import annotations

from html2pptx_mcp.convert.contract import load_contract
from html2pptx_mcp.convert.html import lower

CONTRACT = load_contract()


def _lower(html: str):
    return lower(html, contract=CONTRACT, lossy_fixes="apply")


def test_compat_declarations_absorbed_with_fixes() -> None:
    """display:block 丢弃、margin:0 丢弃、box-sizing 吸收——全部记 fix 不报错。"""
    html = (
        '<div style="width:1280px;height:720px;margin:0;box-sizing:border-box">'
        '<div style="position:absolute;left:0;top:0;width:100px;height:50px;'
        'display:block;box-sizing:border-box;margin:0">文本</div></div>'
    )
    lowered = _lower(html)
    assert not lowered.report.errors
    rules = {fix.rule for fix in lowered.report.fixes}
    assert "fix.compat.display-dropped" in rules
    assert "fix.compat.margin-zero-dropped" in rules
    assert "fix.compat.box-sizing-dropped" in rules


def test_display_none_prunes_subtree() -> None:
    html = (
        '<div style="width:1280px;height:720px">'
        '<div style="position:absolute;left:0;top:0;width:100px;height:50px;'
        'display:none">隐藏</div>'
        '<div style="position:absolute;left:200px;top:0;width:100px;height:50px">可见</div></div>'
    )
    lowered = _lower(html)
    assert not lowered.report.errors
    elements = lowered.document.slides[0].elements
    texts = [
        run.text
        for element in elements
        for paragraph in getattr(element, "paragraphs", ())
        for run in paragraph.runs
    ]
    assert "可见" in "".join(texts)
    assert "隐藏" not in "".join(texts)


def test_validate_and_compile_share_lowering() -> None:
    """校验与生成同源：同一份 HTML 的 errors/fixes 两侧一致。"""
    from html2pptx_mcp.convert import validate_html

    html = (
        '<div style="width:1280px;height:720px">'
        '<div style="position:absolute;left:0;top:0;width:100px;height:50px;'
        'color:#074F92">校验同源</div></div>'
    )
    validated = validate_html(html, options=None)
    lowered = _lower(html)
    assert [e.code for e in validated.errors] == [e.code for e in lowered.report.errors]
    assert [f.rule for f in validated.fixes] == [f.rule for f in lowered.report.fixes]
