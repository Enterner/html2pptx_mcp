"""契约一致性单测：代码里用到的诊断码/修复规则必须在契约里登记。

回归：字体内嵌的 5 个新诊断码忘了登记，运行时才炸 KeyError（Report._add
的防漂移守卫）——这类漂移应该在测试期就抓住。
"""

from __future__ import annotations

import re
from pathlib import Path

from html2pptx_mcp.convert.contract import load_contract

SRC = Path(__file__).resolve().parents[1] / "src"

_CALL_RE = re.compile(r'\.(?:error|warning|fix)\(\s*"([A-Za-z0-9._\-]+)"')
_RULE_KWARG_RE = re.compile(r'rule="([A-Za-z0-9._\-]+)"')


def _py_sources() -> list[Path]:
    return sorted(SRC.rglob("*.py"))


def test_diagnostic_codes_used_are_declared() -> None:
    contract = load_contract()
    declared = set(contract.diagnostics)
    used: dict[str, list[str]] = {}
    for path in _py_sources():
        text = path.read_text(encoding="utf-8", errors="ignore")
        for match in _CALL_RE.finditer(text):
            code = match.group(1)
            if code.isupper():
                used.setdefault(code, []).append(path.name)
    unknown = {code: files for code, files in used.items() if code not in declared}
    assert not unknown, f"代码使用了契约未登记的诊断码：{unknown}"


def test_fix_rules_used_are_declared() -> None:
    contract = load_contract()
    declared = set(contract.fix_rules)
    used: dict[str, list[str]] = {}
    for path in _py_sources():
        text = path.read_text(encoding="utf-8", errors="ignore")
        for match in _RULE_KWARG_RE.finditer(text):
            rule = match.group(1)
            if rule.startswith("fix."):
                used.setdefault(rule, []).append(path.name)
    unknown = {rule: files for rule, files in used.items() if rule not in declared}
    assert not unknown, f"代码使用了契约未登记的修复规则：{unknown}"


def test_diagnostic_severity_known() -> None:
    contract = load_contract()
    for code, spec in contract.diagnostics.items():
        assert spec.severity in ("error", "warning"), f"{code} 的 severity 非法"


def test_reverse_notes_are_not_diagnostic_codes() -> None:
    """逆转的降级记录走自有 Note 通道，不得混入契约诊断码体系。"""
    import tempfile

    from html2pptx_mcp.convert.reverse import reverse_pptx

    example = Path(__file__).resolve().parents[1] / "examples" / "html" / "01_title.html"
    with tempfile.TemporaryDirectory() as tmp:
        # 先由样例编译出一个 pptx，再逆转它（样例始终在仓库里，CI 可跑）
        from html2pptx_mcp.convert import CompileOptions, compile_html

        result = compile_html(
            example.read_text(encoding="utf-8"), options=CompileOptions(base_dir=example.parent)
        )
        assert result.ok
        pptx_path = Path(tmp) / "example.pptx"
        result.save(pptx_path)
        reversed_result = reverse_pptx(pptx_path, Path(tmp) / "reversed")
        for note in reversed_result.notes:
            assert not re.fullmatch(r"[A-Z0-9_]+", note.rule), (
                f"降级规则 {note.rule!r} 长得像诊断码，应走 Note 通道而非契约诊断"
            )
