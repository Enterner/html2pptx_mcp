"""诊断模型：错误、警告与自动修复记录。

诊断码、严重级别、修复规则是否 lossy 全部由契约数据定义；本模块只负责按契约校验后
落账。三个不变量：

- error 必须带 hint（契约里 ``hint_required``），缺了就是实现缺陷，直接抛错；
- 严重级别不能由调用方指定，只能由诊断码决定，避免同码两副面孔；
- 修复规则必须已登记，未知规则直接抛错——静默丢弃从这里就被堵死。
"""

from __future__ import annotations

from dataclasses import dataclass, field
from typing import Any

from .contract import Contract


@dataclass(frozen=True, slots=True)
class Fix:
    line: int
    rule: str
    prop: str
    value_from: str
    value_to: str
    lossless: bool

    def to_dict(self) -> dict[str, Any]:
        return {
            "line": self.line,
            "rule": self.rule,
            "prop": self.prop,
            "from": self.value_from,
            "to": self.value_to,
            "lossless": self.lossless,
        }


@dataclass(frozen=True, slots=True)
class Diagnostic:
    code: str
    severity: str
    line: int
    message: str
    hint: str
    prop: str | None = None
    path: str | None = None

    def to_dict(self) -> dict[str, Any]:
        data: dict[str, Any] = {
            "severity": self.severity,
            "code": self.code,
            "line": self.line,
            "message": self.message,
            "hint": self.hint,
        }
        if self.prop is not None:
            data["prop"] = self.prop
        if self.path is not None:
            data["path"] = self.path
        return data


@dataclass
class Report:
    """一次 lowering 的诊断账本。构造时绑定契约，严重级别与 lossless 都从契约查。

    ``lossy_fixes`` 决定有损修复的记账方式：

    - ``apply``（默认）：应用并记入 ``fixes``，由调用方决定是否接受；
    - ``warn``：仍然应用（否则产不出合法 pptx），但改记 ``FIX_LOSSY_APPLIED`` 警告，
      让严格流水线能拦住「字体回落、透明度丢弃」这类可见差异。
    """

    contract: Contract
    lossy_fixes: str = "apply"
    errors: list[Diagnostic] = field(default_factory=list)
    warnings: list[Diagnostic] = field(default_factory=list)
    fixes: list[Fix] = field(default_factory=list)
    stats: dict[str, int] = field(default_factory=dict)

    def __post_init__(self) -> None:
        if self.lossy_fixes not in ("apply", "warn"):
            raise ValueError(f"未知的有损修复策略 {self.lossy_fixes!r}，可选：apply / warn")

    @property
    def ok(self) -> bool:
        return not self.errors

    def error(
        self,
        code: str,
        *,
        line: int,
        message: str,
        hint: str,
        prop: str | None = None,
        path: str | None = None,
    ) -> None:
        self._add(code, "error", line=line, message=message, hint=hint, prop=prop, path=path)

    def warning(
        self,
        code: str,
        *,
        line: int,
        message: str,
        hint: str,
        prop: str | None = None,
        path: str | None = None,
    ) -> None:
        self._add(code, "warning", line=line, message=message, hint=hint, prop=prop, path=path)

    def fix(self, *, line: int, rule: str, prop: str, value_from: str, value_to: str) -> None:
        spec = self.contract.fix_rules.get(rule)
        if spec is None:
            raise KeyError(f"未登记的修复规则 {rule!r}（契约与实现漂移）")
        if not spec.lossless and self.lossy_fixes == "warn":
            self.warning(
                "FIX_LOSSY_APPLIED",
                line=line,
                message=f"{prop} 的有损修复已应用：{value_from} → {value_to}",
                hint="改用在契约白名单内的等价写法；或确认可以接受该差异后继续",
                prop=prop,
            )
            return
        self.fixes.append(
            Fix(
                line=line,
                rule=rule,
                prop=prop,
                value_from=value_from,
                value_to=value_to,
                lossless=spec.lossless,
            )
        )

    def _add(
        self,
        code: str,
        expected: str,
        *,
        line: int,
        message: str,
        hint: str,
        prop: str | None,
        path: str | None,
    ) -> None:
        spec = self.contract.diagnostics.get(code)
        if spec is None:
            raise KeyError(f"未登记的诊断码 {code!r}（契约与实现漂移）")
        if spec.severity != expected:
            raise ValueError(f"诊断码 {code} 的级别是 {spec.severity}，不能用 {expected} 记录")
        if spec.hint_required and not hint.strip():
            raise ValueError(f"诊断码 {code} 要求 hint，调用方没有提供")
        item = Diagnostic(
            code=code,
            severity=spec.severity,
            line=line,
            message=message,
            hint=hint,
            prop=prop,
            path=path,
        )
        (self.errors if expected == "error" else self.warnings).append(item)

    def to_dict(self) -> dict[str, Any]:
        return {
            "contract": self.contract.name,
            "ok": self.ok,
            "errors": [item.to_dict() for item in self.errors],
            "warnings": [item.to_dict() for item in self.warnings],
            "fixes": [item.to_dict() for item in self.fixes],
            "stats": dict(self.stats),
        }
