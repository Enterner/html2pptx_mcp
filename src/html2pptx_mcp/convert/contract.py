"""受限 HTML 契约：数据模型、加载与自校验。

契约是「哪些 HTML/CSS 写法可以映射到 DrawingML」的唯一事实来源，正文是
``data/contract.json``（外加 colors/fonts 两张附表）。

三个设计决定：

- **不区分版本**：契约只有当前形态，没有版本协商。扩展只允许新增属性/标签/取值/修复规则，
  不允许改动既有语义；要改变语义就是改契约本身，而不是并存两套。
- **加载即校验**：引用完整性（属性→修复规则/诊断码、标签→子标签、字体默认值→白名单）
  任何一项不成立都抛 ContractDataError。白名单的执行依据写错了必须启动就炸。
- **数据与代码分离**：词汇表可被 CLI、文档生成器、模型提示词同时消费，不再有第二份白名单。
"""

from __future__ import annotations

import json
import re
from collections.abc import Mapping
from dataclasses import dataclass, field
from functools import lru_cache
from importlib import resources
from pathlib import Path
from types import MappingProxyType
from typing import Any

SEVERITIES = frozenset({"error", "warning", "info"})
STATUSES = frozenset({"stable", "experimental", "deprecated"})
GROUPS = frozenset({"layout", "paint", "text", "media", "compat"})
HEX_RE = re.compile(r"^[0-9A-F]{6}$")


class ContractDataError(ValueError):
    """契约数据不可用：语法错误、字段缺失或引用不完整。"""


# ---------------------------------------------------------------- 数据模型
@dataclass(frozen=True, slots=True)
class Canvas:
    width_px: float
    height_px: float
    emu_per_px: int
    dpi: int


@dataclass(frozen=True, slots=True)
class Units:
    allowed: tuple[str, ...]
    rejected: tuple[str, ...]


@dataclass(frozen=True, slots=True)
class PropertySpec:
    name: str
    kind: str
    group: str
    syntax: str
    status: str
    maps_to: tuple[str, ...]
    fixes: tuple[str, ...]
    rejects: tuple[str, ...]
    values: tuple[str, ...] | None = None
    # 3.11 的 dataclass 拒绝 mappingproxy 这类不可 hash 默认值，必须走 default_factory
    constraints: Mapping[str, Any] = field(default_factory=lambda: MappingProxyType({}))
    note: str = ""


@dataclass(frozen=True, slots=True)
class TagSpec:
    name: str
    category: str
    status: str
    children: tuple[str, ...] | None = None
    attributes: tuple[str, ...] = ()
    note: str = ""


@dataclass(frozen=True, slots=True)
class FixRule:
    id: str
    lossless: bool
    summary: str


@dataclass(frozen=True, slots=True)
class DiagnosticSpec:
    code: str
    severity: str
    summary: str
    hint_required: bool


@dataclass(frozen=True, slots=True)
class Limits:
    html_bytes: int
    elements_per_slide: int
    nesting_depth: int
    text_chars: int
    paragraphs: int
    table_rows: int
    table_cols: int
    image_bytes: int
    image_pixels: int


@dataclass(frozen=True, slots=True)
class FontSet:
    strategy: str
    default: str
    whitelist: Mapping[str, str]


@dataclass(frozen=True, slots=True)
class Contract:
    name: str
    status: str
    canvas: Canvas
    units: Units
    kinds: Mapping[str, str]
    tags: Mapping[str, TagSpec]
    properties: Mapping[str, PropertySpec]
    fix_rules: Mapping[str, FixRule]
    diagnostics: Mapping[str, DiagnosticSpec]
    limits: Limits
    fonts: Mapping[str, FontSet]
    colors: Mapping[str, str]

    def font_set(self, strategy: str) -> FontSet:
        try:
            return self.fonts[strategy]
        except KeyError:
            raise KeyError(
                f"未知字体策略 {strategy!r}，可选：{'、'.join(sorted(self.fonts))}"
            ) from None

    def capabilities(self) -> dict[str, Any]:
        """能力清单：给客户端与模型做特性探测，不含逐条语法说明。"""
        return {
            "contract": self.name,
            "status": self.status,
            "canvas": _canvas_dict(self.canvas),
            "units": {"allowed": list(self.units.allowed), "rejected": list(self.units.rejected)},
            "kinds": dict(self.kinds),
            "tags": list(self.tags),
            "properties": list(self.properties),
            "fix_rules": list(self.fix_rules),
            "limits": _limits_dict(self.limits),
        }

    def to_dict(self) -> dict[str, Any]:
        """完整契约（文档生成与 CLI 导出用）。键顺序即数据文件顺序，保证输出可 diff。"""
        return {
            "contract": self.name,
            "status": self.status,
            "canvas": _canvas_dict(self.canvas),
            "units": {"allowed": list(self.units.allowed), "rejected": list(self.units.rejected)},
            "kinds": dict(self.kinds),
            "tags": {name: _tag_dict(spec) for name, spec in self.tags.items()},
            "properties": {name: _property_dict(spec) for name, spec in self.properties.items()},
            "fix_rules": {rule_id: _fix_dict(spec) for rule_id, spec in self.fix_rules.items()},
            "diagnostics": {
                code: _diagnostic_dict(spec) for code, spec in self.diagnostics.items()
            },
            "limits": _limits_dict(self.limits),
            "fonts": {
                name: {"default": font.default, "whitelist": dict(font.whitelist)}
                for name, font in self.fonts.items()
            },
            "colors": dict(self.colors),
        }


def _canvas_dict(canvas: Canvas) -> dict[str, Any]:
    return {
        "width_px": canvas.width_px,
        "height_px": canvas.height_px,
        "emu_per_px": canvas.emu_per_px,
        "dpi": canvas.dpi,
    }


def _limits_dict(limits: Limits) -> dict[str, int]:
    return {
        "html_bytes": limits.html_bytes,
        "elements_per_slide": limits.elements_per_slide,
        "nesting_depth": limits.nesting_depth,
        "text_chars": limits.text_chars,
        "paragraphs": limits.paragraphs,
        "table_rows": limits.table_rows,
        "table_cols": limits.table_cols,
        "image_bytes": limits.image_bytes,
        "image_pixels": limits.image_pixels,
    }


def _property_dict(spec: PropertySpec) -> dict[str, Any]:
    data: dict[str, Any] = {
        "kind": spec.kind,
        "group": spec.group,
        "syntax": spec.syntax,
        "status": spec.status,
        "maps_to": list(spec.maps_to),
        "fixes": list(spec.fixes),
        "rejects": list(spec.rejects),
    }
    if spec.values is not None:
        data["values"] = list(spec.values)
    if spec.constraints:
        data["constraints"] = dict(spec.constraints)
    if spec.note:
        data["note"] = spec.note
    return data


def _tag_dict(spec: TagSpec) -> dict[str, Any]:
    data: dict[str, Any] = {
        "category": spec.category,
        "status": spec.status,
    }
    if spec.children is not None:
        data["children"] = list(spec.children)
    if spec.attributes:
        data["attributes"] = list(spec.attributes)
    if spec.note:
        data["note"] = spec.note
    return data


def _fix_dict(spec: FixRule) -> dict[str, Any]:
    return {"lossless": spec.lossless, "summary": spec.summary}


def _diagnostic_dict(spec: DiagnosticSpec) -> dict[str, Any]:
    return {
        "severity": spec.severity,
        "hint_required": spec.hint_required,
        "summary": spec.summary,
    }


# ---------------------------------------------------------------- 读文件
def _no_duplicate_keys(pairs: list[tuple[str, Any]]) -> dict[str, Any]:
    out: dict[str, Any] = {}
    for key, value in pairs:
        if key in out:
            raise ContractDataError(f"契约数据存在重复键 {key!r}")
        out[key] = value
    return out


def _read_json(path: Any) -> Any:
    try:
        text = path.read_text(encoding="utf-8")
    except OSError as exc:  # 打包后缺失文件是部署事故，不能当普通缺数据
        raise ContractDataError(f"无法读取契约数据 {path}：{exc}") from exc
    try:
        return json.loads(text, object_pairs_hook=_no_duplicate_keys)
    except json.JSONDecodeError as exc:
        raise ContractDataError(f"契约数据 {path} 不是合法 JSON：{exc}") from exc


# ---------------------------------------------------------------- 取值断言
def _mapping(value: Any, path: str) -> dict[str, Any]:
    if not isinstance(value, dict):
        raise ContractDataError(f"{path} 必须是对象，实际是 {type(value).__name__}")
    return value


def _string(data: dict[str, Any], key: str, path: str) -> str:
    value = data.get(key)
    if not isinstance(value, str) or not value.strip():
        raise ContractDataError(f"{path}.{key} 必须是非空字符串")
    return value


def _integer(data: dict[str, Any], key: str, path: str) -> int:
    value = data.get(key)
    if not isinstance(value, int) or isinstance(value, bool):
        raise ContractDataError(f"{path}.{key} 必须是整数")
    return value


def _number(data: dict[str, Any], key: str, path: str) -> float:
    value = data.get(key)
    if not isinstance(value, (int, float)) or isinstance(value, bool):
        raise ContractDataError(f"{path}.{key} 必须是数字")
    return float(value)


def _boolean(data: dict[str, Any], key: str, path: str) -> bool:
    value = data.get(key)
    if not isinstance(value, bool):
        raise ContractDataError(f"{path}.{key} 必须是布尔值")
    return value


def _string_tuple(data: dict[str, Any], key: str, path: str) -> tuple[str, ...]:
    value = data.get(key)
    if not isinstance(value, list) or not all(isinstance(item, str) for item in value):
        raise ContractDataError(f"{path}.{key} 必须是字符串数组")
    return tuple(value)


def _optional_string_tuple(data: dict[str, Any], key: str, path: str) -> tuple[str, ...] | None:
    if key not in data:
        return None
    return _string_tuple(data, key, path)


# ---------------------------------------------------------------- 分块构造
def _build_property(name: str, raw: Any, kinds: set[str]) -> PropertySpec:
    path = f"properties.{name}"
    data = _mapping(raw, path)
    kind = _string(data, "kind", path)
    if kind not in kinds:
        raise ContractDataError(f"{path}.kind={kind!r} 未在 kinds 中登记")
    group = _string(data, "group", path)
    if group not in GROUPS:
        raise ContractDataError(f"{path}.group={group!r} 不在 {sorted(GROUPS)}")
    status = _string(data, "status", path)
    if status not in STATUSES:
        raise ContractDataError(f"{path}.status={status!r} 不在 {sorted(STATUSES)}")
    constraints = data.get("constraints", {})
    if not isinstance(constraints, dict):
        raise ContractDataError(f"{path}.constraints 必须是对象")
    note = data.get("note", "")
    return PropertySpec(
        name=name,
        kind=kind,
        group=group,
        syntax=_string(data, "syntax", path),
        status=status,
        maps_to=_string_tuple(data, "maps_to", path),
        fixes=_string_tuple(data, "fixes", path),
        rejects=_string_tuple(data, "rejects", path),
        values=_optional_string_tuple(data, "values", path),
        constraints=constraints,
        note=note if isinstance(note, str) else "",
    )


def _build_tag(name: str, raw: Any) -> TagSpec:
    path = f"tags.{name}"
    data = _mapping(raw, path)
    status = _string(data, "status", path)
    if status not in STATUSES:
        raise ContractDataError(f"{path}.status={status!r} 不在 {sorted(STATUSES)}")
    attributes = data.get("attributes", [])
    if not isinstance(attributes, list) or not all(isinstance(a, str) for a in attributes):
        raise ContractDataError(f"{path}.attributes 必须是字符串数组")
    note = data.get("note", "")
    return TagSpec(
        name=name,
        category=_string(data, "category", path),
        status=status,
        children=_optional_string_tuple(data, "children", path),
        attributes=tuple(attributes),
        note=note if isinstance(note, str) else "",
    )


def _build_fix(rule_id: str, raw: Any) -> FixRule:
    path = f"fix_rules.{rule_id}"
    data = _mapping(raw, path)
    return FixRule(
        id=rule_id,
        lossless=_boolean(data, "lossless", path),
        summary=_string(data, "summary", path),
    )


def _build_diagnostic(code: str, raw: Any) -> DiagnosticSpec:
    path = f"diagnostics.{code}"
    data = _mapping(raw, path)
    severity = _string(data, "severity", path)
    if severity not in SEVERITIES:
        raise ContractDataError(f"{path}.severity={severity!r} 不在 {sorted(SEVERITIES)}")
    return DiagnosticSpec(
        code=code,
        severity=severity,
        summary=_string(data, "summary", path),
        hint_required=_boolean(data, "hint_required", path),
    )


def _build_limits(raw: Any) -> Limits:
    data = _mapping(raw, "limits")
    fields = (
        "html_bytes",
        "elements_per_slide",
        "nesting_depth",
        "text_chars",
        "paragraphs",
        "table_rows",
        "table_cols",
        "image_bytes",
        "image_pixels",
    )
    values = {name: _integer(data, name, "limits") for name in fields}
    for name, value in values.items():
        if value <= 0:
            raise ContractDataError(f"limits.{name} 必须为正数")
    return Limits(**values)


def _build_fonts(raw: Any) -> dict[str, FontSet]:
    data = _mapping(raw, "fonts")
    strategies = _mapping(data.get("strategies"), "fonts.strategies")
    if not strategies:
        raise ContractDataError("fonts.strategies 至少需要一个策略")
    result: dict[str, FontSet] = {}
    for name, entry in strategies.items():
        path = f"fonts.strategies.{name}"
        spec = _mapping(entry, path)
        default = _string(spec, "default", path)
        whitelist_raw = _mapping(spec.get("whitelist"), f"{path}.whitelist")
        whitelist = {str(key).strip().lower(): str(value) for key, value in whitelist_raw.items()}
        if default not in whitelist.values():
            raise ContractDataError(
                f"{path}.default={default!r} 不在该策略的白名单取值中；"
                "默认字体必须是可解析到的实际字体名"
            )
        result[name] = FontSet(strategy=name, default=default, whitelist=whitelist)
    return result


def _build_colors(raw: Any) -> dict[str, str]:
    data = _mapping(raw, "colors")
    named = _mapping(data.get("named"), "colors.named")
    result: dict[str, str] = {}
    for name, value in named.items():
        if not isinstance(value, str) or not HEX_RE.match(value):
            raise ContractDataError(f"colors.named.{name} 必须是 6 位大写十六进制")
        result[str(name).strip().lower()] = value
    return result


# ---------------------------------------------------------------- 引用校验
def _validate_references(
    properties: dict[str, PropertySpec],
    tags: dict[str, TagSpec],
    fixes: dict[str, FixRule],
    diagnostics: dict[str, DiagnosticSpec],
) -> None:
    for name, spec in properties.items():
        for rule_id in spec.fixes:
            if rule_id not in fixes:
                raise ContractDataError(f"属性 {name} 引用了未登记的修复规则 {rule_id}")
        for code in spec.rejects:
            if code not in diagnostics:
                raise ContractDataError(f"属性 {name} 引用了未登记的错误码 {code}")
    for name, tag in tags.items():
        for child in tag.children or ():
            if child not in tags:
                raise ContractDataError(f"标签 {name} 的 children 引用了未登记标签 {child}")


# ---------------------------------------------------------------- 主流程
def _package_data_root() -> Any:
    return resources.files("html2pptx_mcp.convert").joinpath("data")


def _load(root_key: str) -> Contract:
    root = _package_data_root() if root_key == _PACKAGE_ROOT_KEY else Path(root_key)
    raw = _mapping(_read_json(root.joinpath("contract.json")), "contract.json")

    canvas_raw = _mapping(raw.get("canvas"), "canvas")
    canvas = Canvas(
        width_px=_number(canvas_raw, "width_px", "canvas"),
        height_px=_number(canvas_raw, "height_px", "canvas"),
        emu_per_px=_integer(canvas_raw, "emu_per_px", "canvas"),
        dpi=_integer(canvas_raw, "dpi", "canvas"),
    )
    if canvas.width_px <= 0 or canvas.height_px <= 0:
        raise ContractDataError("canvas 宽高必须为正数")

    units_raw = _mapping(raw.get("units"), "units")
    units = Units(
        allowed=_string_tuple(units_raw, "allowed", "units"),
        rejected=_string_tuple(units_raw, "rejected", "units"),
    )

    kinds_raw = _mapping(raw.get("kinds"), "kinds")
    kinds = {str(key): str(value) for key, value in kinds_raw.items()}

    tags = {
        str(name): _build_tag(str(name), spec)
        for name, spec in _mapping(raw.get("tags"), "tags").items()
    }
    properties = {
        str(name): _build_property(str(name), spec, set(kinds))
        for name, spec in _mapping(raw.get("properties"), "properties").items()
    }
    fixes = {
        str(rule_id): _build_fix(str(rule_id), spec)
        for rule_id, spec in _mapping(raw.get("fix_rules"), "fix_rules").items()
    }
    diagnostics = {
        str(code): _build_diagnostic(str(code), spec)
        for code, spec in _mapping(raw.get("diagnostics"), "diagnostics").items()
    }
    _validate_references(properties, tags, fixes, diagnostics)

    colors_file = _string(raw, "colors_file", "contract")
    fonts_file = _string(raw, "fonts_file", "contract")
    colors = _build_colors(_read_json(root.joinpath(colors_file)))
    fonts = _build_fonts(_read_json(root.joinpath(fonts_file)))

    return Contract(
        name=_string(raw, "contract", "contract"),
        status=_string(raw, "status", "contract"),
        canvas=canvas,
        units=units,
        kinds=kinds,
        tags=tags,
        properties=properties,
        fix_rules=fixes,
        diagnostics=diagnostics,
        limits=_build_limits(raw.get("limits")),
        fonts=fonts,
        colors=colors,
    )


_PACKAGE_ROOT_KEY = "<package>"


@lru_cache(maxsize=4)
def _load_cached(root_key: str) -> Contract:
    return _load(root_key)


def load_contract(*, data_root: Path | None = None) -> Contract:
    """加载契约。

    契约只有当前形态，不区分版本：扩展只增不改语义。``data_root`` 仅测试与外部工具使用。
    数据不自洽时抛 ContractDataError——调用方不应吞掉它。
    """
    key = _PACKAGE_ROOT_KEY if data_root is None else str(data_root.resolve())
    return _load_cached(key)


def clear_cache() -> None:
    """清空加载缓存（测试用：换数据根或改数据后需要重新读盘）。"""
    _load_cached.cache_clear()
