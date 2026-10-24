"""受限 HTML → Slide IR：一次遍历完成结构校验、值解析与元素构建。

为什么是「一次遍历」：校验报告必须与实际生成结果同源，否则会出现「校验通过但产物不同」。
上层调用 ``compile_html`` 时复用这里的 output，``validate_html`` 只是不执行发射。

遍历纪律：

- DOM 先序 = 绘制顺序（z 序）；元素 id 按出现顺序分配 ``e0001``；
- 元素级错误只丢弃该元素并记录，继续处理兄弟节点——一次跑完拿到全部问题；
- 任何「不按字面实现」都必须落成 fix（含 lossless 标记）或 error，没有第三种出路；
- 本模块不读文件、不碰 python-pptx，因此 ``validate_html`` 可以零 I/O。
"""

from __future__ import annotations

import re
from collections.abc import Iterator
from contextlib import contextmanager
from dataclasses import dataclass, replace
from typing import Any, TypeGuard

from bs4 import BeautifulSoup, CData, Comment, Doctype, NavigableString, ProcessingInstruction, Tag

from . import ir, textfit
from .contract import Contract, load_contract
from .diagnostics import Report
from .html_svg import SvgConnector, parse_svg_children
from .html_table import parse_table
from .values import (
    ValueContext,
    ValueRejected,
    parse_value,
    resolve_clip_path,
    split_top,
)

_WS_RE = re.compile(r"\s+")
_CANVAS_TOLERANCE_PX = 0.5
_INLINE_ALLOWED = {"color", "font-size", "font-weight", "font-family"}
_NON_TEXT_NODES = (Comment, CData, Doctype, ProcessingInstruction)

_CSS_COMMENT_RE = re.compile(r"/\*.*?\*/", re.S)
_COMPOUND_PART_RE = re.compile(r"([.#]?)([A-Za-z_][\w-]*|\*)")
_SELECTOR_FORBIDDEN = set(":#[]>+~,")


@dataclass(frozen=True, slots=True)
class _Selector:
    """复合选择器链（后代选择器）：从祖先到自身，空格分隔。

    只做「祖先链」匹配，不建模兄弟组合器（`+`/`~`）与直接子代 `>`；
    命中优先级按 CSS 特异性（class 数, 标签数）再按文档序。
    """

    compounds: tuple[_Compound, ...]

    @property
    def specificity(self) -> tuple[int, int]:
        return (
            sum(len(compound.classes) for compound in self.compounds),
            sum(len(compound.tags) for compound in self.compounds),
        )

    def matches(self, el: Tag, classes_of: Any) -> bool:
        index = len(self.compounds) - 1
        if not self.compounds[index].matches(el, classes_of(el)):
            return False
        index -= 1
        node = el.parent
        while index >= 0 and node is not None:
            if isinstance(node, Tag) and self.compounds[index].matches(node, classes_of(node)):
                index -= 1
            node = node.parent
        return index < 0


def _parse_selector(part: str) -> _Selector | None:
    tokens = part.split()
    if not tokens:
        return None
    compounds: list[_Compound] = []
    for token in tokens:
        compound = _parse_compound(token)
        if compound is None:
            return None
        compounds.append(compound)
    return _Selector(tuple(compounds))


@dataclass(frozen=True, slots=True)
class _Compound:
    """单个复合选择器：标签名与 class 的组合，或通配。"""

    tags: frozenset[str]
    classes: frozenset[str]
    wildcard: bool

    def matches(self, el: Tag, classes: frozenset[str]) -> bool:
        if self.wildcard:
            return True
        if self.tags and el.name.lower() not in self.tags:
            return False
        return self.classes <= classes


def _parse_compound(part: str) -> _Compound | None:
    part = part.strip()
    if not part or any(ch in part for ch in _SELECTOR_FORBIDDEN) or " " in part:
        return None
    tags: set[str] = set()
    classes: set[str] = set()
    wildcard = False
    for marker, name in _COMPOUND_PART_RE.findall(part):
        if marker == ".":
            classes.add(name)
        elif name == "*" and not marker:
            wildcard = True
        elif marker:
            return None  # id / 属性等
        else:
            tags.add(name.lower())
    if not tags and not classes and not wildcard:
        return None
    return _Compound(frozenset(tags), frozenset(classes), wildcard)


def _split_css_rules(text: str) -> list[tuple[str, str]]:
    """按大括号切分 CSS 规则；@ 规则（含嵌套）整块跳过，由调用方记诊断。"""
    rules: list[tuple[str, str]] = []
    i, n = 0, len(text)
    while i < n:
        while i < n and text[i].isspace():
            i += 1
        sel_start = i
        while i < n and text[i] != "{":
            i += 1
        if i >= n:
            break
        selector = text[sel_start:i].strip()
        depth, body_start = 1, i + 1
        j = body_start
        while j < n and depth:
            if text[j] == "{":
                depth += 1
            elif text[j] == "}":
                depth -= 1
            j += 1
        rules.append((selector, text[body_start : j - 1].strip()))
        i = j
    return rules


@dataclass(frozen=True, slots=True)
class LoweringResult:
    document: ir.Document
    report: Report


@dataclass(frozen=True, slots=True)
class _Inherited:
    font_px: float
    color: str | None
    font: str | None
    bold: bool | None


def _is_text(node: object) -> TypeGuard[NavigableString]:
    """注释/DOCTYPE/CDATA 也是 NavigableString，但不是可见文本。"""
    return isinstance(node, NavigableString) and not isinstance(node, _NON_TEXT_NODES)


class _Abort(Exception):
    """单个元素放弃构建；诊断已经写入报告，这里只负责停止该子树。"""


def lower(
    html: str,
    *,
    contract: Contract | None = None,
    font_strategy: str = "portable",
    lossy_fixes: str = "apply",
) -> LoweringResult:
    """把一份 HTML 降级成单页 IR，并返回诊断报告。

    页面级错误（体积超限、缺根 div）会返回空页 IR；元素级错误只丢该元素。
    ``font_strategy`` 只影响字体族解析的回落结果，不改变结构；
    ``lossy_fixes="warn"`` 时，有损修复仍会应用，但记成警告而不是 fix。
    """
    active = contract or load_contract()
    return _Lowerer(active, font_strategy, lossy_fixes).run(html)


def validate_html(
    html: str,
    *,
    contract: Contract | None = None,
    font_strategy: str = "portable",
    lossy_fixes: str = "apply",
) -> Report:
    """只做校验：与 ``lower`` 共用同一次遍历，保证「校验通过」等价于「能生成」。

    零 I/O——不读图片、不写文件，因此可以在 HTTP 请求里安全地反复调用。
    """
    return lower(
        html,
        contract=contract,
        font_strategy=font_strategy,
        lossy_fixes=lossy_fixes,
    ).report


class _Lowerer:
    def __init__(self, contract: Contract, font_strategy: str, lossy_fixes: str) -> None:
        self.contract = contract
        self.font_strategy = font_strategy
        self.report = Report(contract, lossy_fixes=lossy_fixes)
        self.limits = contract.limits
        self.canvas = contract.canvas
        self.elements: list[ir.Node] = []
        self.next_id = 1
        self.element_limit_hit = False
        self.inline_tags = {name for name, tag in contract.tags.items() if tag.category == "inline"}
        self.class_rules: list[tuple[_Selector, int, str]] = []
        # flex 布局的两本账：子元素声明缓存（避免二次解析重复报错）与计算盒。
        self.flex_values: dict[int, dict[str, Any]] = {}
        self.flex_boxes: dict[int, ir.Rect] = {}
        self.flex_skip: set[int] = set()

    # ------------------------------------------------------------ 入口
    def run(self, html: str) -> LoweringResult:
        size = len(html.encode("utf-8"))
        if size > self.limits.html_bytes:
            self.report.error(
                "LIMIT_HTML_EXCEEDED",
                line=1,
                message=f"HTML 体积 {size} 字节超过上限 {self.limits.html_bytes}",
                hint="拆成多页（每页一份 HTML），或删减内联样式与内容",
            )
            return self._result(None)

        soup = BeautifulSoup(html, "html.parser")
        body = soup.body or soup
        roots = [
            child
            for child in body.children
            if isinstance(child, Tag) and child.name.lower() not in ("script", "style", "template")
        ]
        if len(roots) != 1 or roots[0].name.lower() != "div":
            line = roots[0].sourceline if roots else 0
            self.report.error(
                "STRUCTURE_ROOT_REQUIRED",
                line=line or 0,
                message=f"body 下需要恰好 1 个根 <div>（画布），发现 {len(roots)} 个元素",
                hint='<body><div style="width:1280px;height:720px;'
                'background-color:#FFFFFF">…</div></body>',
            )
            return self._result(None)

        root = roots[0]
        root_path = "body/div[1]"
        line = root.sourceline or 0
        self.class_rules = self._extract_class_rules(soup)
        stage = self._flatten_viewport(root)
        if stage is not None:
            self.report.fix(
                line=line,
                rule="fix.structure.viewport-flattened",
                prop="display",
                value_from="flex（视口居中壳）",
                value_to="(画布)",
            )
            root = stage
            root_path = "body/div[1]/div[1]"
            line = root.sourceline or 0
        values = self._declarations(
            self._style_text(root, str(root.get("style") or ""), line=line, path=root_path),
            line=line,
            path=root_path,
            ref_w=self.canvas.width_px,
            ref_h=self.canvas.height_px,
            font_px=16.0,
            root=True,
        )

        allowed = {
            "background",
            "background-color",
            "font-size",
            "color",
            "font-family",
            "width",
            "height",
        }
        for prop in values:
            if prop in allowed:
                continue
            if prop in ("left", "top"):
                self.report.error(
                    "STRUCTURE_CANVAS_MISMATCH",
                    line=line,
                    message=f"根画布原点固定为 (0,0)，不接受 {prop}:{float(values[prop]):g}px",
                    hint="根 div 就是画布，去掉 left/top；需要留白请调整子元素坐标",
                    prop=prop,
                    path=root_path,
                )
                continue
            self.report.error(
                "STRUCTURE_ROOT_PROPERTY_FORBIDDEN",
                line=line,
                message=f"根容器不接受属性 “{prop}”",
                hint="根 div 就是画布：只写背景、默认字体与画布尺寸；内容元素请作为其子元素",
                prop=prop,
                path=root_path,
            )

        for key, expected in (("width", self.canvas.width_px), ("height", self.canvas.height_px)):
            if key in values and abs(float(values[key]) - expected) > _CANVAS_TOLERANCE_PX:
                self.report.error(
                    "STRUCTURE_CANVAS_MISMATCH",
                    line=line,
                    message=f"根容器 {key}:{values[key]:g}px 与画布 {expected:g}px 不符",
                    hint=f"画布固定 {self.canvas.width_px:g}×{self.canvas.height_px:g}px，"
                    "该声明可省略",
                    prop=key,
                    path=root_path,
                )

        background = self._root_background(values)
        inherited = _Inherited(
            font_px=float(values.get("font-size", 16.0)),
            color=values["color"].hex if isinstance(values.get("color"), ir.Color) else None,
            font=values.get("font-family"),
            bold=values.get("font-weight"),
        )
        for child in root.children:
            if _is_text(child) and child.strip():
                self.report.error(
                    "STRUCTURE_ORPHAN_TEXT",
                    line=line,
                    message="画布下存在游离文本，没有定位信息",
                    hint='文本必须放在定位的 div 里。示例：<div style="position:absolute;'
                    'left:80px;top:64px;width:600px;height:40px;font-size:24px">标题</div>',
                    path=root_path,
                )
        self._walk(
            root,
            ir.Rect(0.0, 0.0, self.canvas.width_px, self.canvas.height_px),
            root_path,
            inherited,
            depth=0,
        )
        return self._result(background)

    @contextmanager
    def _silent(self) -> Iterator[None]:
        """试解析（探测视口壳）时不污染诊断账本。"""
        original = self.report
        self.report = Report(self.contract)
        try:
            yield
        finally:
            self.report = original

    def _flatten_viewport(self, root: Tag) -> Tag | None:
        """识别「视口居中壳」：100vw/100vh flex 容器里恰好一个画布尺寸的 stage。

        这是演示页常见脚手架（等比缩放由浏览器脚本完成），对固定画布是无损的：
        直接把 stage 当根画布，居中壳整体消失。
        """
        children = [
            child
            for child in root.children
            if isinstance(child, Tag) and child.name.lower() not in ("script", "style", "template")
        ]
        if len(children) != 1 or children[0].name.lower() != "div":
            return None
        stage = children[0]
        with self._silent():
            root_values = self._declarations(
                self._style_text(root, str(root.get("style") or ""), line=0, path="body/div[1]"),
                line=0,
                path="body/div[1]",
                ref_w=self.canvas.width_px,
                ref_h=self.canvas.height_px,
                font_px=16.0,
            )
            stage_values = self._declarations(
                self._style_text(
                    stage,
                    str(stage.get("style") or ""),
                    line=stage.sourceline or 0,
                    path="body/div[1]/div[1]",
                ),
                line=stage.sourceline or 0,
                path="body/div[1]/div[1]",
                ref_w=self.canvas.width_px,
                ref_h=self.canvas.height_px,
                font_px=16.0,
            )
        if root_values.get("display") != "flex":
            return None
        allowed = {
            "display",
            "align-items",
            "justify-content",
            "width",
            "height",
            "padding",
            "margin",
        }
        if set(root_values) - allowed:
            return None
        padding = root_values.get("padding")
        if isinstance(padding, ir.Padding) and any(
            (padding.top, padding.right, padding.bottom, padding.left)
        ):
            return None
        if not self._canvas_sized(root_values) or not self._canvas_sized(stage_values):
            return None
        return stage

    def _canvas_sized(self, values: dict[str, Any]) -> bool:
        width = values.get("width")
        height = values.get("height")
        if width is None or height is None:
            return False
        return (
            abs(float(width) - self.canvas.width_px) <= _CANVAS_TOLERANCE_PX
            and abs(float(height) - self.canvas.height_px) <= _CANVAS_TOLERANCE_PX
        )

    def _result(self, background: ir.Fill | None) -> LoweringResult:
        document = ir.Document(
            canvas=ir.Canvas(width_px=self.canvas.width_px, height_px=self.canvas.height_px),
            slides=(ir.Slide(background=background, elements=tuple(self.elements)),),
        )
        self.report.stats["slides"] = 1
        self.report.stats["elements"] = len(self.elements)
        return LoweringResult(document=document, report=self.report)

    def _root_background(self, values: dict[str, Any]) -> ir.Fill | None:
        background = values.get("background")
        if isinstance(background, (ir.SolidFill, ir.LinearGradient)):
            return background
        color = values.get("background-color")
        if isinstance(color, ir.Color) and color.hex is not None:
            return ir.SolidFill(hex=color.hex, alpha=color.alpha)
        return None

    # ------------------------------------------------------------ 声明解析
    def _extract_class_rules(self, soup: Any) -> list[tuple[_Selector, int, str]]:
        """收集所有 <style> 块内的规则；@ 规则警告后跳过，不支持的选择器报错。

        规则按（特异性, 文档序）排序后合并：后代选择器让同一元素可能命中多条规则，
        没有顺序保证就会随源码书写顺序抖动。
        """
        rules: list[tuple[_Selector, int, str]] = []
        order = 0
        for style_tag in soup.find_all("style"):
            line = style_tag.sourceline or 0
            text = _CSS_COMMENT_RE.sub(" ", style_tag.get_text() or "")
            for selector, body in _split_css_rules(text):
                if not selector:
                    continue
                if selector.startswith("@"):
                    self.report.warning(
                        "STYLE_AT_RULE_IGNORED",
                        line=line,
                        message=f"样式表 @ 规则被忽略：{selector.split('(')[0][:40]}",
                        hint="响应式/打印样式不影响 1280×720 画布导出；如需条件样式请单独成页",
                        path="style",
                    )
                    continue
                for part in split_top(selector, ","):
                    parsed = _parse_selector(part)
                    if parsed is None:
                        self.report.error(
                            "STYLE_SELECTOR_UNSUPPORTED",
                            line=line,
                            message=f"不支持的选择器 “{part.strip()}”",
                            hint="仅支持标签、.class、标签+.class、.a.b 组合与后代选择器"
                            "（空格分隔）；子代 >、兄弟 +/~、伪类、id、属性选择器"
                            "请改为逐元素 class",
                            path="style",
                        )
                        continue
                    rules.append((parsed, order, body))
                    order += 1
        return rules

    def _element_classes(self, el: Tag) -> frozenset[str]:
        raw = el.get("class")
        if not raw:
            return frozenset()
        if isinstance(raw, str):
            return frozenset(raw.split())
        return frozenset(name for item in raw for name in str(item).split())

    def _style_text(self, el: Tag, inline: str, *, line: int, path: str) -> str:
        """命中规则的声明按特异性升序拼接，内联最后（等效内联优先）。"""
        if not self.class_rules:
            return inline
        matched = [
            (selector.specificity, order, declarations)
            for selector, order, declarations in self.class_rules
            if selector.matches(el, self._element_classes)
        ]
        if not matched:
            return inline
        matched.sort(key=lambda item: (item[0], item[1]))
        return "; ".join(declarations for _, _, declarations in matched) + "; " + inline

    def _declarations(
        self,
        raw: str,
        *,
        line: int,
        path: str,
        ref_w: float | None,
        ref_h: float | None,
        font_px: float,
        root: bool = False,
    ) -> dict[str, Any]:
        result: dict[str, Any] = {}
        if not raw.strip():
            return result
        context = ValueContext(
            contract=self.contract,
            font_strategy=self.font_strategy,
            ref_width_px=ref_w,
            ref_height_px=ref_h,
            font_px=font_px,
            viewport_width_px=self.canvas.width_px,
            viewport_height_px=self.canvas.height_px,
        )
        for declaration in split_top(raw, ";"):
            if ":" not in declaration:
                self.report.error(
                    "CONTRACT_UNKNOWN_PROPERTY",
                    line=line,
                    message=f"无法解析的声明 “{declaration}”",
                    hint="写法：属性: 值;（示例：left: 64px）",
                    path=path,
                )
                continue
            raw_prop, raw_value = declaration.split(":", 1)
            prop = raw_prop.strip().lower()
            value = raw_value.strip()
            if root and prop == "position":
                if value.lower() in ("absolute", "relative", "static"):
                    self.report.fix(
                        line=line,
                        rule="fix.structure.root-origin",
                        prop=prop,
                        value_from=value,
                        value_to="(忽略)",
                    )
                    continue
                self.report.error(
                    "CONTRACT_UNSUPPORTED_VALUE",
                    line=line,
                    message=f"根容器 position 仅支持 absolute/relative/static"
                    f"（视为噪声），得到 “{value}”",
                    hint="根 div 即画布，position 不影响布局",
                    prop=prop,
                    path=path,
                )
                continue
            if prop == "box-sizing":
                self.report.fix(
                    line=line,
                    rule="fix.compat.box-sizing-dropped",
                    prop=prop,
                    value_from=value,
                    value_to="(border-box 即契约盒语义)",
                )
                continue
            if prop == "overflow":
                if value.lower() in ("hidden", "clip"):
                    self.report.fix(
                        line=line,
                        rule="fix.compat.overflow-dropped",
                        prop=prop,
                        value_from=value,
                        value_to="(管线不裁剪，声明为无操作)",
                    )
                    continue
                self.report.error(
                    "CONTRACT_UNSUPPORTED_VALUE",
                    line=line,
                    message=f"overflow 仅支持 hidden/clip（按无操作处理），得到 “{value}”",
                    hint="契约不做滚动/可见性裁剪",
                    prop=prop,
                    path=path,
                )
                continue
            spec = self.contract.properties.get(prop)
            if spec is None:
                self.report.error(
                    "CONTRACT_UNKNOWN_PROPERTY",
                    line=line,
                    message=f"CSS 属性 “{prop}” 不在契约白名单",
                    hint=f"允许的属性：{'、'.join(self.contract.properties)}",
                    prop=prop,
                    path=path,
                )
                continue
            try:
                parsed = parse_value(spec, value, context)
            except ValueRejected as error:
                self.report.error(
                    error.code,
                    line=line,
                    message=error.message,
                    hint=error.hint,
                    prop=prop,
                    path=path,
                )
                continue
            for fix in parsed.fixes:
                self.report.fix(
                    line=line,
                    rule=fix.rule,
                    prop=prop,
                    value_from=fix.value_from,
                    value_to=fix.value_to,
                )
            if root and self._drop_root_noise(prop, parsed.value, value, line=line):
                continue
            result[prop] = parsed.value
        return result

    def _drop_root_noise(self, prop: str, value: Any, raw: str, *, line: int) -> bool:
        """根画布是固定原点：position 与 left/top=0 是模型常写的噪声，忽略并记录。"""
        if prop == "position":
            self.report.fix(
                line=line,
                rule="fix.structure.root-origin",
                prop=prop,
                value_from=raw,
                value_to="(忽略)",
            )
            return True
        if prop in ("left", "top") and float(value) == 0.0:
            self.report.fix(
                line=line,
                rule="fix.structure.root-origin",
                prop=prop,
                value_from=raw,
                value_to="(忽略)",
            )
            return True
        if prop == "display" and value != "none":
            self.report.fix(
                line=line,
                rule="fix.compat.display-dropped",
                prop=prop,
                value_from=raw,
                value_to="(丢弃)",
            )
            return True
        if prop == "flex":
            self.report.fix(
                line=line,
                rule="fix.structure.root-noise-dropped",
                prop=prop,
                value_from=raw,
                value_to="(忽略)",
            )
            return True
        if (
            prop == "padding"
            and isinstance(value, ir.Padding)
            and not any((value.top, value.right, value.bottom, value.left))
        ):
            self.report.fix(
                line=line,
                rule="fix.structure.root-noise-dropped",
                prop=prop,
                value_from=raw,
                value_to="(忽略)",
            )
            return True
        # 解析器已经记录了丢弃类修复，这里只避免根节点报错
        return prop in ("margin", "letter-spacing", "transform-origin")

    # ------------------------------------------------------------ 结构遍历
    def _walk(
        self,
        parent: Tag,
        parent_box: ir.Rect,
        path: str,
        inherited: _Inherited,
        *,
        depth: int,
    ) -> None:
        if depth > self.limits.nesting_depth:
            self.report.error(
                "STRUCTURE_DEPTH_EXCEEDED",
                line=parent.sourceline or 0,
                message=f"嵌套深度超过 {self.limits.nesting_depth} 层",
                hint="扁平化结构：把子元素提升到定位父级，减少容器层级",
                path=path,
            )
            return
        counts: dict[str, int] = {}
        for child in parent.children:
            if not isinstance(child, Tag):
                # 直接文本由 _extract_text 处理；只有根画布的直接文本算游离
                continue
            name = child.name.lower()
            if name in ("script", "style", "template"):
                continue  # 行为与样式来源，不产生画面元素
            counts[name] = counts.get(name, 0) + 1
            child_path = f"{path}/{name}[{counts[name]}]"
            try:
                self._visit(child, parent_box, child_path, inherited, depth=depth + 1)
            except _Abort:
                continue

    def _visit(
        self,
        el: Tag,
        parent_box: ir.Rect,
        path: str,
        inherited: _Inherited,
        *,
        depth: int,
    ) -> None:
        name = el.name.lower()
        if name == "div":
            self._visit_div(el, parent_box, path, inherited, depth=depth)
        elif name == "img":
            self._visit_image(el, parent_box, path)
        elif name == "table":
            self._visit_table(el, parent_box, path)
        elif name == "svg":
            self._visit_svg(el, parent_box, path, inherited)
        else:
            hint = "允许：div / img / table / svg(rect|circle|line)"
            if name in self.inline_tags:
                hint = "内联标签只能出现在文本 div 内部；块级内容请用定位的 <div>"
            self.report.error(
                "CONTRACT_UNKNOWN_TAG",
                line=el.sourceline or 0,
                message=f"标签 <{name}> 不在白名单或位置不合法",
                hint=hint,
                path=path,
            )
            raise _Abort

    # ------------------------------------------------------------ div
    def _visit_div(
        self,
        el: Tag,
        parent_box: ir.Rect,
        path: str,
        inherited: _Inherited,
        *,
        depth: int,
    ) -> None:
        line = el.sourceline or 0
        if id(el) in self.flex_skip:
            self.flex_skip.discard(id(el))
            return
        cached = self.flex_values.pop(id(el), None)
        if cached is not None:
            values = cached
        else:
            values = self._declarations(
                self._style_text(el, str(el.get("style") or ""), line=line, path=path),
                line=line,
                path=path,
                ref_w=parent_box.width,
                ref_h=parent_box.height,
                font_px=inherited.font_px,
            )

        if values.get("display") == "none":
            self.report.fix(
                line=line,
                rule="fix.structure.hidden-subtree",
                prop="display",
                value_from="none",
                value_to="(剪掉子树)",
            )
            return
        forced = self.flex_boxes.pop(id(el), None)
        display = values.get("display")
        if display in ("block", "inline", "inline-block"):
            self.report.fix(
                line=line,
                rule="fix.compat.display-dropped",
                prop="display",
                value_from=str(display),
                value_to="(丢弃)",
            )
        if values.get("position") == "relative":
            self.report.fix(
                line=line,
                rule="fix.structure.relative-normalized",
                prop="position",
                value_from="relative",
                value_to="absolute（按声明盒）",
            )
        if ("left" in values or "top" in values) and "position" not in values and forced is None:
            self.report.fix(
                line=line,
                rule="fix.structure.implicit-absolute",
                prop="position",
                value_from="(缺省)",
                value_to="absolute",
            )

        font_px = float(values.get("font-size", inherited.font_px))
        color_value = values.get("color")
        color = color_value.hex if isinstance(color_value, ir.Color) else inherited.color
        font = values.get("font-family", inherited.font)
        bold = values.get("font-weight", inherited.bold)
        align = values.get("text-align")
        line_height = values.get("line-height")
        padding = values.get("padding") or ir.Padding()
        word_wrap = values.get("white-space") != "nowrap"

        is_flex = values.get("display") == "flex"
        flex_text = is_flex and not self._has_structural_children(el)
        valign: str | None = None
        if flex_text:
            valign = {"center": "middle", "flex-end": "bottom"}.get(
                str(values.get("align-items", "stretch")), "top"
            )
            if align is None:
                align = {"center": "center", "flex-end": "right"}.get(
                    str(values.get("justify-content", "flex-start"))
                )
            self.report.fix(
                line=line,
                rule="fix.layout.flex-text-absorbed",
                prop="display",
                value_from="flex",
                value_to="(文本对齐)",
            )

        base_run = {
            "bold": bold,
            "italic": None,
            "color": color,
            "size_px": font_px,
            "font": font,
        }
        spacing_px = values.get("letter-spacing")
        indent_px = values.get("text-indent")
        vertical = values.get("writing-mode") is True
        paragraphs, structural = self._extract_text(
            el,
            base_run,
            line=line,
            path=path,
            align=align,
            line_height=line_height,
            spacing_px=spacing_px,
            indent_px=indent_px,
        )
        if paragraphs and structural:
            self.report.error(
                "STRUCTURE_MIXED_CONTENT",
                line=line,
                message="同一 div 内混排了文本与结构子元素",
                hint="文本单独放一个定位 div，与结构子元素平级。示例："
                '<div style="position:absolute;left:24px;top:24px">标题</div>',
                path=path,
            )
            paragraphs = []

        fill = self._fill_of(values)
        stroke = values.get("border") if isinstance(values.get("border"), ir.Stroke) else None
        shadow = values.get("box-shadow")
        has_visual = bool(fill or stroke or shadow)
        box = forced if forced is not None else self._box_of(values, parent_box)

        if is_flex and not flex_text and box is not None:
            self._plan_flex(el, box, padding, values, font_px=font_px, line=line, path=path)

        if not paragraphs and not has_visual:
            self._walk_div_children(
                el,
                values,
                box if box is not None else parent_box,
                path,
                inherited,
                font_px,
                color,
                font,
                bold,
                depth,
            )
            return

        if box is None:
            missing = [key for key in ("left", "top", "width", "height") if key not in values]
            self.report.error(
                "STRUCTURE_MISSING_BOX",
                line=line,
                message=f"元素缺少定位/尺寸：{'、'.join(missing)}",
                hint='绝对定位模型要求完整盒。示例：<div style="position:absolute;'
                'left:64px;top:120px;width:560px;height:420px">',
                path=path,
            )
            self._walk_div_children(
                el, values, parent_box, path, inherited, font_px, color, font, bold, depth
            )
            return

        self._check_text_limits(paragraphs, line=line, path=path)

        opacity = float(values.get("opacity", 1.0))
        if paragraphs and not has_visual and "opacity" in values:
            self.report.fix(
                line=line,
                rule="fix.paint.opacity-dropped",
                prop="opacity",
                value_from=str(values["opacity"]),
                value_to="(丢弃)",
            )
            opacity = 1.0

        rotation = float(values.get("transform", 0.0))
        source = ir.SourceRef(line=line, path=path)

        if "clip-path" in values:
            geometry: ir.Geometry = ir.PolygonGeometry(
                resolve_clip_path(values["clip-path"], box.width, box.height)
            )
        else:
            geometry = self._geometry(values.get("border-radius"), box)

        if has_visual:
            self._add(
                ir.Shape(
                    id="",
                    box=box,
                    source=source,
                    rotation_deg=rotation,
                    geometry=geometry,
                    fill=fill,
                    stroke=stroke,
                    shadow=shadow if isinstance(shadow, ir.Shadow) else None,
                    opacity=opacity,
                    paragraphs=tuple(paragraphs),
                    padding=padding,
                    word_wrap=word_wrap,
                    vertical=vertical,
                    valign=valign,
                ),
                line=line,
                path=path,
            )
        else:
            self._add(
                ir.TextBox(
                    id="",
                    box=box,
                    source=source,
                    rotation_deg=rotation,
                    paragraphs=tuple(paragraphs),
                    padding=padding,
                    word_wrap=word_wrap,
                    vertical=vertical,
                    valign=valign,
                ),
                line=line,
                path=path,
            )

        self._walk_div_children(
            el,
            values,
            self._child_origin(values, parent_box, box),
            path,
            inherited,
            font_px,
            color,
            font,
            bold,
            depth,
        )

    def _walk_div_children(
        self,
        el: Tag,
        values: dict[str, Any],
        origin: ir.Rect,
        path: str,
        inherited: _Inherited,
        font_px: float,
        color: str | None,
        font: str | None,
        bold: bool | None,
        depth: int,
    ) -> None:
        if not any(
            isinstance(child, Tag) and child.name.lower() not in self.inline_tags
            for child in el.children
        ):
            return
        child_inherited = _Inherited(font_px=font_px, color=color, font=font, bold=bold)
        self._walk(el, origin, path, child_inherited, depth=depth)

    def _child_origin(
        self, values: dict[str, Any], parent_box: ir.Rect, box: ir.Rect | None
    ) -> ir.Rect:
        """子元素绝对定位的参照原点：有盒用盒，只有 left/top 时按声明偏移，否则继承父盒。

        有边框时参照 padding box：绝对子元素相对边框内侧定位（CSS 语义），
        静态填充的图片也因此拿到正确的内容区尺寸。
        """
        border = values.get("border")
        inset = float(border.width_px) if isinstance(border, ir.Stroke) else 0.0
        if box is not None:
            return ir.Rect(
                box.x + inset,
                box.y + inset,
                max(0.0, box.width - inset * 2),
                max(0.0, box.height - inset * 2),
            )
        if "left" in values or "top" in values:
            return ir.Rect(
                parent_box.x + float(values.get("left", 0.0)) + inset,
                parent_box.y + float(values.get("top", 0.0)) + inset,
                max(0.0, float(values.get("width", 0.0)) - inset * 2),
                max(0.0, float(values.get("height", 0.0)) - inset * 2),
            )
        return parent_box

    # ------------------------------------------------------------ flex 布局
    def _has_structural_children(self, el: Tag) -> bool:
        return any(
            isinstance(child, Tag)
            and child.name.lower() not in self.inline_tags
            and child.name.lower() not in ("script", "style", "template")
            for child in el.children
        )

    @staticmethod
    def _text_lines(el: Tag) -> list[str]:
        """按 <br> 切行、忽略内联标签边界：用于 flex 自动尺寸的内容估算。"""
        lines: list[list[str]] = [[]]

        def walk(node: Tag) -> None:
            for child in node.children:
                if isinstance(child, Tag):
                    name = child.name.lower()
                    if name == "br":
                        lines.append([])
                    elif name in ("script", "style", "template"):
                        continue
                    else:
                        walk(child)
                elif _is_text(child):
                    lines[-1].append(_WS_RE.sub(" ", str(child)))

        walk(el)
        return ["".join(parts) for parts in lines]

    def _estimate_text_size(
        self, el: Tag, values: dict[str, Any], font_px: float
    ) -> tuple[float, float]:
        """自动尺寸的内容估算；度量与 textfit 同源（CJK≈1em、拉丁≈0.55em）。

        宽度额外留 2px：PPT 的 CJK 字宽取整略大于浏览器，正好卡边会提前换行。
        """
        padding = values.get("padding") or ir.Padding()
        size = float(values.get("font-size", font_px))
        spacing = float(values.get("letter-spacing") or 0.0)
        indent = float(values.get("text-indent") or 0.0)
        line_height = values.get("line-height")
        multiple = 1.2
        if isinstance(line_height, ir.LineHeight):
            if line_height.multiple is not None:
                multiple = line_height.multiple
            elif line_height.px is not None and size > 0:
                multiple = line_height.px / size
        metrics = textfit.HeuristicMetrics()
        lines = [line.strip() for line in self._text_lines(el) if line.strip()]
        if not lines:
            return padding.left + padding.right, padding.top + padding.bottom
        widths = [
            metrics.width_px(line, size)
            # PPT 不裁剪行尾字距：自动宽度把末尾 spacing 也算进去，否则紧盒会折行。
            + spacing * len(line)
            + (indent if index == 0 else 0.0)
            for index, line in enumerate(lines)
        ]
        return (
            max(widths) + padding.left + padding.right + 2.0,
            len(lines) * size * multiple + padding.top + padding.bottom,
        )

    def _plan_flex(
        self,
        el: Tag,
        box: ir.Rect,
        padding: ir.Padding,
        values: dict[str, Any],
        *,
        font_px: float,
        line: int,
        path: str,
    ) -> None:
        """单行 flex 求值：把子元素的声明尺寸与剩余空间解成绝对盒。

        只支持单行、无 wrap、无嵌套 margin/overflow 的子集；任何无法确定的尺寸都报
        STRUCTURE_FLEX_INVALID（一次定位，不产生连锁缺盒错误）。
        """
        row = str(values.get("flex-direction", "row")) == "row"
        justify = str(values.get("justify-content", "flex-start"))
        align = str(values.get("align-items", "stretch"))
        gap = float(values.get("gap") or 0.0)
        border = values.get("border")
        inset = float(border.width_px) if isinstance(border, ir.Stroke) else 0.0
        content = ir.Rect(
            box.x + inset + padding.left,
            box.y + inset + padding.top,
            box.width - inset * 2 - padding.left - padding.right,
            box.height - inset * 2 - padding.top - padding.bottom,
        )
        if content.width <= 0 or content.height <= 0:
            self.report.error(
                "STRUCTURE_FLEX_INVALID",
                line=line,
                message=f"flex 容器内容区 {content.width:g}×{content.height:g}px 无法布局",
                hint="增大容器 width/height 或减小 padding",
                path=path,
            )
            return
        main_len = content.width if row else content.height
        cross_len = content.height if row else content.width

        entries: list[dict[str, Any]] = []
        counts: dict[str, int] = {}
        for child in el.children:
            if not isinstance(child, Tag):
                continue
            name = child.name.lower()
            if name in ("script", "style", "template"):
                continue
            counts[name] = counts.get(name, 0) + 1
            cpath = f"{path}/{name}[{counts[name]}]"
            cline = child.sourceline or line
            child_values = self._declarations(
                self._style_text(child, str(child.get("style") or ""), line=cline, path=cpath),
                line=cline,
                path=cpath,
                ref_w=content.width,
                ref_h=content.height,
                font_px=font_px,
            )
            self.flex_values[id(child)] = child_values
            if child_values.get("display") == "none":
                continue
            if child_values.get("position") == "absolute":
                continue  # 离流元素按自身 left/top 走常规绝对定位
            if "left" in child_values or "top" in child_values:
                self.report.fix(
                    line=cline,
                    rule="fix.layout.flex-offset-dropped",
                    prop="left/top",
                    value_from="(声明)",
                    value_to="(忽略)",
                )
            spec = child_values.get("flex")
            flex_spec = spec if isinstance(spec, ir.FlexSpec) else ir.FlexSpec()

            main_key = "width" if row else "height"
            cross_key = "height" if row else "width"
            declared_main = child_values.get(main_key)
            declared_cross = child_values.get(cross_key)
            auto_w, auto_h = self._estimate_text_size(child, child_values, font_px)
            auto_main = auto_w if row else auto_h
            auto_cross = auto_h if row else auto_w

            # 图片与结构容器无法从内容反推尺寸：必须两轴显式声明。
            if (name != "div" or self._has_structural_children(child)) and (
                declared_main is None or declared_cross is None
            ):
                missing = [
                    key
                    for key, value in (
                        (main_key, declared_main),
                        (cross_key, declared_cross),
                    )
                    if value is None
                ]
                self.report.error(
                    "STRUCTURE_FLEX_INVALID",
                    line=cline,
                    message=f"flex 子元素 <{name}> 缺少尺寸：{'、'.join(missing)}",
                    hint="flex 子项需要显式宽高（图片尺寸决定裁剪），"
                    '示例：style="width:262px;height:216px"',
                    path=cpath,
                )
                self.flex_skip.add(id(child))
                continue

            entries.append(
                {
                    "child": child,
                    "grow": flex_spec.grow,
                    "shrink": flex_spec.shrink,
                    "main": float(declared_main) if declared_main is not None else auto_main,
                    "cross": float(declared_cross) if declared_cross is not None else auto_cross,
                    "cross_auto": declared_cross is None,
                    "text_auto": declared_main is None and name == "div",
                }
            )

        if not entries:
            return
        sizes = [entry["main"] for entry in entries]
        gaps = gap * max(0, len(entries) - 1)
        free = main_len - gaps - sum(sizes)
        if free < 0:
            shrinkable = [
                index
                for index, entry in enumerate(entries)
                if entry["shrink"] > 0 and sizes[index] > 0
            ]
            weight = sum(entries[index]["shrink"] * sizes[index] for index in shrinkable)
            if weight > 0:
                deficit = -free
                for index in shrinkable:
                    cut = deficit * (entries[index]["shrink"] * sizes[index]) / weight
                    sizes[index] = max(0.0, sizes[index] - cut)
                free = main_len - gaps - sum(sizes)
        elif free > 0:
            growers = [index for index, entry in enumerate(entries) if entry["grow"] > 0]
            weight = sum(entries[index]["grow"] for index in growers)
            if weight > 0:
                for index in growers:
                    sizes[index] += free * entries[index]["grow"] / weight
                free = main_len - gaps - sum(sizes)
            elif justify == "flex-start" and entries[-1]["text_auto"]:
                # 单行文本项的 max-content 只能估算；把剩余空间给它，换行位置更接近浏览器。
                sizes[-1] += free
                free = 0.0

        cursor = 0.0
        extra = 0.0
        if free > 0:
            if justify == "center":
                cursor = free / 2
            elif justify == "flex-end":
                cursor = free
            elif justify == "space-between" and len(entries) > 1:
                extra = free / (len(entries) - 1)
            elif justify == "space-around":
                extra = free / len(entries)
                cursor = extra / 2

        for index, entry in enumerate(entries):
            if index:
                cursor += gap + extra
            cross_size = float(entry["cross"])
            if entry["cross_auto"] and align == "stretch":
                cross_size = cross_len
            if align == "center":
                cross_offset = (cross_len - cross_size) / 2
            elif align == "flex-end":
                cross_offset = cross_len - cross_size
            else:
                cross_offset = 0.0
            if row:
                rect = ir.Rect(
                    content.x + cursor,
                    content.y + cross_offset,
                    sizes[index],
                    cross_size,
                )
            else:
                rect = ir.Rect(
                    content.x + cross_offset,
                    content.y + cursor,
                    cross_size,
                    sizes[index],
                )
            self.flex_boxes[id(entry["child"])] = rect
            cursor += sizes[index]

    def _fill_of(self, values: dict[str, Any]) -> ir.Fill | None:
        background = values.get("background")
        if background is None:
            color = values.get("background-color")
            if isinstance(color, ir.Color) and color.hex is not None:
                return ir.SolidFill(hex=color.hex, alpha=color.alpha)
            return None
        if isinstance(background, (ir.SolidFill, ir.LinearGradient)):
            return background
        return None

    def _box_of(self, values: dict[str, Any], parent_box: ir.Rect) -> ir.Rect | None:
        """声明坐标 + 父级原点 = 画布绝对坐标。

        绝对定位的参照是「最近的定位祖先」，因此这里必须叠加 parent_box——丢掉它，
        嵌套容器里的子元素会整体漂到页面左上角（真实试用 R5 暴露的缺陷）。
        """
        if not all(key in values for key in ("left", "top", "width", "height")):
            return None
        return ir.Rect(
            x=parent_box.x + float(values["left"]),
            y=parent_box.y + float(values["top"]),
            width=float(values["width"]),
            height=float(values["height"]),
        )

    def _geometry(self, radius: Any, box: ir.Rect) -> ir.Geometry:
        """CSS 圆角 → 形状语义：px 半径封顶短边一半；正方形满圆与 50% 走椭圆。"""
        if not isinstance(radius, ir.Radius):
            return ir.RectGeometry()
        if radius.percent is not None:
            rx, ry = box.width * radius.percent, box.height * radius.percent
            if rx >= box.width / 2 - 0.01 and ry >= box.height / 2 - 0.01:
                return ir.EllipseGeometry()
            return ir.RoundRectGeometry(radius_px=min(rx, ry))
        value = radius.px or 0.0
        if value <= 0:
            return ir.RectGeometry()
        if abs(box.width - box.height) < 0.01 and value >= box.width / 2 - 0.01:
            return ir.EllipseGeometry()
        return ir.RoundRectGeometry(radius_px=min(value, min(box.width, box.height) / 2))

    # ------------------------------------------------------------ 文本
    def _extract_text(
        self,
        el: Tag,
        base: dict[str, Any],
        *,
        line: int,
        path: str,
        align: str | None,
        line_height: ir.LineHeight | None,
        spacing_px: float | None,
        indent_px: float | None,
    ) -> tuple[list[ir.Paragraph], bool]:
        chunks: list[list[ir.TextRun]] = [[]]
        structural = False

        def walk(node: Tag, style: dict[str, Any]) -> None:
            nonlocal structural
            for child in node.children:
                if _is_text(child):
                    text = _WS_RE.sub(" ", str(child))
                    if not text:
                        continue
                    if not chunks[-1] and text == " ":
                        continue
                    chunks[-1].append(
                        ir.TextRun(
                            text=text,
                            bold=style.get("bold"),
                            italic=style.get("italic"),
                            color=style.get("color"),
                            size_px=style.get("size_px"),
                            font=style.get("font"),
                        )
                    )
                elif isinstance(child, Tag):
                    name = child.name.lower()
                    if name == "br":
                        if chunks[-1]:
                            chunks.append([])
                    elif name in ("b", "strong"):
                        walk(child, {**style, "bold": True})
                    elif name in ("em", "i"):
                        walk(child, {**style, "italic": True})
                    elif name == "span":
                        walk(child, self._span_style(child, style, line, path))
                    else:
                        structural = True

        walk(el, dict(base))

        paragraphs: list[ir.Paragraph] = []
        for runs in chunks:
            if not runs:
                continue
            runs[0] = replace(runs[0], text=runs[0].text.lstrip())
            runs[-1] = replace(runs[-1], text=runs[-1].text.rstrip())
            kept = [run for run in runs if run.text]
            if kept:
                paragraphs.append(
                    ir.Paragraph(
                        runs=tuple(kept),
                        align=align,
                        line_height=line_height,
                        spacing_px=spacing_px,
                        indent_px=indent_px if not paragraphs else None,
                    )
                )
        return paragraphs, structural

    def _span_style(
        self,
        el: Tag,
        style: dict[str, Any],
        line: int,
        path: str,
    ) -> dict[str, Any]:
        updated = dict(style)
        raw = str(el.get("style") or "")
        if not raw.strip():
            return updated
        context = ValueContext(
            contract=self.contract,
            font_strategy=self.font_strategy,
            font_px=float(style.get("size_px") or 16.0),
        )
        for declaration in split_top(raw, ";"):
            if ":" not in declaration:
                self.report.error(
                    "CONTRACT_UNKNOWN_PROPERTY",
                    line=line,
                    message=f"<span> 内无法解析的声明 “{declaration}”",
                    hint="写法：属性: 值;",
                    path=path,
                )
                continue
            prop, value = declaration.split(":", 1)
            prop, value = prop.strip().lower(), value.strip()
            if prop not in _INLINE_ALLOWED:
                self.report.error(
                    "CONTRACT_UNKNOWN_PROPERTY",
                    line=line,
                    message=f"<span> 内仅支持 {'、'.join(sorted(_INLINE_ALLOWED))}，得到 “{prop}”",
                    hint='示例：<span style="color:#4F46E5">',
                    prop=prop,
                    path=path,
                )
                continue
            spec = self.contract.properties[prop]
            try:
                parsed = parse_value(spec, value, context)
            except ValueRejected as error:
                self.report.error(
                    error.code,
                    line=line,
                    message=error.message,
                    hint=error.hint,
                    prop=prop,
                    path=path,
                )
                continue
            for fix in parsed.fixes:
                self.report.fix(
                    line=line,
                    rule=fix.rule,
                    prop=prop,
                    value_from=fix.value_from,
                    value_to=fix.value_to,
                )
            if prop == "color" and isinstance(parsed.value, ir.Color):
                updated["color"] = parsed.value.hex
            elif prop == "font-size":
                updated["size_px"] = float(parsed.value)
            elif prop == "font-weight":
                updated["bold"] = bool(parsed.value)
            elif prop == "font-family":
                updated["font"] = str(parsed.value)
        return updated

    def _check_text_limits(self, paragraphs: list[ir.Paragraph], *, line: int, path: str) -> None:
        chars = sum(len(run.text) for para in paragraphs for run in para.runs)
        if chars > self.limits.text_chars:
            self.report.error(
                "LIMIT_TEXT_EXCEEDED",
                line=line,
                message=f"单元素文本 {chars} 字符超过上限 {self.limits.text_chars}",
                hint="拆成多个文本元素，或精简文案",
                path=path,
            )
        if len(paragraphs) > self.limits.paragraphs:
            self.report.error(
                "LIMIT_TEXT_EXCEEDED",
                line=line,
                message=f"段落数 {len(paragraphs)} 超过上限 {self.limits.paragraphs}",
                hint="精简段落，或拆到多页",
                path=path,
            )

    # ------------------------------------------------------------ 图片
    def _visit_image(self, el: Tag, parent_box: ir.Rect, path: str) -> None:
        line = el.sourceline or 0
        if id(el) in self.flex_skip:
            self.flex_skip.discard(id(el))
            return
        cached = self.flex_values.pop(id(el), None)
        if cached is not None:
            values = cached
        else:
            values = self._declarations(
                self._style_text(el, str(el.get("style") or ""), line=line, path=path),
                line=line,
                path=path,
                ref_w=parent_box.width,
                ref_h=parent_box.height,
                font_px=16.0,
            )
        forced = self.flex_boxes.pop(id(el), None)
        box = forced if forced is not None else self._box_of(values, parent_box)
        if (
            box is None
            and "width" in values
            and "height" in values
            and "left" not in values
            and "top" not in values
            and "position" not in values
        ):
            # 静态块：无定位的图片按流式规则从父内容区左上角开始（.frame img 惯用法）。
            box = ir.Rect(
                parent_box.x,
                parent_box.y,
                float(values["width"]),
                float(values["height"]),
            )
            self.report.fix(
                line=line,
                rule="fix.layout.static-fill",
                prop="position",
                value_from="(缺省)",
                value_to="静态块左上角",
            )
        if box is None:
            self.report.error(
                "STRUCTURE_MISSING_BOX",
                line=line,
                message="<img> 缺少定位/尺寸",
                hint='示例：<img style="position:absolute;left:88px;top:144px;'
                'width:512px;height:300px" src="assets/photo.png">',
                path=path,
            )
            raise _Abort
        src = str(el.get("src") or "").strip()
        if not src:
            self.report.error(
                "ASSET_SRC_MISSING",
                line=line,
                message="<img> 缺少 src",
                hint="使用相对 base_dir 的路径（如 assets/photo.png）或 data:image/...;base64",
                path=path,
            )
            raise _Abort
        if "opacity" in values:
            self.report.fix(
                line=line,
                rule="fix.media.opacity-dropped",
                prop="opacity",
                value_from=str(values["opacity"]),
                value_to="(丢弃)",
            )
        if "display" in values:
            self.report.fix(
                line=line,
                rule="fix.compat.display-dropped",
                prop="display",
                value_from=str(values["display"]),
                value_to="(图片按绝对定位发射)",
            )
        if "border-radius" in values:
            self.report.fix(
                line=line,
                rule="fix.media.border-radius-dropped",
                prop="border-radius",
                value_from=str(values["border-radius"]),
                value_to="(直角)",
            )
        source = ir.SourceRef(line=line, path=path)
        alt = str(el.get("alt") or "")
        fit = str(values.get("object-fit", "contain"))
        self._add(
            ir.Picture(
                id="",
                box=box,
                source=source,
                src=src,
                fit=fit,
                alt=alt,
            ),
            line=line,
            path=path,
        )

    # ------------------------------------------------------------ 表格
    def _visit_table(self, el: Tag, parent_box: ir.Rect, path: str) -> None:
        line = el.sourceline or 0
        values = self._declarations(
            self._style_text(el, str(el.get("style") or ""), line=line, path=path),
            line=line,
            path=path,
            ref_w=parent_box.width,
            ref_h=parent_box.height,
            font_px=16.0,
        )
        box = self._box_of(values, parent_box)
        if box is None:
            self.report.error(
                "STRUCTURE_MISSING_BOX",
                line=line,
                message="<table> 缺少定位/尺寸",
                hint='示例：<table style="position:absolute;left:88px;top:144px;'
                'width:600px;height:240px">',
                path=path,
            )
            raise _Abort
        model = parse_table(
            el,
            box_width=box.width,
            report=self.report,
            contract=self.contract,
            font_strategy=self.font_strategy,
            path=path,
        )
        if model is None:
            raise _Abort
        color_value = values.get("color")
        style = ir.TableTextStyle(
            size_px=float(values.get("font-size", 16.0)),
            color=color_value.hex if isinstance(color_value, ir.Color) else None,
            font=values.get("font-family"),
        )
        rotation = float(values.get("transform", 0.0))
        source = ir.SourceRef(line=line, path=path)
        self._add(
            ir.Table(
                id="",
                box=box,
                source=source,
                rotation_deg=rotation,
                rows=model.rows,
                col_width_px=model.col_width_px,
                header=model.header,
                text_style=style,
                cells=model.cells,
            ),
            line=line,
            path=path,
        )

    # ------------------------------------------------------------ SVG
    def _visit_svg(
        self,
        el: Tag,
        parent_box: ir.Rect,
        path: str,
        inherited: _Inherited,
    ) -> None:
        line = el.sourceline or 0
        values = self._declarations(
            self._style_text(el, str(el.get("style") or ""), line=line, path=path),
            line=line,
            path=path,
            ref_w=parent_box.width,
            ref_h=parent_box.height,
            font_px=inherited.font_px,
        )
        for attribute in ("width", "height"):
            if attribute in values or not el.has_attr(attribute):
                continue
            raw = str(el.get(attribute))
            try:
                parsed = parse_value(
                    self.contract.properties[attribute],
                    raw,
                    ValueContext(
                        contract=self.contract,
                        font_strategy=self.font_strategy,
                        ref_width_px=parent_box.width,
                        ref_height_px=parent_box.height,
                        font_px=inherited.font_px,
                    ),
                )
            except ValueRejected as error:
                self.report.error(
                    error.code,
                    line=line,
                    message=f"<svg> {attribute} 非法：{error.message}",
                    hint=error.hint,
                    prop=attribute,
                    path=path,
                )
                raise _Abort from None
            for fix in parsed.fixes:
                self.report.fix(
                    line=line,
                    rule=fix.rule,
                    prop=attribute,
                    value_from=fix.value_from,
                    value_to=fix.value_to,
                )
            values[attribute] = parsed.value

        box = self._box_of(values, parent_box)
        if box is None:
            self.report.error(
                "STRUCTURE_MISSING_BOX",
                line=line,
                message="<svg> 缺少定位/尺寸",
                hint='示例：<svg style="position:absolute;left:0;top:0;width:1280px;height:720px">',
                path=path,
            )
            raise _Abort

        items = parse_svg_children(
            el,
            box=box,
            report=self.report,
            contract=self.contract,
            font_strategy=self.font_strategy,
            path=path,
        )
        for item in items:
            source = ir.SourceRef(line=item.line, path=path)
            if isinstance(item, SvgConnector):
                self._add(
                    ir.Connector(
                        id="",
                        box=item.box,
                        source=source,
                        start=item.start,
                        end=item.end,
                        stroke=item.stroke,
                    ),
                    line=item.line,
                    path=path,
                )
            else:
                self._add(
                    ir.Shape(
                        id="",
                        box=item.box,
                        source=source,
                        geometry=item.geometry,
                        fill=item.fill,
                        stroke=item.stroke,
                        opacity=item.opacity,
                    ),
                    line=item.line,
                    path=path,
                )

    # ------------------------------------------------------------ 追加与限额
    def _add(self, node: ir.Node, *, line: int, path: str) -> None:
        if len(self.elements) >= self.limits.elements_per_slide:
            if not self.element_limit_hit:
                self.report.error(
                    "LIMIT_ELEMENTS_EXCEEDED",
                    line=line,
                    message=f"单页元素数超过上限 {self.limits.elements_per_slide}",
                    hint="拆到多页，或合并重复的形状/文本",
                    path=path,
                )
                self.element_limit_hit = True
            return
        self.elements.append(replace(node, id=f"e{self.next_id:04d}"))
        self.next_id += 1
