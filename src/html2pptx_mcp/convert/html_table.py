"""简单表格解析：<table> → TableModel。

只处理契约允许的子集：无合并单元格、无嵌套、单元格纯文本。列宽取首行单元格的显式
width（相对表格宽度），任何一格缺失就整表均分——混合部分宽度会让版式不可预期。

本模块不分配元素 id、不判限额，也不碰 IR 节点：输入 DOM 与盒子宽度，输出纯数据，
所有诊断写进传入的 Report。
"""

from __future__ import annotations

import re
from dataclasses import dataclass

from bs4 import NavigableString, Tag

from . import ir
from .contract import Contract
from .diagnostics import Report
from .values import ValueContext, ValueRejected, parse_value, split_top

_WS_RE = re.compile(r"\s+")


@dataclass(frozen=True, slots=True)
class TableModel:
    rows: tuple[tuple[str, ...], ...]
    col_width_px: tuple[float, ...]
    header: bool
    cells: tuple[tuple[ir.CellStyle | None, ...], ...]


def parse_table(
    el: Tag,
    *,
    box_width: float,
    report: Report,
    contract: Contract,
    font_strategy: str,
    path: str,
) -> TableModel | None:
    """解析表格；遇到契约禁止的形态时记录诊断并返回 None。"""
    line = el.sourceline or 0

    if el.find("table") is not None:
        report.error(
            "TABLE_NESTED",
            line=line,
            message="不支持嵌套表格",
            hint="拆分为多个平级表格",
            path=path,
        )
        return None

    for cell in el.find_all(["td", "th"]):
        if cell.has_attr("rowspan") or cell.has_attr("colspan"):
            report.error(
                "TABLE_MERGED_CELLS",
                line=cell.sourceline or line,
                message="表格包含 rowspan/colspan（合并单元格）",
                hint="仅支持简单表格。示例：<tr><td>A</td><td>B</td></tr>",
                path=path,
            )
            return None

    rows_tags = el.find_all("tr")
    if not rows_tags:
        report.error(
            "TABLE_EMPTY",
            line=line,
            message="表格没有数据行",
            hint="至少写一行：<tr><td>内容</td></tr>",
            path=path,
        )
        return None

    rows: list[list[str]] = []
    row_tags: list[list[Tag]] = []
    style_rows: list[list[ir.CellStyle | None]] = []
    for row_tag in rows_tags:
        cells = row_tag.find_all(["td", "th"], recursive=False)
        texts: list[str] = []
        row_tags.append(cells)
        row_styles: list[ir.CellStyle | None] = []
        for column_index, cell in enumerate(cells):
            text, style = _cell_content(
                cell,
                row_index=column_index,
                report=report,
                contract=contract,
                font_strategy=font_strategy,
                path=path,
                line=cell.sourceline or line,
            )
            texts.append(text)
            row_styles.append(style)
        rows.append(texts)
        style_rows.append(row_styles)

    columns = max(len(row) for row in rows)
    if any(len(row) < columns for row in rows):
        for row in rows:
            row.extend([""] * (columns - len(row)))
        report.fix(
            line=line,
            rule="fix.table.pad-columns",
            prop="table",
            value_from="列数不齐",
            value_to=f"补齐到 {columns} 列",
        )
    if len(rows) > contract.limits.table_rows or columns > contract.limits.table_cols:
        report.error(
            "LIMIT_TABLE_EXCEEDED",
            line=line,
            message=f"表格 {len(rows)}×{columns} 超过上限 "
            f"{contract.limits.table_rows}×{contract.limits.table_cols}",
            hint="拆分表格，或把明细放到附录页",
            path=path,
        )
        return None
    header = el.find("thead") is not None
    first_cells = rows_tags[0].find_all(["td", "th"], recursive=False)
    if not header and first_cells and all(cell.name.lower() == "th" for cell in first_cells):
        header = True

    widths = _column_widths(
        first_cells,
        table_width=box_width,
        columns=columns,
        report=report,
        contract=contract,
        font_strategy=font_strategy,
        line=line,
        path=path,
    )
    cell_styles: list[tuple[ir.CellStyle | None, ...]] = []
    for row_styles in style_rows:
        padded = list(row_styles)
        while len(padded) < columns:
            padded.append(None)
        cell_styles.append(tuple(padded))
    return TableModel(
        rows=tuple(tuple(row) for row in rows),
        col_width_px=tuple(widths),
        header=header,
        cells=tuple(cell_styles),
    )


def _column_widths(
    cells: list[Tag],
    *,
    table_width: float,
    columns: int,
    report: Report,
    contract: Contract,
    font_strategy: str,
    line: int,
    path: str,
) -> list[float]:
    widths: list[float] = []
    context = ValueContext(
        contract=contract,
        font_strategy=font_strategy,
        ref_width_px=table_width,
        ref_height_px=0.0,
        font_px=16.0,
    )
    spec = contract.properties["width"]
    for cell in cells:
        raw = ""
        for declaration in split_top(str(cell.get("style") or ""), ";"):
            if ":" not in declaration:
                continue
            prop, value = declaration.split(":", 1)
            if prop.strip().lower() == "width":
                raw = value.strip()
                break
        if not raw:
            return [table_width / columns] * columns
        try:
            parsed = parse_value(spec, raw, context)
        except ValueRejected as error:
            report.error(
                error.code,
                line=line,
                message=f"单元格列宽非法：{error.message}",
                hint=error.hint,
                prop="width",
                path=path,
            )
            return [table_width / columns] * columns
        for fix in parsed.fixes:
            report.fix(
                line=line,
                rule=fix.rule,
                prop="width",
                value_from=fix.value_from,
                value_to=fix.value_to,
            )
        widths.append(float(parsed.value))
    if len(widths) != columns:
        return [table_width / columns] * columns
    return widths


_CELL_ALLOWED = {
    "background-color",
    "color",
    "font-weight",
    "text-align",
    "border",
    "padding",
}


def _cell_content(
    cell: Tag,
    *,
    row_index: int,
    report: Report,
    contract: Contract,
    font_strategy: str,
    path: str,
    line: int,
) -> tuple[str, ir.CellStyle | None]:
    """取单元格文本并收集 b/i 与内联样式；白名单外的标签丢弃并记 fix。"""
    parts: list[str] = []
    marks: dict[str, bool | None] = {"bold": None, "italic": None}

    def walk(node: Tag, context: dict[str, bool | None]) -> None:
        for child in node.children:
            if isinstance(child, NavigableString):
                parts.append(str(child))
                continue
            if not isinstance(child, Tag):
                continue
            name = child.name.lower()
            if name in ("b", "strong"):
                context["bold"] = True
                walk(child, context)
            elif name in ("em", "i"):
                context["italic"] = True
                walk(child, context)
            else:
                report.fix(
                    line=line,
                    rule="fix.table.cell-formatting-dropped",
                    prop=name,
                    value_from=f"<{name}>",
                    value_to="(丢弃)",
                )
                walk(child, context)

    walk(cell, marks)
    text = _WS_RE.sub(" ", "".join(parts)).strip()
    style = _cell_style(
        cell,
        row_index=row_index,
        marks=marks,
        report=report,
        contract=contract,
        font_strategy=font_strategy,
        path=path,
        line=line,
    )
    return text, style


def _cell_style(
    cell: Tag,
    *,
    row_index: int,
    marks: dict[str, bool | None],
    report: Report,
    contract: Contract,
    font_strategy: str,
    path: str,
    line: int,
) -> ir.CellStyle | None:
    fill: ir.SolidFill | None = None
    color: str | None = None
    bold: bool | None = marks["bold"]
    italic: bool | None = marks["italic"]
    align: str | None = None
    border: ir.Stroke | None = None
    padding: ir.Padding | None = None

    raw = str(cell.get("style") or "")
    if not raw.strip():
        if not any((bold, italic)):
            return None
        return ir.CellStyle(bold=bold, italic=italic)

    context = ValueContext(contract=contract, font_strategy=font_strategy, font_px=16.0)
    for declaration in split_top(raw, ";"):
        if ":" not in declaration:
            continue
        prop, value = declaration.split(":", 1)
        prop, value = prop.strip().lower(), value.strip()
        if prop == "width":
            if row_index == 0:
                continue  # ?? width ?? _column_widths ??
            report.fix(
                line=line,
                rule="fix.table.cell-width-ignored",
                prop="width",
                value_from=value,
                value_to="(??)",
            )
            continue
        if prop not in _CELL_ALLOWED:
            report.error(
                "CONTRACT_UNKNOWN_PROPERTY",
                line=line,
                message=f"单元格样式 “{prop}” 不在白名单",
                hint=f"允许：{'、'.join(sorted(_CELL_ALLOWED))}",
                prop=prop,
                path=path,
            )
            continue
        spec = contract.properties[prop]
        try:
            parsed = parse_value(spec, value, context)
        except ValueRejected as error:
            report.error(
                error.code,
                line=line,
                message=f"单元格样式 {prop} 非法：{error.message}",
                hint=error.hint,
                prop=prop,
                path=path,
            )
            continue
        for fix in parsed.fixes:
            report.fix(
                line=line,
                rule=fix.rule,
                prop=prop,
                value_from=fix.value_from,
                value_to=fix.value_to,
            )
        if prop == "background-color":
            cell_color = parsed.value
            if cell_color.hex is not None:
                fill = ir.SolidFill(hex=cell_color.hex, alpha=cell_color.alpha)
        elif prop == "color":
            color = parsed.value.hex
        elif prop == "font-weight":
            bold = bool(parsed.value)
        elif prop == "text-align":
            align = str(parsed.value)
        elif prop == "border":
            border = parsed.value
        elif prop == "padding":
            padding = parsed.value

    if not any((fill, color, bold, italic, align, border, padding)):
        return None
    return ir.CellStyle(
        fill=fill,
        color=color,
        bold=bold,
        italic=italic,
        align=align,
        border=border,
        padding=padding,
    )
