"""IR → 原生 .pptx。

职责边界：只做机械翻译，不做业务判断——契约校验、fix 记录、限额都在遍历层完成。
两类东西必须手写 DrawingML：渐变填充、单层外阴影、填充/描边的 alpha。python-pptx
没有对应 API，且 ``spPr`` 的子元素顺序与唯一性会被 PowerPoint 严格检查（重复的
``effectLst`` 会让文件直接打不开），因此每次写入前先删同名旧节点。

坐标与字号换算统一走契约画布的 ``emu_per_px`` 与 1px=0.75pt。
"""

from __future__ import annotations

import io
import math
from dataclasses import dataclass
from pathlib import Path
from typing import Any

from lxml import etree
from pptx import Presentation
from pptx.dml.color import RGBColor
from pptx.enum.dml import MSO_LINE_DASH_STYLE
from pptx.enum.shapes import MSO_CONNECTOR, MSO_SHAPE
from pptx.enum.text import MSO_ANCHOR, MSO_AUTO_SIZE, PP_ALIGN
from pptx.oxml.ns import qn
from pptx.util import Emu, Pt

from . import ir, textfit
from .assets import resolve_image
from .contract import Contract, load_contract
from .diagnostics import Report

_A_NS = "http://schemas.openxmlformats.org/drawingml/2006/main"
_ALIGN = {
    "left": PP_ALIGN.LEFT,
    "center": PP_ALIGN.CENTER,
    "right": PP_ALIGN.RIGHT,
    "justify": PP_ALIGN.JUSTIFY,
}
# 浏览器默认表格无框线、无斑马纹；样式全部显式控制
_PLAIN_TABLE_STYLE = "{2D5ABB26-0587-4C30-8999-92F81FD0307C}"  # No Style, No Grid
_DEFAULT_FONT_PX = 16.0
_PX_TO_PT = 0.75


@dataclass(frozen=True, slots=True)
class EmitResult:
    presentation: Any
    report: Report


def emit(
    document: ir.Document,
    *,
    contract: Contract | None = None,
    base_dir: Path | None = None,
    font_strategy: str = "portable",
    metrics: textfit.Metrics | None = None,
    report: Report | None = None,
) -> EmitResult:
    """把 IR 渲染成 python-pptx 演示文稿。

    ``report`` 传入遍历层的报告即可把发射期诊断（图片、溢出）追加到同一份账本；
    图片解析失败只跳过该图片，其余元素照常生成。
    """
    ir.validate_document(document)
    active = contract or load_contract()
    active_report = report or Report(active)
    presentation = Presentation()
    presentation.slide_width = Emu(_emu(document.canvas.width_px, active))
    presentation.slide_height = Emu(_emu(document.canvas.height_px, active))

    emitter = _Emitter(active, base_dir or Path("."), font_strategy, active_report)
    for slide in document.slides:
        emitter.add_slide(presentation, slide)

    for overflow in textfit.estimate_overflow(document, metrics):
        active_report.warning(
            "TEXT_OVERFLOW_ESTIMATED",
            line=overflow.line,
            message=f"{overflow.element_id} 估算文本高度为盒高的 {overflow.ratio:.2f} 倍",
            hint="扩大盒高或减小 font-size；预览仅作排版检查",
            path=overflow.path,
        )
    return EmitResult(presentation=presentation, report=active_report)


def save(document: ir.Document, path: Path, **kwargs: Any) -> EmitResult:
    """便捷入口：生成并原子写入磁盘（先写临时文件再替换）。"""
    result = emit(document, **kwargs)
    target = Path(path)
    target.parent.mkdir(parents=True, exist_ok=True)
    buffer = io.BytesIO()
    result.presentation.save(buffer)
    temporary = target.with_name(f".{target.name}.tmp")
    temporary.write_bytes(buffer.getvalue())
    temporary.replace(target)
    return result


class _Emitter:
    def __init__(
        self,
        contract: Contract,
        base_dir: Path,
        font_strategy: str,
        report: Report,
    ) -> None:
        self.contract = contract
        self.base_dir = base_dir
        self.report = report
        self.default_font = contract.font_set(font_strategy).default

    def add_slide(self, presentation: Any, slide: ir.Slide) -> None:
        target = presentation.slides.add_slide(presentation.slide_layouts[6])
        if isinstance(slide.background, ir.SolidFill):
            fill = target.background.fill
            fill.solid()
            fill.fore_color.rgb = RGBColor.from_string(slide.background.hex)
        elif isinstance(slide.background, ir.LinearGradient):
            width = Emu(_emu(self.contract.canvas.width_px, self.contract))
            height = Emu(_emu(self.contract.canvas.height_px, self.contract))
            shape = target.shapes.add_shape(MSO_SHAPE.RECTANGLE, 0, 0, width, height)
            shape.shadow.inherit = False
            shape.line.fill.background()
            self._apply_gradient(shape, slide.background)

        for element in slide.elements:
            if isinstance(element, ir.TextBox):
                self._add_text_box(target, element)
            elif isinstance(element, ir.Shape):
                self._add_shape(target, element)
            elif isinstance(element, ir.Picture):
                self._add_picture(target, element)
            elif isinstance(element, ir.Table):
                self._add_table(target, element)
            elif isinstance(element, ir.Connector):
                self._add_connector(target, element)

    # ------------------------------------------------------------ 元素
    def _add_text_box(self, slide: Any, element: ir.TextBox) -> None:
        box = element.box
        shape = slide.shapes.add_textbox(
            Emu(_emu(box.x, self.contract)),
            Emu(_emu(box.y, self.contract)),
            Emu(_emu(box.width, self.contract)),
            Emu(_emu(box.height, self.contract)),
        )
        if element.rotation_deg:
            shape.rotation = element.rotation_deg
        frame = shape.text_frame
        frame.word_wrap = element.word_wrap
        frame.auto_size = MSO_AUTO_SIZE.NONE
        if element.vertical:
            self._set_vertical(frame)
        self._set_margins(frame, element.padding)
        self._fill_paragraphs(frame, element.paragraphs, box.height)

    def _add_shape(self, slide: Any, element: ir.Shape) -> None:
        box = element.box
        if isinstance(element.geometry, ir.PolygonGeometry):
            shape = self._add_freeform(slide, element)
        else:
            if isinstance(element.geometry, ir.EllipseGeometry):
                preset = MSO_SHAPE.OVAL
            elif isinstance(element.geometry, ir.RoundRectGeometry):
                preset = MSO_SHAPE.ROUNDED_RECTANGLE
            else:
                preset = MSO_SHAPE.RECTANGLE
            shape = slide.shapes.add_shape(
                preset,
                Emu(_emu(box.x, self.contract)),
                Emu(_emu(box.y, self.contract)),
                Emu(_emu(box.width, self.contract)),
                Emu(_emu(box.height, self.contract)),
            )
        # 先禁掉继承阴影：会生成一个空 effectLst，后面复用它而不是再追加一个
        shape.shadow.inherit = False
        self._apply_shape_style(shape, element, box)

        if element.rotation_deg:
            shape.rotation = element.rotation_deg
        if element.shadow is not None:
            self._apply_shadow(shape, element.shadow)

        if element.paragraphs:
            frame = shape.text_frame
            frame.word_wrap = element.word_wrap
            frame.auto_size = MSO_AUTO_SIZE.NONE
            self._set_margins(frame, element.padding)
            self._fill_paragraphs(frame, element.paragraphs, box.height)
            if element.vertical:
                self._set_vertical(frame)

    def _apply_shape_style(self, shape: Any, element: ir.Shape, box: ir.Rect) -> None:
        """填充 / 描边 / 圆角 / 阴影：预设形状与自由多边形共用。"""
        if isinstance(element.fill, ir.LinearGradient):
            self._apply_gradient(shape, element.fill)
        elif isinstance(element.fill, ir.SolidFill):
            shape.fill.solid()
            shape.fill.fore_color.rgb = RGBColor.from_string(element.fill.hex)
            alpha = element.fill.alpha * element.opacity
            if alpha < 1.0:
                self._set_fill_alpha(shape._element.spPr.find(qn("a:solidFill")), alpha)
        else:
            shape.fill.background()

        if element.stroke is not None:
            shape.line.color.rgb = RGBColor.from_string(element.stroke.color.hex or "000000")
            shape.line.width = Pt(element.stroke.width_px * _PX_TO_PT)
            if element.stroke.style != "solid":
                shape.line.dash_style = (
                    MSO_LINE_DASH_STYLE.DASH
                    if element.stroke.style == "dashed"
                    else MSO_LINE_DASH_STYLE.ROUND_DOT
                )
            alpha = element.stroke.color.alpha * element.opacity
            if alpha < 1.0:
                line = shape._element.spPr.find(qn("a:ln"))
                if line is not None:
                    self._set_fill_alpha(line.find(qn("a:solidFill")), alpha)
        else:
            shape.line.fill.background()

        if isinstance(element.geometry, ir.RoundRectGeometry):
            short_side = min(box.width, box.height) or 1.0
            shape.adjustments[0] = min(0.5, element.geometry.radius_px / short_side)

    def _add_picture(self, slide: Any, element: ir.Picture) -> None:
        asset = resolve_image(
            element.src,
            base_dir=self.base_dir,
            contract=self.contract,
            report=self.report,
            line=element.source.line,
            path=element.source.path,
        )
        if asset is None:
            return
        box = element.box
        if element.fit == "cover":
            scale = max(box.width / asset.width_px, box.height / asset.height_px)
            shape_w, shape_h = box.width, box.height
            left, top = box.x, box.y
        else:
            scale = min(box.width / asset.width_px, box.height / asset.height_px)
            shape_w, shape_h = asset.width_px * scale, asset.height_px * scale
            left = box.x + (box.width - shape_w) / 2
            top = box.y + (box.height - shape_h) / 2
        shape = slide.shapes.add_picture(
            io.BytesIO(asset.data),
            Emu(_emu(left, self.contract)),
            Emu(_emu(top, self.contract)),
            Emu(_emu(shape_w, self.contract)),
            Emu(_emu(shape_h, self.contract)),
        )
        if element.fit == "cover":
            # 源图超出声明盒的部分按 srcRect 裁掉，裁剪后的图恰好填满形状
            crop_x = (asset.width_px - box.width / scale) / (2 * asset.width_px)
            crop_y = (asset.height_px - box.height / scale) / (2 * asset.height_px)
            if crop_x > 0:
                shape.crop_left = crop_x
                shape.crop_right = crop_x
            if crop_y > 0:
                shape.crop_top = crop_y
                shape.crop_bottom = crop_y
        if element.rotation_deg:
            shape.rotation = element.rotation_deg

    def _add_table(self, slide: Any, element: ir.Table) -> None:
        rows = len(element.rows)
        columns = len(element.col_width_px) or (len(element.rows[0]) if element.rows else 0)
        if rows == 0 or columns == 0:
            return
        box = element.box
        graphic = slide.shapes.add_table(
            rows,
            columns,
            Emu(_emu(box.x, self.contract)),
            Emu(_emu(box.y, self.contract)),
            Emu(_emu(box.width, self.contract)),
            Emu(_emu(box.height, self.contract)),
        )
        table = graphic.table
        table.first_row = element.header
        table.horz_banding = False
        self._set_plain_table_style(table)
        for index, width in enumerate(element.col_width_px):
            table.columns[index].width = Emu(_emu(width, self.contract))

        style = element.text_style
        font_name = style.font or self.default_font
        font_size = Pt((style.size_px or _DEFAULT_FONT_PX) * _PX_TO_PT)
        for row_index, row in enumerate(element.rows):
            for column_index, text in enumerate(row):
                if column_index >= columns:
                    continue
                cell = table.cell(row_index, column_index)
                cell_style = self._cell_style_of(element, row_index, column_index)
                header = element.header and row_index == 0

                padding = (cell_style.padding if cell_style else None) or ir.Padding(
                    top=4.0, right=8.0, bottom=4.0, left=8.0
                )
                cell.margin_left = Emu(_emu(padding.left, self.contract))
                cell.margin_right = Emu(_emu(padding.right, self.contract))
                cell.margin_top = Emu(_emu(padding.top, self.contract))
                cell.margin_bottom = Emu(_emu(padding.bottom, self.contract))
                cell.vertical_anchor = MSO_ANCHOR.MIDDLE

                paragraph = cell.text_frame.paragraphs[0]
                align = (cell_style.align if cell_style else None) or (
                    "center" if header else "left"
                )
                paragraph.alignment = _ALIGN.get(align, PP_ALIGN.LEFT)

                run = paragraph.add_run()
                run.text = text
                font = run.font
                font.name = font_name
                self._set_ea_font(run, font_name)
                font.size = font_size
                color = (cell_style.color if cell_style else None) or style.color
                if color:
                    font.color.rgb = RGBColor.from_string(color)
                bold = (
                    cell_style.bold
                    if cell_style and cell_style.bold is not None
                    else (True if header else None)
                )
                if bold is not None:
                    font.bold = bold
                if cell_style and cell_style.italic:
                    font.italic = True

                fill = cell_style.fill if cell_style else None
                if fill is not None:
                    cell.fill.solid()
                    cell.fill.fore_color.rgb = RGBColor.from_string(fill.hex)
                    if fill.alpha < 1.0:
                        properties = cell._tc.get_or_add_tcPr()
                        solid = properties.find(qn("a:solidFill"))
                        if solid is not None:
                            self._set_fill_alpha(solid, fill.alpha)
                if cell_style and cell_style.border is not None:
                    self._set_cell_borders(cell, cell_style.border)

        if element.rotation_deg:
            graphic.rotation = element.rotation_deg

    @staticmethod
    def _cell_style_of(element: ir.Table, row_index: int, column_index: int) -> ir.CellStyle | None:
        if row_index >= len(element.cells):
            return None
        row = element.cells[row_index]
        if column_index >= len(row):
            return None
        return row[column_index]

    def _set_cell_borders(self, cell: Any, stroke: ir.Stroke) -> None:
        """单元格四边框：lnL/lnR/lnT/lnB 必须是 tcPr 的前四个子元素（schema 顺序）。"""
        properties = cell._tc.get_or_add_tcPr()
        for tag in ("a:lnL", "a:lnR", "a:lnT", "a:lnB"):
            for existing in properties.findall(qn(tag)):
                properties.remove(existing)
        for index, tag in enumerate(("a:lnL", "a:lnR", "a:lnT", "a:lnB")):
            width = _emu(stroke.width_px, self.contract)
            color = stroke.color.hex or "000000"
            xml = (
                f'<{tag} xmlns:a="{_A_NS}" w="{width}" cap="flat" cmpd="sng" algn="ctr">'
                f'<a:solidFill><a:srgbClr val="{color}">'
                f'<a:alpha val="{round(stroke.color.alpha * 100000)}"/>'
                f"</a:srgbClr></a:solidFill></{tag}>"
            )
            properties.insert(index, etree.fromstring(xml))

    def _add_connector(self, slide: Any, element: ir.Connector) -> None:
        start, end = element.start, element.end
        connector = slide.shapes.add_connector(
            MSO_CONNECTOR.STRAIGHT,
            Emu(_emu(start[0], self.contract)),
            Emu(_emu(start[1], self.contract)),
            Emu(_emu(end[0], self.contract)),
            Emu(_emu(end[1], self.contract)),
        )
        connector.line.color.rgb = RGBColor.from_string(element.stroke.color.hex or "000000")
        connector.line.width = Pt(element.stroke.width_px * _PX_TO_PT)
        if element.stroke.color.alpha < 1.0:
            line = connector._element.spPr.find(qn("a:ln"))
            if line is not None:
                self._set_fill_alpha(line.find(qn("a:solidFill")), element.stroke.color.alpha)

    # ------------------------------------------------------------ 文本
    def _fill_paragraphs(
        self,
        frame: Any,
        paragraphs: tuple[ir.Paragraph, ...],
        box_height: float,
    ) -> None:
        center_idiom = False
        if len(paragraphs) == 1:
            line_height = self._line_height_px(paragraphs[0])
            if (
                line_height
                and box_height
                and abs(line_height - box_height) <= max(2.0, box_height * 0.02)
            ):
                center_idiom = True
        frame.vertical_anchor = MSO_ANCHOR.MIDDLE if center_idiom else MSO_ANCHOR.TOP

        for index, paragraph in enumerate(paragraphs):
            target = frame.paragraphs[0] if index == 0 else frame.add_paragraph()
            target.alignment = _ALIGN.get(paragraph.align or "left", PP_ALIGN.LEFT)
            if not center_idiom:
                line_height = self._line_height_px(paragraph)
                if line_height:
                    target.line_spacing = Emu(_emu(line_height, self.contract))
            # NOTE: 行尾字距裁不裁是两个口径的共同前提，先按段判定一次：
            # 末 run 去不留 spc、居中段 indent 的抵消量都取决于它；None = 行尾保留。
            tail_spacing_px = paragraph.spacing_px if len(paragraph.runs[-1].text) > 1 else None
            if paragraph.indent_px:
                effective_indent = paragraph.indent_px
                if target.alignment == PP_ALIGN.CENTER and tail_spacing_px is not None:
                    # 浏览器把行尾字距计入居中行盒（字面左移 字距/2），源页惯用等值
                    # text-indent 补回；PPT 侧行尾字距已裁掉，indent 原样保留会把居中
                    # 行右推 字距/2——把已被抵消的那份字距同步减掉才是浏览器口径。
                    effective_indent = max(0.0, effective_indent - tail_spacing_px)
                target._p.get_or_add_pPr().set("indent", str(_emu(effective_indent, self.contract)))
            for run_index, text_run in enumerate(paragraph.runs):
                # 末字符不给 spc：CSS 行尾字距不可见，PPT 却按“字符后留白”判断换行，
                # 紧盒里会把最后一字挤到下一行。
                trim_tail = tail_spacing_px is not None and run_index == len(paragraph.runs) - 1
                segments = (
                    [(text_run.text[:-1], True), (text_run.text[-1:], False)]
                    if trim_tail
                    else [(text_run.text, True)]
                )
                for text, spaced in segments:
                    run = target.add_run()
                    run.text = text
                    font = run.font
                    font.name = text_run.font or self.default_font
                    self._set_ea_font(run, font.name)
                    if text_run.bold is not None:
                        font.bold = text_run.bold
                    if text_run.italic is not None:
                        font.italic = text_run.italic
                    if text_run.color:
                        font.color.rgb = RGBColor.from_string(text_run.color)
                    font.size = Pt((text_run.size_px or _DEFAULT_FONT_PX) * _PX_TO_PT)
                    if paragraph.spacing_px and spaced:
                        run._r.get_or_add_rPr().set(
                            "spc", str(round(paragraph.spacing_px * _PX_TO_PT * 100))
                        )

    def _line_height_px(self, paragraph: ir.Paragraph) -> float | None:
        line_height = paragraph.line_height
        if line_height is None:
            return None
        if line_height.px is not None:
            return line_height.px
        if line_height.multiple is not None:
            size = max((run.size_px or _DEFAULT_FONT_PX) for run in paragraph.runs)
            return line_height.multiple * size
        return None

    def _set_margins(self, frame: Any, padding: ir.Padding) -> None:
        frame.margin_left = Emu(_emu(padding.left, self.contract))
        frame.margin_right = Emu(_emu(padding.right, self.contract))
        frame.margin_top = Emu(_emu(padding.top, self.contract))
        frame.margin_bottom = Emu(_emu(padding.bottom, self.contract))

    def _add_freeform(self, slide: Any, element: ir.Shape) -> Any:
        """clip-path polygon → 自由多边形（顶点为声明盒内相对分数）。"""
        assert isinstance(element.geometry, ir.PolygonGeometry)
        box = element.box
        vertices = [
            (
                Emu(_emu(box.x + fx * box.width, self.contract)),
                Emu(_emu(box.y + fy * box.height, self.contract)),
            )
            for fx, fy in element.geometry.vertices
        ]
        builder = slide.shapes.build_freeform(vertices[0][0], vertices[0][1], 1.0)
        builder.add_line_segments(vertices[1:], close=True)
        return builder.convert_to_shape()

    @staticmethod
    def _set_vertical(frame: Any) -> None:
        """中文竖排：字符直立、列从右到左。

        NOTE: eaVert 下 anchor 的作用轴随文流旋转，管的不再是上下而是列的横向
        落位。竖排单列的横向位置在两边都是布局伪影：CSS 贴块起点（右侧留槽位）、
        PowerPoint 对 TOP 锚点同样贴块起点——都不等于居中徽章的设计意图。锚点
        统一取中心：单列即形状横向中央，多列整组居中。
        """
        frame._txBody.bodyPr.set("vert", "eaVert")
        frame.vertical_anchor = MSO_ANCHOR.MIDDLE

    @staticmethod
    def _set_ea_font(run: Any, name: str) -> None:
        """python-pptx 的 font.name 只写 <a:latin>；中文需要 <a:ea>。"""
        properties = run._r.get_or_add_rPr()
        for existing in properties.findall(qn("a:ea")):
            properties.remove(existing)
        element = properties.makeelement(qn("a:ea"), {"typeface": name})
        properties.append(element)

    # ------------------------------------------------------------ DrawingML
    def _apply_gradient(self, shape: Any, gradient: ir.LinearGradient) -> None:
        sp_pr = shape._element.spPr
        for tag in (
            "a:solidFill",
            "a:noFill",
            "a:gradFill",
            "a:blipFill",
            "a:pattFill",
            "a:grpFill",
        ):
            for existing in sp_pr.findall(qn(tag)):
                sp_pr.remove(existing)
        stops = "".join(
            f'<a:gs pos="{int(stop.position * 100000)}">'
            f'<a:srgbClr val="{stop.hex}">'
            f'<a:alpha val="{round(stop.alpha * 100000)}"/></a:srgbClr></a:gs>'
            for stop in gradient.stops
        )
        # CSS 角度 0deg=向上、顺时针；OOXML ang 0=向右、顺时针 → ang = css - 90
        angle = round(((gradient.angle_css_deg - 90.0) % 360.0) * 60000)
        xml = (
            f'<a:gradFill xmlns:a="{_A_NS}" rotWithShape="1">'
            f"<a:gsLst>{stops}</a:gsLst>"
            f'<a:lin ang="{angle}" scaled="0"/></a:gradFill>'
        )
        geometry = sp_pr.find(qn("a:prstGeom"))
        if geometry is None:
            sp_pr.append(etree.fromstring(xml))
        else:
            geometry.addnext(etree.fromstring(xml))

    def _apply_shadow(self, shape: Any, shadow: ir.Shadow) -> None:
        sp_pr = shape._element.spPr
        # shadow.inherit=False 已建空 effectLst；必须复用，重复元素会被 PowerPoint 拒绝打开
        for existing in sp_pr.findall(qn("a:effectLst")):
            sp_pr.remove(existing)
        distance = math.hypot(shadow.dx_px, shadow.dy_px)
        direction = math.degrees(math.atan2(shadow.dy_px, shadow.dx_px)) % 360.0
        xml = (
            f'<a:effectLst xmlns:a="{_A_NS}">'
            f'<a:outerShdw blurRad="{_emu(shadow.blur_px, self.contract)}" '
            f'dist="{_emu(distance, self.contract)}" '
            f'dir="{round(direction * 60000)}" rotWithShape="0">'
            f'<a:srgbClr val="{shadow.color.hex or "000000"}">'
            f'<a:alpha val="{round(shadow.color.alpha * 100000)}"/></a:srgbClr>'
            f"</a:outerShdw></a:effectLst>"
        )
        sp_pr.append(etree.fromstring(xml))

    @staticmethod
    def _set_fill_alpha(fill: Any, alpha: float) -> None:
        if fill is None:
            return
        color = fill.find(qn("a:srgbClr"))
        if color is None:
            return
        existing = color.find(qn("a:alpha"))
        if existing is None:
            existing = etree.SubElement(color, qn("a:alpha"))
        existing.set("val", str(round(alpha * 100000)))

    @staticmethod
    def _set_plain_table_style(table: Any) -> None:
        table_properties = table._tbl.tblPr
        for existing in table_properties.findall(qn("a:tableStyleId")):
            table_properties.remove(existing)
        style = table_properties.makeelement(qn("a:tableStyleId"), {})
        style.text = _PLAIN_TABLE_STYLE
        table_properties.append(style)


def _emu(px: float, contract: Contract) -> int:
    return round(px * contract.canvas.emu_per_px)
