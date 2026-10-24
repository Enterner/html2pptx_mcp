"""字体子集内嵌：把 deck 实际用到的字形以子集 TTF 写进 pptx，机器无字体也能正确渲染。

口径：

- 用量按 (家族, 粗体) 收集——PowerPoint 的内嵌单元就是这两轴；
- 字体文件经 ``fonts.FontResolver`` 解析，缺文件记 warning 跳过：逆转场景下
  重建产物与原版走同一回退链，行为天然一致；
- fsType 位 0-1 = 01（受限许可）跳过并记 warning，避免嵌入违规；
- 子集只保留出现过的码点（fontTools.subset），等线/雅黑这类大字体通常仅数十 KB；
- 部件写入走 python-pptx OPC：fntdata 部件 + p:embeddedFontLst（notesSz 之后，
  defaultTextStyle 之前——CT_Presentation 的 schema 顺序），并置
  @embedTrueTypeFonts/@saveSubsetFonts。
"""

from __future__ import annotations

from collections import defaultdict
from typing import Any

from lxml import etree

from . import ir
from .diagnostics import Report
from .fonts import FontResolver

_R_NS = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
_P_NS = "http://schemas.openxmlformats.org/presentationml/2006/main"
_FONT_REL = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/font"
_FONT_CONTENT_TYPE = "application/x-fontdata"
_EMBEDDED_FONTS_LIMIT = 12


def collect_usage(document: ir.Document, default_font: str) -> dict[tuple[str, bool], set[str]]:
    """按 (家族, 粗体) 收集全册用到的字符集；无显式字体的 run 落策略默认字体。"""
    usage: dict[tuple[str, bool], set[str]] = defaultdict(set)
    for slide in document.slides:
        for element in slide.elements:
            paragraphs: tuple[ir.Paragraph, ...] | None = getattr(element, "paragraphs", None)
            if not paragraphs:
                continue
            for paragraph in paragraphs:
                for run in paragraph.runs:
                    if not run.text:
                        continue
                    family = run.font or default_font
                    usage[(family, bool(run.bold))].update(run.text)
    return usage


def embed_used_fonts(
    document: ir.Document,
    presentation: Any,
    report: Report,
    default_font: str,
    resolver: FontResolver | None = None,
) -> None:
    """把用量字体内嵌进演示文稿；失败/跳过一律记 warning，不阻断出片。"""
    usage = collect_usage(document, default_font)
    if not usage:
        return
    resolver = resolver or FontResolver()

    part = presentation.part
    pr_element = part._element
    notes_sz = pr_element.find(f"{{{_P_NS}}}notesSz")
    if notes_sz is None:
        report.warning(
            "FONT_EMBED_SKIPPED",
            line=0,
            message="presentation.xml 结构异常（无 notesSz），跳过字体内嵌",
            hint="无需处理；如需内嵌请用 PowerPoint 另存一次",
        )
        return

    font_list = etree.SubElement(pr_element, f"{{{_P_NS}}}embeddedFontLst")
    part_seq = 0
    embedded = 0
    # 超量保护：内嵌字体过多会撑爆文件体积，命中即停并提示
    for (family, bold), chars in sorted(usage.items()):
        if embedded >= _EMBEDDED_FONTS_LIMIT:
            report.warning(
                "FONT_EMBED_LIMIT",
                line=0,
                message=f"内嵌字体超过 {_EMBEDDED_FONTS_LIMIT} 个，剩余家族跳过",
                hint="减少正文使用的字体家族数",
            )
            break
        text = "".join(sorted(chars))
        face = resolver.resolve(family, bold)
        if face is None:
            # 缺字体不内嵌：重建产物与原版走同一回退链，渲染行为一致
            report.warning(
                "FONT_EMBED_FILE_MISSING",
                line=0,
                message=f"字体 {family!r}（bold={bold}）在系统里未找到，未内嵌",
                hint="目标机器需自行安装该字体；或改用 fonts.json 白名单内的字体",
            )
            continue
        fs_type = face.fs_type() & 0b11
        if fs_type == 0b01:
            report.warning(
                "FONT_EMBED_LICENSE_RESTRICTED",
                line=0,
                message=f"字体 {family!r} 许可为受限嵌入（fsType=1），未内嵌",
                hint="目标机器需自行安装该字体",
            )
            continue
        data = face.subset_bytes(text)
        if data is None:
            report.warning(
                "FONT_EMBED_SUBSET_FAILED",
                line=0,
                message=f"字体 {family!r} 子集化失败，未内嵌",
                hint="字体文件可能损坏；无碍生成，仅影响无字体机器的渲染",
            )
            continue

        part_seq += 1
        part_name = f"/ppt/fonts/font{part_seq}.fntdata"
        from pptx.opc.package import Part
        from pptx.opc.packuri import PackURI

        font_part = Part(PackURI(part_name), _FONT_CONTENT_TYPE, package=part.package, blob=data)
        r_id = part.relate_to(font_part, _FONT_REL)

        embedded_font = etree.SubElement(font_list, f"{{{_P_NS}}}embeddedFont")
        etree.SubElement(
            embedded_font,
            f"{{{_P_NS}}}font",
            typeface=family,
        )
        style_tag = f"{{{_P_NS}}}{'bold' if bold else 'regular'}"
        style_el = etree.SubElement(embedded_font, style_tag)
        style_el.set(f"{{{_R_NS}}}id", r_id)
        embedded += 1

    if embedded:
        # schema 顺序：embeddedFontLst 在 notesSz 之后、defaultTextStyle 之前
        notes_sz.addnext(font_list)
        pr_element.set("embedTrueTypeFonts", "1")
        pr_element.set("saveSubsetFonts", "1")
        report.stats["fonts_embedded"] = embedded
    else:
        pr_element.remove(font_list)
