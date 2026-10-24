"""契约文档生成器。

三份产物共用同一份契约数据，避免「文档写 A、实现做 B」：

- ``html.json``：完整机读契约（CLI 与外部工具消费）
- ``html.md``：人读参考（评审属性表时看这份）
- ``html.llm.md``：压缩版规范，给生成 HTML 的模型做提示词，控制在千余 token

``--check`` 模式供 CI 使用：提交物与数据不一致时失败，而不是悄悄过期。
"""

from __future__ import annotations

import json
import sys
from pathlib import Path

from .contract import Contract, PropertySpec

GROUP_LABELS: dict[str, str] = {
    "layout": "布局",
    "paint": "绘制",
    "text": "文本",
    "media": "媒体",
    "compat": "兼容性收口",
}

STATUS_MARK: dict[str, str] = {
    "stable": "",
    "experimental": "（实验）",
    "deprecated": "（废弃）",
}


def render_json(contract: Contract) -> str:
    return json.dumps(contract.to_dict(), ensure_ascii=False, indent=2) + "\n"


def _constraints(spec: PropertySpec) -> str:
    if not spec.constraints:
        return "—"
    return json.dumps(dict(spec.constraints), ensure_ascii=False, separators=(",", ":"))


def _property_row(spec: PropertySpec) -> str:
    values = f"`{'` / `'.join(spec.values)}`" if spec.values else "—"
    fixes = "<br>".join(f"`{rule}`" for rule in spec.fixes) or "—"
    mark = STATUS_MARK.get(spec.status, "")
    note = spec.note or "—"
    return (
        f"| `{spec.name}` | {spec.kind} | {spec.syntax} | {values} | "
        f"{_constraints(spec)} | {fixes} | {note}{mark} |"
    )


def render_markdown(contract: Contract) -> str:
    canvas = contract.canvas
    lines: list[str] = [
        f"# 受限 HTML 契约 `{contract.name}`",
        "",
        "> 由 `scripts/gen_contract_docs.py` 从 `convert/data/` 生成，请勿手工编辑。",
        "",
        "契约之外的写法一律报错并给出可用示例，不做静默丢弃。",
        "",
        "## 画布与单位",
        "",
        f"- 画布：{canvas.width_px:g}×{canvas.height_px:g}px（{canvas.width_px / 96:g}×"
        f"{canvas.height_px / 96:g}in，{canvas.dpi}dpi）",
        f"- 换算：1px = {canvas.emu_per_px} EMU",
        f"- 允许单位：{', '.join(contract.units.allowed)}",
        f"- 拒绝单位：{', '.join(contract.units.rejected)}",
        "",
        "## 文档骨架",
        "",
        "```html",
        "<!DOCTYPE html>",
        "<html>",
        "<body>",
        f'  <div style="width:{canvas.width_px:g}px;height:{canvas.height_px:g}px;'
        'background-color:#FFFFFF">',
        '    <div style="position:absolute;left:80px;top:72px;width:600px;height:64px;',
        '                font-size:40px;font-weight:bold;color:#0F172A">标题</div>',
        "  </div>",
        "</body>",
        "</html>",
        "```",
        "",
        "`body` 下必须恰好一个根 `<div>`，它就是画布；根上只接受背景、默认字体与画布尺寸。",
        "",
        "## 标签白名单",
        "",
        "| 标签 | 类别 | 说明 |",
        "|---|---|---|",
    ]
    for tag in contract.tags.values():
        lines.append(f"| `<{tag.name}>` | {tag.category} | {tag.note or '—'} |")

    lines += [
        "",
        "## 属性白名单",
        "",
        "| 属性 | kind | 写法 | 枚举值 | 约束 | 自动修复 | 说明 |",
        "|---|---|---|---|---|---|---|",
    ]
    for group, label in GROUP_LABELS.items():
        specs = [spec for spec in contract.properties.values() if spec.group == group]
        if not specs:
            continue
        lines.append(f"| **{label}** | | | | | | |")
        lines.extend(_property_row(spec) for spec in specs)

    lines += [
        "",
        "## 自动修复规则",
        "",
        "| 规则 id | 有损 | 说明 |",
        "|---|---|---|",
    ]
    for rule in contract.fix_rules.values():
        loss = "是" if not rule.lossless else "否"
        lines.append(f"| `{rule.id}` | {loss} | {rule.summary} |")

    lines += [
        "",
        "## 诊断码",
        "",
        "| code | 级别 | 说明 |",
        "|---|---|---|",
    ]
    for diag in contract.diagnostics.values():
        lines.append(f"| `{diag.code}` | {diag.severity} | {diag.summary} |")

    limits = contract.limits
    lines += [
        "",
        "## 限额",
        "",
        "| 项 | 上限 |",
        "|---|---|",
        f"| HTML 字节数 | {limits.html_bytes} |",
        f"| 单页元素数 | {limits.elements_per_slide} |",
        f"| 嵌套深度 | {limits.nesting_depth} |",
        f"| 单元素文本字符数 | {limits.text_chars} |",
        f"| 单元素段落数 | {limits.paragraphs} |",
        f"| 表格行 × 列 | {limits.table_rows} × {limits.table_cols} |",
        f"| 单张图片字节数 | {limits.image_bytes} |",
        "",
        "## 扩展约定",
        "",
        "- 契约只有当前形态，不做版本协商。",
        "- 扩展只允许新增属性/标签/取值/修复规则；已有写法的语义不变。",
        "- 表外写法一律拒绝，绝不静默丢弃或近似处理。",
        "",
    ]
    return "\n".join(lines)


def render_llm(contract: Contract) -> str:
    """给生成模型的压缩规范：只保留会改变产物或导致拒绝的信息。"""
    canvas = contract.canvas
    tags = " / ".join(f"<{name}>" for name in contract.tags)
    lines: list[str] = [
        f"# {contract.name}（受限 HTML → PPTX）",
        "",
        f"画布固定 {canvas.width_px:g}×{canvas.height_px:g}px，1px = {canvas.emu_per_px} EMU；"
        "只允许绝对定位，禁止 flex/grid/文档流/z-index。",
        "",
        "## 硬规则",
        "",
        "1. `body` 下恰好一个根 `<div>` 作为画布，写 width/height 与背景；"
        "根 div 不需要 position/left/top。",
        "2. 每个有文本或视觉的元素必须写全 "
        "`position:absolute;left;top;width;height`（px 最省事）。",
        "3. 坐标相对最近的定位祖先；z 序 = DOM 顺序，后写的盖住先写的。",
        "4. 文本块内只允许 `<br>` 分段与 "
        '`<b>/<i>/<span style="color|font-size|font-weight|font-family">`。',
        "5. 同一 div 里不要同时放文本和子元素：文本单独一个 div，子元素另一个 div。",
        "6. 样式可写在 `<style>` 的 class 里（仅 `.class` 单选择器，无嵌套/伪类/后代）"
        "或元素内联 style；同名属性内联优先。",
        "7. 契约之外的属性或取值会被拒绝；报错带 hint，照 hint 改。",
        "",
        "## 优先写法（能减少自动修复）",
        "",
        "- 颜色优先 `#RRGGBB`；只有需要半透明时才用 `rgba()`。",
        "- 坐标与尺寸优先 px；只有需要按父级比例缩放时才用 `%`。",
        "- 行高优先倍数（如 1.5）；只有需要固定像素行高时才用 px。",
        "- 字体用白名单里的名字（如 微软雅黑 / 思源黑体 / Segoe UI）。",
        "",
        f"## 允许标签\n\n{tags}",
        "",
        "## 允许属性",
        "",
    ]
    for group, label in GROUP_LABELS.items():
        specs = [spec for spec in contract.properties.values() if spec.group == group]
        if not specs:
            continue
        body = "；".join(f"`{spec.name}`={spec.syntax}" for spec in specs)
        lines.append(f"- **{label}**：{body}")
    lines.append(
        "- **表格单元格**：`background-color` / `color` / `border` / `text-align` / "
        "`font-weight` / `padding`（作用于单个 td/th）；`<b>` / `<i>` 生效。"
    )
    lines += [
        "",
        "## 示例",
        "",
        "```html",
        '<div style="width:1280px;height:720px;background-color:#FFFFFF">',
        '  <div style="position:absolute;left:80px;top:64px;width:800px;height:72px;',
        '              font-size:48px;font-weight:bold;color:#0F172A">季度经营回顾</div>',
        '  <div style="position:absolute;left:80px;top:160px;width:360px;height:220px;',
        "              background-color:#EEF2FF;border-radius:16px;",
        '              box-shadow:0 4px 12px rgba(15,23,42,0.08)"></div>',
        '  <img style="position:absolute;left:480px;top:160px;width:720px;height:420px;',
        '       object-fit:contain" src="assets/photo.png">',
        "</div>",
        "```",
        "",
        "## 常见拒绝与改法",
        "",
        "- `display:flex` → 用 left/top 逐个定位子元素。",
        "- 缺 width/height → 补齐盒（盒是权威，不做自动缩放）。",
        "- `calc()` / `%` 无参照 / `vw` → 换成 px。",
        "- `text-shadow` / `filter` / 多层阴影 → 不支持，去掉。",
        "- `<section>/<p>/<h1>` → 换成定位的 `<div>`，段落用 `<br>`。",
        "",
    ]
    return "\n".join(lines)


ARTIFACTS = {
    "html.json": render_json,
    "html.md": render_markdown,
    "html.llm.md": render_llm,
}


def write_docs(contract: Contract, out_dir: Path, *, check: bool = False) -> int:
    """生成或校验三份契约文档。返回 0 表示一致，1 表示存在差异（check 模式）。"""
    out_dir.mkdir(parents=True, exist_ok=True)
    stale: list[str] = []
    for name, render in ARTIFACTS.items():
        target = out_dir / name
        content = render(contract)
        if check:
            current = target.read_text(encoding="utf-8") if target.is_file() else None
            if current != content:
                stale.append(name)
        else:
            target.write_text(content, encoding="utf-8", newline="\n")
    if check:
        for name in stale:
            print(f"契约文档已过期：{out_dir / name}", file=sys.stderr)
        if stale:
            print(
                "运行 `uv run python scripts/gen_contract_docs.py` 重新生成并提交。",
                file=sys.stderr,
            )
            return 1
    return 0
