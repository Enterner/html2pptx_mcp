# pptx-html（受限 HTML → PPTX）

画布固定 1280×720px，1px = 9525 EMU；只允许绝对定位，禁止 flex/grid/文档流/z-index。

## 硬规则

1. `body` 下恰好一个根 `<div>` 作为画布，写 width/height 与背景；根 div 不需要 position/left/top。
2. 每个有文本或视觉的元素必须写全 `position:absolute;left;top;width;height`（px 最省事）。
3. 坐标相对最近的定位祖先；z 序 = DOM 顺序，后写的盖住先写的。
4. 文本块内只允许 `<br>` 分段与 `<b>/<i>/<span style="color|font-size|font-weight|font-family">`。
5. 同一 div 里不要同时放文本和子元素：文本单独一个 div，子元素另一个 div。
6. 样式可写在 `<style>` 的 class 里（仅 `.class` 单选择器，无嵌套/伪类/后代）或元素内联 style；同名属性内联优先。
7. 契约之外的属性或取值会被拒绝；报错带 hint，照 hint 改。

## 优先写法（能减少自动修复）

- 颜色优先 `#RRGGBB`；只有需要半透明时才用 `rgba()`。
- 坐标与尺寸优先 px；只有需要按父级比例缩放时才用 `%`。
- 行高优先倍数（如 1.5）；只有需要固定像素行高时才用 px。
- 字体用白名单里的名字（如 微软雅黑 / 思源黑体 / Segoe UI）。

## 允许标签

<div> / <img> / <table> / <thead> / <tbody> / <tr> / <td> / <th> / <svg> / <rect> / <circle> / <line> / <b> / <strong> / <em> / <i> / <span> / <br>

## 允许属性

- **布局**：`position`=absolute / relative；`left`=长度（px / em / % / pt）；`top`=长度（px / em / % / pt）；`width`=长度（px / em / % / pt）；`height`=长度（px / em / % / pt）；`transform`=rotate(<N>deg)；`transform-origin`=center / center center；`flex-direction`=row / column；`justify-content`=flex-start / center / flex-end / space-between / space-around；`align-items`=flex-start / center / flex-end / stretch；`gap`=长度（px / em / pt）；`flex`=none / auto / <grow> [<shrink> [auto]]
- **绘制**：`background`=纯色 或 linear-gradient(<角度或方向>, <色标> 0%, <色标> 100%)；`background-color`=颜色（#hex / rgb() / rgba() / 具名色）；`border`=<宽度>px solid|dashed|dotted <颜色> 或 none；`border-radius`=单值：<长度> 或 <百分比>；`box-shadow`=dx dy [blur] <颜色>；`opacity`=0–1；`clip-path`=polygon(x y, ...)，坐标 %/px，至少 3 个顶点
- **文本**：`font-size`=长度（px / em / % / pt），6–200px；`font-weight`=normal / bold / 100–900；`font-family`=白名单字体名，逗号分隔；`color`=颜色（#hex / rgb() / rgba() / 具名色）；`line-height`=倍数 / <长度> / <百分比>；`text-align`=left / center / right / justify；`padding`=1–4 个长度；`white-space`=normal / nowrap；`letter-spacing`=长度（px，可负），0 表示无；`writing-mode`=vertical-rl；`text-indent`=长度（px，可负）；`padding-top`=长度（px / em / pt）；`padding-right`=长度（px / em / pt）；`padding-bottom`=长度（px / em / pt）；`padding-left`=长度（px / em / pt）
- **媒体**：`object-fit`=contain / cover
- **兼容性收口**：`display`=none / block / inline / inline-block / flex；`margin`=0（1–4 个 0）
- **表格单元格**：`background-color` / `color` / `border` / `text-align` / `font-weight` / `padding`（作用于单个 td/th）；`<b>` / `<i>` 生效。

## 示例

```html
<div style="width:1280px;height:720px;background-color:#FFFFFF">
  <div style="position:absolute;left:80px;top:64px;width:800px;height:72px;
              font-size:48px;font-weight:bold;color:#0F172A">季度经营回顾</div>
  <div style="position:absolute;left:80px;top:160px;width:360px;height:220px;
              background-color:#EEF2FF;border-radius:16px;
              box-shadow:0 4px 12px rgba(15,23,42,0.08)"></div>
  <img style="position:absolute;left:480px;top:160px;width:720px;height:420px;
       object-fit:contain" src="assets/photo.png">
</div>
```

## 常见拒绝与改法

- `display:flex` → 用 left/top 逐个定位子元素。
- 缺 width/height → 补齐盒（盒是权威，不做自动缩放）。
- `calc()` / `%` 无参照 / `vw` → 换成 px。
- `text-shadow` / `filter` / 多层阴影 → 不支持，去掉。
- `<section>/<p>/<h1>` → 换成定位的 `<div>`，段落用 `<br>`。
