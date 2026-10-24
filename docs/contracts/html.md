# 受限 HTML 契约 `pptx-html`

> 由 `scripts/gen_contract_docs.py` 从 `convert/data/` 生成，请勿手工编辑。

契约之外的写法一律报错并给出可用示例，不做静默丢弃。

## 画布与单位

- 画布：1280×720px（13.3333×7.5in，96dpi）
- 换算：1px = 9525 EMU
- 允许单位：px, em, %, pt, vw, vh
- 拒绝单位：rem, vmin, vmax, ex, ch, cm, mm, in, q

## 文档骨架

```html
<!DOCTYPE html>
<html>
<body>
  <div style="width:1280px;height:720px;background-color:#FFFFFF">
    <div style="position:absolute;left:80px;top:72px;width:600px;height:64px;
                font-size:40px;font-weight:bold;color:#0F172A">标题</div>
  </div>
</body>
</html>
```

`body` 下必须恰好一个根 `<div>`，它就是画布；根上只接受背景、默认字体与画布尺寸。

## 标签白名单

| 标签 | 类别 | 说明 |
|---|---|---|
| `<div>` | structural | 唯一的块级元素：承载文本、形状，或作为定位容器 |
| `<img>` | media | src 仅接受 base_dir 内的相对路径或 data:image/*;base64 |
| `<table>` | table | 简单表格：无合并单元格、无嵌套 |
| `<thead>` | table | — |
| `<tbody>` | table | — |
| `<tr>` | table | — |
| `<td>` | table | 单元格支持 background-color / color / border / text-align / font-weight / padding；<b>/<i> 生效 |
| `<th>` | table | — |
| `<svg>` | vector | — |
| `<rect>` | vector | — |
| `<circle>` | vector | — |
| `<line>` | vector | — |
| `<b>` | inline | — |
| `<strong>` | inline | — |
| `<em>` | inline | — |
| `<i>` | inline | — |
| `<span>` | inline | 内联样式仅接受 color / font-size / font-weight / font-family |
| `<br>` | inline | — |

## 属性白名单

| 属性 | kind | 写法 | 枚举值 | 约束 | 自动修复 | 说明 |
|---|---|---|---|---|---|---|
| **布局** | | | | | | |
| `position` | enum | absolute / relative | `absolute` / `relative` | — | `fix.structure.implicit-absolute` | absolute 为唯一布局模型；relative 在盒已确定时等价（作为定位祖先），缺 left/top 的 relative 元素由 flex 布局给盒 |
| `left` | length | 长度（px / em / % / pt） | — | {"percent_ref":"nearest_positioned_ancestor.width","allow_negative":true} | `fix.unit.percent-to-px`<br>`fix.unit.em-to-px`<br>`fix.unit.pt-to-px` | — |
| `top` | length | 长度（px / em / % / pt） | — | {"percent_ref":"nearest_positioned_ancestor.height","allow_negative":true} | `fix.unit.percent-to-px`<br>`fix.unit.em-to-px`<br>`fix.unit.pt-to-px` | — |
| `width` | length | 长度（px / em / % / pt） | — | {"percent_ref":"nearest_positioned_ancestor.width","min_px":0} | `fix.unit.percent-to-px`<br>`fix.unit.em-to-px`<br>`fix.unit.pt-to-px` | — |
| `height` | length | 长度（px / em / % / pt） | — | {"percent_ref":"nearest_positioned_ancestor.height","min_px":0} | `fix.unit.percent-to-px`<br>`fix.unit.em-to-px`<br>`fix.unit.pt-to-px` | — |
| `transform` | transform | rotate(<N>deg) | — | — | — | 只支持旋转；translate/scale/skew 会改变盒语义，不属于盒模型 |
| `transform-origin` | transform-origin | center / center center | `center` / `center center` | — | `fix.style.transform-origin-dropped` | DrawingML 旋转枢轴恒为形状中心；非 center 原点会改变画面，直接拒绝 |
| `flex-direction` | enum | row / column | `row` / `column` | — | — | 单行 flex：不支持 wrap 与 *-reverse |
| `justify-content` | enum | flex-start / center / flex-end / space-between / space-around | `flex-start` / `center` / `flex-end` / `space-between` / `space-around` | — | — | — |
| `align-items` | enum | flex-start / center / flex-end / stretch | `flex-start` / `center` / `flex-end` / `stretch` | — | — | — |
| `gap` | length | 长度（px / em / pt） | — | {"min_px":0} | `fix.unit.percent-to-px`<br>`fix.unit.em-to-px`<br>`fix.unit.pt-to-px` | 单值：主轴与交叉轴同距（row-gap/column-gap 长写不支持） |
| `flex` | flex | none / auto / <grow> [<shrink> [auto]] | — | — | — | flex-basis 仅 auto；主轴尺寸由声明盒或内容估算给出 |
| **绘制** | | | | | | |
| `background` | fill | 纯色 或 linear-gradient(<角度或方向>, <色标> 0%, <色标> 100%) | — | {"gradient_stops":2} | `fix.color.rgb-to-hex`<br>`fix.color.named-to-hex` | — |
| `background-color` | color | 颜色（#hex / rgb() / rgba() / 具名色） | — | — | `fix.color.rgb-to-hex`<br>`fix.color.named-to-hex` | — |
| `border` | border | <宽度>px solid|dashed|dotted <颜色> 或 none | — | {"styles":["solid","dashed","dotted"]} | `fix.unit.percent-to-px`<br>`fix.unit.em-to-px`<br>`fix.unit.pt-to-px`<br>`fix.color.rgb-to-hex`<br>`fix.color.named-to-hex` | — |
| `border-radius` | radius | 单值：<长度> 或 <百分比> | — | {"min_px":0,"percent_max":100} | — | px 半径封顶短边一半（stadium）；正方形满圆与 50% 走椭圆 |
| `box-shadow` | shadow | dx dy [blur] <颜色> | — | {"layers":1,"allow_inset":false,"allow_spread":false} | `fix.color.rgb-to-hex`<br>`fix.color.named-to-hex` | — |
| `opacity` | opacity | 0–1 | — | {"min":0,"max":1} | `fix.paint.opacity-dropped`<br>`fix.media.opacity-dropped` | 无填充/描边/阴影的元素上无意义，会被丢弃并记录 |
| `clip-path` | clip-path | polygon(x y, ...)，坐标 %/px，至少 3 个顶点 | — | {"shapes":["polygon"]} | — | 仅 polygon；顶点按声明盒换算为自由多边形，其余 clip-path 形态拒绝 |
| **文本** | | | | | | |
| `font-size` | length | 长度（px / em / % / pt），6–200px | — | {"min_px":6,"max_px":200,"percent_ref":"parent.font_size"} | `fix.unit.percent-to-px`<br>`fix.unit.em-to-px`<br>`fix.unit.pt-to-px` | — |
| `font-weight` | font-weight | normal / bold / 100–900 | — | {"min":100,"max":900,"bold_threshold":600} | — | — |
| `font-family` | font | 白名单字体名，逗号分隔 | — | — | `fix.font.whitelist` | 取首个命中白名单的字体名；无命中时回落到策略默认字体并记 lossy 修复 |
| `color` | color | 颜色（#hex / rgb() / rgba() / 具名色） | — | — | `fix.color.rgb-to-hex`<br>`fix.color.named-to-hex` | — |
| `line-height` | line-height | 倍数 / <长度> / <百分比> | — | {"min_multiple":0,"max_multiple":5} | `fix.unit.percent-to-multiple`<br>`fix.unit.em-to-px`<br>`fix.unit.pt-to-px` | — |
| `text-align` | enum | left / center / right / justify | `left` / `center` / `right` / `justify` | — | — | — |
| `padding` | padding | 1–4 个长度 | — | — | `fix.unit.percent-to-px`<br>`fix.unit.em-to-px`<br>`fix.unit.pt-to-px` | — |
| `white-space` | enum | normal / nowrap | `normal` / `nowrap` | — | — | — |
| `letter-spacing` | letter-spacing | 长度（px，可负），0 表示无 | — | — | — | 任意 px 字距映射为 run 级 spc（1/100 pt） |
| `writing-mode` | writing-mode | vertical-rl | `vertical-rl` / `horizontal-tb` | — | — | horizontal-tb 为默认，按无操作处理 |
| `text-indent` | text-indent | 长度（px，可负） | — | — | — | 按 CSS 语义只作用于首段首行 |
| `padding-top` | length | 长度（px / em / pt） | — | {"min_px":0} | `fix.unit.em-to-px`<br>`fix.unit.pt-to-px` | padding 长写，按书写顺序覆盖简写对应边 |
| `padding-right` | length | 长度（px / em / pt） | — | {"min_px":0} | `fix.unit.em-to-px`<br>`fix.unit.pt-to-px` | padding 长写，按书写顺序覆盖简写对应边 |
| `padding-bottom` | length | 长度（px / em / pt） | — | {"min_px":0} | `fix.unit.em-to-px`<br>`fix.unit.pt-to-px` | padding 长写，按书写顺序覆盖简写对应边 |
| `padding-left` | length | 长度（px / em / pt） | — | {"min_px":0} | `fix.unit.em-to-px`<br>`fix.unit.pt-to-px` | padding 长写，按书写顺序覆盖简写对应边 |
| **媒体** | | | | | | |
| `object-fit` | enum | contain / cover | `contain` / `cover` | — | — | contain 等比缩放并留白；cover 等比缩放并裁剪超出部分 |
| **兼容性收口** | | | | | | |
| `display` | display | none / block / inline / inline-block / flex | `none` / `block` / `inline` / `inline-block` / `flex` | — | `fix.structure.hidden-subtree`<br>`fix.compat.display-dropped` | none 剪掉整棵子树；block/inline 在绝对定位模型下无意义，丢弃；flex 触发单行 flex 布局（见 flex-direction 等） |
| `margin` | margin | 0（1–4 个 0） | — | {"only":0} | `fix.compat.margin-zero-dropped` | 定位用 left/top；margin:0 是模型的常见噪声，直接丢弃 |

## 自动修复规则

| 规则 id | 有损 | 说明 |
|---|---|---|
| `fix.structure.implicit-absolute` | 否 | 元素声明了 left/top 却缺 position 时补 absolute |
| `fix.structure.root-origin` | 否 | 根画布上的 position 与 left/top=0 属噪声，忽略不影响语义 |
| `fix.structure.hidden-subtree` | 否 | display:none 的元素连同子树不进入 IR |
| `fix.unit.percent-to-px` | 否 | 百分比按最近定位祖先的对应轴换算为 px |
| `fix.unit.percent-to-multiple` | 否 | line-height 的百分比换算为倍数 |
| `fix.unit.em-to-px` | 否 | em 按当前继承字号换算为 px |
| `fix.unit.pt-to-px` | 否 | pt 按 96/72 换算为 px |
| `fix.color.rgb-to-hex` | 否 | rgb()/rgba() 归一为 #RRGGBB，alpha 单独保留 |
| `fix.color.named-to-hex` | 否 | 具名颜色归一为 #RRGGBB |
| `fix.font.whitelist` | 是 | 字体族回落到当前策略白名单中的命中项或默认字体，字形可能变化 |
| `fix.compat.display-dropped` | 否 | display 的 block/inline 值在绝对定位模型下丢弃 |
| `fix.compat.margin-zero-dropped` | 否 | margin:0 丢弃 |
| `fix.table.pad-columns` | 否 | 行列数不齐时补空单元格到最大列数 |
| `fix.table.cell-width-ignored` | 否 | 单元格 width 由表格列宽决定，声明被忽略 |
| `fix.table.cell-formatting-dropped` | 是 | 单元格内除 <b>/<i> 外的内联标签样式丢弃，只保留文本 |
| `fix.paint.opacity-dropped` | 是 | 无填充/描边/阴影的元素上的 opacity 无意义，丢弃 |
| `fix.media.opacity-dropped` | 是 | 图片透明度 v1 不支持，丢弃 |
| `fix.compat.box-sizing-dropped` | 否 | box-sizing:border-box 与直写宽高语义一致，声明被吸收 |
| `fix.compat.overflow-dropped` | 否 | 文本溢出由自动缩字处理，overflow:hidden 被吸收 |
| `fix.style.transform-origin-dropped` | 否 | transform-origin 恒为中心（旋转枢轴），声明丢弃 |
| `fix.structure.relative-normalized` | 否 | position:relative 按声明盒定位处理（绝对盒模型无文档流） |
| `fix.unit.vw-to-px` | 否 | vw 按画布宽换算（视口即 1280×720 画布） |
| `fix.unit.vh-to-px` | 否 | vh 按画布高换算（视口即 1280×720 画布） |
| `fix.layout.flex-text-absorbed` | 否 | 纯文本 flex 容器：居中/对齐吸收为文本对齐 |
| `fix.layout.flex-offset-dropped` | 否 | flex 子项的 left/top 在静态流中本就不生效，丢弃 |
| `fix.structure.viewport-flattened` | 否 | 视口居中壳（100vw/100vh flex + 画布尺寸 stage）合并为画布 |
| `fix.structure.root-noise-dropped` | 否 | 根画布上无意义的 flex 声明丢弃 |
| `fix.layout.static-fill` | 否 | 无定位的图片按流式规则放在父内容区左上角 |
| `fix.media.border-radius-dropped` | 是 | 图片圆角 v1 不支持，按直角渲染 |

## 诊断码

| code | 级别 | 说明 |
|---|---|---|
| `CONTRACT_UNKNOWN_PROPERTY` | error | CSS 属性不在 v1 白名单 |
| `CONTRACT_UNKNOWN_TAG` | error | HTML 标签不在 v1 白名单 |
| `CONTRACT_UNKNOWN_ATTRIBUTE` | error | SVG 属性不在白名单 |
| `CONTRACT_UNSUPPORTED_VALUE` | error | 属性取值不受支持 |
| `CONTRACT_UNSUPPORTED_UNIT` | error | 长度单位或写法不受支持 |
| `STRUCTURE_ROOT_REQUIRED` | error | body 下必须恰好有一个根 div（画布） |
| `STRUCTURE_ROOT_PROPERTY_FORBIDDEN` | error | 根容器只接受背景、默认字体与画布尺寸 |
| `STRUCTURE_CANVAS_MISMATCH` | error | 根容器声明的宽高与固定画布不一致 |
| `STRUCTURE_ORPHAN_TEXT` | error | 游离文本没有定位盒 |
| `STRUCTURE_MISSING_BOX` | error | 元素缺少 left/top/width/height |
| `STRUCTURE_MIXED_CONTENT` | error | 同一 div 内混排文本与结构子元素 |
| `STRUCTURE_DEPTH_EXCEEDED` | error | 嵌套深度超过限额 |
| `STRUCTURE_TAG_FORBIDDEN_CHILD` | error | 标签出现在不允许的父元素下 |
| `TABLE_MERGED_CELLS` | error | 表格包含 rowspan/colspan |
| `TABLE_NESTED` | error | 表格内嵌表格 |
| `TABLE_EMPTY` | error | 表格没有数据行 |
| `TABLE_CELL_FORMATTING_DROPPED` | warning | 单元格内的内联样式被丢弃，只保留文本 |
| `SVG_UNSUPPORTED_ELEMENT` | error | SVG 图元不在白名单（仅 rect/circle/line） |
| `SVG_INVALID_GEOMETRY` | error | SVG 图元缺少必要几何属性或取值为非正 |
| `ASSET_SRC_MISSING` | error | 图片缺少 src 或 src 为空 |
| `ASSET_NOT_FOUND` | error | 图片缺少 src 或 src 为空 |
| `ASSET_PATH_DENIED` | error | 图片路径越出 base_dir 或使用绝对路径/UNC |
| `ASSET_TOO_LARGE` | error | 图片体积超过限额 |
| `ASSET_UNSUPPORTED_TYPE` | error | 图片 MIME 不在 image/* 或无法识别 |
| `ASSET_DECODE_FAILED` | error | 图片字节无法解码 |
| `LIMIT_ELEMENTS_EXCEEDED` | error | 单页元素数超过限额 |
| `LIMIT_HTML_EXCEEDED` | error | 单页 HTML 字节数超过限额 |
| `LIMIT_TEXT_EXCEEDED` | error | 单元素文本量超过限额 |
| `LIMIT_TABLE_EXCEEDED` | error | 表格行列数超过限额 |
| `LIMIT_IMAGE_EXCEEDED` | error | 图片像素尺寸超过限额 |
| `FONT_EMBED_FILE_MISSING` | warning | 字体文件在系统里未找到，未内嵌 |
| `FONT_EMBED_LICENSE_RESTRICTED` | warning | 字体许可为受限嵌入（fsType=1），未内嵌 |
| `FONT_EMBED_LIMIT` | warning | 内嵌字体超过上限，剩余家族跳过 |
| `FONT_EMBED_SUBSET_FAILED` | warning | 字体子集化失败，未内嵌 |
| `FONT_EMBED_SKIPPED` | warning | presentation.xml 结构异常，跳过字体内嵌 |
| `TEXT_OVERFLOW_ESTIMATED` | warning | 估算文本高度超出盒子，可能溢出 |
| `FONT_FALLBACK_APPLIED` | warning | 字体在目标机器可能缺失，回退后换行可能变化 |
| `FIX_LOSSY_APPLIED` | warning | 表格行列数超过限额 |
| `STYLE_AT_RULE_IGNORED` | warning | 样式表 @ 规则（font-face/media 等）被忽略 |
| `STYLE_SELECTOR_UNSUPPORTED` | error | 选择器仅支持标签 / .class / 标签.class 组合 |
| `STRUCTURE_FLEX_INVALID` | error | flex 容器/子项无法求值（缺尺寸、内容区为零等） |

## 限额

| 项 | 上限 |
|---|---|
| HTML 字节数 | 2000000 |
| 单页元素数 | 150 |
| 嵌套深度 | 12 |
| 单元素文本字符数 | 20000 |
| 单元素段落数 | 200 |
| 表格行 × 列 | 50 × 12 |
| 单张图片字节数 | 10485760 |

## 扩展约定

- 契约只有当前形态，不做版本协商。
- 扩展只允许新增属性/标签/取值/修复规则；已有写法的语义不变。
- 表外写法一律拒绝，绝不静默丢弃或近似处理。
