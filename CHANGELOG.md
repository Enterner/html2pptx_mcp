# Changelog

格式参考 [Keep a Changelog](https://keepachangelog.com/zh-CN/1.1.0/)，
版本号遵循 [Semantic Versioning](https://semver.org/lang/zh-CN/)。

## [0.4.0a1] - 2026-10-25

首个预览版（初次发布）。

### Added

- **受限 HTML 契约 v1**：37 属性 / 21 类值解析 / 29 条修复规则 / 41 个诊断码，
  契约即数据（contract.json + fonts.json + colors.json），参考文档由脚本生成并在
  CI 强制与数据一致；契约外写法一律报错并附可照抄的 hint。
- **HTML → Slide IR → 原生 DrawingML**：单次遍历，校验与生成同源（validate 零 I/O）；
  渐变、单层外阴影、clip-path、连接线直写 DrawingML，产物 100% 原生可编辑，无截图。
- **MCP 服务（5 工具）**：`validate_html` / `convert_html`（字体子集内嵌，默认开）/
  `get_contract` / `ppt_outline` / `ppt_reverse`；stdio 与 streamable-http 双传输。
- **参考 pptx 逆转**：`ppt_outline` 角色盘点（主题色板/字体体系/形状角色初判）+
  按页/按元素选择性还原；保真要点：主题色 shade/tint/lumMod 变换、ea 字体槽优先、
  组合 chOff/chExt 展平、flip/rotate 叠加、box-sizing 边到边口径。实测答辩模板页
  与原版渲染 diff 2.43%。
- **字体子集内嵌**：按 (家族, 粗体) 收集用量，fontTools.subset 只留出现过的码点，
  fntdata 部件 + p:embeddedFontLst 写入；fsType 许可位与缺字体护栏（缺字体时与
  原版走同一回退链）。
- **textfit 真实字体度量**：hmtx 前进宽 + hhea 行高（等线≈1.04 / 雅黑≈1.32），
  逐 run 家族/粗细分别计量；缺字体回退启发式，度量精度只升不降。
- **版式能力**：绝对定位、单行 flex 子集（justify-content / align-items / gap /
  flex 简写）、竖排、简单表格、SVG 三图元、图片 cover 裁剪、旋转、透明度、
  行高居中惯用法、视口壳折叠。
- **测试**：65 个单测（值解析 / textfit / 字体解析 / 契约一致性 / 逆转断言）+
  视觉验收（真实 PowerPoint 渲染像素 diff，compare_visual）。
- **分发**：PyPI trusted publishing（tag 触发），`uvx html2pptx-mcp` 一行接入；
  wheel 自检脚本随仓库提供。

### Removed

- `poc/` 冻结原型（已被 src 正式实现取代；旧历史备份于仓库外的
  ppt-history.bundle），连同 `scripts/trial.py` 试用采集脚本。
- 旧施工文档：docs/PPT-MCP实施方案.md、docs/temp/阶段一施工方案.md
  （结论已由 README / CHANGELOG / docs/adr 承载）。
