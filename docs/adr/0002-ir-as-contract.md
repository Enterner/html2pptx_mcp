# ADR-0002：Slide IR 作为独立契约，校验即 lowering

- 状态：已接受
- 日期：2026-10-24

## 背景

最小验证代码里，validator 产出 `DeckModel`，mapper 再消费它。两者各自持有对「合法」的假设：
validator 在解析圆角时就把 `style` 改写成形状语义（`_ellipse` 这类下划线私有键），mapper
又往同一个可变对象里回填图片实际落位。语义靠 dict 键传递，校验与生成之间没有单一判定点，
一旦两条路径分叉，就会出现「校验通过但产物不同」。

同时，HTML 只是众多潜在输入之一。后续的 native 编辑工具、模板系统、Markdown 前端都需要
产出同一种中间表示，否则每加一个前端就要重写一遍发射逻辑。

## 决定

1. 定义 **Slide IR**（frozen dataclass，可 JSON 序列化，浮点固定舍入），作为前端与发射器
   之间唯一接口。IR 有自己的版本号，与契约版本解耦。
2. **校验即 lowering**：`validate_html` 与 `compile_html` 共用同一次遍历，validate 只是
   不执行 emit。报告与产物由同一套判定产生。
3. 语义显式建模：`geometry = Rect | RoundRect | Ellipse`、`fill/stroke/shadow`、`Connector`，
   不再用 dict 私有键承载语义。
4. 图片字节不进 IR：IR 只存 `src + fit`；字节、自然尺寸与实际落位由 `AssetStore` 在
   发射阶段解析，保证 `validate_html` 完全零 I/O。

## 后果

- 报告与产物天然一致；IR 可以逐字节对比，回归测试从「比对最终 pptx」前移到「比对 IR」。
- 发射器可替换（将来接 PDF/PNG 预览），前端可增加（native 工具、Markdown）。
- 代价：IR 需要版本与迁移测试；v1 只建模能被发射器消费的字段，不做布局求解、不做样式继承树，
  避免过度设计。
