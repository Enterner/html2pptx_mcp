# html2pptx-mcp

用**受限 HTML 契约**生成**原生可编辑**的 PowerPoint（`.pptx`），并提供 MCP 服务接口。

> 现状：`0.1.0a1` 预览版——转换器、MCP 服务（5 工具）、参考 pptx 逆转与字体子集内嵌均可用。

## 安装

```bash
# 发布后：一行接入任意 MCP 客户端（无需克隆仓库）
uvx html2pptx-mcp

# 或 pip 安装
pip install html2pptx-mcp
```

## 为什么不是「HTML 直接转 PPT」

自由 HTML/CSS 到 PPTX 的有损转换，会让你在浏览器里看到的和交付出去的不是同一个东西。
本项目反过来：**先用契约把可表达的写法收窄到能一一映射到 DrawingML 的子集**，契约之外的
写法一律报错并给出可用写法，而不是静默丢弃。

产物是真正的 PPT 对象（文本框、形状、表格、图片），可以在 PowerPoint 里继续编辑，不是截图。

## 架构

```
受限 HTML --lower--> Slide IR(JSON) --emit--> .pptx
                  \-> 诊断报告（error / fix / warning，全部机器可读）
```

- **契约是数据**：属性、标签、取值、修复规则、错误码写在 `convert/data/*.json`；解析器按 kind 注册，
  CI 双向对账，杜绝「文档说有、实现没做」。
- **IR 是第二个契约**：HTML 只是一个前端；发射器只认 IR，后续 native 工具 / 模板系统复用同一层。
- **单次遍历**：校验与生成共用一条 lowering 路径，不会出现「校验通过但生成不一致」。

## 快速开始（开发）

## 快速开始

需要 Python 3.11+ 与 [uv](https://docs.astral.sh/uv/)。

```bash
uv sync --extra dev
uv run pytest
uv run ruff check .
uv run mypy
```

### 命令行

```bash
# 校验（不读图片、不生成文件）
uv run html2pptx-cli validate examples/html/04_bad.html

# 生成单页 / 多页
uv run html2pptx-cli build examples/html/01_title.html -o out/01.pptx
uv run html2pptx-cli build examples/html/01_title.html examples/html/03_data.html \
    -o out/deck.pptx --report out/report.json

# 导出契约 / 回读产物几何
uv run html2pptx-cli contract --format json
uv run html2pptx-cli inspect out/deck.pptx
```

退出码：`0` 通过；`2` 存在契约错误（`--strict` 时警告也算失败）；`1` 文件或内部错误。

### Python API

```python
from pathlib import Path

from html2pptx_mcp.convert import CompileOptions, compile_html

result = compile_html(
    Path("examples/html/01_title.html").read_text(encoding="utf-8"),
    options=CompileOptions(base_dir=Path("examples/html")),
)
if result.ok:
    result.save("out/01.pptx")
else:
    for error in result.report.errors:
        print(error.code, error.message, error.hint)
```

`compile_html` 接受单份 HTML 或多份（按顺序拼页），返回结果对象而不是抛异常；
`validate_html` 与它共用同一次遍历，只是不执行发射。

### MCP 服务

```bash
uv run html2pptx-mcp                                             # stdio（默认）
uv run html2pptx-mcp --transport streamable-http --host 127.0.0.1 --port 8000
```

客户端（stdio）配置：

```json
{
  "mcpServers": {
    "ppt": {
      "command": "uv",
      "args": ["run", "--directory", "/path/to/ppt", "html2pptx-mcp"]
    }
  }
}
```

| 类型 | 名称 | 说明 |
|---|---|---|
| Tool | `validate_html` | 零副作用校验，返回 errors / warnings / fixes |
| Tool | `convert_html` | 生成并写出 pptx；失败时以 `isError` 返回完整诊断 |
| Tool | `get_contract` | 导出契约（`llm` 精简规范 / `json` 完整契约） |
| Resource | `pptx://contract` | 完整契约 JSON |
| Resource | `pptx://contract/llm` | 给模型看的精简规范 |
| Prompt | `design_html_slide` | 契约规范 + 任务要求，拼成可直接使用的提示词 |

### 试用采集

```bash
python scripts/trial.py --dir <候选 HTML 目录> --json-out <报告路径>
```

统计一次通过率、错误码分布、自动修复条数与 token 估算。候选稿应在**没跑过校验**的
前提下一次性生成，否则指标失真。首轮自测见 `docs/temp/trial/试用报告.md`。

### 视觉对比

```bash
python scripts/compare_visual.py --dir tests/fixtures/review --out test/review
```

对每页 HTML：Edge 无头截图（基准）vs PowerPoint COM 渲染（产物），输出逐像素差异
（`diff>16` 占比、均值）、差异热点图与上下并排图；`report.json` 汇总。依赖本机
Edge 与 PowerPoint；几何与样式一致性由 pytest 冒烟保证，本脚本回答「看起来像不像」。

## 仓库结构

| 路径 | 职责 |
|---|---|
| `src/html2pptx_mcp/server.py` | MCP 适配层：tools / resources / prompt |
| `src/html2pptx_mcp/__main__.py` | `html2pptx-mcp` 入口（stdio / streamable-http） |
| `src/html2pptx_mcp/convert/contract.py` | 受限 HTML 契约：模型、加载、引用校验（数据在 `data/`） |
| `src/html2pptx_mcp/convert/ir.py` | Slide IR：可序列化的中间表示，发射器的唯一输入 |
| `src/html2pptx_mcp/convert/diagnostics.py` | 诊断账本：error / warning / fix，级别与 lossless 由契约决定 |
| `src/html2pptx_mcp/convert/values.py` | CSS 单值解析（纯函数，无 I/O） |
| `src/html2pptx_mcp/convert/html.py` | DOM + 样式 → Slide IR，唯一的校验判定点 |
| `src/html2pptx_mcp/convert/html_table.py` | 简单表格解析（列宽 / 表头 / 单元格） |
| `src/html2pptx_mcp/convert/html_svg.py` | SVG 子集（rect / circle / line）解析 |
| `src/html2pptx_mcp/convert/emit.py` | IR → python-pptx / DrawingML |
| `src/html2pptx_mcp/convert/textfit.py` | 文本溢出估算（hmtx/hhea 真实字体度量，可注入） |
| `src/html2pptx_mcp/convert/fonts.py` | 系统字体定位与度量（家族别名 → 文件，含中文别名） |
| `src/html2pptx_mcp/convert/fontembed.py` | 字体子集内嵌（fntdata 部件 + embeddedFontLst） |
| `src/html2pptx_mcp/convert/assets.py` | 资源定位与路径沙箱 |
| `src/html2pptx_mcp/convert/reverse.py` | 参考 pptx 逆转：outline 盘点 + 按页/按元素选择性还原为契约 HTML |
| `tests/<名称>/` | 成对测试产物（同名 `.html` + `.pptx` + `assets/`，不入库） |
| `examples/html/` | 契约样例（标题页 / 图文页 / 数据页 / 违规样本） |
| `docs/` | 规范与总方案；契约参考在 `docs/contracts/`（生成物，CI 校验）、ADR 在 `docs/adr/`；施工期文档在 `docs/temp/` |

## 能力边界

- **不做**：动画、页面切换、SmartArt、音视频、OLE；这些对象即使出现在第三方 deck 里也只
  做原样保留，不编辑（打开/编辑第三方 deck 不在本阶段范围）。
- **流式布局只支持单行 flex 子集**：`display:flex` + `row/column` + `justify-content` + `align-items` + `gap` + `flex` 简写（无 wrap、无 `*-reverse`、无嵌套文档流）；其余元素仍须绝对定位并写全盒。
- **表格是简单表**：不支持合并单元格、嵌套表格；单元格只承载纯文本。
- **SVG 只有三个图元**：`rect` / `circle` / `line`；`path`、文本、渐变一律拒绝。
- **不读网络图片**：只接受 `base_dir` 内的相对路径与 `data:image/*;base64`；URL 取图
  属服务化阶段的安全项。
- **预览不是交付判据**：`textfit` 是估算器，真实换行由 PowerPoint 决定；排版请以
  PowerPoint 打开结果为准。
- **单一画布**：1280×720（16:9），其他尺寸属后续扩展；一份 HTML 一页。

## 文档

- [注释规范](docs/注释规范.md)
- [总体实施方案](docs/PPT-MCP实施方案.md)
- [阶段一施工方案](docs/temp/阶段一施工方案.md)

## 许可证

Apache-2.0，见 [LICENSE](LICENSE)。
