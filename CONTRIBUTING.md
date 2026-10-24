# 贡献指南

## 环境

- Python 3.11+
- [uv](https://docs.astral.sh/uv/)（依赖与虚拟环境）

```bash
uv sync --extra dev
uv run pre-commit install   # 可选：提交前自动跑 ruff
```

## 提交前必须通过

```bash
uv run ruff check .
uv run ruff format --check .
uv run mypy
uv run pytest
```

渲染相关用例标记为 `slow`，需要 LibreOffice 的镜像环境，不在普通 runner 上跑。

## 代码约定

### 注释

写代码前先读 [注释规范](../docs/注释规范.md)。要点：

- 中文注释，写 why 与不变式，不复述代码。
- 注释**自包含**：不写「详见 docs/xxx」，把约束本身写进注释。
- 标记统一 `NOTE:` / `踩过:` / `TODO:` / `HACK:`。
- 公开面（导出函数、类、端点）必须有文档注释：职责 + 副作用时机 + 失败行为。
- 闭源期不写 license 头；公开时统一补 SPDX。

### 分层

`src/html2pptx_mcp/convert/` 内部依赖方向由 `tests/test_architecture.py` 强制：

```
contract, ir, diagnostics  ←  values  ←  css  ←  lower  →  emit
                                                  ↘ assets ↗
```

- `contract/`、`ir.py`、`diagnostics.py` 只依赖标准库。
- `lower/` 不读文件、不碰 python-pptx。
- `emit/` 不碰 HTML 解析。
- `convert/` 任何时候都不 import MCP 层。

### 契约变更

契约数据是唯一事实来源。新增属性/标签/修复规则时：

1. 在 `contract/data/<version>/contract.json` 登记（带 `since` / `status`）。
2. 若引入新 `kind`，在 `values/registry.py` 注册解析器。
3. 在 `tests/contract/<version>/` 加 accept + reject 用例。
4. 重新生成契约文档。

缺用例的契约变更不允许合并。

### 文档放置

新建 doc 默认放 `docs/temp/`；只有长期定论（规范、总方案、契约参考）放 `docs/` 根。
更新既有 `docs/` 文档不受此限。

## 提交信息

采用 [Conventional Commits](https://www.conventionalcommits.org/)：

```
feat(contract): 支持 border-radius 百分比
fix(lower): display:none 时剪掉整棵子树
docs(adr): 记录 IR 作为契约
```

一个提交只做一件事；行为变化必须带测试。
