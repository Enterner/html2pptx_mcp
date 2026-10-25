"""发布前检查：wheel 必须带齐契约数据与类型标记。

契约 JSON 是运行时资源，被 .gitignore 或打包配置漏掉时，源码测试全绿但安装后启动即崩。
"""

from __future__ import annotations

import glob
import sys
import zipfile

# CI 的 Windows runner 默认 stdout 是 cp1252，中文 print 会直接 UnicodeEncodeError；
# 强制 UTF-8 让脚本在任何代码页下都能输出（Python 3.7+ 支持 reconfigure）。
if hasattr(sys.stdout, "reconfigure"):
    sys.stdout.reconfigure(encoding="utf-8", errors="replace")

REQUIRED = (
    "html2pptx_mcp/py.typed",
    "html2pptx_mcp/server.py",
    "html2pptx_mcp/__main__.py",
    "html2pptx_mcp/convert/data/contract.json",
    "html2pptx_mcp/convert/data/colors.json",
    "html2pptx_mcp/convert/data/fonts.json",
)


def main() -> int:
    wheels = glob.glob("dist/*.whl")
    if not wheels:
        print("dist/ 下没有 wheel，先运行 python -m build", file=sys.stderr)
        return 1
    archive = zipfile.ZipFile(wheels[0])
    names = archive.namelist()
    missing = [name for name in REQUIRED if name not in names]
    if missing:
        print(f"{wheels[0]} 缺少必要文件：{', '.join(missing)}", file=sys.stderr)
        return 1
    entry_points = [name for name in names if name.endswith(".dist-info/entry_points.txt")]
    if not entry_points:
        print(f"{wheels[0]} 缺少 entry_points.txt（CLI 入口未注册）", file=sys.stderr)
        return 1
    text = archive.read(entry_points[0]).decode("utf-8")
    if "html2pptx-cli" not in text:
        print(f"{wheels[0]} 的 entry_points.txt 没有 html2pptx-cli", file=sys.stderr)
        return 1
    if "html2pptx-mcp" not in text:
        print(f"{wheels[0]} 的 entry_points.txt 没有 html2pptx-mcp", file=sys.stderr)
        return 1
    print(f"{wheels[0]} 打包完整性检查通过")
    return 0


if __name__ == "__main__":
    sys.exit(main())
