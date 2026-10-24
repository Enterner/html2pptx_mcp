"""系统字体定位与度量：家族名（含中文别名）→ 字体文件 → 像素级度量。

textfit 的换行估算与字体内嵌共用这一层。设计约束：

- 扫描懒加载：注册表/目录索引先建（不打开字体文件），文件解析与 name 表
  别名补全推迟到该文件首次被引用；解析结果进程内缓存。
- 找不到家族 → 返回 None，调用方回退启发式（textfit 既有口径）——
  度量缺失绝不阻断生成。
- 别名来自字体 name 表（nameID 1/16，Windows 平台的 en-US 与 zh-CN 两条），
  因此「等线」「DengXian」「微软雅黑」「Microsoft YaHei」都能命中同一文件。
"""

from __future__ import annotations

import contextlib
import io
import re
from dataclasses import dataclass, field
from pathlib import Path
from typing import Any

from fontTools.ttLib import TTFont

with contextlib.suppress(ImportError):
    import winreg

_REGISTRY_PATH = r"SOFTWARE\Microsoft\Windows NT\CurrentVersion\Fonts"
_FONT_DIRS = (
    Path("C:/Windows/Fonts"),
    Path.home() / "AppData/Local/Microsoft/Windows/Fonts",
    Path("/usr/share/fonts"),
    Path("/usr/local/share/fonts"),
    Path.home() / ".fonts",
)
_STYLE_RE = re.compile(
    r"\s*\((TrueType|OpenType|TrueTypeOutline|Type1|Raster Font|Vector Font)\)\s*$"
)
_CJK_RE = re.compile(r"[\u2e80-\u9fff\uf900-\ufaff\uff00-\uffef]")


def _heuristic_em(char: str) -> float:
    return 1.0 if _CJK_RE.match(char) else 0.55


@dataclass
class Face:
    """一个已定位的字体文件面：度量惰性求值，字形宽度进程内缓存。"""

    path: Path
    index: int = 0
    bold: bool = False
    _display_name: str = field(default="", repr=False, compare=False)
    _font: Any = field(default=None, repr=False, compare=False)
    _upm: int = field(default=0, repr=False, compare=False)
    _asc: int = field(default=0, repr=False, compare=False)
    _desc: int = field(default=0, repr=False, compare=False)
    _gap: int = field(default=0, repr=False, compare=False)
    _advances: dict[int, float] = field(default_factory=dict, repr=False, compare=False)
    _loaded: bool = field(default=False, repr=False, compare=False)

    def _load(self) -> bool:
        """打开字体并取度量；损坏/不可读只影响本 face，不抛出。"""
        if self._loaded:
            return self._upm > 0
        self._loaded = True
        try:
            font = TTFont(str(self.path), fontNumber=self.index, lazy=True)
            self._font = font
            self._upm = font["head"].unitsPerEm
            hhea = font["hhea"]
            self._asc, self._desc, self._gap = hhea.ascent, hhea.descent, hhea.lineGap
        except Exception:
            return False
        return self._upm > 0

    def registry_aliases(self) -> set[str]:
        """注册表展示名派生的家族别名（不打开字体文件）。"""
        name = _STYLE_RE.sub("", self._display_name) if self._display_name else ""
        families = [f.strip() for f in name.split("&")] if "&" in name else [name]
        return {f.lower() for f in families if f}

    def name_aliases(self) -> set[str]:
        """name 表里的家族名（nameID 1/16 × en-US/zh-CN），小写。"""
        aliases: set[str] = set()
        if not self._load():
            return aliases
        with contextlib.suppress(Exception):
            for record in self._font["name"].names:
                if record.nameID not in (1, 16):
                    continue
                if (record.platformID, record.platEncID) not in ((3, 1), (3, 10)):
                    continue
                value = record.toUnicode().strip()
                if value:
                    aliases.add(value.lower())
        return aliases

    def _advance_units(self, char: str) -> float | None:
        """字符前进宽（字体单位）；cmap 缺字返回 None，由调用方回退启发式。"""
        if not self._load():
            return None
        code = ord(char)
        if code in self._advances:
            cached = self._advances[code]
            return None if cached < 0 else cached
        advance: float | None = None
        with contextlib.suppress(Exception):
            glyph = self._font.getBestCmap().get(code)
            if glyph is not None:
                advance = self._font["hmtx"][glyph][0]
        self._advances[code] = -1.0 if advance is None else advance
        return advance

    def width_px(self, text: str, size_px: float) -> float:
        if not self._load():
            return sum(_heuristic_em(ch) * size_px for ch in text)
        total = 0.0
        for ch in text:
            advance = self._advance_units(ch)
            em = advance / self._upm if advance is not None else _heuristic_em(ch)
            total += em * size_px
        return total

    def line_height_px(self, size_px: float) -> float | None:
        """PPT 单倍行距口径：hhea ascender+|descender|+lineGap。"""
        if not self._load():
            return None
        return (self._asc - self._desc + self._gap) / self._upm * size_px

    def fs_type(self) -> int:
        """OS/2 fsType 嵌入许可位；无 OS/2 表视为可嵌入（0）。"""
        if not self._load():
            return 0
        with contextlib.suppress(Exception, KeyError):
            return int(self._font["OS/2"].fsType)
        return 0

    def subset_bytes(self, text: str) -> bytes | None:
        """按出现过的码点做子集，返回独立 TTF 字节；失败返回 None。"""
        if not self._load():
            return None
        from fontTools import subset

        try:
            options = subset.Options()
            options.layout_features = ["*"]
            options.name_IDs = ["*"]
            options.notdef_outline = True
            subsetter = subset.Subsetter(options)
            subsetter.populate(text=text)
            subsetter.subset(self._font)
            buffer = io.BytesIO()
            self._font.save(buffer)
            return buffer.getvalue()
        except Exception:
            return None


def _parse_display_name(display: str) -> tuple[list[str], bool]:
    """注册表展示名 → (家族别名列表, bold)。

    组内每个子家族各带风格后缀（「A Bold & B Bold」），逐段解析——只看串尾
    会把「YaHei Bold & UI Bold」错拆成 UI 的 regular。纯函数，便于测试。
    """
    display = _STYLE_RE.sub("", display)
    parts = [p.strip() for p in display.split("&")] if "&" in display else [display]
    families: list[str] = []
    bold = False
    for part in parts:
        part_bold = False
        for suffix in (" Bold Italic", " Bold", " Italic"):
            if part.endswith(suffix):
                part = part[: -len(suffix)]
                part_bold = suffix in (" Bold", " Bold Italic")
                break
        bold = bold or part_bold
        if part:
            families.append(part)
    return families, bold


class FontResolver:
    """家族别名 → Face 的进程内索引；Windows 注册表 + 字体目录双来源。"""

    def __init__(self) -> None:
        # 别名(小写) → {bold: Face}；同族 regular/bold 分别登记
        self._by_alias: dict[str, dict[bool, Face]] = {}
        self._registry_done = False
        self._aliases_done = False

    def resolve(self, family: str, bold: bool = False) -> Face | None:
        key = (family or "").lower().strip()
        if not key:
            return None
        self._ensure_registry()
        faces = self._by_alias.get(key)
        if faces is None and not self._aliases_done:
            # 中文名（等线/微软雅黑）在注册表里可能只登记了英文名，首次 miss
            # 时打开文件读 name 表补别名，之后全部命中缓存
            self._ensure_aliases()
            faces = self._by_alias.get(key)
        if faces is None:
            return None
        return faces.get(bold) or faces.get(False)

    def _register(self, aliases: set[str], face: Face) -> None:
        for alias in aliases:
            bucket = self._by_alias.setdefault(alias, {})
            # 同名冲突保留首个：注册表枚举顺序即 PowerPoint 的替换优先级
            bucket.setdefault(face.bold, face)

    def _ensure_registry(self) -> None:
        if self._registry_done:
            return
        self._registry_done = True
        if "winreg" in globals():
            self._from_registry()
        if not self._by_alias:
            self._from_dirs()

    def _from_registry(self) -> None:
        for hive in (winreg.HKEY_LOCAL_MACHINE, winreg.HKEY_CURRENT_USER):
            try:
                key = winreg.OpenKey(hive, _REGISTRY_PATH)
            except OSError:
                continue
            with key:
                i = 0
                while True:
                    try:
                        display, file_name, _ = winreg.EnumValue(key, i)
                    except OSError:
                        break
                    i += 1
                    # 只登记注册表别名（不开文件）；name 表别名延后到首次 miss
                    face = self._face_from_registry(display, file_name)
                    if face is not None:
                        self._register(face.registry_aliases(), face)

    def _face_from_registry(self, display: str, file_name: str) -> Face | None:
        families, bold = _parse_display_name(display)
        path = Path(file_name)
        if not path.is_absolute():
            path = _FONT_DIRS[0] / path
        if not path.exists():
            for directory in _FONT_DIRS[1:]:
                candidate = directory / file_name
                if candidate.exists():
                    path = candidate
                    break
        if not path.exists():
            return None
        face = Face(path=path, bold=bold, _display_name=display)
        self._register({f.lower() for f in families if f}, face)
        return face

    def _from_dirs(self) -> None:
        """非 Windows 退目录扫描；TTC 只取第 0 面（度量场景足够）。"""
        seen: set[Path] = set()
        for directory in _FONT_DIRS:
            if not directory.is_dir():
                continue
            for pattern in ("*.ttf", "*.otf", "*.ttc"):
                for path in sorted(directory.rglob(pattern)):
                    if path in seen:
                        continue
                    seen.add(path)
                    face = Face(path=path, index=0, bold=False)
                    self._register(face.name_aliases() or {path.stem.lower()}, face)

    def _ensure_aliases(self) -> None:
        if self._aliases_done:
            return
        self._aliases_done = True
        for faces in list(self._by_alias.values()):
            for face in list(faces.values()):
                self._register(face.name_aliases(), face)


class FontFileMetrics:
    """textfit.Metrics 的真字体实现：逐字符 hmtx 前进宽 + hhea 行高。

    家族解析失败或缺字时按字符回退启发式（CJK≈1em / 拉丁≈0.55em），
    度量精度只升不降；行高回退 None → textfit 用 1.2 倍默认。
    """

    def __init__(self, resolver: FontResolver | None = None) -> None:
        self._resolver = resolver or FontResolver()

    def width_px(
        self, text: str, size_px: float, *, family: str | None = None, bold: bool = False
    ) -> float:
        if not text:
            return 0.0
        face = self._resolver.resolve(family or "", bold)
        if face is None:
            return sum(_heuristic_em(ch) * size_px for ch in text)
        return face.width_px(text, size_px)

    def line_height_px(
        self, size_px: float, *, family: str | None = None, bold: bool = False
    ) -> float | None:
        face = self._resolver.resolve(family or "", bold)
        if face is None:
            return None
        return face.line_height_px(size_px)


_system_metrics: FontFileMetrics | None = None
_system_resolved = False


def system_metrics() -> FontFileMetrics | None:
    """进程级单例：系统字体度量；扫描失败返回 None（调用方用启发式）。"""
    global _system_metrics, _system_resolved
    if not _system_resolved:
        _system_resolved = True
        with contextlib.suppress(Exception):
            _system_metrics = FontFileMetrics(FontResolver())
    return _system_metrics
