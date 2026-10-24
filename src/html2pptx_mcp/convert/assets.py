"""图片资源解析与安全校验。

这里是唯一读用户文件的地方，因此把边界一次性收紧：

- 只接受 ``base_dir`` 内的相对路径，或 ``data:image/*;base64``；绝对路径、UNC、盘符、
  ``..``、符号链接逃逸一律拒绝（不先尝试读，避免把「能不能读到」变成探测手段）；
- 体积与像素双重上限：压缩后字节数防大文件，像素数防解压炸弹；
- 格式以 Pillow 识别结果为准，扩展名与声明的 MIME 只做辅助校验；
- 网络地址（http/https）在服务化阶段才允许，届时另做 SSRF 防护，现在直接拒绝。

解析失败返回 None 并写诊断，不抛业务异常；调用方据此跳过该图片、继续生成其余元素。
"""

from __future__ import annotations

import base64
import binascii
import hashlib
import io
import re
from dataclasses import dataclass
from pathlib import Path, PurePosixPath

from PIL import Image

from .contract import Contract
from .diagnostics import Report

_DRIVE_RE = re.compile(r"^[a-zA-Z]:")
_DATA_URI_RE = re.compile(r"^data:(?P<mime>image/[a-z0-9.+-]+);base64,(?P<body>.*)$", re.I | re.S)
_MIME_BY_FORMAT = {
    "PNG": "image/png",
    "JPEG": "image/jpeg",
    "GIF": "image/gif",
    "BMP": "image/bmp",
    "WEBP": "image/webp",
}


@dataclass(frozen=True, slots=True)
class ImageAsset:
    data: bytes
    width_px: int
    height_px: int
    mime: str
    digest: str


def resolve_image(
    src: str,
    *,
    base_dir: Path,
    contract: Contract,
    report: Report,
    line: int,
    path: str,
) -> ImageAsset | None:
    """把 ``src`` 解析成可嵌入的图片；失败写诊断并返回 None。"""
    if src.startswith("data:"):
        data = _from_data_uri(src, contract=contract, report=report, line=line, path=path)
    elif src.lower().startswith(("http://", "https://")):
        report.error(
            "ASSET_PATH_DENIED",
            line=line,
            message=f"不支持网络图片 “{src}”",
            hint="先把图片下载到 base_dir 内再用相对路径引用",
            path=path,
        )
        return None
    else:
        data = _from_file(
            src, base_dir=base_dir, contract=contract, report=report, line=line, path=path
        )
    if data is None:
        return None
    return _decode(data, contract=contract, report=report, line=line, path=path)


def _from_data_uri(
    src: str,
    *,
    contract: Contract,
    report: Report,
    line: int,
    path: str,
) -> bytes | None:
    match = _DATA_URI_RE.match(src.strip())
    if not match:
        report.error(
            "ASSET_UNSUPPORTED_TYPE",
            line=line,
            message="data URI 必须是 image/*;base64",
            hint="示例：data:image/png;base64,iVBORw0KGgo...",
            path=path,
        )
        return None
    try:
        data = base64.b64decode(match.group("body"), validate=True)
    except (binascii.Error, ValueError):
        report.error(
            "ASSET_DECODE_FAILED",
            line=line,
            message="data URI 的 base64 无法解码",
            hint="确认 base64 内容完整，或改用相对路径引用图片",
            path=path,
        )
        return None
    if len(data) > contract.limits.image_bytes:
        report.error(
            "ASSET_TOO_LARGE",
            line=line,
            message=f"图片 {len(data)} 字节超过上限 {contract.limits.image_bytes}",
            hint="压缩图片，或改用更小的图片",
            path=path,
        )
        return None
    return data


def _from_file(
    src: str,
    *,
    base_dir: Path,
    contract: Contract,
    report: Report,
    line: int,
    path: str,
) -> bytes | None:
    normalized = src.strip().replace("\\", "/")
    if not normalized or "\x00" in normalized:
        report.error(
            "ASSET_SRC_MISSING",
            line=line,
            message="图片 src 为空",
            hint='示例：src="assets/photo.png"',
            path=path,
        )
        return None
    if (
        normalized.startswith("/")
        or normalized.startswith("//")
        or _DRIVE_RE.match(normalized)
        or ".." in PurePosixPath(normalized).parts
    ):
        report.error(
            "ASSET_PATH_DENIED",
            line=line,
            message=f"图片路径越界：“{src}”",
            hint="只允许 base_dir 内的相对路径，例如 assets/photo.png",
            path=path,
        )
        return None

    root = base_dir.resolve()
    try:
        resolved = (root / normalized).resolve(strict=False)
    except OSError:
        report.error(
            "ASSET_NOT_FOUND",
            line=line,
            message=f"无法解析图片路径：“{src}”",
            hint="检查路径是否相对 base_dir 正确",
            path=path,
        )
        return None
    if not resolved.is_relative_to(root):
        report.error(
            "ASSET_PATH_DENIED",
            line=line,
            message=f"图片路径逃出 base_dir：“{src}”",
            hint="只允许 base_dir 内的相对路径，例如 assets/photo.png",
            path=path,
        )
        return None
    try:
        size = resolved.stat().st_size
    except OSError:
        report.error(
            "ASSET_NOT_FOUND",
            line=line,
            message=f"图片不存在：“{src}”",
            hint="确认文件已放在 base_dir 内，且大小写与路径一致",
            path=path,
        )
        return None
    if size > contract.limits.image_bytes:
        report.error(
            "ASSET_TOO_LARGE",
            line=line,
            message=f"图片 {size} 字节超过上限 {contract.limits.image_bytes}",
            hint="压缩图片，或改用更小的图片",
            path=path,
        )
        return None
    try:
        return resolved.read_bytes()
    except OSError:
        report.error(
            "ASSET_NOT_FOUND",
            line=line,
            message=f"图片不可读：“{src}”",
            hint="检查文件权限，或改用 base_dir 内的另一张图",
            path=path,
        )
        return None


def _decode(
    data: bytes,
    *,
    contract: Contract,
    report: Report,
    line: int,
    path: str,
) -> ImageAsset | None:
    try:
        with Image.open(io.BytesIO(data)) as image:
            pixel_count = image.width * image.height
            if pixel_count > contract.limits.image_pixels:
                report.error(
                    "LIMIT_IMAGE_EXCEEDED",
                    line=line,
                    message=f"图片 {image.width}×{image.height} 像素超过上限 "
                    f"{contract.limits.image_pixels}",
                    hint="缩小图片分辨率后再引用",
                    path=path,
                )
                return None
            image_format = image.format or ""
            width, height = image.width, image.height
            image.verify()
    except Image.DecompressionBombError:
        report.error(
            "LIMIT_IMAGE_EXCEEDED",
            line=line,
            message="图片像素数触发放大炸弹保护",
            hint="缩小图片分辨率后再引用",
            path=path,
        )
        return None
    except (OSError, ValueError):
        report.error(
            "ASSET_DECODE_FAILED",
            line=line,
            message="图片无法解码",
            hint="确认文件确实是 PNG/JPEG/GIF/BMP/WebP，且未损坏",
            path=path,
        )
        return None

    mime = _MIME_BY_FORMAT.get(image_format.upper())
    if mime is None:
        report.error(
            "ASSET_UNSUPPORTED_TYPE",
            line=line,
            message=f"不支持的图片格式：{image_format or '未知'}",
            hint="支持 PNG / JPEG / GIF / BMP / WebP",
            path=path,
        )
        return None
    return ImageAsset(
        data=data,
        width_px=width,
        height_px=height,
        mime=mime,
        digest=hashlib.sha256(data).hexdigest(),
    )
