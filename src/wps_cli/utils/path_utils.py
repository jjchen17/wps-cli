"""路径处理工具

集中所有路径校验逻辑，CLI 层入口必须经过 ``ensure_safe_input_path`` 与
``ensure_safe_output_path`` 校验，避免裸 ``Path(...)`` 透传到 COM。
"""

from __future__ import annotations

import os
import pathlib
import re

from wps_cli.exceptions import FileNotFoundErrorCli, ValidationError


def validate_path(
    path: str | pathlib.Path,
    allowed_dirs: list[str] | None = None,
) -> pathlib.Path:
    """验证并规范化路径，防止路径遍历攻击

    Args:
        path: 待校验的路径
        allowed_dirs: 允许的根目录白名单。为空时不限制根目录，仅做符号链接和 UNC 检查。

    Raises:
        ValidationError: 路径包含符号链接、UNC 路径或越出白名单
    """
    p = pathlib.Path(path)
    text = str(path)

    if text.startswith("\\\\") or text.startswith("//"):
        raise ValidationError("不允许使用 UNC 路径（\\\\server\\share）")

    if p.is_symlink():
        raise ValidationError("不允许符号链接")

    abs_path = p.resolve()

    if allowed_dirs:
        allowed = [pathlib.Path(d).resolve() for d in allowed_dirs]
        if not any(_is_relative_to(abs_path, d) for d in allowed):
            raise ValidationError(f"路径 {path} 不在允许的目录内")

    return abs_path


def _is_relative_to(path: pathlib.Path, parent: pathlib.Path) -> bool:
    """Path.is_relative_to 在 3.10+ 可用，这里写兼容版"""
    try:
        path.relative_to(parent)
        return True
    except ValueError:
        return False


def ensure_safe_input_path(path: str | pathlib.Path) -> pathlib.Path:
    """校验输入文件路径：必须存在、非符号链接、非 UNC

    用于所有读取已有文件的命令。
    """
    p = validate_path(path)
    if not p.exists():
        raise FileNotFoundErrorCli(str(path))
    if not p.is_file():
        raise ValidationError(f"路径不是文件: {path}")
    return p


def ensure_safe_output_path(path: str | pathlib.Path) -> pathlib.Path:
    """校验输出文件路径：父目录可写、非符号链接、非 UNC

    用于所有写入文件的命令。父目录会自动创建。
    """
    p = validate_path(path)
    p.parent.mkdir(parents=True, exist_ok=True)
    return p


def ensure_safe_glob(pattern: str, base_dir: pathlib.Path | None = None) -> list[pathlib.Path]:
    """安全地展开 glob 模式

    限制：
      - 拒绝绝对路径模式
      - 拒绝 UNC 路径
      - 结果必须位于 base_dir 之下（默认当前工作目录）

    用于 batch 操作，防止攻击者扫描整盘。
    """
    if not pattern:
        raise ValidationError("glob 模式不能为空")
    if pattern.startswith("\\\\") or pattern.startswith("//"):
        raise ValidationError("不允许使用 UNC 路径")
    if pathlib.Path(pattern).is_absolute():
        raise ValidationError("不允许使用绝对路径模式（请改用相对当前目录的模式）")

    import glob

    base = (base_dir or pathlib.Path.cwd()).resolve()
    matched = glob.glob(pattern, recursive=True)
    safe: list[pathlib.Path] = []
    for m in matched:
        candidate = pathlib.Path(m).resolve()
        if candidate.is_symlink():
            continue
        if _is_relative_to(candidate, base) and candidate.is_file():
            safe.append(candidate)
    return safe


def sanitize_filename(name: str) -> str:
    """清理文件名中的非法字符"""
    cleaned = re.sub(r'[<>:"/\\|?*]', "_", name)
    cleaned = cleaned.strip(". ")
    if not cleaned:
        raise ValidationError("文件名不能为空")
    return cleaned


def ensure_parent_dir(path: pathlib.Path) -> None:
    """确保父目录存在"""
    path.parent.mkdir(parents=True, exist_ok=True)


def redact_path(text: str) -> str:
    """对错误消息中的本地路径做脱敏处理

    将用户主目录、Windows 盘符路径替换为占位符，避免 JSON 错误响应泄露文件系统结构。
    """
    if not text:
        return text
    home = os.path.expanduser("~")
    if home and home != "~":
        text = text.replace(home, "~")
        text = text.replace(home.replace("\\", "/"), "~")
    text = re.sub(r"[A-Za-z]:[\\/][^\s'\"]*", "<path>", text)
    return text
