"""路径处理工具

集中所有路径校验逻辑。CLI 层入口必须经过 ``ensure_safe_input_path`` 与
``ensure_safe_output_path`` 校验，避免裸 ``Path(...)`` 透传到 COM。
"""

from __future__ import annotations

import os
import pathlib
import re

from wps_cli.consts import MAX_GLOB_RESULTS
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


def ensure_safe_input_path(
    path: str | pathlib.Path,
    *,
    allowed_extensions: frozenset[str] | set[str] | None = None,
) -> pathlib.Path:
    """校验输入文件路径：必须存在、非符号链接、非 UNC

    Args:
        path: 待校验路径
        allowed_extensions: 允许的扩展名集合（小写，含点，如 ``{".xlsx", ".csv"}``）。
            为空时不限制扩展名。**所有写操作命令应当传入对应应用类型的白名单**，
            防止 ``wps calc cell-set README.md ...`` 这类把任意文件作为工作簿打开
            并覆盖的攻击。
    """
    p = validate_path(path)
    if not p.exists():
        raise FileNotFoundErrorCli(str(path))
    if not p.is_file():
        raise ValidationError(f"路径不是文件: {path}")
    if allowed_extensions is not None:
        suffix = p.suffix.lower()
        if suffix not in allowed_extensions:
            raise ValidationError(
                f"不支持的文件扩展名 {suffix!r}，允许: {sorted(allowed_extensions)}"
            )
    return p


def ensure_safe_output_path(path: str | pathlib.Path) -> pathlib.Path:
    """校验输出文件路径：父目录可写、非符号链接、非 UNC

    用于所有写入文件的命令。父目录会自动创建。
    """
    p = validate_path(path)
    p.parent.mkdir(parents=True, exist_ok=True)
    return p


def ensure_safe_glob(
    pattern: str,
    base_dir: pathlib.Path | None = None,
    *,
    max_results: int = MAX_GLOB_RESULTS,
) -> list[pathlib.Path]:
    """安全地展开 glob 模式

    限制：
      - 拒绝绝对路径模式
      - 拒绝 UNC 路径
      - 拒绝包含 ``..`` 的模式（避免依赖 resolve 后的最终防线）
      - 结果必须位于 base_dir 之下（默认当前工作目录）
      - 结果数量不超过 ``max_results``，防止 ``**`` 触发大量 COM 操作（DoS）

    用于 batch 操作，防止攻击者扫描整盘。
    """
    if not pattern:
        raise ValidationError("glob 模式不能为空")
    if pattern.startswith("\\\\") or pattern.startswith("//"):
        raise ValidationError("不允许使用 UNC 路径")
    if pathlib.Path(pattern).is_absolute():
        raise ValidationError("不允许使用绝对路径模式（请改用相对当前目录的模式）")
    if ".." in pathlib.PurePosixPath(pattern.replace("\\", "/")).parts:
        raise ValidationError("glob 模式不允许包含 '..' 路径跳转")

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
        if len(safe) > max_results:
            raise ValidationError(f"glob 匹配结果超出上限 {max_results}，请缩小匹配范围")
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

    覆盖：用户主目录、Windows 盘符路径、UNC 路径、相对路径（./ ../）。
    """
    if not text:
        return text
    home = os.path.expanduser("~")
    if home and home != "~":
        text = text.replace(home, "~")
        text = text.replace(home.replace("\\", "/"), "~")
    # UNC 路径 \\server\share\...
    text = re.sub(r"\\\\[^\s'\"]+", "<unc-path>", text)
    # Windows 盘符路径 C:\... 或 c:/...
    text = re.sub(r"[A-Za-z]:[\\/][^\s'\"]*", "<path>", text)
    # 相对路径 ./xxx 或 ../xxx
    text = re.sub(r"(?:^|(?<=[\s'\"(]))\.{1,2}[\\/][^\s'\")]*", "<rel-path>", text)
    return text
