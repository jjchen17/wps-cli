"""路径处理工具"""

import pathlib
import re


def validate_path(path: str, allowed_dirs: list[str] | None = None) -> pathlib.Path:
    """验证并规范化路径，防止路径遍历攻击"""
    p = pathlib.Path(path)

    # 先检查符号链接（resolve 前）
    if p.is_symlink():
        raise ValueError("不允许符号链接")

    abs_path = p.resolve()

    if allowed_dirs:
        allowed = [pathlib.Path(d).resolve() for d in allowed_dirs]
        if not any(abs_path.is_relative_to(d) for d in allowed):
            raise ValueError(f"路径 {path} 不在允许的目录内")

    return abs_path


def sanitize_filename(name: str) -> str:
    """清理文件名中的非法字符"""
    cleaned = re.sub(r'[<>:"/\\|?*]', "_", name)
    cleaned = cleaned.strip(". ")
    if not cleaned:
        raise ValueError("文件名不能为空")
    return cleaned


def ensure_parent_dir(path: pathlib.Path) -> None:
    """确保父目录存在"""
    path.parent.mkdir(parents=True, exist_ok=True)
