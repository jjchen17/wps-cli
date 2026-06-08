# 设计参考: iOfficeAI/OfficeCLI (Apache 2.0, https://github.com/iOfficeAI/OfficeCLI)
"""AI 工具集成安装命令

将 SKILL.md 和 MCP 配置安装到 AI 工具的配置目录，
让 AI 工具能够识别并调用 wps-cli 的能力。
"""

from __future__ import annotations

import json
import os
import shutil
from pathlib import Path
from typing import Any

import typer

from wps_cli.cli.common import handle_error, success

app = typer.Typer(help="AI 工具集成安装")


# ── SKILL.md 安装目标 ──

_SKILL_TARGETS: dict[str, dict[str, str]] = {
    "claude": {
        "description": "Claude Code",
        "path": "~/.claude/skills/wps-cli.md",
    },
    "cursor": {
        "description": "Cursor IDE",
        "path": ".cursor/skills/wps-cli.md",
    },
    "vscode": {
        "description": "VS Code / Cline",
        "path": "~/.vscode/skills/wps-cli.md",
    },
    "windsurf": {
        "description": "Windsurf",
        "path": ".windsurf/skills/wps-cli.md",
    },
    "codex": {
        "description": "Codex CLI",
        "path": ".agents/skills/wps-cli.md",
    },
    "hermes": {
        "description": "Hermes Agent",
        "path": ".hermes/skills/wps-cli.md",
    },
    "minimax": {
        "description": "MiniMax CLI",
        "path": ".minimax/skills/wps-cli.md",
    },
    "opencode": {
        "description": "OpenCode",
        "path": ".opencode/skills/wps-cli.md",
    },
    "nanobot": {
        "description": "NanoBot",
        "path": ".nanobot/skills/wps-cli.md",
    },
    "zeroclaw": {
        "description": "ZeroClaw",
        "path": ".zeroclaw/skills/wps-cli.md",
    },
    "openclaw": {
        "description": "OpenClaw",
        "path": ".openclaw/skills/wps-cli.md",
    },
}

# ── MCP 配置模板（与 mcp_cmd 中一致） ──

_MCP_ENTRY = {
    "wps-cli": {
        "command": "wps",
        "args": ["mcp", "serve"],
    },
}

_MCP_TARGETS: dict[str, dict[str, Any]] = {
    "claude": {
        "description": "Claude Code",
        "config_path": "~/.claude/mcp.json",
        "entry": _MCP_ENTRY,
    },
    "cursor": {
        "description": "Cursor IDE",
        "config_path": ".cursor/mcp.json",
        "entry": _MCP_ENTRY,
    },
    "vscode": {
        "description": "VS Code / Cline",
        "config_path": "~/.vscode/mcp.json",
        "entry": _MCP_ENTRY,
    },
    "windsurf": {
        "description": "Windsurf",
        "config_path": ".windsurf/mcp.json",
        "entry": _MCP_ENTRY,
    },
    "codex": {
        "description": "Codex CLI",
        "config_path": ".agents/mcp.json",
        "entry": _MCP_ENTRY,
    },
    "hermes": {
        "description": "Hermes Agent",
        "config_path": ".hermes/mcp.json",
        "entry": _MCP_ENTRY,
    },
    "minimax": {
        "description": "MiniMax CLI",
        "config_path": ".minimax/mcp.json",
        "entry": _MCP_ENTRY,
    },
    "opencode": {
        "description": "OpenCode",
        "config_path": ".opencode/mcp.json",
        "entry": _MCP_ENTRY,
    },
    "nanobot": {
        "description": "NanoBot",
        "config_path": ".nanobot/mcp.json",
        "entry": _MCP_ENTRY,
    },
    "zeroclaw": {
        "description": "ZeroClaw",
        "config_path": ".zeroclaw/mcp.json",
        "entry": _MCP_ENTRY,
    },
    "openclaw": {
        "description": "OpenClaw",
        "config_path": ".openclaw/mcp.json",
        "entry": _MCP_ENTRY,
    },
}


def _resolve_path(path_str: str) -> Path:
    return Path(os.path.expanduser(path_str))


def _find_skill_dir() -> Path | None:
    """查找项目的 skills/wps-cli/ 目录（标准 skill 包）"""
    try:
        import wps_cli

        pkg_dir = Path(wps_cli.__file__).parent.parent  # src/
    except Exception:
        pkg_dir = Path.cwd() / "src"

    candidates = [
        pkg_dir / ".." / "skills" / "wps-cli",        # 项目根目录/skills/wps-cli
        pkg_dir / ".." / ".." / "skills" / "wps-cli", # 进一步向上
        Path.cwd() / "skills" / "wps-cli",             # 当前工作目录
    ]
    for c in candidates:
        resolved = c.resolve()
        if resolved.is_dir() and (resolved / "SKILL.md").exists():
            return resolved
    return None


def _find_skill_md() -> Path | None:
    """查找项目的 SKILL.md 文件（单文件回退）"""
    # 优先从安装位置查找
    try:
        import wps_cli

        pkg_dir = Path(wps_cli.__file__).parent.parent  # src/
    except Exception:
        pkg_dir = Path.cwd() / "src"

    candidates = [
        pkg_dir / ".." / "SKILL.md",              # 项目根目录
        pkg_dir / ".." / ".." / "SKILL.md",       # 进一步向上
        Path.cwd() / "SKILL.md",                   # 当前工作目录
    ]
    for c in candidates:
        resolved = c.resolve()
        if resolved.exists():
            return resolved
    return None


def _copy_skill_dir(src_dir: Path, dest_parent: Path) -> Path:
    """递归复制 skill 目录到目标位置"""
    dest_dir = dest_parent  # 直接使用目标目录路径
    if dest_dir.exists():
        shutil.rmtree(dest_dir)
    shutil.copytree(src_dir, dest_dir)
    return dest_dir / "SKILL.md"


def _get_skill_target_dir(target: str) -> Path:
    """获取指定 AI 工具的 skill 安装目录路径"""
    # skill 目录安装路径（不含 .md 后缀）
    if target == "claude":
        return _resolve_path("~/.claude/skills/wps-cli")
    elif target == "cursor":
        return _resolve_path(".cursor/skills/wps-cli")
    elif target == "vscode":
        return _resolve_path("~/.vscode/skills/wps-cli")
    else:
        # 其他工具使用统一模式
        tpl = _SKILL_TARGETS.get(target, {})
        path_str = tpl.get("path", f"~/.{target}/skills/wps-cli.md")
        # 如果路径以 .md 结尾，改为目录
        if path_str.endswith(".md"):
            path_str = path_str[:-3]
        return _resolve_path(path_str)


@app.command()
def skill(
    target: str = typer.Option("all", "--target", "-t", help="目标 AI 工具: claude/cursor/vscode/windsurf/codex/hermes/minimax/opencode/nanobot/zeroclaw/openclaw/all"),
):
    """安装 SKILL.md 到 AI 工具配置目录

    优先使用标准 skill 包（skills/wps-cli/ 目录，含模块化参考文档），
    如不可用则回退到单文件安装。

    示例::

        wps install skill                 # 安装到所有支持的 AI 工具
        wps install skill -t claude       # 仅安装到 Claude Code
        wps install skill -t cursor       # 仅安装到 Cursor
    """
    cmd = "install.skill"
    try:
        if target == "all":
            target_list = list(_SKILL_TARGETS.keys())
        elif target not in _SKILL_TARGETS:
            valid = ", ".join(_SKILL_TARGETS.keys())
            typer.echo(f"不支持的目标: {target}，可选: all/{valid}", err=True)
            raise typer.Exit(1)
        else:
            target_list = [target]

        # 优先查找标准 skill 目录
        skill_dir = _find_skill_dir()
        skill_file = _find_skill_md() if skill_dir is None else None

        if skill_dir is None and skill_file is None:
            typer.echo("错误: 找不到 SKILL.md 或 skills/wps-cli/ 目录。请确保在 wps-cli 项目目录下运行。", err=True)
            raise typer.Exit(1)

        installed = []
        for t in target_list:
            tpl = _SKILL_TARGETS[t]

            if skill_dir is not None:
                # 标准 skill 包：目录级安装
                dest_dir = _get_skill_target_dir(t)
                dest_dir.parent.mkdir(parents=True, exist_ok=True)
                _copy_skill_dir(skill_dir, dest_dir)
                installed.append({
                    "target": t,
                    "description": tpl["description"],
                    "path": str(dest_dir),
                    "type": "directory",
                })
            else:
                # 单文件回退
                dest = _resolve_path(tpl["path"])
                dest.parent.mkdir(parents=True, exist_ok=True)
                shutil.copy2(skill_file, dest)  # type: ignore[arg-type]
                installed.append({
                    "target": t,
                    "description": tpl["description"],
                    "path": str(dest),
                    "type": "file",
                })

        success(
            {
                "source": str(skill_dir if skill_dir else skill_file),
                "installed": installed,
            },
            command=cmd,
            json_mode=True,
        )
        typer.echo(f"\n安装模式: {'skill 包目录（含模块化参考文档）' if skill_dir else '单文件（回退模式）'}\n")
        for item in installed:
            typer.echo(f"  {item['target']}: {item['path']} ({item['type']})")
    except Exception as e:
        handle_error(e, command=cmd, json_mode=True)


@app.command()
def mcp(
    target: str = typer.Option("claude", "--target", "-t", help="目标 AI 工具：claude/cursor/vscode/windsurf/codex/hermes/minimax/opencode/nanobot/zeroclaw/openclaw"),
):
    """安装 MCP 配置到 AI 工具

    将 wps-cli 注册为指定 AI 工具的 MCP 服务器，
    使 AI 工具能够通过 MCP 协议调用 wps-cli。

    示例::

        wps install mcp                  # 安装到 Claude Code
        wps install mcp -t cursor        # 安装到 Cursor
    """
    cmd = "install.mcp"
    try:
        target = target.lower()
        if target not in _MCP_TARGETS:
            valid = ", ".join(_MCP_TARGETS.keys())
            typer.echo(f"不支持的目标: {target}，可选: {valid}", err=True)
            raise typer.Exit(1)

        tpl = _MCP_TARGETS[target]
        config_path = _resolve_path(tpl["config_path"])
        config_path.parent.mkdir(parents=True, exist_ok=True)

        existing = {}
        if config_path.exists():
            try:
                existing = json.loads(config_path.read_text(encoding="utf-8"))
            except json.JSONDecodeError:
                typer.echo(f"警告: {config_path} 并非合法 JSON，将创建新文件", err=True)

        existing.update(tpl["entry"])
        config_path.write_text(json.dumps(existing, ensure_ascii=False, indent=2) + "\n", encoding="utf-8")

        success(
            {"target": target, "config_path": str(config_path), "description": tpl["description"]},
            command=cmd,
            json_mode=True,
        )
        typer.echo(f"已写入: {config_path}")
    except Exception as e:
        handle_error(e, command=cmd, json_mode=True)


@app.command()
def all_tools(
    target: str = typer.Option("all", "--target", "-t", help="目标 AI 工具: claude/cursor/vscode/windsurf/codex/hermes/minimax/opencode/nanobot/zeroclaw/openclaw/all"),
):
    """一键安装 SKILL.md + MCP 配置

    同时安装技能文件和 MCP 服务器配置到指定 AI 工具。

    示例::

        wps install all-tools            # 全部安装到所有 AI 工具
        wps install all-tools -t claude  # 安装到 Claude Code
    """
    cmd = "install.all_tools"
    try:
        typer.echo("=== 安装 SKILL.md ===")
        skill(target=target)

        typer.echo("\n=== 安装 MCP 配置 ===")
        if target == "all":
            for t in _MCP_TARGETS:
                typer.echo(f"\n--- {t} ---")
                mcp(target=t)
        else:
            mcp(target=target)

        success({"status": "ok"}, command=cmd, json_mode=True)
        typer.echo("\n安装完成！重启 AI 工具后即可使用 wps-cli 能力。")
    except Exception as e:
        handle_error(e, command=cmd, json_mode=True)
