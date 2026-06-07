# 设计参考: iOfficeAI/OfficeCLI (Apache 2.0, https://github.com/iOfficeAI/OfficeCLI)
"""MCP 服务器 CLI 命令"""

from __future__ import annotations

import json
import os
from pathlib import Path
from typing import Any

import typer

from wps_cli.cli.common import handle_error, success

app = typer.Typer(help="MCP 服务器管理")


@app.command()
def serve():
    """启动 MCP stdio 服务器（供 AI Agent 调用）

    此命令启动一个 JSON-RPC 2.0 over stdio 的 MCP 服务器，
    Claude Code、Cursor 等 AI 工具可以通过此服务器调用 wps-cli 的所有文档操作能力。

    使用方式::

        wps mcp serve
    """
    try:
        from wps_cli.mcp.server import WpsMcpServer

        server = WpsMcpServer()
        server.run()
    except KeyboardInterrupt:
        pass
    except Exception as e:
        typer.echo(f"错误: {e}", err=True)
        raise typer.Exit(1) from e


_MCP_CONFIG_TEMPLATES: dict[str, dict] = {
    "claude": {
        "description": "Claude Code MCP 配置",
        "config_path": "~/.claude/mcp.json",
        "entry": {
            "wps-cli": {
                "command": "wps",
                "args": ["mcp", "serve"],
            },
        },
    },
    "cursor": {
        "description": "Cursor IDE MCP 配置",
        "config_path": ".cursor/mcp.json",
        "entry": {
            "wps-cli": {
                "command": "wps",
                "args": ["mcp", "serve"],
            },
        },
    },
    "vscode": {
        "description": "VS Code / Cline MCP 配置",
        "config_path": "~/.vscode/mcp.json",
        "entry": {
            "wps-cli": {
                "command": "wps",
                "args": ["mcp", "serve"],
            },
        },
    },
}


def _resolve_path(path_str: str) -> Path:
    """将 ~ 展开为用户主目录"""
    return Path(os.path.expanduser(path_str))


def _read_json_config(config_path: Path) -> dict:
    """读取 JSON 配置文件，不存在时返回空字典"""
    if config_path.exists():
        try:
            return json.loads(config_path.read_text(encoding="utf-8"))
        except json.JSONDecodeError:
            typer.echo(f"警告: {config_path} 并非合法 JSON，将创建新文件", err=True)
            return {}
    return {}


def _merge_mcp_config(existing: dict, entry: dict) -> dict:
    """合并 MCP 配置，保留已有的其他条目"""
    result = dict(existing)
    result.update(entry)
    return result


@app.command()
def install(
    target: str = typer.Option("claude", "--target", "-t", help="目标 AI 工具：claude/cursor/vscode"),
):
    """安装 MCP 配置到 AI 工具

    将 wps-cli 注册为指定 AI 工具的 MCP 服务器。

    示例::

        wps mcp install              # 默认安装到 Claude Code
        wps mcp install -t cursor    # 安装到 Cursor
        wps mcp install -t vscode    # 安装到 VS Code
    """
    cmd = "mcp.install"
    try:
        target = target.lower()
        if target not in _MCP_CONFIG_TEMPLATES:
            valid = ", ".join(_MCP_CONFIG_TEMPLATES.keys())
            typer.echo(f"不支持的目标: {target}，可选: {valid}", err=True)
            raise typer.Exit(1)

        template = _MCP_CONFIG_TEMPLATES[target]
        config_path = _resolve_path(template["config_path"])
        config_path.parent.mkdir(parents=True, exist_ok=True)

        existing = _read_json_config(config_path)
        merged = _merge_mcp_config(existing, template["entry"])
        config_path.write_text(json.dumps(merged, ensure_ascii=False, indent=2) + "\n", encoding="utf-8")

        success(
            {"target": target, "config_path": str(config_path), "description": template["description"]},
            command=cmd,
            json_mode=True,
        )
        typer.echo(f"已写入: {config_path}")
    except Exception as e:
        handle_error(e, command=cmd, json_mode=True)


@app.command()
def status():
    """检查 MCP 注册状态

    检测各 AI 工具的 MCP 配置文件是否已注册 wps-cli。
    """
    cmd = "mcp.status"
    results: list[dict[str, Any]] = []
    for target, template in _MCP_CONFIG_TEMPLATES.items():
        config_path = _resolve_path(template["config_path"])
        registered = False
        config_file_exists = config_path.exists()
        if config_file_exists:
            try:
                config = json.loads(config_path.read_text(encoding="utf-8"))
                registered = "wps-cli" in config
            except json.JSONDecodeError:
                registered = False

        results.append({
            "target": target,
            "config_path": str(config_path),
            "config_exists": config_file_exists,
            "registered": registered,
        })

    success(
        {"statuses": results},
        command=cmd,
        json_mode=True,
    )

    # 终端友好的表格输出
    if not any(r["registered"] for r in results):
        typer.echo("\n提示: 运行 'wps mcp install' 安装到 Claude Code", err=True)
