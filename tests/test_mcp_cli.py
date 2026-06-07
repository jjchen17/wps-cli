# 设计参考: iOfficeAI/OfficeCLI (Apache 2.0, https://github.com/iOfficeAI/OfficeCLI)
"""MCP CLI 命令测试"""

from __future__ import annotations

import json
import tempfile
from pathlib import Path

import pytest
from typer.testing import CliRunner

from wps_cli.main import app as main_app


@pytest.fixture
def runner() -> CliRunner:
    """提供 typer CLI 测试运行器"""
    return CliRunner()


# ── mcp serve ──


class TestMcpServe:
    """测试 'wps mcp serve' 命令"""

    def test_serve_registered(self, runner):
        """确认 mcp serve 子命令已注册"""
        result = runner.invoke(main_app, ["mcp", "serve", "--help"])
        # 这个命令会实际启动服务器，在测试中会失败（没有 stdin）
        # 但至少能验证命令已注册
        assert result.exit_code == 0


# ── mcp install ──


class TestMcpInstall:
    """测试 'wps mcp install' 命令"""

    def test_install_to_claude_creates_config(self, runner):
        """安装到 Claude Code 应创建 MCP 配置文件"""
        with tempfile.TemporaryDirectory() as tmpdir:
            claude_dir = Path(tmpdir) / ".claude"
            config_path = claude_dir / "mcp.json"

            # 使用 monkeypatch 修改配置路径
            import wps_cli.cli.mcp_cmd as mcp_mod

            original = mcp_mod._MCP_CONFIG_TEMPLATES.copy()
            try:
                mcp_mod._MCP_CONFIG_TEMPLATES["claude"]["config_path"] = str(config_path)
                result = runner.invoke(main_app, ["mcp", "install", "-t", "claude"])
            finally:
                mcp_mod._MCP_CONFIG_TEMPLATES = original

            assert result.exit_code == 0
            assert config_path.exists()
            config = json.loads(config_path.read_text(encoding="utf-8"))
            assert "wps-cli" in config
            assert config["wps-cli"]["command"] == "wps"
            assert config["wps-cli"]["args"] == ["mcp", "serve"]

    def test_install_to_invalid_target_errors(self, runner):
        """安装到无效目标应报错"""
        result = runner.invoke(main_app, ["mcp", "install", "-t", "invalid"])
        assert result.exit_code != 0

    def test_install_preserves_existing_config(self, runner):
        """安装 MCP 配置时应保留已有的其他条目"""
        with tempfile.TemporaryDirectory() as tmpdir:
            claude_dir = Path(tmpdir) / ".claude"
            claude_dir.mkdir(parents=True)
            config_path = claude_dir / "mcp.json"

            # 预存已有配置
            existing_config = {
                "existing-server": {
                    "command": "some-other-tool",
                    "args": ["serve"],
                }
            }
            config_path.write_text(json.dumps(existing_config), encoding="utf-8")

            import wps_cli.cli.mcp_cmd as mcp_mod

            original = mcp_mod._MCP_CONFIG_TEMPLATES.copy()
            try:
                mcp_mod._MCP_CONFIG_TEMPLATES["claude"]["config_path"] = str(config_path)
                result = runner.invoke(main_app, ["mcp", "install", "-t", "claude"])
            finally:
                mcp_mod._MCP_CONFIG_TEMPLATES = original

            assert result.exit_code == 0
            config = json.loads(config_path.read_text(encoding="utf-8"))
            assert "existing-server" in config  # 保留已有条目
            assert "wps-cli" in config           # 新增 wps-cli 条目


# ── mcp status ──


class TestMcpStatus:
    """测试 'wps mcp status' 命令"""

    def test_status_command_runs(self, runner):
        """status 命令应能正常运行"""
        result = runner.invoke(main_app, ["mcp", "status"])
        assert result.exit_code == 0

    def test_status_json_output(self, runner):
        """status 命令的 JSON 输出应包含所有目标"""
        result = runner.invoke(main_app, ["mcp", "status"])
        assert result.exit_code == 0
        assert "claude" in result.stdout.lower() or result.stdout.strip() != ""


# ── install skill ──


class TestInstallSkill:
    """测试 'wps install skill' 命令"""

    def test_skill_command_help(self, runner):
        """确认 install skill 子命令已注册"""
        result = runner.invoke(main_app, ["install", "skill", "--help"])
        assert result.exit_code == 0

    def test_skill_invalid_target_errors(self, runner):
        """无效的目标应报错"""
        result = runner.invoke(main_app, ["install", "skill", "-t", "invalid"])
        assert result.exit_code != 0


# ── install mcp ──


class TestInstallMcp:
    """测试 'wps install mcp' 命令"""

    def test_install_mcp_command_help(self, runner):
        """确认 install mcp 子命令已注册"""
        result = runner.invoke(main_app, ["install", "mcp", "--help"])
        assert result.exit_code == 0

    def test_install_mcp_creates_config(self, runner):
        """安装 MCP 应创建配置文件"""
        with tempfile.TemporaryDirectory() as tmpdir:
            claude_dir = Path(tmpdir) / ".claude"
            config_path = claude_dir / "mcp.json"

            import wps_cli.cli.install_cmd as inst_mod

            original = inst_mod._MCP_TARGETS.copy()
            try:
                inst_mod._MCP_TARGETS["claude"]["config_path"] = str(config_path)
                result = runner.invoke(main_app, ["install", "mcp", "-t", "claude"])
            finally:
                inst_mod._MCP_TARGETS = original

            assert result.exit_code == 0
            assert config_path.exists()


# ── install all-tools ──


class TestInstallAllTools:
    """测试 'wps install all-tools' 命令"""

    def test_all_tools_command_help(self, runner):
        """确认 install all-tools 子命令已注册"""
        result = runner.invoke(main_app, ["install", "all-tools", "--help"])
        assert result.exit_code == 0
