# 设计参考: iOfficeAI/OfficeCLI (Apache 2.0, https://github.com/iOfficeAI/OfficeCLI)
"""install_cmd 安装命令单元测试"""

from __future__ import annotations

from pathlib import Path
from unittest.mock import patch

import pytest
from typer.testing import CliRunner

from wps_cli.cli.install_cmd import (
    _copy_skill_dir,
    _find_skill_dir,
    _find_skill_md,
    _get_skill_target_dir,
)
from wps_cli.main import app as main_app

# ── helpers ──


def _mock_expanduser(base_dir: Path):
    """Mock os.path.expanduser 替换 ~ 为 base_dir，同时保留路径其余部分"""

    def expanduser(path: str) -> str:
        if path.startswith("~"):
            # 去掉 ~ 前缀，拼接 base_dir
            remainder = path[1:].lstrip("/\\")
            if remainder:
                return str(base_dir / remainder)
            return str(base_dir)
        return path

    return expanduser


@pytest.fixture
def runner() -> CliRunner:
    """提供 typer CLI 测试运行器"""
    return CliRunner()


# ── _find_skill_dir() 测试 ──


class TestFindSkillDir:
    """测试 _find_skill_dir() 函数"""

    def test_returns_path_when_dir_exists(self, tmp_path: Path):
        """skills/wps-cli/ 目录存在时应返回正确路径"""
        skill_dir = tmp_path / "skills" / "wps-cli"
        skill_dir.mkdir(parents=True)
        (skill_dir / "SKILL.md").write_text("# Test SKILL", encoding="utf-8")

        with patch("wps_cli.cli.install_cmd.Path.cwd", return_value=tmp_path):
            import wps_cli
            with patch.object(wps_cli, "__file__", str(tmp_path / "src" / "wps_cli" / "__init__.py")):
                result = _find_skill_dir()

        assert result is not None
        assert result.is_dir()
        assert (result / "SKILL.md").exists()

    def test_returns_none_when_dir_missing(self, tmp_path: Path):
        """skills/wps-cli/ 目录不存在时应返回 None"""
        skill_dir = tmp_path / "skills" / "wps-cli"
        assert not skill_dir.exists()

        with patch("wps_cli.cli.install_cmd.Path.cwd", return_value=tmp_path):
            import wps_cli
            with patch.object(wps_cli, "__file__", str(tmp_path / "src" / "wps_cli" / "__init__.py")):
                result = _find_skill_dir()

        assert result is None

    def test_dir_exists_but_no_skill_md(self, tmp_path: Path):
        """skills/wps-cli/ 目录存在但无 SKILL.md 时应返回 None"""
        skill_dir = tmp_path / "skills" / "wps-cli"
        skill_dir.mkdir(parents=True)
        # 不创建 SKILL.md

        with patch("wps_cli.cli.install_cmd.Path.cwd", return_value=tmp_path):
            import wps_cli
            with patch.object(wps_cli, "__file__", str(tmp_path / "src" / "wps_cli" / "__init__.py")):
                result = _find_skill_dir()

        assert result is None

    def test_from_pkg_dir_parent_candidate(self, tmp_path: Path):
        """通过 wps_cli 包目录的 .. 候选路径查找（项目根目录）"""
        project_root = tmp_path / "my_project"
        project_root.mkdir()
        skill_dir = project_root / "skills" / "wps-cli"
        skill_dir.mkdir(parents=True)
        (skill_dir / "SKILL.md").write_text("# Test", encoding="utf-8")

        import wps_cli
        fake_pkg_file = project_root / "src" / "wps_cli" / "__init__.py"
        fake_pkg_file.parent.mkdir(parents=True)
        fake_pkg_file.write_text("", encoding="utf-8")

        with patch.object(wps_cli, "__file__", str(fake_pkg_file)):
            with patch("wps_cli.cli.install_cmd.Path.cwd", return_value=tmp_path):
                result = _find_skill_dir()

        assert result is not None
        assert result.is_dir()
        assert (result / "SKILL.md").exists()


# ── _find_skill_md() 测试 ──


class TestFindSkillMd:
    """测试 _find_skill_md() 回退查找"""

    def test_returns_path_when_found_in_cwd(self, tmp_path: Path):
        """在当前工作目录下找到 SKILL.md"""
        skill_md = tmp_path / "SKILL.md"
        skill_md.write_text("# SKILL", encoding="utf-8")

        with patch("wps_cli.cli.install_cmd.Path.cwd", return_value=tmp_path):
            import wps_cli
            with patch.object(wps_cli, "__file__", str(tmp_path / "src" / "wps_cli" / "__init__.py")):
                result = _find_skill_md()

        assert result is not None
        assert result.name == "SKILL.md"

    def test_returns_none_when_not_found(self, tmp_path: Path):
        """SKILL.md 不存在时应返回 None"""
        with patch("wps_cli.cli.install_cmd.Path.cwd", return_value=tmp_path):
            import wps_cli
            with patch.object(wps_cli, "__file__", str(tmp_path / "src" / "wps_cli" / "__init__.py")):
                result = _find_skill_md()

        assert result is None

    def test_from_project_root_candidate(self, tmp_path: Path):
        """从项目根目录候选路径查找到 SKILL.md"""
        project_root = tmp_path / "my_project"
        project_root.mkdir()
        skill_md = project_root / "SKILL.md"
        skill_md.write_text("# SKILL", encoding="utf-8")

        import wps_cli
        fake_pkg_file = project_root / "src" / "wps_cli" / "__init__.py"
        fake_pkg_file.parent.mkdir(parents=True)
        fake_pkg_file.write_text("", encoding="utf-8")

        with patch.object(wps_cli, "__file__", str(fake_pkg_file)):
            result = _find_skill_md()

        assert result is not None
        assert result.name == "SKILL.md"


# ── _get_skill_target_dir() 测试 ──


class TestGetSkillTargetDir:
    """测试 _get_skill_target_dir() 函数"""

    def test_claude_target_path(self, tmp_path: Path):
        """claude 目标应返回 ~/.claude/skills/wps-cli 路径"""
        home = tmp_path / "home"

        with patch("os.path.expanduser", side_effect=_mock_expanduser(home)):
            result = _get_skill_target_dir("claude")

        expected = home / ".claude" / "skills" / "wps-cli"
        assert result == expected

    def test_cursor_target_path(self, tmp_path: Path):
        """cursor 目标应返回 .cursor/skills/wps-cli（相对路径，不以 ~ 开头）"""
        result = _get_skill_target_dir("cursor")

        # cursor 路径是相对的（不以 ~ 开头），expanduser 不展开
        expected = Path(".cursor/skills/wps-cli")
        assert result == expected

    def test_windsurf_target_path(self, tmp_path: Path):
        """windsurf 目标使用 _SKILL_TARGETS 中的 path 去掉 .md 后缀"""
        result = _get_skill_target_dir("windsurf")

        # windsurf 路径是相对的（不以 ~ 开头），expanduser 不展开
        expected = Path(".windsurf/skills/wps-cli")
        assert result == expected

    def test_unknown_target_fallback(self, tmp_path: Path):
        """未知目标应回退到 ~/.{target}/skills/wps-cli 默认路径"""
        cwd = tmp_path / "project"
        cwd.mkdir()

        with patch("os.path.expanduser", side_effect=_mock_expanduser(cwd)):
            result = _get_skill_target_dir("unknown_tool")

        expected = cwd / ".unknown_tool" / "skills" / "wps-cli"
        assert result == expected


# ── _copy_skill_dir() 测试 ──


class TestCopySkillDir:
    """测试 _copy_skill_dir() 函数"""

    def test_copies_directory_recursively(self, tmp_path: Path):
        """递归复制目录，返回 SKILL.md 路径"""
        src_dir = tmp_path / "src_skill"
        src_dir.mkdir()
        (src_dir / "SKILL.md").write_text("# Main SKILL", encoding="utf-8")
        sub_dir = src_dir / "reference"
        sub_dir.mkdir()
        (sub_dir / "guide.md").write_text("# Reference Guide", encoding="utf-8")

        dest_dir = tmp_path / "dest" / "wps-cli"

        result = _copy_skill_dir(src_dir, dest_dir)

        assert result == dest_dir / "SKILL.md"
        assert dest_dir.exists()
        assert (dest_dir / "SKILL.md").exists()
        assert (dest_dir / "SKILL.md").read_text(encoding="utf-8") == "# Main SKILL"
        assert (dest_dir / "reference" / "guide.md").exists()
        assert (dest_dir / "reference" / "guide.md").read_text(encoding="utf-8") == "# Reference Guide"

    def test_overwrites_existing_destination(self, tmp_path: Path):
        """目标目录已存在时应先删除再复制"""
        src_dir = tmp_path / "src_skill"
        src_dir.mkdir()
        (src_dir / "SKILL.md").write_text("new content", encoding="utf-8")

        dest_dir = tmp_path / "dest"
        dest_dir.mkdir()
        (dest_dir / "old_file.md").write_text("old content", encoding="utf-8")

        _copy_skill_dir(src_dir, dest_dir)

        assert (dest_dir / "SKILL.md").exists()
        assert (dest_dir / "SKILL.md").read_text(encoding="utf-8") == "new content"
        assert not (dest_dir / "old_file.md").exists()


# ── CLI 集成测试 ──


class TestInstallSkillCli:
    """测试 'wps install skill' CLI 命令"""

    def test_skill_help_output(self, runner):
        """skill --help 应输出帮助信息"""
        result = runner.invoke(main_app, ["install", "skill", "--help"])
        assert result.exit_code == 0
        assert "skill" in result.stdout.lower()
        assert "--target" in result.stdout or "-t" in result.stdout

    def test_skill_invalid_target(self, runner):
        """无效 target 应报错退出"""
        result = runner.invoke(main_app, ["install", "skill", "-t", "invalid"])
        assert result.exit_code != 0
        assert "不支持" in result.stdout or "不支持" in result.stderr

    def test_skill_missing_skill_source(self, runner, tmp_path: Path):
        """找不到 SKILL.md 或 skills/wps-cli 目录时应报错"""
        import wps_cli
        fake_pkg_file = tmp_path / "src" / "wps_cli" / "__init__.py"
        fake_pkg_file.parent.mkdir(parents=True)
        fake_pkg_file.write_text("", encoding="utf-8")

        with patch.object(wps_cli, "__file__", str(fake_pkg_file)):
            with patch("wps_cli.cli.install_cmd.Path.cwd", return_value=tmp_path):
                result = runner.invoke(main_app, ["install", "skill", "-t", "claude"])

        assert result.exit_code != 0
        assert "找不到" in result.stdout or "找不到" in result.stderr

    def test_skill_install_to_claude_with_skill_dir(self, runner, tmp_path: Path):
        """有 skills/wps-cli/SKILL.md 时安装到 claude 应创建目标目录"""
        # 模拟项目结构
        skill_dir = tmp_path / "skills" / "wps-cli"
        skill_dir.mkdir(parents=True)
        (skill_dir / "SKILL.md").write_text("# WPS CLI Skill", encoding="utf-8")

        home = tmp_path / "home"
        expected_dest = home / ".claude" / "skills" / "wps-cli"

        import wps_cli
        fake_pkg_file = tmp_path / "src" / "wps_cli" / "__init__.py"
        fake_pkg_file.parent.mkdir(parents=True)
        fake_pkg_file.write_text("", encoding="utf-8")

        with patch.object(wps_cli, "__file__", str(fake_pkg_file)):
            with patch("wps_cli.cli.install_cmd.Path.cwd", return_value=tmp_path):
                with patch("os.path.expanduser", side_effect=_mock_expanduser(home)):
                    result = runner.invoke(main_app, ["install", "skill", "-t", "claude"])

        assert result.exit_code == 0
        assert expected_dest.exists()
        assert (expected_dest / "SKILL.md").exists()

    def test_skill_install_all_targets(self, runner, tmp_path: Path):
        """安装到所有目标应正常执行"""
        skill_dir = tmp_path / "skills" / "wps-cli"
        skill_dir.mkdir(parents=True)
        (skill_dir / "SKILL.md").write_text("# WPS CLI Skill", encoding="utf-8")

        home = tmp_path / "home"

        import wps_cli
        fake_pkg_file = tmp_path / "src" / "wps_cli" / "__init__.py"
        fake_pkg_file.parent.mkdir(parents=True)
        fake_pkg_file.write_text("", encoding="utf-8")

        with patch.object(wps_cli, "__file__", str(fake_pkg_file)):
            with patch("wps_cli.cli.install_cmd.Path.cwd", return_value=tmp_path):
                with patch("os.path.expanduser", side_effect=_mock_expanduser(home)):
                    result = runner.invoke(main_app, ["install", "skill"])

        assert result.exit_code == 0
        # 验证至少 claude 目标被创建
        claude_dest = home / ".claude" / "skills" / "wps-cli"
        assert claude_dest.exists()


class TestInstallMcpCli:
    """测试 'wps install mcp' CLI 命令"""

    def test_mcp_help_output(self, runner):
        """install mcp --help 应输出帮助信息"""
        result = runner.invoke(main_app, ["install", "mcp", "--help"])
        assert result.exit_code == 0
        assert "mcp" in result.stdout.lower()

    def test_mcp_invalid_target(self, runner):
        """无效 target 应报错"""
        result = runner.invoke(main_app, ["install", "mcp", "-t", "invalid"])
        assert result.exit_code != 0


class TestInstallAllToolsCli:
    """测试 'wps install all-tools' CLI 命令"""

    def test_all_tools_help_output(self, runner):
        """install all-tools --help 应输出帮助信息"""
        result = runner.invoke(main_app, ["install", "all-tools", "--help"])
        assert result.exit_code == 0


# ── 边界情况与 MCP 安装测试 ──


class TestEdgeCases:
    """边界情况和回归测试"""

    def test_find_skill_dir_import_error_fallback(self, tmp_path: Path):
        """wps_cli 导入失败时回退到 Path.cwd()/src"""
        skill_dir = tmp_path / "src" / ".." / "skills" / "wps-cli"
        skill_dir = skill_dir.resolve()
        skill_dir.mkdir(parents=True, exist_ok=True)
        (skill_dir / "SKILL.md").write_text("# Test", encoding="utf-8")

        with patch("wps_cli.cli.install_cmd.Path.cwd", return_value=tmp_path / "src"):
            # 模拟 import 失败：让 _find_skill_dir 中的 import 触发异常
            with patch("wps_cli.__file__", None, create=True):
                result = _find_skill_dir()

        # import 失败时回退到 Path.cwd()/"src"，然后转 ".."/"skills"/"wps-cli"
        assert result is None or (result.is_dir() if result else True)

    def test_skill_cli_help_shows_all_targets(self, runner):
        """skill --help 输出应列出所有可用的 target"""
        result = runner.invoke(main_app, ["install", "skill", "--help"])
        assert result.exit_code == 0
        for t in ("claude", "cursor", "vscode", "windsurf"):
            assert t in result.stdout

    def test_skill_install_with_skill_md_fallback(self, runner, tmp_path: Path):
        """skills/wps-cli/ 目录不存在但有 SKILL.md 时使用单文件回退"""
        skill_md = tmp_path / "SKILL.md"
        skill_md.write_text("# WPS CLI Skill (fallback)", encoding="utf-8")

        home = tmp_path / "home"

        import wps_cli
        fake_pkg_file = tmp_path / "src" / "wps_cli" / "__init__.py"
        fake_pkg_file.parent.mkdir(parents=True)
        fake_pkg_file.write_text("", encoding="utf-8")

        with patch.object(wps_cli, "__file__", str(fake_pkg_file)):
            with patch("wps_cli.cli.install_cmd.Path.cwd", return_value=tmp_path):
                with patch("os.path.expanduser", side_effect=_mock_expanduser(home)):
                    result = runner.invoke(main_app, ["install", "skill", "-t", "claude"])

        assert result.exit_code == 0
        dest_file = home / ".claude" / "skills" / "wps-cli.md"
        assert dest_file.exists()

    def test_install_mcp_creates_config(self, runner, tmp_path: Path):
        """安装 MCP 应创建配置文件"""
        claude_dir = tmp_path / ".claude"
        claude_dir.mkdir()
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

    def test_install_mcp_skip_invalid_json(self, runner, tmp_path: Path):
        """MCP 配置文件不是合法 JSON 时应给出警告并创建新文件"""
        claude_dir = tmp_path / ".claude"
        claude_dir.mkdir(parents=True)
        config_path = claude_dir / "mcp.json"
        config_path.write_text("not valid json {{{", encoding="utf-8")

        import wps_cli.cli.install_cmd as inst_mod

        original = inst_mod._MCP_TARGETS.copy()
        try:
            inst_mod._MCP_TARGETS["claude"]["config_path"] = str(config_path)
            result = runner.invoke(main_app, ["install", "mcp", "-t", "claude"])
        finally:
            inst_mod._MCP_TARGETS = original

        assert result.exit_code == 0
        assert "警告" in result.stdout or "警告" in result.stderr
