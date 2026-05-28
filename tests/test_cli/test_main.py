"""主入口与 doctor 命令测试"""

from __future__ import annotations

from unittest.mock import patch

from typer.testing import CliRunner

from wps_cli.main import app

runner = CliRunner()


def test_version():
    result = runner.invoke(app, ["version"])
    assert result.exit_code == 0
    assert "wps-cli" in result.output


def test_help():
    result = runner.invoke(app, ["--help"])
    assert result.exit_code == 0
    assert "WPS Office" in result.output


def test_subcommands_registered():
    """五个子命令组都在 --help 输出里"""
    result = runner.invoke(app, ["--help"])
    for sub in ("writer", "calc", "impress", "pdf", "export"):
        assert sub in result.output


class TestDoctorReport:
    """doctor --report 必须输出脱敏 markdown"""

    def test_report_includes_required_fields(self):
        with patch("wps_cli.main._detect_wps_versions", return_value={}):
            result = runner.invoke(app, ["doctor", "--report"])
        assert result.exit_code == 0
        out = result.output
        assert "Environment Report" in out
        assert "wps-cli:" in out
        assert "Python:" in out
        assert "Platform:" in out

    def test_report_no_user_path_leak(self, monkeypatch, tmp_path):
        """--report 输出不应包含 HOME 路径或盘符"""
        monkeypatch.setenv("HOME", str(tmp_path))
        monkeypatch.setenv("USERPROFILE", str(tmp_path))
        with patch("wps_cli.main._detect_wps_versions", return_value={"Writer": "12.1"}):
            result = runner.invoke(app, ["doctor", "--report"])
        assert result.exit_code == 0
        # 报告中不应有 / 或 \\ 跟随用户名的形式
        assert str(tmp_path) not in result.output
