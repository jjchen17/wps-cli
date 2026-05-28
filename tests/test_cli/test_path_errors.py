"""CLI 路径校验类测试 — 不依赖 COM

验证：
* 文件不存在 → exit code 21
* 扩展名不在白名单 → exit code 50
* 参数本身非法 → exit code 50
* 错误响应是合法的统一 JSON Schema
"""

from __future__ import annotations

import json

import pytest
from typer.testing import CliRunner

from wps_cli.main import app

runner = CliRunner()


def _json(out: str) -> dict:
    return json.loads(out)


@pytest.mark.parametrize(
    "cmd_args",
    [
        ["writer", "info", "ghost.docx"],
        ["writer", "count", "ghost.docx"],
        ["writer", "table-get", "ghost.docx"],
        ["calc", "info", "ghost.xlsx"],
        ["calc", "sheet-list", "ghost.xlsx"],
        ["calc", "cell-get", "ghost.xlsx", "A1"],
        ["impress", "info", "ghost.pptx"],
        ["impress", "slide-list", "ghost.pptx"],
        ["pdf", "info", "ghost.pdf"],
    ],
)
def test_missing_input_file_exits_21(cmd_args):
    result = runner.invoke(app, [*cmd_args, "--json"])
    assert result.exit_code == 21
    payload = _json(result.output)
    assert payload["success"] is False
    assert payload["error"]["code"] == 21
    assert payload["error"]["type"] == "FileNotFoundErrorCli"


class TestExtensionWhitelist:
    """H-1: 拒绝非目标应用类型的文件，防止 README.md 被覆盖为 xlsx"""

    def test_calc_rejects_md_file(self, tmp_path):
        f = tmp_path / "README.md"
        f.write_text("# hello", encoding="utf-8")
        result = runner.invoke(app, ["calc", "cell-set", str(f), "A1", "value", "--json"])
        assert result.exit_code == 50
        payload = _json(result.output)
        assert "扩展名" in payload["error"]["message"]

    def test_writer_rejects_xlsx(self, tmp_path):
        f = tmp_path / "data.xlsx"
        f.write_bytes(b"fake")
        result = runner.invoke(app, ["writer", "info", str(f), "--json"])
        assert result.exit_code == 50

    def test_pdf_rejects_docx(self, tmp_path):
        f = tmp_path / "data.docx"
        f.write_bytes(b"fake")
        result = runner.invoke(app, ["pdf", "info", str(f), "--json"])
        assert result.exit_code == 50


class TestWriterTableInsertValidation:
    """table_insert 的纯 Python 校验路径，不需要 COM"""

    def test_rows_zero_rejected(self, tmp_path):
        f = tmp_path / "doc.docx"
        f.write_bytes(b"fake")
        result = runner.invoke(
            app,
            ["writer", "table-insert", str(f), "--rows", "0", "--cols", "3", "--json"],
        )
        assert result.exit_code == 50

    def test_cols_negative_rejected(self, tmp_path):
        f = tmp_path / "doc.docx"
        f.write_bytes(b"fake")
        result = runner.invoke(
            app,
            ["writer", "table-insert", str(f), "--rows", "2", "--cols", "-1", "--json"],
        )
        assert result.exit_code == 50

    def test_invalid_json_data(self, tmp_path):
        f = tmp_path / "doc.docx"
        f.write_bytes(b"fake")
        result = runner.invoke(
            app,
            [
                "writer",
                "table-insert",
                str(f),
                "--rows",
                "2",
                "--cols",
                "2",
                "--data",
                "{bad json",
                "--json",
            ],
        )
        assert result.exit_code == 50


class TestWriterReplaceValidation:
    """H-3: 通配符反向引用 + 长度限制"""

    def test_empty_old_rejected(self, tmp_path):
        f = tmp_path / "doc.docx"
        f.write_bytes(b"fake")
        result = runner.invoke(app, ["writer", "replace", str(f), "", "new", "--json"])
        assert result.exit_code == 50

    def test_too_long_old_rejected(self, tmp_path):
        f = tmp_path / "doc.docx"
        f.write_bytes(b"fake")
        result = runner.invoke(app, ["writer", "replace", str(f), "a" * 2000, "x", "--json"])
        assert result.exit_code == 50

    def test_wildcard_backref_rejected(self, tmp_path):
        f = tmp_path / "doc.docx"
        f.write_bytes(b"fake")
        result = runner.invoke(
            app,
            [
                "writer",
                "replace",
                str(f),
                r"(\w+)",
                r"\1\1\1\1",
                "--wildcard",
                "--json",
            ],
        )
        assert result.exit_code == 50
        payload = _json(result.output)
        assert "反向引用" in payload["error"]["message"]


class TestStyleApplyListShortCircuit:
    """style-apply --list 不需要 COM 也不需要 file 存在"""

    def test_list_presets_no_com(self):
        result = runner.invoke(
            app,
            ["writer", "style-apply", "nonexistent.docx", "公文标题", "--list", "--json"],
        )
        assert result.exit_code == 0
        payload = _json(result.output)
        assert payload["success"] is True
        assert "公文标题" in payload["data"]["presets"]


class TestExportBatchValidation:
    """export batch 的 glob 安全性"""

    def test_absolute_glob_rejected(self, tmp_path):
        result = runner.invoke(
            app,
            [
                "export",
                "batch",
                "C:\\*.docx",
                "--to",
                "pdf",
                "--output-dir",
                str(tmp_path),
                "--json",
            ],
        )
        assert result.exit_code == 50

    def test_dotdot_glob_rejected(self, tmp_path, monkeypatch):
        monkeypatch.chdir(tmp_path)
        result = runner.invoke(
            app,
            [
                "export",
                "batch",
                "../**/*.docx",
                "--to",
                "pdf",
                "--output-dir",
                str(tmp_path / "out"),
                "--json",
            ],
        )
        assert result.exit_code == 50
