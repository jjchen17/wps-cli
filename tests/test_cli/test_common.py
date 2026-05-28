"""CLI 公共工具测试 — 验证统一 JSON Schema 与错误响应"""

import json

import click
import pytest

from wps_cli.cli.common import handle_error, success
from wps_cli.exceptions import ValidationError, WpsNotFoundError


def _read_json(captured) -> dict:
    return json.loads(captured.out)


class TestSuccessOutput:
    def test_json_envelope(self, capsys):
        success({"x": 1}, command="calc.cell_get", json_mode=True)
        captured = capsys.readouterr()
        payload = _read_json(captured)
        assert payload["success"] is True
        assert payload["command"] == "calc.cell_get"
        assert payload["data"] == {"x": 1}

    def test_text_dict_output(self, capsys):
        success({"path": "out.pdf"}, command="x", json_mode=False)
        captured = capsys.readouterr()
        assert "path: out.pdf" in captured.out


class TestErrorOutput:
    def test_validation_error_exit_code(self, capsys):
        with pytest.raises(click.exceptions.Exit) as exc:
            handle_error(
                ValidationError("bad arg"),
                command="calc.cell_set",
                json_mode=True,
            )
        # ValidationError.exit_code == 50
        assert exc.value.exit_code == 50
        captured = capsys.readouterr()
        payload = _read_json(captured)
        assert payload["success"] is False
        assert payload["command"] == "calc.cell_set"
        assert payload["error"]["type"] == "ValidationError"
        assert payload["error"]["code"] == 50

    def test_wps_not_found_includes_suggestion(self, capsys):
        with pytest.raises(click.exceptions.Exit) as exc:
            handle_error(WpsNotFoundError("writer"), command="writer.new", json_mode=True)
        assert exc.value.exit_code == 10
        captured = capsys.readouterr()
        payload = _read_json(captured)
        assert payload["error"]["suggestion"]
        assert "doctor" in payload["error"]["suggestion"]

    def test_generic_exception_falls_back(self, capsys):
        with pytest.raises(click.exceptions.Exit) as exc:
            handle_error(RuntimeError("boom"), command="x", json_mode=True)
        assert exc.value.exit_code == 1
        captured = capsys.readouterr()
        payload = _read_json(captured)
        assert payload["error"]["type"] == "RuntimeError"

    def test_path_redaction_in_error(self, capsys):
        err = ValidationError("failed to open C:\\Users\\alice\\secret.docx")
        with pytest.raises(click.exceptions.Exit):
            handle_error(err, command="x", json_mode=True)
        captured = capsys.readouterr()
        payload = _read_json(captured)
        assert "alice" not in payload["error"]["message"]
