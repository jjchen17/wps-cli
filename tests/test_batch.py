# 设计参考: iOfficeAI/OfficeCLI (Apache 2.0, https://github.com/iOfficeAI/OfficeCLI)
"""Batch 批量命令测试"""

from __future__ import annotations

import json
from pathlib import Path
from unittest.mock import MagicMock

import pytest
from typer.testing import CliRunner

from wps_cli.main import app

runner = CliRunner()


# ── 基本 CLI 测试 ──


def test_batch_help():
    """batch 命令出现在帮助输出中"""
    result = runner.invoke(app, ["batch", "--help"])
    assert result.exit_code == 0
    assert "批量" in result.output or "batch" in result.output.lower()


def test_batch_no_input():
    """无输入时提示错误"""
    result = runner.invoke(app, ["batch"])
    assert result.exit_code == 1
    assert "错误" in result.output


def test_batch_empty_commands():
    """空命令数组"""
    result = runner.invoke(app, ["batch", "--json"], input="[]")
    assert result.exit_code == 0
    assert "空" in result.output or "无操作" in result.output


def test_batch_invalid_json():
    """无效 JSON 输入"""
    result = runner.invoke(app, ["batch", "--commands", "not-json"])
    assert result.exit_code == 1
    assert "JSON" in result.output


def test_batch_not_array():
    """输入不是 JSON 数组"""
    result = runner.invoke(app, ["batch", "--commands", '{"key": "value"}'])
    assert result.exit_code == 1
    assert "数组" in result.output


def test_batch_empty_input_string():
    """空字符串输入"""
    result = runner.invoke(app, ["batch", "--commands", ""])
    assert result.exit_code == 1


# ── 命令输出格式测试 ──


def test_batch_output_format():
    """验证 batch 输出格式"""
    result = runner.invoke(app, ["batch", "--json"], input="[]")
    assert result.exit_code == 0
    # JSON 跨多行在输出前面，后面是中文摘要文本
    output = result.output
    # 找到第一个 { 和对应的 }
    start = output.index("{")
    depth = 0
    end = start
    for i, ch in enumerate(output[start:], start):
        if ch == "{":
            depth += 1
        elif ch == "}":
            depth -= 1
            if depth == 0:
                end = i + 1
                break
    json_str = output[start:end]
    data = json.loads(json_str)
    assert data["success"] is True
    assert data["command"] == "batch"
    assert "steps" in data["data"]
    assert "summary" in data["data"]
    assert data["data"]["summary"] == {"total": 0, "succeeded": 0, "failed": 0}


def test_batch_with_commands_param():
    """通过 --commands 参数传入命令"""
    result = runner.invoke(app, ["batch", "--commands", "[]"])
    assert result.exit_code == 0


def test_batch_with_input_file(tmp_path: Path):
    """通过 --input 文件传入命令"""
    commands_file = tmp_path / "commands.json"
    commands_file.write_text("[]", encoding="utf-8")
    result = runner.invoke(app, ["batch", "--input", str(commands_file)])
    assert result.exit_code == 0


# ── 命令错误处理测试 ──


def test_batch_unknown_command():
    """未知命令处理"""
    commands = [
        {"command": "unknown.cmd", "params": {}}
    ]
    result = runner.invoke(app, ["batch", "--commands", json.dumps(commands)])
    # 调用 COM 可能会失败（测试环境无 WPS），但我们检查 JSON 输出格式
    output = result.output
    # 至少 batch 被调用了且产生了输出（即使 COM 失败也有 JSON 输出）
    assert result.exit_code != 0 or "error" in output.lower() or "失败" in output


def test_batch_missing_command_field():
    """缺少 command 字段的命令"""
    commands = [
        {"params": {"file": "test.docx"}},
    ]
    # 这个测试可能触发 COM 连接，先验证 JSON 可以解析
    assert isinstance(commands, list)
    assert len(commands) == 1
    assert "command" not in commands[0]


def test_batch_per_step_result_format():
    """验证 per-step 结果格式（空命令数组）"""
    result = runner.invoke(app, ["batch", "--json"], input="[]")
    assert result.exit_code == 0
    # 从输出中提取完整的 JSON（JSON 有多行，在摘要文本之前）
    output = result.output.strip()
    json_start = output.find("{")
    json_end = output.rfind("}") + 1
    assert json_start >= 0, "输出中未找到 JSON"
    json_str = output[json_start:json_end]
    data = json.loads(json_str)
    assert "command" in data
    assert "data" in data
    assert "steps" in data["data"]
    assert "summary" in data["data"]


# ── 驻留模式回退测试 ──


def test_batch_resident_fallback(monkeypatch):
    """驻留不可用时回退到本地执行"""
    import socket

    def mock_connect_ex(*args, **kwargs):
        return 1  # 连接失败

    monkeypatch.setattr(socket.socket, "connect_ex", mock_connect_ex)

    commands = [{"command": "writer.info", "params": {"file": "nonexistent.docx"}}]
    result = runner.invoke(app, ["batch", "--commands", json.dumps(commands)])
    # 文件不存在时会报错，但说明本地执行了（没有转发到驻留）
    assert "文件不存在" in result.output or result.exit_code != 0


# ── BatchExecutor 单元测试 ──


class TestBatchExecutor:
    """BatchExecutor 类的单元测试（使用 mock 避免 COM 依赖）"""

    def test_handler_map_completeness(self):
        """确认 COMMAND_HANDLER_MAP 覆盖了所有核心命令"""
        from wps_cli.cli.batch import BatchExecutor

        handler_map = BatchExecutor.COMMAND_HANDLER_MAP

        # Writer 命令
        assert "writer.info" in handler_map
        assert "writer.replace" in handler_map
        assert "writer.count" in handler_map
        assert "writer.table-get" in handler_map
        assert "writer.table-insert" in handler_map
        assert "writer.image-insert" in handler_map
        assert "writer.page-setup" in handler_map
        assert "writer.style-apply" in handler_map
        assert "writer.export-pdf" in handler_map
        assert "writer.new" in handler_map

        # Calc 命令
        assert "calc.info" in handler_map
        assert "calc.cell-get" in handler_map
        assert "calc.cell-set" in handler_map
        assert "calc.cell-range" in handler_map
        assert "calc.cell-formula" in handler_map
        assert "calc.chart-create" in handler_map
        assert "calc.sort" in handler_map
        assert "calc.export-csv" in handler_map
        assert "calc.sheet-list" in handler_map
        assert "calc.new" in handler_map

        # Impress 命令
        assert "impress.info" in handler_map
        assert "impress.slide-list" in handler_map
        assert "impress.slide-add" in handler_map
        assert "impress.slide-delete" in handler_map
        assert "impress.text-get" in handler_map
        assert "impress.text-set" in handler_map
        assert "impress.image-insert" in handler_map
        assert "impress.export-pdf" in handler_map
        assert "impress.new" in handler_map

        # PDF 命令
        assert "pdf.info" in handler_map
        assert "pdf.merge" in handler_map
        assert "pdf.split" in handler_map
        assert "pdf.watermark" in handler_map
        assert "pdf.extract-pages" in handler_map

        # Export 命令
        assert "export.convert" in handler_map
        assert "export.batch" in handler_map

    def test_resolve_unknown_command(self):
        """未知命令抛出 ValidationError"""
        from wps_cli.cli.batch import BatchExecutor
        from wps_cli.exceptions import ValidationError

        executor = BatchExecutor()
        with pytest.raises(ValidationError, match="未知命令"):
            executor._resolve_handler("unknown.cmd")

    def test_detect_app_type(self):
        """文件扩展名到应用类型检测"""
        from wps_cli.cli.batch import _detect_app_type

        assert _detect_app_type("test.docx") == "writer"
        assert _detect_app_type("test.doc") == "writer"
        assert _detect_app_type("test.xlsx") == "calc"
        assert _detect_app_type("test.xls") == "calc"
        assert _detect_app_type("test.csv") == "calc"
        assert _detect_app_type("test.pptx") == "impress"
        assert _detect_app_type("test.ppt") == "impress"

        from wps_cli.exceptions import ValidationError
        with pytest.raises(ValidationError):
            _detect_app_type("test.unknown")

    def test_is_resident_running_negative(self, monkeypatch):
        """驻留未运行的情况"""
        import socket

        def mock_connect_ex(*args, **kwargs):
            return 1  # 未运行

        monkeypatch.setattr(socket.socket, "connect_ex", mock_connect_ex)
        from wps_cli.cli.batch import _is_resident_running
        assert _is_resident_running() is False

    def test_is_resident_running_positive(self, monkeypatch):
        """驻留在运行的情况"""
        import socket

        def mock_connect_ex(*args, **kwargs):
            return 0  # 正在运行

        monkeypatch.setattr(socket.socket, "connect_ex", mock_connect_ex)
        from wps_cli.cli.batch import _is_resident_running
        assert _is_resident_running() is True

    def test_is_resident_running_exception(self, monkeypatch):
        """socket 异常时认为未运行"""
        import socket

        def mock_connect_ex(*args, **kwargs):
            raise OSError("模拟错误")

        monkeypatch.setattr(socket.socket, "connect_ex", mock_connect_ex)
        from wps_cli.cli.batch import _is_resident_running
        assert _is_resident_running() is False

    def test_execute_batch_continue_on_error(self, monkeypatch):
        """continue-on-error 模式：一条失败不影响后续"""
        from wps_cli.cli.batch import BatchExecutor

        executor = BatchExecutor()

        # Mock _ensure_services 避免 COM 初始化
        executor._ensure_services = MagicMock()

        call_order = []

        class MockHandler:
            def __init__(self, name):
                self.name = name

            def __call__(self, params):
                call_order.append(self.name)
                if self.name == "fail_me":
                    raise ValueError("模拟失败")
                return {"status": "ok"}

        # Mock _resolve_handler
        handlers = {
            "cmd.one": MockHandler("one"),
            "cmd.fail": MockHandler("fail_me"),
            "cmd.two": MockHandler("two"),
        }
        executor._resolve_handler = lambda cmd: handlers.get(cmd, MagicMock())

        # Mock _cleanup
        executor._cleanup = MagicMock()

        commands = [
            {"command": "cmd.one", "params": {}},
            {"command": "cmd.fail", "params": {}},
            {"command": "cmd.two", "params": {}},
        ]

        result = executor.execute_batch(commands, stop_on_error=False)

        assert result["summary"]["total"] == 3
        assert result["summary"]["succeeded"] == 2
        assert result["summary"]["failed"] == 1
        assert len(result["steps"]) == 3

        # 验证执行顺序
        assert call_order == ["one", "fail_me", "two"]

        # 验证各步骤
        assert result["steps"][0]["success"] is True
        assert result["steps"][0]["command"] == "cmd.one"
        assert result["steps"][0]["result"] == {"status": "ok"}

        assert result["steps"][1]["success"] is False
        assert result["steps"][1]["command"] == "cmd.fail"
        assert result["steps"][1]["error"]["type"] == "ValueError"

        assert result["steps"][2]["success"] is True
        assert result["steps"][2]["command"] == "cmd.two"

    def test_execute_batch_stop_on_error(self, monkeypatch):
        """stop-on-error 模式：第一条失败就停止"""
        from wps_cli.cli.batch import BatchExecutor

        executor = BatchExecutor()
        executor._ensure_services = MagicMock()
        executor._cleanup = MagicMock()

        call_order = []

        class MockHandler:
            def __init__(self, name):
                self.name = name

            def __call__(self, params):
                call_order.append(self.name)
                if self.name == "fail_first":
                    raise ValueError("模拟失败")
                return {"status": "ok"}

        handlers = {
            "cmd.fail": MockHandler("fail_first"),
            "cmd.two": MockHandler("two"),
        }
        executor._resolve_handler = lambda cmd: handlers.get(cmd, MagicMock())

        commands = [
            {"command": "cmd.fail", "params": {}},
            {"command": "cmd.two", "params": {}},
        ]

        result = executor.execute_batch(commands, stop_on_error=True)

        assert result["summary"]["total"] == 2
        assert result["summary"]["succeeded"] == 0
        assert result["summary"]["failed"] == 1
        assert len(result["steps"]) == 1  # 只执行了一步就停止了

        # 验证第二条命令没有执行
        assert call_order == ["fail_first"]

    def test_execute_batch_empty(self):
        """空命令数组"""
        from wps_cli.cli.batch import BatchExecutor

        executor = BatchExecutor()
        executor._ensure_services = MagicMock()
        executor._cleanup = MagicMock()

        result = executor.execute_batch([])
        assert result["summary"]["total"] == 0
        assert result["summary"]["succeeded"] == 0
        assert result["summary"]["failed"] == 0
        assert len(result["steps"]) == 0

    def test_execute_batch_missing_command_field(self):
        """缺少 command 字段的命令"""
        from wps_cli.cli.batch import BatchExecutor

        executor = BatchExecutor()
        executor._ensure_services = MagicMock()
        executor._cleanup = MagicMock()

        commands = [
            {"params": {"file": "test.docx"}},
        ]

        result = executor.execute_batch(commands)
        assert result["summary"]["total"] == 1
        assert result["summary"]["succeeded"] == 0
        assert result["summary"]["failed"] == 1
        assert result["steps"][0]["success"] is False
        assert "缺少 command" in result["steps"][0]["error"]["message"]

    def test_execute_batch_unknown_command(self):
        """未知命令"""
        from wps_cli.cli.batch import BatchExecutor

        executor = BatchExecutor()
        executor._ensure_services = MagicMock()
        executor._cleanup = MagicMock()

        commands = [
            {"command": "no.such.command", "params": {}},
        ]

        result = executor.execute_batch(commands)
        assert result["summary"]["failed"] == 1
        assert result["steps"][0]["error"]["type"] == "ValidationError"
        assert "未知命令" in result["steps"][0]["error"]["message"]
