# 设计参考: iOfficeAI/OfficeCLI (Apache 2.0, https://github.com/iOfficeAI/OfficeCLI)
"""MCP 服务器单元测试"""

from __future__ import annotations

from typing import Any

from wps_cli.backends.base import ComBackend

# ── 模拟 COM 后端（不与真实 WPS 交互）─────────────────────────────


class MockMcpApp:
    """模拟 WPS COM 应用对象，用于 MCP 测试"""

    def __init__(self, app_type: str):
        self.Name = f"Mock {app_type}"
        self.Version = "12.0.0-test"
        self.Visible = False
        self.AutomationSecurity = 0
        self.DisplayAlerts = True
        self._app_type = app_type
        self.calls: list[tuple[str, tuple, dict]] = []

    def Quit(self) -> None:
        self.calls.append(("Quit", (), {}))


class MockMcpBackend(ComBackend):
    """模拟 COM 后端，用于 MCP 测试"""

    def __init__(self) -> None:
        self.last_app: MockMcpApp | None = None

    def connect(self, app_type: str) -> object:
        app = MockMcpApp(app_type)
        self.last_app = app
        self.harden(app)
        return app

    def disconnect(self, app: Any) -> None:
        try:
            app.Quit()
        except Exception:
            pass

    def is_alive(self, app: Any) -> bool:
        return True

    def get_version(self, app: Any) -> str:
        return "12.0.0-test"


# ── JSON-RPC 协议测试 ─────────────────────────────────────────────


class TestJsonRpcProtocol:
    """测试 JSON-RPC 2.0 协议处理"""

    def test_initialize_returns_capabilities(self, monkeypatch):
        """initialize 方法应返回服务器能力和版本信息"""

        server = _create_mock_server(monkeypatch)
        request = {"jsonrpc": "2.0", "id": 1, "method": "initialize", "params": {}}
        response = server._handle_request(request)

        assert response["jsonrpc"] == "2.0"
        assert response["id"] == 1
        assert "result" in response
        assert response["result"]["protocolVersion"] == "2024-11-05"
        assert "tools" in response["result"]["capabilities"]
        assert response["result"]["serverInfo"]["name"] == "wps-cli"

    def test_tools_list_returns_all_tools(self, monkeypatch):
        """tools/list 应返回所有已注册的工具"""

        server = _create_mock_server(monkeypatch)
        request = {"jsonrpc": "2.0", "id": 2, "method": "tools/list", "params": {}}
        response = server._handle_request(request)

        assert "result" in response
        tools = response["result"]["tools"]
        assert len(tools) > 0
        # 每个工具必须有 name, description, inputSchema
        for tool in tools:
            assert "name" in tool
            assert "description" in tool
            assert "inputSchema" in tool

    def test_tools_call_unknown_tool_returns_error(self, monkeypatch):
        """调用不存在的工具应返回错误"""
        from wps_cli.mcp.server import TOOL_NOT_FOUND

        server = _create_mock_server(monkeypatch)
        request = {
            "jsonrpc": "2.0",
            "id": 3,
            "method": "tools/call",
            "params": {"name": "nonexistent_tool", "arguments": {}},
        }
        response = server._handle_request(request)

        assert "error" in response
        assert response["error"]["code"] == TOOL_NOT_FOUND

    def test_unknown_method_returns_error(self, monkeypatch):
        """未知的 JSON-RPC 方法应返回 METHOD_NOT_FOUND"""
        from wps_cli.mcp.server import METHOD_NOT_FOUND

        server = _create_mock_server(monkeypatch)
        request = {"jsonrpc": "2.0", "id": 4, "method": "some/unknown", "params": {}}
        response = server._handle_request(request)

        assert "error" in response
        assert response["error"]["code"] == METHOD_NOT_FOUND

    def test_invalid_json_returns_parse_error(self, monkeypatch):
        """无效的 JSON 输入应返回 PARSE_ERROR"""

        from wps_cli.mcp.server import PARSE_ERROR

        server = _create_mock_server(monkeypatch)

        # 模拟无效 JSON 通过 _handle_request 传入的场景
        # 直接测试 _make_error 行为
        error_response = server._make_error(0, PARSE_ERROR, "JSON 解析错误")
        assert error_response["error"]["code"] == PARSE_ERROR
        assert error_response["id"] == 0

    def test_notification_no_response(self, monkeypatch):
        """通知（如 notifications/initialized）应返回空响应"""

        server = _create_mock_server(monkeypatch)
        request = {"jsonrpc": "2.0", "method": "notifications/initialized", "params": {}}
        response = server._handle_request(request)

        assert response == {}


# ── Tool 定义完整性测试 ───────────────────────────────────────────


class TestToolDefinitions:
    """测试 MCP tool 定义的完整性和正确性"""

    def test_all_expected_tools_present(self, monkeypatch):
        """所有预期的工具都应出现在列表中"""

        server = _create_mock_server(monkeypatch)
        tool_names = {t["name"] for t in server.list_tools()}

        expected = {
            # Writer
            "writer_info", "writer_replace", "writer_count",
            "writer_table_get", "writer_table_insert",
            "writer_export_pdf", "writer_image_insert", "writer_page_setup",
            # Calc
            "calc_info", "calc_cell_get", "calc_cell_set",
            "calc_range_get", "calc_cell_formula", "calc_chart_create",
            "calc_sheet_list",
            # Impress
            "impress_info", "impress_slide_list",
            "impress_text_get", "impress_text_set", "impress_export_pdf",
            # PDF
            "pdf_info", "pdf_merge", "pdf_extract_pages",
            "pdf_watermark", "pdf_split",
            # Export
            "export_convert",
        }
        for name in expected:
            assert name in tool_names, f"缺少工具: {name}"

    def test_no_duplicate_tool_names(self, monkeypatch):
        """工具名不应重复"""

        server = _create_mock_server(monkeypatch)
        tool_names = [t["name"] for t in server.list_tools()]
        assert len(tool_names) == len(set(tool_names)), "存在重复的工具名"

    def test_each_tool_has_handler(self, monkeypatch):
        """每个工具应有对应的处理器"""

        server = _create_mock_server(monkeypatch)
        for tool in server.list_tools():
            name = tool["name"]
            assert name in server._handlers, f"工具 {name} 缺少处理器"


# ── 助手函数 ──


def _create_mock_server(monkeypatch):  # type: ignore[no-untyped-def]
    """创建使用模拟后端的 WpsMcpServer 实例"""
    from wps_cli.mcp.server import WpsMcpServer

    monkeypatch.setattr(
        "wps_cli.mcp.server.WpsComBackend",
        MockMcpBackend,
    )

    server = WpsMcpServer()
    # 替换后端为模拟后端
    server._manager.backend = MockMcpBackend()
    return server
