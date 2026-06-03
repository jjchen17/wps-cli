"""bridge client 单元测试"""

from __future__ import annotations

import json
import urllib.error
import urllib.request
from unittest.mock import Mock, patch

import pytest

from wps_cli.bridge.client import BridgeClient, BridgeError

# ── Helpers ──────────────────────────────────────────────────


class _FakeResponse:
    def __init__(self, data: dict, status: int = 200) -> None:
        self._data = json.dumps(data).encode()
        self.status = status

    def read(self) -> bytes:
        return self._data

    def __enter__(self) -> "_FakeResponse":
        return self

    def __exit__(self, *args: object) -> None:
        pass


def _fake_urlopen_success(req: urllib.request.Request, timeout: int = 30) -> _FakeResponse:
    return _FakeResponse({"id": "test1", "ok": True, "data": {"key": "value"}})


def _fake_urlopen_error(req: urllib.request.Request, timeout: int = 30) -> _FakeResponse:
    return _FakeResponse({"id": "test1", "ok": False, "error": "WPS not connected"})


# ── Tests ────────────────────────────────────────────────────


class TestBridgeClient:
    """测试 BridgeClient 请求/响应处理"""

    def test_send_success(self) -> None:
        client = BridgeClient()
        client.ensure_server = Mock()  # 跳过 daemon 启动

        with patch("urllib.request.urlopen", _fake_urlopen_success):
            result = client.send("writer.info", {"path": "/tmp/test.docx"})
        assert result == {"key": "value"}

    def test_send_error_response(self) -> None:
        client = BridgeClient()
        client.ensure_server = Mock()

        with patch("urllib.request.urlopen", _fake_urlopen_error):
            with pytest.raises(BridgeError, match="WPS not connected"):
                client.send("writer.info", {"path": "/tmp/test.docx"})

    def test_send_connection_refused(self) -> None:
        client = BridgeClient()
        client.ensure_server = Mock()

        with patch("urllib.request.urlopen", side_effect=OSError("Connection refused")):
            with pytest.raises(BridgeError, match="Connection refused"):
                client.send("writer.info")

    def test_send_invalid_json_response(self) -> None:
        client = BridgeClient()
        client.ensure_server = Mock()

        class BadResponse:
            def read(self) -> bytes:
                return b"not json"
            def __enter__(self) -> "BadResponse":
                return self
            def __exit__(self, *args: object) -> None:
                pass

        with patch("urllib.request.urlopen", return_value=BadResponse()):
            with pytest.raises(BridgeError, match="Invalid bridge server response"):
                client.send("writer.info")

    def test_send_uses_unique_request_ids(self) -> None:
        """每次 send 应该生成不同的 request id"""
        client = BridgeClient()
        client.ensure_server = Mock()

        captured = []

        def capture(req: urllib.request.Request, timeout: int = 30) -> _FakeResponse:
            captured.append(json.loads(req.data))
            return _FakeResponse({"id": captured[-1]["id"], "ok": True, "data": {}})

        with patch("urllib.request.urlopen", capture):
            client.send("writer.info")
            client.send("calc.cell_get")

        assert len(captured) == 2
        assert captured[0]["id"] != captured[1]["id"]
        assert captured[0]["method"] == "writer.info"
        assert captured[1]["method"] == "calc.cell_get"

    def test_send_passes_params(self) -> None:
        """确保 params 被正确序列化"""
        client = BridgeClient()
        client.ensure_server = Mock()

        captured: dict = {}

        def capture(req: urllib.request.Request, timeout: int = 30) -> _FakeResponse:
            nonlocal captured
            captured = json.loads(req.data)
            return _FakeResponse({"id": captured["id"], "ok": True, "data": {}})

        with patch("urllib.request.urlopen", capture):
            client.send("writer.replace", {"path": "/f.docx", "old": "x", "new": "y"})

        assert captured["params"]["path"] == "/f.docx"
        assert captured["params"]["old"] == "x"


class TestBridgeClientTimeout:
    """超时测试"""

    def test_send_http_timeout(self) -> None:
        client = BridgeClient()
        client.ensure_server = Mock()

        with patch("urllib.request.urlopen", side_effect=urllib.error.URLError("timed out")):
            with pytest.raises(BridgeError, match="timed out"):
                client.send("writer.info", timeout=1)
