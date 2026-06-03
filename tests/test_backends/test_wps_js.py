"""WPS JS 后端测试"""

from __future__ import annotations

from unittest.mock import Mock, patch

import pytest

from wps_cli.backends.wps_js import WpsJsBackend
from wps_cli.bridge.client import BridgeError
from wps_cli.exceptions import WpsNotFoundError


class TestWpsJsBackend:
    """测试 WpsJsBackend 基本接口"""

    def test_connect_returns_client(self) -> None:
        backend = WpsJsBackend()
        backend._client.ensure_server = Mock()

        client = backend.connect("writer")
        assert client is backend._client

    def test_connect_invalid_app_type(self) -> None:
        backend = WpsJsBackend()
        with pytest.raises(ValueError, match="不支持的应用类型"):
            backend.connect("photoshop")

    def test_connect_bridge_unavailable(self) -> None:
        backend = WpsJsBackend()
        backend._client.ensure_server = Mock(
            side_effect=BridgeError("Cannot reach bridge")
        )
        with pytest.raises(WpsNotFoundError, match="writer"):
            backend.connect("writer")

    def test_is_alive_true(self) -> None:
        backend = WpsJsBackend()
        backend._client.ensure_server = Mock()
        assert backend.is_alive(None) is True

    def test_is_alive_false(self) -> None:
        backend = WpsJsBackend()
        backend._client.ensure_server = Mock(
            side_effect=BridgeError("down")
        )
        assert backend.is_alive(None) is False

    def test_disconnect_noop(self) -> None:
        backend = WpsJsBackend()
        backend.disconnect(None)  # 不应抛异常

    def test_get_version_returns_string(self) -> None:
        backend = WpsJsBackend()
        backend._client.ensure_server = Mock()
        with patch.object(backend._client, "send", return_value={"version": "12.1.2"}):
            assert backend.get_version(None) == "12.1.2"

    def test_get_version_unknown_on_error(self) -> None:
        backend = WpsJsBackend()
        backend._client.ensure_server = Mock()
        with patch.object(backend._client, "send", side_effect=BridgeError("err")):
            assert backend.get_version(None) == "unknown"
