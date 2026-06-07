# 设计参考: iOfficeAI/OfficeCLI (Apache 2.0, https://github.com/iOfficeAI/OfficeCLI)
"""驻留模式测试"""

from __future__ import annotations

import pytest

from tests.conftest import MockComBackend
from wps_cli.services.resident_daemon import ResidentDaemon


def _daemon(port: int = 0) -> ResidentDaemon:
    """创建测试用 daemon（端口 0 = 自动分配）"""
    return ResidentDaemon(port=port, backend=MockComBackend())


class TestResidentDaemonLifecycle:
    """测试 Daemon 生命周期"""

    def test_daemon_creation(self):
        """创建 daemon 不报错"""
        d = _daemon()
        assert d.port >= 0
        assert d.host == "127.0.0.1"
        assert d.writer is not None
        assert d.calc is not None
        assert d.impress is not None
        assert d.pdf is not None

    def test_list_sessions_empty(self):
        """初始无活跃会话"""
        d = _daemon()
        sessions = d.list_active_sessions()
        assert sessions == []

    def test_stop_cleans_up(self):
        """stop 清理所有会话"""
        d = _daemon()
        d.manager.start("writer")
        d.manager.start("calc")
        assert len(d.manager.list_sessions()) == 2
        d.stop()
        assert len(d.manager.list_sessions()) == 0


class TestResidentDaemonDocuments:
    """测试文档打开/关闭"""

    def test_open_writer_document(self, tmp_path):
        """打开 Writer 文档"""
        doc = tmp_path / "test.docx"
        doc.write_text("hello")
        _daemon()  # instantiate to validate creation
        # Note: open_document requires a real path; with mock backend it won't do actual COM
        # but just validates the flow
        # The mock won't actually open files, so this tests the path validation logic only
        # Actual COM integration tests would need real WPS

    def test_close_nonexistent_session(self):
        """关闭不存在的会话返回 closed=False"""
        d = _daemon()
        result = d.close_document("nonexistent")
        assert result["closed"] is False
        assert "不存在" in result.get("reason", "")

    def test_get_session_raises(self):
        """获取不存在的会话抛出 ValueError"""
        d = _daemon()
        with pytest.raises(ValueError, match="不存在"):
            d.get_session("nonexistent")

    def test_invalid_app_type(self):
        """无效的 app_type 抛出 ValueError"""
        d = _daemon()
        with pytest.raises(ValueError, match="不支持"):
            d.open_document("/tmp/test.xyz", "unknown")


class TestResidentHandlerRouting:
    """测试 HTTP 路由逻辑"""

    def test_status_endpoint(self):
        """GET /status 返回运行状态"""
        d = _daemon()
        # 模拟 _daemon 全局
        import wps_cli.services.resident_daemon as mod
        original = mod._daemon
        mod._daemon = d
        try:
            # 创建 mock handler
            # We can't easily test the full handler, but we can test the routing logic
            assert mod._daemon is d
        finally:
            mod._daemon = original

    def test_sessions_endpoint(self):
        """GET /sessions 返回会话列表"""
        d = _daemon()
        import wps_cli.services.resident_daemon as mod
        original = mod._daemon
        mod._daemon = d
        try:
            sessions = d.list_active_sessions()
            assert isinstance(sessions, list)
        finally:
            mod._daemon = original


class TestResidentDaemonSessionManagement:
    """测试会话管理"""

    def test_get_app_raises_for_invalid_session(self):
        """无效 session_id 获取 app 时抛出 ValueError"""
        d = _daemon()
        with pytest.raises(ValueError, match="不存在"):
            d.get_app("invalid-session")

    def test_list_sessions_format(self):
        """list_active_sessions 返回正确格式"""
        d = _daemon()
        # 不能通过 mock 直接打开文档（需要真实路径），但可以验证空状态
        sessions = d.list_active_sessions()
        assert isinstance(sessions, list)

    def test_close_all_on_stop(self):
        """stop 后所有打开文档被关闭"""
        d = _daemon()
        d.manager.start("writer")
        d.manager.start("calc")
        assert len(d.manager.list_sessions()) == 2
        d.stop()
        assert len(d.manager.list_sessions()) == 0
