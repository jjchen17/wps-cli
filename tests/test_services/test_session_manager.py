"""会话管理器测试"""

from wps_cli.services.session_manager import SessionManager
from tests.conftest import MockComBackend


class TestSessionManager:
    """测试会话管理"""

    def test_start_session(self):
        mgr = SessionManager(backend=MockComBackend())
        session = mgr.start("writer")
        assert session.session_id == "s1"
        assert session.app_type == "writer"
        mgr.stop_all()

    def test_list_sessions(self):
        mgr = SessionManager(backend=MockComBackend())
        mgr.start("writer")
        mgr.start("calc")
        sessions = mgr.list_sessions()
        assert len(sessions) == 2
        mgr.stop_all()

    def test_stop_session(self):
        mgr = SessionManager(backend=MockComBackend())
        session = mgr.start("writer")
        mgr.stop(session.session_id)
        assert mgr.get(session.session_id) is None

    def test_stop_all(self):
        mgr = SessionManager(backend=MockComBackend())
        mgr.start("writer")
        mgr.start("calc")
        mgr.stop_all()
        assert len(mgr.list_sessions()) == 0

    def test_context_manager(self):
        mgr = SessionManager(backend=MockComBackend())
        with mgr.session("writer") as app:
            assert app.Name == "Mock writer"
        assert len(mgr.list_sessions()) == 0
