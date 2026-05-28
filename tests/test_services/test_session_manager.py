"""会话管理器测试"""

from tests.conftest import MockComBackend
from wps_cli.services.session_manager import SessionManager


class TestSessionManager:
    """测试会话管理"""

    def test_start_session(self):
        mgr = SessionManager(backend=MockComBackend())
        session = mgr.start("writer")
        assert session.session_id.startswith("s")
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

    def test_session_id_is_unique(self):
        mgr = SessionManager(backend=MockComBackend())
        ids = {mgr.start("writer").session_id for _ in range(20)}
        assert len(ids) == 20
        mgr.stop_all()

    def test_context_manager(self):
        mgr = SessionManager(backend=MockComBackend())
        with mgr.session("writer") as app:
            assert app.Name == "Mock writer"
        assert len(mgr.list_sessions()) == 0

    def test_backend_harden_sets_automation_security(self):
        mgr = SessionManager(backend=MockComBackend())
        session = mgr.start("writer")
        # MSO_AUTOMATION_SECURITY_FORCE_DISABLE = 3
        assert session.app.AutomationSecurity == 3
        assert session.app.DisplayAlerts is False
        mgr.stop_all()

    def test_manager_as_context_manager(self):
        with SessionManager(backend=MockComBackend()) as mgr:
            mgr.start("writer")
            mgr.start("calc")
            assert len(mgr.list_sessions()) == 2
        assert len(mgr.list_sessions()) == 0
