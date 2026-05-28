"""WPS 进程生命周期管理"""

from contextlib import contextmanager
from dataclasses import dataclass, field
from typing import Any, Generator

from wps_cli.backends.base import ComBackend


@dataclass
class Session:
    """活跃会话"""

    session_id: str
    app_type: str
    app: Any
    backend: ComBackend


@dataclass
class SessionManager:
    """管理多个 WPS 应用实例的生命周期"""

    backend: ComBackend
    _sessions: dict[str, Session] = field(default_factory=dict)
    _counter: int = 0

    def start(self, app_type: str) -> Session:
        """启动新的 WPS 应用实例"""
        app = self.backend.connect(app_type)
        app.Visible = False  # 后台运行
        self._counter += 1
        session_id = f"s{self._counter}"
        session = Session(
            session_id=session_id,
            app_type=app_type,
            app=app,
            backend=self.backend,
        )
        self._sessions[session_id] = session
        return session

    def stop(self, session_id: str) -> None:
        """停止指定会话"""
        session = self._sessions.pop(session_id, None)
        if session:
            self.backend.disconnect(session.app)

    def stop_all(self) -> None:
        """停止所有会话"""
        for sid in list(self._sessions):
            self.stop(sid)

    def get(self, session_id: str) -> Session | None:
        """获取指定会话"""
        return self._sessions.get(session_id)

    def list_sessions(self) -> list[Session]:
        """列出所有活跃会话"""
        return list(self._sessions.values())

    @contextmanager
    def session(self, app_type: str) -> Generator[object, None, None]:
        """上下文管理器，自动启动和关闭 WPS 实例

        Usage:
            with manager.session("writer") as app:
                doc = app.Documents.Add()
        """
        session = self.start(app_type)
        try:
            yield session.app
        finally:
            self.stop(session.session_id)

    def __del__(self) -> None:
        self.stop_all()
