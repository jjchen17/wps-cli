"""WPS 进程生命周期管理

线程安全说明：``SessionManager`` 内部使用 ``threading.RLock`` 保护
``_sessions`` 与 ``_counter``，可以在多线程环境复用，但同一会话对象不应
跨线程并发使用（COM 调用本身是 STA 模型）。
"""

from __future__ import annotations

import threading
import uuid
from collections.abc import Generator
from contextlib import contextmanager
from dataclasses import dataclass, field
from typing import Any

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
    _lock: threading.RLock = field(default_factory=threading.RLock)

    def start(self, app_type: str) -> Session:
        """启动新的 WPS 应用实例"""
        app = self.backend.connect(app_type)
        try:
            app.Visible = False
        except Exception:
            pass
        session_id = f"s{uuid.uuid4().hex[:12]}"
        session = Session(
            session_id=session_id,
            app_type=app_type,
            app=app,
            backend=self.backend,
        )
        with self._lock:
            self._sessions[session_id] = session
        return session

    def stop(self, session_id: str) -> None:
        """停止指定会话"""
        with self._lock:
            session = self._sessions.pop(session_id, None)
        if session:
            self.backend.disconnect(session.app)

    def stop_all(self) -> None:
        """停止所有会话"""
        with self._lock:
            ids = list(self._sessions)
        for sid in ids:
            self.stop(sid)

    def get(self, session_id: str) -> Session | None:
        """获取指定会话"""
        with self._lock:
            return self._sessions.get(session_id)

    def list_sessions(self) -> list[Session]:
        """列出所有活跃会话"""
        with self._lock:
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

    def __enter__(self) -> SessionManager:
        return self

    def __exit__(self, exc_type, exc, tb) -> None:
        self.stop_all()

    def __del__(self) -> None:
        try:
            self.stop_all()
        except Exception:
            pass
