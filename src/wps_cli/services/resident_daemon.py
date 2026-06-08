# 设计参考: iOfficeAI/OfficeCLI (Apache 2.0, https://github.com/iOfficeAI/OfficeCLI)
"""驻留模式 Daemon

启动后台 COM 进程，通过 HTTP localhost API 暴露文档操作能力。
连续操作性能提升 5-10x（省略每次的 WPS 进程启动和文档 Open/Close）。
"""

from __future__ import annotations

import json as json_mod
import logging
import secrets
import tempfile
import threading
from dataclasses import dataclass, field
from http.server import BaseHTTPRequestHandler, HTTPServer
from pathlib import Path
from typing import Any

from wps_cli.backends.base import ComBackend
from wps_cli.backends.wps_com import WpsComBackend
from wps_cli.services.calc_service import CalcService
from wps_cli.services.impress_service import ImpressService
from wps_cli.services.pdf_service import PdfService
from wps_cli.services.session_manager import Session, SessionManager
from wps_cli.services.writer_service import WriterService
from wps_cli.utils.path_utils import ensure_safe_input_path

logger = logging.getLogger("wps_cli.resident")

# 模块级 daemon 引用，供 Handler 访问
_daemon: ResidentDaemon | None = None


@dataclass
class ResidentDaemon:
    """驻留进程管理器

    启动 HTTP 服务器，通过 localhost API 暴露文档操作能力。
    所有操作共享同一个 COM 会话池，避免重复启动 WPS 进程。

    端点设计::

        POST /open            {"path": "...", "type": "writer"}  -> {"session_id": "s..."}
        POST /close           {"session_id": "s..."}
        GET  /sessions        -> {"sessions": [...]}
        POST /writer/info     {"session_id": "s..."}   -> {...}
        POST /writer/count    {"session_id": "s..."}   -> {...}
        POST /writer/replace  {"session_id": "s...", "old": "", "new": ""}  -> {...}
        POST /writer/template-fill {"session_id": "s...", "data": {...}}  -> {...}
        POST /calc/cell-get   {"session_id": "s...", "ref": "A1"}  -> {...}
        POST /calc/cell-set   {"session_id": "s...", "ref": "A1", "value": ...}  -> {...}
        POST /calc/range-get  {"session_id": "s...", "ref": "A1:B2"}  -> {...}
        POST /shutdown        -> 停止服务器
    """

    port: int = 9123
    host: str = "127.0.0.1"
    backend: ComBackend = field(default_factory=WpsComBackend)

    def __post_init__(self) -> None:
        self.manager = SessionManager(backend=self.backend)
        self.writer = WriterService(manager=self.manager)
        self.calc = CalcService(manager=self.manager)
        self.impress = ImpressService(manager=self.manager)
        self.pdf = PdfService(manager=self.manager)
        self._opened_docs: dict[str, Session] = {}
        self._lock = threading.RLock()
        self._server: HTTPServer | None = None
        self._auth_token: str = ""  # 启动时生成

    # ── 生命周期 ──

    def start(self) -> None:
        """启动 HTTP 服务器（阻塞当前线程）"""
        global _daemon
        _daemon = self
        self._auth_token = secrets.token_hex(32)  # 256-bit 随机 token
        # 将 token 写入临时文件供 CLI 客户端自动读取
        token_file = Path(tempfile.gettempdir()) / "wps-cli-resident-token"
        token_file.write_text(self._auth_token)
        self._server = HTTPServer((self.host, self.port), ResidentHandler)
        logger.info("驻留服务已启动: http://%s:%d", self.host, self.port)
        print(f"\n  驻留服务已启动 (端口 {self.port})，Token 已保存到 {token_file}\n")
        try:
            self._server.serve_forever()
        except KeyboardInterrupt:
            logger.info("收到中断信号，正在关闭...")
        finally:
            self.stop()

    def stop(self) -> None:
        """停止服务器并关闭所有 COM 进程"""
        global _daemon
        if self._server:
            try:
                self._server.shutdown()
            except Exception:
                pass
        self.manager.stop_all()
        # 清理 token 文件
        try:
            token_file = Path(tempfile.gettempdir()) / "wps-cli-resident-token"
            if token_file.exists():
                token_file.unlink()
        except Exception:
            pass
        _daemon = None
        logger.info("驻留服务已停止")

    # ── 文档生命周期 ──

    def open_document(self, path: str, app_type: str) -> dict:
        """打开文档并创建会话"""
        if app_type not in ("writer", "calc", "impress"):
            raise ValueError(f"不支持的应用类型: {app_type}，可选: writer/calc/impress")
        p = ensure_safe_input_path(path)  # 路径安全校验（在 app_type 校验之后）
        if app_type == "writer":
            session = self.writer.open_document(p)
        elif app_type == "calc":
            session = self._open_calc(p)
        else:
            session = self._open_impress(p)

        with self._lock:
            self._opened_docs[session.session_id] = session
        logger.info("文档已打开: %s [%s] -> %s", p.name, app_type, session.session_id)
        return {"session_id": session.session_id, "app_type": app_type, "path": str(p)}

    def close_document(self, session_id: str) -> dict:
        """关闭文档并释放会话"""
        with self._lock:
            session = self._opened_docs.pop(session_id, None)
        if session is None:
            return {"session_id": session_id, "closed": False, "reason": "会话不存在"}
        self.manager.stop(session_id)
        return {"session_id": session_id, "closed": True}

    def list_active_sessions(self) -> list[dict]:
        """列出活跃会话"""
        with self._lock:
            return [
                {"session_id": s.session_id, "app_type": s.app_type}
                for s in self._opened_docs.values()
            ]

    def get_session(self, session_id: str) -> Session:
        """获取并校验会话"""
        with self._lock:
            session = self._opened_docs.get(session_id)
        if session is None:
            raise ValueError(f"会话不存在: {session_id}")
        return session

    def get_app(self, session_id: str) -> Any:
        """根据 session_id 获取 COM app 对象"""
        return self.get_session(session_id).app

    def _open_calc(self, path: Path) -> Session:
        session = self.manager.start("calc")
        try:
            session.app.Workbooks.Open(str(path), UpdateLinks=0, ReadOnly=False)
            return session
        except Exception:
            self.manager.stop(session.session_id)
            raise

    def _open_impress(self, path: Path) -> Session:
        session = self.manager.start("impress")
        try:
            session.app.Presentations.Open(str(path), ReadOnly=False)
            return session
        except Exception:
            self.manager.stop(session.session_id)
            raise


# ── HTTP 请求处理器 ──


class ResidentHandler(BaseHTTPRequestHandler):
    """驻留模式 HTTP 请求处理器

    每个请求创建新实例，通过模块级 ``_daemon`` 访问 daemon 状态。
    """

    def log_message(self, fmt: str, *args: object) -> None:
        logger.debug("HTTP %s", fmt % args)

    # ── 认证检查 ──

    def _check_auth(self) -> bool:
        """验证请求携带的 Authorization Token 是否匹配"""
        if not _daemon._auth_token:
            return True  # 未启用认证时放行（向后兼容旧测试）
        expected = f"Bearer {_daemon._auth_token}"
        auth_header = self.headers.get("Authorization", "")
        return auth_header == expected

    def _require_auth(self) -> None:
        if not self._check_auth():
            self._respond_error("未授权：缺少或无效的 Authorization Token", 401)

    # ── 路由调度 ──

    def do_GET(self) -> None:
        self._require_auth()
        if self.path == "/sessions":
            self._respond(_daemon.list_active_sessions())
        elif self.path == "/status":
            self._respond({"status": "running", "port": _daemon.port})
        else:
            self._respond_error(f"未知端点: GET {self.path}", 404)

    def do_POST(self) -> None:
        self._require_auth()
        body = self._read_body()

        try:
            result = self._route_post(body)
            self._respond(result)
        except ValueError as e:
            self._respond_error(str(e), 400)
        except Exception as e:
            logger.exception("请求处理异常: %s", self.path)
            self._respond_error(f"{type(e).__name__}: {e}", 500)

    def _route_post(self, body: dict) -> dict:
        path = self.path
        sid = body.get("session_id", "")

        # ── 全局端点 ──
        if path == "/open":
            return _daemon.open_document(
                path=body["path"],
                app_type=body.get("type", "writer"),
            )
        if path == "/close":
            return _daemon.close_document(body["session_id"])
        if path == "/shutdown":
            result = {"status": "shutting_down"}
            threading.Thread(target=_daemon.stop, daemon=True).start()
            return result

        if not sid:
            raise ValueError("缺少 session_id")

        app = _daemon.get_app(sid)

        # ── Writer 端点 ──
        if path == "/writer/info":
            return _daemon.writer.text_count(app)
        if path == "/writer/count":
            return _daemon.writer.text_count(app)
        if path == "/writer/text-get":
            return {"text": _daemon.writer.text_get(app, body.get("start", 0), body.get("end", -1))}
        if path == "/writer/replace":
            return {"replaced": _daemon.writer.text_replace(
                app,
                body["old"],
                body["new"],
                body.get("wildcard", False),
                body.get("case", False),
            )}
        if path == "/writer/template-fill":
            return _daemon.writer.template_fill(app, body["data"])
        if path == "/writer/table-get":
            return {"data": _daemon.writer.table_get(app, body.get("index", 1))}
        if path == "/writer/table-insert":
            return {"table_index": _daemon.writer.table_insert(
                app, body["rows"], body["cols"], body.get("data"),
            )}
        if path == "/writer/save":
            out = body.get("output")
            p = Path(out) if out else None
            return {"path": str(_daemon.writer.save(app, p))}

        # ── Calc 端点 ──
        if path == "/calc/info":
            return _daemon.calc.info(Path(body["path"]))
        if path == "/calc/cell-get":
            return {"value": _daemon.calc.cell_get(app, body["ref"], body.get("sheet"))}
        if path == "/calc/cell-set":
            _daemon.calc.cell_set(app, body["ref"], body["value"], body.get("sheet"))
            return {"status": "ok"}
        if path == "/calc/cell-formula":
            _daemon.calc.cell_formula(app, body["ref"], body["formula"], body.get("sheet"))
            return {"status": "ok"}
        if path == "/calc/range-get":
            return {"data": _daemon.calc.range_get(app, body["ref"], body.get("sheet"))}
        if path == "/calc/range-set":
            _daemon.calc.range_set(app, body["ref"], body["data"], body.get("sheet"))
            return {"status": "ok"}
        if path == "/calc/sheet-list":
            return {"sheets": _daemon.calc.sheet_list(app)}
        if path == "/calc/sheet-add":
            return {"name": _daemon.calc.sheet_add(app, body.get("name"))}
        if path == "/calc/save":
            out = body.get("output")
            p = Path(out) if out else None
            return {"path": str(_daemon.calc.save(app, p))}

        raise ValueError(f"未知端点: POST {path}")

    # ── 请求/响应工具 ──

    # 请求体大小上限（防止 OOM/DoS）
    _MAX_BODY_SIZE: int = 10 * 1024 * 1024  # 10MB

    def _read_body(self) -> dict:
        content_length = int(self.headers.get("Content-Length", 0))
        if content_length == 0:
            return {}
        if content_length > self._MAX_BODY_SIZE:
            raise ValueError(
                f"请求体过大: {content_length} bytes (上限 {self._MAX_BODY_SIZE})"
            )
        raw = self.rfile.read(content_length)
        return json_mod.loads(raw)

    def _respond(self, data: Any, status: int = 200) -> None:
        body = json_mod.dumps(data, ensure_ascii=False, indent=2).encode("utf-8")
        self.send_response(status)
        self.send_header("Content-Type", "application/json; charset=utf-8")
        self.send_header("Content-Length", str(len(body)))
        self.end_headers()
        self.wfile.write(body)

    def _respond_error(self, message: str, status: int) -> None:
        self._respond({"success": False, "error": message}, status)
