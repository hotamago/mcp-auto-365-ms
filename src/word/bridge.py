"""Local bridge server for Word Companion Add-in.

Provides:
- HTTPS server on ``127.0.0.1:<port>`` serving the taskpane UI and add-in assets.
- REST endpoints for Word add-in session registration, heartbeat and job polling.
- Thread-safe job queue and event barrier for MCP tools to execute live Word comments.
"""

from __future__ import annotations

import http.server
import json
import logging
import queue
import socketserver
import ssl
import threading
import time
import urllib.parse
import uuid
from dataclasses import dataclass, field
from pathlib import Path
from typing import Any

from common.errors import Mcp365Error
from word.certs import get_or_create_dev_certificate

logger = logging.getLogger(__name__)

_ADDIN_DIR = Path(__file__).resolve().parent / "addin"


@dataclass
class CommentJob:
    job_id: str
    session_id: str
    action: str
    comments: list[dict[str, str]]
    author: str
    created_at: float = field(default_factory=time.time)
    event: threading.Event = field(default_factory=threading.Event)
    status: str = "pending"
    results: list[dict[str, Any]] = field(default_factory=list)
    error: str = ""


@dataclass
class WordSession:
    session_id: str
    token: str
    doc_url: str
    doc_title: str
    platform: str
    client_id: str
    connected_at: float = field(default_factory=time.time)
    last_heartbeat: float = field(default_factory=time.time)
    job_queue: queue.Queue[CommentJob] = field(default_factory=queue.Queue)
    jobs: dict[str, CommentJob] = field(default_factory=dict)

    def is_alive(self, max_idle: float = 45.0) -> bool:
        return (time.time() - self.last_heartbeat) < max_idle


class _ThreadedHTTPServer(socketserver.ThreadingMixIn, http.server.HTTPServer):
    daemon_threads = True


class _BridgeRequestHandler(http.server.BaseHTTPRequestHandler):
    server: _BridgeServerWrapper

    def log_message(self, format: str, *args: Any) -> None:
        # Suppress noisy standard HTTP access logs
        pass

    def _send_cors_headers(self) -> None:
        self.send_header("Access-Control-Allow-Origin", "*")
        self.send_header("Access-Control-Allow-Methods", "GET, POST, OPTIONS")
        self.send_header("Access-Control-Allow-Headers", "Content-Type, Authorization, X-Session-Token")
        self.send_header("Cache-Control", "no-store, no-cache, must-revalidate")

    def do_OPTIONS(self) -> None:
        self.send_response(204)
        self._send_cors_headers()
        self.end_headers()

    def _send_json(self, status: int, data: Any) -> None:
        payload = json.dumps(data, ensure_ascii=False).encode("utf-8")
        self.send_response(status)
        self._send_cors_headers()
        self.send_header("Content-Type", "application/json; charset=utf-8")
        self.send_header("Content-Length", str(len(payload)))
        self.end_headers()
        self.wfile.write(payload)

    def _send_file(self, path: Path, content_type: str) -> None:
        if not path.is_file():
            self.send_response(404)
            self._send_cors_headers()
            self.end_headers()
            self.wfile.write(b"Not Found")
            return
        data = path.read_bytes()
        self.send_response(200)
        self._send_cors_headers()
        self.send_header("Content-Type", content_type)
        self.send_header("Content-Length", str(len(data)))
        self.end_headers()
        self.wfile.write(data)

    def do_GET(self) -> None:
        parsed = urllib.parse.urlparse(self.path)
        path = parsed.path.rstrip("/")
        query = urllib.parse.parse_qs(parsed.query)

        if path in ("", "/"):
            self._serve_welcome_page()
            return
        if path == "/taskpane.html":
            self._send_file(_ADDIN_DIR / "taskpane.html", "text/html; charset=utf-8")
            return
        if path == "/taskpane.js":
            self._send_file(_ADDIN_DIR / "taskpane.js", "application/javascript; charset=utf-8")
            return
        if path == "/taskpane.css":
            self._send_file(_ADDIN_DIR / "taskpane.css", "text/css; charset=utf-8")
            return
        if path == "/icon.png":
            self._send_file(_ADDIN_DIR / "icon.png", "image/png")
            return
        if path == "/manifest.xml":
            self._serve_manifest()
            return
        if path in ("/cert.crt", "/ca.crt"):
            cert_path = self.server.cert_path
            if cert_path and cert_path.is_file():
                self._send_file(cert_path, "application/x-x509-ca-cert")
            else:
                self.send_response(404)
                self.end_headers()
            return

        if path == "/api/status":
            self._send_json(200, self.server.bridge.get_status_summary())
            return
        if path == "/api/sessions":
            self._send_json(200, self.server.bridge.get_active_sessions_info())
            return

        # Polling for jobs: /api/sessions/<id>/jobs/poll
        if path.startswith("/api/sessions/") and path.endswith("/jobs/poll"):
            parts = path.split("/")
            if len(parts) == 6:
                session_id = parts[3]
                self._handle_job_poll(session_id, query)
                return

        self.send_response(404)
        self._send_cors_headers()
        self.end_headers()
        self.wfile.write(b"Not Found")

    def _serve_manifest(self) -> None:
        manifest_file = _ADDIN_DIR / "manifest.xml"
        if not manifest_file.is_file():
            self.send_response(404)
            self.end_headers()
            return
        text = manifest_file.read_text(encoding="utf-8")
        # Ensure host and port match the current server
        proto = "https" if self.server.ssl_enabled else "http"
        port_str = f":{self.server.port}" if self.server.port not in (80, 443) else ""
        base_url = f"{proto}://{self.server.host}{port_str}"
        text = text.replace("https://127.0.0.1:3650", base_url)
        text = text.replace("https://localhost:3650", base_url)
        data = text.encode("utf-8")

        self.send_response(200)
        self._send_cors_headers()
        self.send_header("Content-Type", "application/xml; charset=utf-8")
        self.send_header("Content-Disposition", 'attachment; filename="mcp-auto-365-word-manifest.xml"')
        self.send_header("Content-Length", str(len(data)))
        self.end_headers()
        self.wfile.write(data)

    def _serve_welcome_page(self) -> None:
        status = self.server.bridge.get_status_summary()
        sessions = self.server.bridge.get_active_sessions_info()
        sessions_html = "".join(
            f"<li><b>{s['doc_title']}</b> — <code>{s['doc_url'][:70]}...</code> (kết nối {s['last_seen_ago_s']}s trước)</li>"
            for s in sessions
        ) or "<li><i>Chưa có tài liệu nào kết nối. Mở Word và add-in để kết nối.</i></li>"

        html = f"""<!DOCTYPE html>
<html lang="vi">
<head>
  <meta charset="UTF-8">
  <title>MCP Auto 365 MS - Word Companion Bridge</title>
  <style>
    body {{ font-family: -apple-system, BlinkMacSystemFont, "Segoe UI", Roboto, sans-serif; max-width: 720px; margin: 40px auto; padding: 0 20px; line-height: 1.6; color: #1f2937; background: #f9fafb; }}
    .card {{ background: #fff; border: 1px solid #e5e7eb; border-radius: 8px; padding: 20px; margin-bottom: 20px; box-shadow: 0 1px 3px rgba(0,0,0,0.05); }}
    h1 {{ font-size: 20px; margin-top: 0; }}
    h2 {{ font-size: 15px; border-bottom: 1px solid #e5e7eb; padding-bottom: 6px; }}
    .badge {{ display: inline-block; padding: 2px 8px; border-radius: 9999px; background: #d1fae5; color: #065f46; font-size: 12px; font-weight: 600; }}
    a.btn {{ display: inline-block; background: #2563eb; color: #fff; text-decoration: none; padding: 8px 14px; border-radius: 6px; font-size: 13px; font-weight: 500; margin-right: 8px; }}
    a.btn-secondary {{ background: #f3f4f6; color: #1f2937; border: 1px solid #d1d5db; }}
    code {{ background: #f3f4f6; padding: 2px 5px; border-radius: 4px; font-size: 12px; }}
  </style>
</head>
<body>
  <div class="card">
    <h1>📝 Word Companion Bridge <span class="badge">Đang chạy</span></h1>
    <p>Bridge local đang hoạt động tại <code>{status['base_url']}</code>.</p>
    <p>
      <a href="/manifest.xml" class="btn" download>Tải file Manifest XML</a>
      <a href="/cert.crt" class="btn btn-secondary" download>Tải Chứng chỉ Dev SSL</a>
      <a href="/taskpane.html" class="btn btn-secondary" target="_blank">Mở Task Pane trực tiếp</a>
    </p>
  </div>

  <div class="card">
    <h2>Tài liệu Word đang kết nối ({len(sessions)})</h2>
    <ul>{sessions_html}</ul>
  </div>

  <div class="card">
    <h2>Hướng dẫn cài đặt (Sideload vào Word Online)</h2>
    <ol>
      <li>Mở tài liệu Word trên trình duyệt (Word Online / SharePoint).</li>
      <li>Vào tab <b>Home</b> &gt; chọn <b>Add-ins</b> &gt; chọn <b>More Settings</b> (hoặc <b>Upload My Add-in</b>).</li>
      <li>Chọn <b>Upload My Add-in</b> &gt; tải lên file <code>mcp-auto-365-word-manifest.xml</code> vừa tải ở trên.</li>
      <li>Task pane <b>Auto 365 Companion</b> sẽ xuất hiện và tự động kết nối với coding agent.</li>
    </ol>
  </div>
</body>
</html>"""
        data = html.encode("utf-8")
        self.send_response(200)
        self._send_cors_headers()
        self.send_header("Content-Type", "text/html; charset=utf-8")
        self.send_header("Content-Length", str(len(data)))
        self.end_headers()
        self.wfile.write(data)

    def do_POST(self) -> None:
        parsed = urllib.parse.urlparse(self.path)
        path = parsed.path.rstrip("/")

        content_len = int(self.headers.get("Content-Length", "0"))
        body_bytes = self.rfile.read(content_len) if content_len > 0 else b"{}"
        try:
            body = json.loads(body_bytes.decode("utf-8")) if body_bytes else {}
        except Exception:
            body = {}

        if path == "/api/sessions/register":
            session = self.server.bridge.register_session(
                doc_url=str(body.get("doc_url", "")),
                doc_title=str(body.get("doc_title", "")),
                platform=str(body.get("platform", "web")),
                client_id=str(body.get("client_id", "")),
            )
            self._send_json(200, {"session_id": session.session_id, "token": session.token})
            return

        if path.startswith("/api/sessions/") and path.endswith("/heartbeat"):
            parts = path.split("/")
            if len(parts) == 5:
                session_id = parts[3]
                ok = self.server.bridge.heartbeat_session(
                    session_id,
                    doc_url=str(body.get("doc_url", "")),
                    doc_title=str(body.get("doc_title", "")),
                )
                if ok:
                    self._send_json(200, {"ok": True})
                else:
                    self._send_json(404, {"error": "Session not found or expired"})
                return

        if path.startswith("/api/sessions/") and "/jobs/" in path and path.endswith("/result"):
            parts = path.split("/")
            # /api/sessions/<session_id>/jobs/<job_id>/result
            if len(parts) == 7:
                session_id = parts[3]
                job_id = parts[5]
                ok = self.server.bridge.record_job_result(
                    session_id,
                    job_id,
                    status=str(body.get("status", "ok")),
                    results=body.get("results", []),
                    error=str(body.get("error", "")),
                )
                if ok:
                    self._send_json(200, {"ok": True})
                else:
                    self._send_json(404, {"error": "Job or session not found"})
                return

        self.send_response(404)
        self._send_cors_headers()
        self.end_headers()
        self.wfile.write(b"Not Found")

    def _handle_job_poll(self, session_id: str, query: dict[str, list[str]]) -> None:
        session = self.server.bridge.get_session(session_id)
        if not session:
            self._send_json(404, {"error": "Session not found"})
            return

        timeout = 15.0
        if "timeout" in query and query["timeout"]:
            try:
                timeout = min(30.0, max(1.0, float(query["timeout"][0])))
            except ValueError:
                pass

        try:
            job = session.job_queue.get(block=True, timeout=timeout)
            self._send_json(200, {
                "job_id": job.job_id,
                "action": job.action,
                "comments": job.comments,
                "author": job.author,
            })
        except queue.Empty:
            self.send_response(204)
            self._send_cors_headers()
            self.end_headers()


class _BridgeServerWrapper(_ThreadedHTTPServer):
    bridge: WordBridge
    port: int
    host: str
    ssl_enabled: bool
    cert_path: Path | None = None


class WordBridge:
    """Manages the Word Companion background server and active sessions."""

    def __init__(self) -> None:
        self._lock = threading.Lock()
        self._sessions: dict[str, WordSession] = {}
        self._server: _BridgeServerWrapper | None = None
        self._server_thread: threading.Thread | None = None
        self._start_time: float = 0.0
        self._port: int = 3650
        self._host: str = "127.0.0.1"
        self._ssl_enabled: bool = True
        self._cert_path: Path | None = None
        self._key_path: Path | None = None

    @property
    def is_running(self) -> bool:
        return self._server is not None and self._server_thread is not None and self._server_thread.is_alive()

    @property
    def port(self) -> int:
        return self._port

    @property
    def base_url(self) -> str:
        proto = "https" if self._ssl_enabled else "http"
        port_str = f":{self._port}" if self._port not in (80, 443) else ""
        return f"{proto}://{self._host}{port_str}"

    def ensure_running(
        self,
        host: str = "127.0.0.1",
        port: int = 3650,
        ssl_enabled: bool = True,
        cert_file: str = "",
        key_file: str = "",
    ) -> WordBridge:
        with self._lock:
            if self.is_running:
                return self

            self._host = host
            self._port = port
            self._ssl_enabled = ssl_enabled

            if ssl_enabled:
                if cert_file and key_file and Path(cert_file).is_file() and Path(key_file).is_file():
                    self._cert_path = Path(cert_file)
                    self._key_path = Path(key_file)
                else:
                    self._cert_path, self._key_path = get_or_create_dev_certificate()

            try:
                server = _BridgeServerWrapper((host, port), _BridgeRequestHandler)
            except OSError as err:
                # If configured port is busy, attempt fallback if default
                if port == 3650:
                    try:
                        port = 3651
                        self._port = port
                        server = _BridgeServerWrapper((host, port), _BridgeRequestHandler)
                    except OSError:
                        raise Mcp365Error(
                            f"Không thể mở cổng {self._port} cho Word Companion Bridge: {err}",
                            "Kiểm tra xem có tiến trình nào đang chiếm cổng hoặc đổi `port` trong [word] config.toml.",
                        ) from err
                else:
                    raise Mcp365Error(
                        f"Không thể mở cổng {self._port} cho Word Companion Bridge: {err}",
                        "Kiểm tra cổng hoặc đổi `port` trong [word] config.toml.",
                    ) from err

            server.bridge = self
            server.port = self._port
            server.host = self._host
            server.ssl_enabled = self._ssl_enabled
            server.cert_path = self._cert_path

            if ssl_enabled and self._cert_path and self._key_path:
                ssl_ctx = ssl.SSLContext(ssl.PROTOCOL_TLS_SERVER)
                ssl_ctx.load_cert_chain(certfile=str(self._cert_path), keyfile=str(self._key_path))
                server.socket = ssl_ctx.wrap_socket(server.socket, server_side=True)

            self._server = server
            self._start_time = time.time()
            thread = threading.Thread(target=server.serve_forever, daemon=True, name="WordBridgeServer")
            thread.start()
            self._server_thread = thread

            logger.info("Word Companion Bridge started at %s", self.base_url)
            return self

    def stop(self) -> None:
        with self._lock:
            if self._server:
                self._server.shutdown()
                self._server.server_close()
                self._server = None
                self._server_thread = None
                logger.info("Word Companion Bridge stopped")

    def register_session(self, doc_url: str, doc_title: str, platform: str, client_id: str) -> WordSession:
        with self._lock:
            session_id = f"w_{uuid.uuid4().hex[:8]}"
            token = uuid.uuid4().hex
            session = WordSession(
                session_id=session_id,
                token=token,
                doc_url=doc_url,
                doc_title=doc_title or "Document.docx",
                platform=platform,
                client_id=client_id,
            )
            self._sessions[session_id] = session
            return session

    def heartbeat_session(self, session_id: str, doc_url: str = "", doc_title: str = "") -> bool:
        with self._lock:
            session = self._sessions.get(session_id)
            if not session:
                return False
            session.last_heartbeat = time.time()
            if doc_url:
                session.doc_url = doc_url
            if doc_title:
                session.doc_title = doc_title
            return True

    def get_session(self, session_id: str) -> WordSession | None:
        with self._lock:
            return self._sessions.get(session_id)

    def get_active_sessions(self) -> list[WordSession]:
        with self._lock:
            now = time.time()
            return [s for s in self._sessions.values() if (now - s.last_heartbeat) < 45.0]

    def get_active_sessions_info(self) -> list[dict[str, Any]]:
        active = self.get_active_sessions()
        now = time.time()
        return [
            {
                "session_id": s.session_id,
                "doc_title": s.doc_title,
                "doc_url": s.doc_url,
                "platform": s.platform,
                "connected_seconds_ago": int(now - s.connected_at),
                "last_seen_ago_s": int(now - s.last_heartbeat),
            }
            for s in active
        ]

    def get_status_summary(self) -> dict[str, Any]:
        with self._lock:
            active = [s for s in self._sessions.values() if s.is_alive()]
            uptime = (time.time() - self._start_time) if self._start_time else 0.0
            return {
                "status": "running" if self.is_running else "stopped",
                "base_url": self.base_url,
                "port": self._port,
                "ssl": self._ssl_enabled,
                "uptime": uptime,
                "active_sessions": len(active),
            }

    def find_session_for_document(self, file_url_or_guid: str, item_name: str = "") -> WordSession | None:
        """Find a live connected Word session matching the file name or URL."""
        active = self.get_active_sessions()
        if not active:
            return None

        # 1. Exact or partial URL match
        norm_query = urllib.parse.unquote(file_url_or_guid).lower()
        query_filename = item_name.lower() or Path(norm_query.split("?")[0]).name.lower()

        # If file name contains something like .docx, try matching
        for s in active:
            s_url = urllib.parse.unquote(s.doc_url).lower()
            s_title = s.doc_title.lower()

            if query_filename and (query_filename in s_title or query_filename in s_url):
                return s
            if norm_query and (norm_query in s_url or (len(norm_query) > 15 and norm_query in s_url)):
                return s

        # 2. If exactly one session is active and query_filename is not completely different
        if len(active) == 1:
            return active[0]

        return None

    def record_job_result(
        self,
        session_id: str,
        job_id: str,
        status: str,
        results: list[dict[str, Any]],
        error: str = "",
    ) -> bool:
        with self._lock:
            session = self._sessions.get(session_id)
            if not session:
                return False
            job = session.jobs.get(job_id)
            if not job:
                return False
            job.status = status
            job.results = results
            job.error = error
            job.event.set()
            return True

    def execute_comments_live(
        self,
        session: WordSession,
        comments: list[dict[str, str]],
        author: str = "",
        timeout: float = 35.0,
    ) -> list[dict[str, Any]]:
        """Submit an add_comments job to the live Word session and wait for result."""
        job_id = f"job_{uuid.uuid4().hex[:8]}"
        job = CommentJob(
            job_id=job_id,
            session_id=session.session_id,
            action="add_comments",
            comments=comments,
            author=author,
        )
        with self._lock:
            session.jobs[job_id] = job
            session.job_queue.put(job)

        # Wait for add-in to finish
        finished = job.event.wait(timeout=timeout)
        if not finished:
            raise Mcp365Error(
                f"Word Companion Add-in không phản hồi sau {int(timeout)} giây.",
                "Kiểm tra xem taskpane add-in trong Word có đang mở không và bấm 'Làm mới kết nối'.",
            )

        if job.status != "ok" and job.error:
            raise Mcp365Error(
                f"Lỗi khi thêm comment qua Word Add-in: {job.error}",
                "Kiểm tra nội dung tài liệu và quyền chỉnh sửa của bạn trong Word.",
            )

        return job.results


_global_bridge: WordBridge | None = None


def get_bridge() -> WordBridge:
    """Return the global WordBridge instance."""
    global _global_bridge
    if _global_bridge is None:
        _global_bridge = WordBridge()
    return _global_bridge
