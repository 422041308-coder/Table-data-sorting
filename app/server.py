# -*- coding: utf-8 -*-
"""HTTP 服务骨架（Lead 维护，契约冻结）。

* 只依赖标准库；
* 路由表由 ``app/api.py`` 在启动时注册；
* 静态资源来自 ``app/web/``；
* 长任务与下载令牌在此统一管理。

启动：``python -m app.server`` 或 ``python app/server.py``（见 启动.bat）。
"""
from __future__ import annotations

import json
import mimetypes
import os
import re
import sys
import threading
import time
import traceback
import uuid
import webbrowser
from http.server import BaseHTTPRequestHandler, ThreadingHTTPServer
from urllib.parse import parse_qs, unquote, urlparse

APP_DIR = os.path.dirname(os.path.abspath(__file__))
# 允许通过环境变量指定"程序根目录"（打包成单文件 EXE 时，data/ 需要放到 exe 旁边）
PROJECT_ROOT = os.environ.get("EXCEL_TOOLKIT_HOME") or os.path.dirname(APP_DIR)
WEB_DIR = os.path.join(APP_DIR, "web")
DATA_DIR = os.path.join(PROJECT_ROOT, "data")

DEFAULT_PORT = 8765
MAX_BODY = 64 * 1024 * 1024  # 64 MB，上传模板用


class ApiError(Exception):
    """业务异常，带 HTTP 状态码与稳定错误码。"""

    def __init__(self, code: str, message: str, status: int = 400, detail=None):
        super().__init__(message)
        self.code = code
        self.message = message
        self.status = status
        self.detail = detail


class FileResponse(Exception):
    """handler 抛出它表示"直接回一个文件"。"""

    def __init__(self, data: bytes, filename: str):
        super().__init__(filename)
        self.data = data
        self.filename = filename


class Task:
    def __init__(self, task_id: str, kind: str):
        self.id = task_id
        self.kind = kind
        self.status = "running"  # running | done | error
        self.progress = 0
        self.message = ""
        self.result = None
        self.error = ""
        self.created = time.time()

    def to_dict(self) -> dict:
        return {
            "id": self.id,
            "kind": self.kind,
            "status": self.status,
            "progress": self.progress,
            "message": self.message,
            "result": self.result,
            "error": self.error,
        }


class App:
    """路由 + 任务 + 下载令牌。"""

    def __init__(self, root: str):
        self.root = root
        self.routes: list[tuple[str, re.Pattern, object]] = []
        self.tasks: dict[str, Task] = {}
        self.downloads: dict[str, str] = {}
        self.tokens_by_task: dict[str, str] = {}
        self.lock = threading.Lock()
        self.started_at = time.time()

    # ---------------- routing ----------------
    def route(self, method: str, pattern: str):
        def deco(fn):
            self.routes.append((method.upper(), re.compile("^" + pattern + "$"), fn))
            return fn
        return deco

    def get(self, pattern):
        return self.route("GET", pattern)

    def post(self, pattern):
        return self.route("POST", pattern)

    def put(self, pattern):
        return self.route("PUT", pattern)

    def delete(self, pattern):
        return self.route("DELETE", pattern)

    def match(self, method: str, path: str):
        for m, rx, fn in self.routes:
            if m != method:
                continue
            mm = rx.match(path)
            if mm:
                return fn, mm
        return None, None

    # ---------------- tasks ----------------
    def start_task(self, kind: str, fn, *args, **kwargs) -> Task:
        task = Task(uuid.uuid4().hex[:12], kind)
        with self.lock:
            self.tasks[task.id] = task

        def runner():
            try:
                task.result = fn(task, *args, **kwargs)
                task.status = "done"
                task.progress = 100
            except ApiError as e:
                task.status = "error"
                task.error = f"{e.code}: {e.message}"
            except Exception as e:  # noqa: BLE001
                task.status = "error"
                task.error = f"{type(e).__name__}: {e}"
                traceback.print_exc()

        threading.Thread(target=runner, name=f"task-{kind}", daemon=True).start()
        return task

    def get_task(self, task_id: str) -> Task | None:
        return self.tasks.get(task_id)

    # ---------------- downloads ----------------
    def add_download(self, path: str) -> str:
        path = os.path.abspath(path)
        token = uuid.uuid4().hex
        self.downloads[token] = path
        return token


class Request:
    """一次请求的上下文。"""

    def __init__(self, app: App, method: str, path: str, query: dict, body, match):
        self.app = app
        self.method = method
        self.path = path
        self.query = query
        self.body = body
        self.match = match

    def param(self, name: str, default=None):
        vals = self.query.get(name)
        return vals[0] if vals else default

    def q_int(self, name: str, default=None):
        v = self.param(name)
        if v in (None, ""):
            return default
        try:
            return int(v)
        except ValueError:
            raise ApiError("BAD_PARAM", f"参数 {name} 需要整数", 400)

    def group(self, idx: int, default=None):
        try:
            return self.match.group(idx)
        except (IndexError, AttributeError):
            return default

    def json(self) -> dict:
        if not isinstance(self.body, dict):
            raise ApiError("BAD_BODY", "请求体必须是 JSON 对象", 400)
        return self.body


def _json_bytes(obj) -> bytes:
    return json.dumps(obj, ensure_ascii=False, default=str).encode("utf-8")


class Handler(BaseHTTPRequestHandler):
    server_version = "ExcelToolkit/0.1"
    protocol_version = "HTTP/1.1"

    # 安静一点：只记录错误
    def log_message(self, fmt, *args):  # noqa: A003
        if getattr(self, "_quiet", False):
            return
        sys.stderr.write("[%s] %s\n" % (time.strftime("%H:%M:%S"), fmt % args))

    # ---------------- helpers ----------------
    def _send(self, status: int, body: bytes, ctype: str, extra: dict | None = None):
        self.send_response(status)
        self.send_header("Content-Type", ctype)
        self.send_header("Content-Length", str(len(body)))
        self.send_header("Cache-Control", "no-store")
        for k, v in (extra or {}).items():
            self.send_header(k, v)
        self.end_headers()
        if self.command != "HEAD":
            self.wfile.write(body)

    def _send_json(self, obj, status: int = 200):
        self._send(status, _json_bytes(obj), "application/json; charset=utf-8")

    def _error(self, status: int, code: str, message: str, detail=None):
        self._send_json({"ok": False, "code": code, "message": message, "detail": detail}, status)

    def _read_body(self):
        length = int(self.headers.get("Content-Length") or 0)
        if length <= 0:
            return None
        if length > MAX_BODY:
            raise ApiError("BODY_TOO_LARGE", "请求体过大", 413)
        raw = self.rfile.read(length)
        ctype = (self.headers.get("Content-Type") or "").lower()
        if "application/json" in ctype or raw[:1] in (b"{", b"["):
            try:
                return json.loads(raw.decode("utf-8"))
            except Exception:  # noqa: BLE001
                raise ApiError("BAD_JSON", "JSON 解析失败", 400)
        # 文件上传：返回原始字节的封装
        return {"__raw__": raw, "__filename__": self.headers.get("X-Filename") or "upload.bin"}

    # ---------------- dispatch ----------------
    def _handle(self, method: str):
        parsed = urlparse(self.path)
        path = unquote(parsed.path)
        try:
            if path.startswith("/api/"):
                self._handle_api(method, path, parsed.query)
            else:
                self._handle_static(path)
        except FileResponse as fr:
            from urllib.parse import quote

            disp = "attachment; filename*=UTF-8''" + quote(fr.filename)
            self._send(200, fr.data, "application/octet-stream", {"Content-Disposition": disp})
        except ApiError as e:
            self._error(e.status, e.code, e.message, e.detail)
        except BrokenPipeError:
            pass
        except Exception as e:  # noqa: BLE001
            traceback.print_exc()
            self._error(500, "INTERNAL", f"{type(e).__name__}: {e}")

    def _handle_api(self, method: str, path: str, raw_query: str):
        app: App = self.server.app  # type: ignore[attr-defined]
        query = parse_qs(raw_query or "")
        body = None
        if method in ("POST", "PUT", "DELETE"):
            body = self._read_body()
        fn, match = app.match(method, path)
        if fn is None:
            # 方法不匹配时给出更明确的提示
            for m, rx, _ in app.routes:
                if rx.match(path):
                    raise ApiError("METHOD_NOT_ALLOWED", f"该路径不支持 {method}", 405)
            raise ApiError("NOT_FOUND", f"接口不存在: {path}", 404)
        req = Request(app, method, path, query, body, match)
        result = fn(req)
        if result is None:
            result = {"ok": True}
        elif not isinstance(result, dict) or "ok" not in result:
            result = {"ok": True, "data": result}
        self._send_json(result)

    def _handle_static(self, path: str):
        if path in ("/", "/index.html"):
            rel = "index.html"
        else:
            rel = path.lstrip("/")
        full = os.path.abspath(os.path.join(WEB_DIR, rel))
        if not full.startswith(os.path.abspath(WEB_DIR) + os.sep) or not os.path.isfile(full):
            self._error(404, "NOT_FOUND", "页面不存在")
            return
        ctype = mimetypes.guess_type(full)[0] or "application/octet-stream"
        if ctype.startswith("text/") or ctype in ("application/javascript", "application/json"):
            ctype += "; charset=utf-8"
        with open(full, "rb") as f:
            data = f.read()
        self._send(200, data, ctype)

    def do_GET(self):
        self._handle("GET")

    def do_POST(self):
        self._handle("POST")

    def do_PUT(self):
        self._handle("PUT")

    def do_DELETE(self):
        self._handle("DELETE")

    def do_HEAD(self):
        self._handle("GET")


class Server(ThreadingHTTPServer):
    daemon_threads = True
    allow_reuse_address = True

    def __init__(self, addr, handler, app: App):
        super().__init__(addr, handler)
        self.app = app


def create_app(root: str = PROJECT_ROOT) -> App:
    app = App(root)
    from . import api  # 延迟导入，避免循环

    api.register(app)
    return app


def run(port: int = DEFAULT_PORT, host: str = "127.0.0.1", open_browser: bool = True, root: str = PROJECT_ROOT):
    app = create_app(root)
    httpd = Server((host, port), Handler, app)
    url = f"http://{host}:{port}/"
    try:
        from .services import workspace_service

        ws_root = workspace_service.get_root()
    except Exception:  # noqa: BLE001
        ws_root = root
    print("=" * 62)
    print("  表格检索 / 模板改造 / 数据核对")
    print(f"  程序目录: {root}")
    print(f"  工作区:   {ws_root}")
    print(f"  服务地址: {url}")
    print("  关闭窗口或按 Ctrl+C 停止")
    print("=" * 62, flush=True)
    if open_browser:
        threading.Timer(0.8, lambda: webbrowser.open(url)).start()
    try:
        httpd.serve_forever()
    except KeyboardInterrupt:
        print("\n已停止。")
    finally:
        httpd.server_close()


class _Tee:
    """把输出同时写到控制台与日志文件（EXE 启动器下没有控制台）。"""

    def __init__(self, *streams):
        self.streams = [s for s in streams if s is not None]

    def write(self, data):
        for s in self.streams:
            try:
                s.write(data)
                s.flush()
            except Exception:  # noqa: BLE001
                pass
        return len(data or "")

    def flush(self):
        for s in self.streams:
            try:
                s.flush()
            except Exception:  # noqa: BLE001
                pass


def _setup_log(path: str | None):
    if not path:
        return
    try:
        os.makedirs(os.path.dirname(os.path.abspath(path)) or ".", exist_ok=True)
        f = open(path, "a", encoding="utf-8", buffering=1)
    except OSError:
        return
    sys.stdout = _Tee(getattr(sys, "__stdout__", None), f)
    sys.stderr = _Tee(getattr(sys, "__stderr__", None), f)
    try:
        with open(os.path.join(os.path.dirname(os.path.abspath(path)), "server.pid"), "w",
                  encoding="utf-8") as pf:
            pf.write(str(os.getpid()))
    except OSError:
        pass


def main(argv: list[str] | None = None):
    argv = list(sys.argv[1:] if argv is None else argv)
    port = DEFAULT_PORT
    host = "127.0.0.1"
    no_browser = False
    log_path = None
    i = 0
    while i < len(argv):
        a = argv[i]
        if a in ("--port", "-p") and i + 1 < len(argv):
            port = int(argv[i + 1])
            i += 2
            continue
        if a == "--host" and i + 1 < len(argv):
            host = argv[i + 1]
            i += 2
            continue
        if a in ("--no-browser", "-n"):
            no_browser = True
            i += 1
            continue
        if a in ("--log", "-l") and i + 1 < len(argv):
            log_path = argv[i + 1]
            i += 2
            continue
        i += 1
    _setup_log(log_path)
    run(port=port, host=host, open_browser=not no_browser)


if __name__ == "__main__":
    # 统一委派给 app.server 模块实例：
    # 无论 `python -m app.server` 还是 `python app/server.py`，都只应存在一份
    # app.server（否则 api.py 里的 FileResponse/ApiError 会与 __main__ 副本不是同一类）。
    sys.path.insert(0, PROJECT_ROOT)
    from app import server as _srv

    _srv.main()
