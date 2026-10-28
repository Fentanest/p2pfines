from __future__ import annotations

import argparse
import json
import multiprocessing
import queue
import threading
import time
import tkinter as tk
import webbrowser
from tkinter import messagebox, ttk
from typing import Callable
from urllib.error import URLError
from urllib.request import urlopen

from werkzeug.serving import WSGIRequestHandler, make_server

from app import app as flask_app
from report_generator_app.paths import RESOURCE_ROOT


DEFAULT_HOST = "127.0.0.1"
DEFAULT_PORT = 8787
STARTUP_TIMEOUT_SECONDS = 5.0
SHUTDOWN_TIMEOUT_SECONDS = 5.0


class ServerStartError(RuntimeError):
    pass


class QuietRequestHandler(WSGIRequestHandler):
    def log(self, type: str, message: str, *args: object) -> None:
        return


class LocalWebServer:
    def __init__(self, host: str = DEFAULT_HOST, port: int = DEFAULT_PORT) -> None:
        self.host = host
        self.port = port
        self._server = None
        self._thread: threading.Thread | None = None
        self._lock = threading.RLock()

    @property
    def url(self) -> str:
        return f"http://{self.host}:{self.port}"

    @property
    def running(self) -> bool:
        with self._lock:
            return bool(self._server and self._thread and self._thread.is_alive())

    def start(self, timeout: float = STARTUP_TIMEOUT_SECONDS) -> str:
        with self._lock:
            if self.running:
                return self.url
            self._close_stale_server()
            try:
                server = make_server(
                    self.host,
                    self.port,
                    flask_app,
                    threaded=True,
                    request_handler=QuietRequestHandler,
                )
            except SystemExit as exc:
                raise ServerStartError(self._port_error_message()) from exc
            except OSError as exc:
                raise ServerStartError(self._port_error_message()) from exc

            self.port = int(server.server_port)
            thread = threading.Thread(
                target=server.serve_forever,
                name="report-generator-web-server",
                daemon=True,
            )
            self._server = server
            self._thread = thread
            thread.start()

        try:
            self._wait_until_ready(timeout)
        except Exception:
            self.stop()
            raise
        return self.url

    def stop(self, timeout: float = SHUTDOWN_TIMEOUT_SECONDS) -> None:
        with self._lock:
            server = self._server
            thread = self._thread
        if server is None:
            return

        server.shutdown()
        if thread and thread is not threading.current_thread():
            thread.join(timeout)
            if thread.is_alive():
                raise RuntimeError("웹서버가 제한 시간 안에 종료되지 않았습니다.")
        server.server_close()

        with self._lock:
            if self._server is server:
                self._server = None
                self._thread = None

    def _wait_until_ready(self, timeout: float) -> None:
        deadline = time.monotonic() + timeout
        last_error: Exception | None = None
        while time.monotonic() < deadline:
            try:
                with urlopen(f"{self.url}/api/health", timeout=0.5) as response:
                    if response.status == 200:
                        return
            except (OSError, URLError) as exc:
                last_error = exc
            time.sleep(0.05)
        raise ServerStartError(f"웹서버 시작을 확인하지 못했습니다: {last_error or '시간 초과'}")

    def _close_stale_server(self) -> None:
        if self._server is not None:
            try:
                self._server.server_close()
            finally:
                self._server = None
                self._thread = None

    def _port_error_message(self) -> str:
        return f"{self.host}:{self.port} 포트를 사용할 수 없습니다. 이미 실행 중인 프로그램이 있는지 확인해주세요."


class LauncherWindow:
    def __init__(self, root: tk.Tk, open_browser: bool = True) -> None:
        self.root = root
        self.open_browser = open_browser
        self.server = LocalWebServer()
        self._busy = False
        self._closing = False
        self._operation_lock = threading.Lock()
        self._ui_events: queue.SimpleQueue[Callable[[], None]] = queue.SimpleQueue()

        root.title("보고서 작성기 웹서버")
        root.geometry("420x245")
        root.minsize(380, 225)
        root.protocol("WM_DELETE_WINDOW", self.close)
        root.after(50, self._drain_ui_events)

        frame = ttk.Frame(root, padding=28)
        frame.pack(fill="both", expand=True)

        ttk.Label(frame, text="보고서 작성기", font=("Malgun Gothic", 20, "bold")).pack(pady=(2, 8))
        ttk.Label(
            frame,
            text="로컬 보고서 업로더 웹페이지를 실행합니다.",
            font=("Malgun Gothic", 10),
        ).pack(pady=(0, 16))

        self.status_var = tk.StringVar(value="중지됨")
        self.status_label = ttk.Label(frame, textvariable=self.status_var, foreground="#52606d")
        self.status_label.pack(pady=(0, 12))

        self.toggle_button = ttk.Button(frame, text="웹서버 시작", command=self.toggle_server)
        self.toggle_button.pack(ipadx=32, ipady=9)

        ttk.Label(
            frame,
            text=f"주소: http://{DEFAULT_HOST}:{DEFAULT_PORT}",
            foreground="#7b8794",
            font=("Malgun Gothic", 9),
        ).pack(pady=(14, 0))

    def toggle_server(self) -> None:
        if self._busy or self._closing:
            return
        if self.server.running:
            self._begin_stop()
        else:
            self._begin_start()

    def _begin_start(self) -> None:
        self._set_busy("시작 중...", "시작 중...")
        threading.Thread(target=self._start_worker, name="report-generator-start", daemon=True).start()

    def _start_worker(self) -> None:
        try:
            with self._operation_lock:
                url = self.server.start()
                if self.open_browser and not self._closing:
                    webbrowser.open_new_tab(url)
            self._post(lambda: self._finish_start(url))
        except BaseException as exc:
            self._post(lambda error=exc: self._show_error("웹서버 시작 실패", error))

    def _finish_start(self, url: str) -> None:
        if self._closing:
            return
        self._busy = False
        self.status_var.set(f"실행 중 · {url}")
        self.status_label.configure(foreground="#137333")
        self.toggle_button.configure(text="종료", state="normal")

    def _begin_stop(self) -> None:
        self._set_busy("종료 중...", "종료 중...")
        threading.Thread(target=self._stop_worker, name="report-generator-stop", daemon=True).start()

    def _stop_worker(self) -> None:
        try:
            with self._operation_lock:
                self.server.stop()
            self._post(self._finish_stop)
        except BaseException as exc:
            self._post(lambda error=exc: self._show_error("웹서버 종료 실패", error))

    def _finish_stop(self) -> None:
        if self._closing:
            self.root.destroy()
            return
        self._busy = False
        self.status_var.set("중지됨")
        self.status_label.configure(foreground="#52606d")
        self.toggle_button.configure(text="웹서버 시작", state="normal")

    def _set_busy(self, status: str, button_text: str) -> None:
        self._busy = True
        self.status_var.set(status)
        self.status_label.configure(foreground="#b06000")
        self.toggle_button.configure(text=button_text, state="disabled")

    def _show_error(self, title: str, error: BaseException) -> None:
        if self._closing:
            self.root.destroy()
            return
        self._busy = False
        self.status_var.set("중지됨")
        self.status_label.configure(foreground="#b3261e")
        self.toggle_button.configure(text="웹서버 시작", state="normal")
        messagebox.showerror(title, str(error) or error.__class__.__name__, parent=self.root)

    def close(self) -> None:
        if self._closing:
            return
        self._closing = True
        self._set_busy("종료 중...", "종료 중...")
        threading.Thread(target=self._close_worker, name="report-generator-close", daemon=True).start()

    def _close_worker(self) -> None:
        try:
            with self._operation_lock:
                self.server.stop()
        finally:
            self._post(self.root.destroy)

    def _post(self, callback: Callable[[], None]) -> None:
        self._ui_events.put(callback)

    def _drain_ui_events(self) -> None:
        try:
            while True:
                self._ui_events.get_nowait()()
        except queue.Empty:
            pass
        except tk.TclError:
            return
        try:
            self.root.after(50, self._drain_ui_events)
        except tk.TclError:
            return


def run_smoke_test(port: int) -> int:
    required_resources = (
        RESOURCE_ROOT / "templates" / "index.html",
        RESOURCE_ROOT / "apps_script" / "Code_work.gs",
        RESOURCE_ROOT / "apps_script" / "Code_watch.gs",
        RESOURCE_ROOT / "june.xlsx",
        RESOURCE_ROOT / "업무보고서 작성기_v5.4_통합.xlsx",
    )
    missing = [str(path) for path in required_resources if not path.exists()]
    if missing:
        raise RuntimeError("필수 리소스를 찾지 못했습니다: " + ", ".join(missing))

    server = LocalWebServer(port=port)
    try:
        url = server.start()
        with urlopen(url, timeout=5.0) as response:
            html = response.read()
        if response.status != 200 or "보고서 작성기".encode("utf-8") not in html:
            raise RuntimeError("웹 UI 응답이 올바르지 않습니다.")
        with urlopen(f"{url}/api/apps-script?target=work", timeout=5.0) as response:
            script_payload = json.loads(response.read().decode("utf-8"))
        if response.status != 200 or "recalculateAllReports" not in script_payload.get("code", ""):
            raise RuntimeError("Apps Script 리소스 응답이 올바르지 않습니다.")
        return 0
    finally:
        server.stop()


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(description="보고서 작성기 Windows launcher")
    parser.add_argument("--no-browser", action="store_true", help="서버 시작 시 브라우저를 자동으로 열지 않습니다.")
    parser.add_argument("--smoke-test", action="store_true", help="서버와 웹 UI를 검사한 뒤 즉시 종료합니다.")
    parser.add_argument("--port", type=int, default=DEFAULT_PORT, help=argparse.SUPPRESS)
    return parser.parse_args()


def main() -> int:
    multiprocessing.freeze_support()
    args = parse_args()
    if args.smoke_test:
        return run_smoke_test(args.port)

    root = tk.Tk()
    LauncherWindow(root, open_browser=not args.no_browser)
    root.mainloop()
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
