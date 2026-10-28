from __future__ import annotations

import socket
import time
import tkinter as tk
import unittest
from urllib.request import urlopen

from desktop_launcher import LauncherWindow, LocalWebServer, ServerStartError


class LocalWebServerTests(unittest.TestCase):
    def test_start_serves_health_and_page_then_stops(self) -> None:
        server = LocalWebServer(port=0)
        try:
            url = server.start()
            self.assertTrue(server.running)
            with urlopen(f"{url}/api/health", timeout=2) as response:
                self.assertEqual(response.status, 200)
                self.assertIn(b'"ok":true', response.read())
            with urlopen(url, timeout=2) as response:
                self.assertEqual(response.status, 200)
                self.assertIn("보고서 작성기".encode("utf-8"), response.read())
        finally:
            server.stop()
        self.assertFalse(server.running)

    def test_server_can_restart_without_leaving_an_orphan_port(self) -> None:
        server = LocalWebServer(port=0)
        first_url = server.start()
        selected_port = server.port
        server.stop()

        second_url = server.start()
        try:
            self.assertEqual(first_url, second_url)
            self.assertEqual(server.port, selected_port)
            self.assertTrue(server.running)
        finally:
            server.stop()

    def test_occupied_port_reports_a_start_error(self) -> None:
        listener = socket.socket(socket.AF_INET, socket.SOCK_STREAM)
        listener.bind(("127.0.0.1", 0))
        listener.listen(1)
        port = listener.getsockname()[1]
        try:
            with self.assertRaises(ServerStartError):
                LocalWebServer(port=port).start()
        finally:
            listener.close()


class LauncherWindowTests(unittest.TestCase):
    def setUp(self) -> None:
        self.root = tk.Tk()
        self.root.withdraw()
        self.launcher = LauncherWindow(self.root, open_browser=False)
        self.launcher.server = LocalWebServer(port=0)

    def tearDown(self) -> None:
        self.launcher.server.stop()
        try:
            self.root.destroy()
        except tk.TclError:
            pass

    def test_toggle_button_starts_and_stops_server(self) -> None:
        self.assertEqual(self.launcher.toggle_button.cget("text"), "웹서버 시작")

        self.launcher.toggle_server()
        self._pump_until(lambda: self.launcher.toggle_button.cget("text") == "종료")
        self.assertTrue(self.launcher.server.running)

        self.launcher.toggle_server()
        self._pump_until(lambda: self.launcher.toggle_button.cget("text") == "웹서버 시작")
        self.assertFalse(self.launcher.server.running)

    def test_closing_window_stops_running_server(self) -> None:
        self.launcher.toggle_server()
        self._pump_until(lambda: self.launcher.server.running)

        self.launcher.close()
        deadline = time.monotonic() + 5.0
        while time.monotonic() < deadline and self.launcher.server.running:
            try:
                self.root.update()
            except tk.TclError:
                break
            time.sleep(0.02)
        self.assertFalse(self.launcher.server.running)

    def _pump_until(self, predicate, timeout: float = 5.0) -> None:
        deadline = time.monotonic() + timeout
        while time.monotonic() < deadline:
            self.root.update()
            if predicate():
                return
            time.sleep(0.02)
        self.fail("GUI 상태 변경이 제한 시간 안에 완료되지 않았습니다.")


if __name__ == "__main__":
    unittest.main()
