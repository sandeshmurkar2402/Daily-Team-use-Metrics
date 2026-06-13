#!/usr/bin/env python3
"""Local dashboard server with on-demand data refresh.
Usage: python server.py [port]   (default port: 8080)
"""
import json
import os
import subprocess
import sys
from http.server import HTTPServer, SimpleHTTPRequestHandler

ROOT = os.path.dirname(os.path.abspath(__file__))


class DashboardHandler(SimpleHTTPRequestHandler):
    def do_GET(self):
        if self.path.split("?")[0] == "/api/refresh":
            self._handle_refresh()
        else:
            super().do_GET()

    def _handle_refresh(self):
        print("  --> Fetching latest data from Google Sheets...")
        try:
            result = subprocess.run(
                [sys.executable, os.path.join(ROOT, "scripts", "fetch.py")],
                capture_output=True,
                text=True,
                cwd=ROOT,
                timeout=60,
            )
            if result.returncode == 0:
                with open(os.path.join(ROOT, "data", "metrics.json"), "rb") as f:
                    body = f.read()
                print("  --> OK: data refreshed")
                self._json_response(200, body)
            else:
                msg = (result.stderr or result.stdout or "fetch.py exited with error").strip()
                print(f"  --> FAIL: {msg}")
                self._json_response(500, json.dumps({"error": msg}).encode())
        except subprocess.TimeoutExpired:
            self._json_response(504, json.dumps({"error": "Fetch timed out after 60s"}).encode())
        except Exception as e:
            self._json_response(500, json.dumps({"error": str(e)}).encode())

    def _json_response(self, status, body: bytes):
        self.send_response(status)
        self.send_header("Content-Type", "application/json")
        self.send_header("Content-Length", str(len(body)))
        self.send_header("Cache-Control", "no-store")
        self.end_headers()
        self.wfile.write(body)

    def log_message(self, fmt, *args):
        # Suppress noisy static-file logs; keep API logs
        if "/api/" in str(args[0]):
            print(f"[{self.log_date_time_string()}] {args[0]}")


if __name__ == "__main__":
    port = int(sys.argv[1]) if len(sys.argv) > 1 else 8080
    os.chdir(ROOT)
    httpd = HTTPServer(("", port), DashboardHandler)
    print(f"Dashboard server running at http://localhost:{port}")
    print("Press Ctrl+C to stop.\n")
    try:
        httpd.serve_forever()
    except KeyboardInterrupt:
        print("\nServer stopped.")
