#!/usr/bin/env python3
"""
Callable sidecar over update_timesheet.py.

Turns the CLI into one endpoint the ai-automation backend calls over the compose
network — the image stays the standalone tool, this just exposes it as HTTP. The
two files travel as base64 in a small JSON body (no multipart machinery needed):

    POST /generate
      {"template": "<base64 xlsx>", "csv": "<base64 csv>"}
    → 200 {"xlsx": "<base64 xlsx>"}   |   400/500 {"error": "..."}

    GET /ready → 200 {"ok": true}

Run with `python server.py` (PORT env, default 8080). Both the CLI image mode
(`docker run image <template> <csv>`) and this server live in the same image.
"""
import base64
import json
import os
import subprocess
import sys
import tempfile
from http.server import BaseHTTPRequestHandler, ThreadingHTTPServer
from urllib.parse import urlparse

HERE = os.path.dirname(os.path.abspath(__file__))
SCRIPT = os.path.join(HERE, "update_timesheet.py")
# A generous cap so a stray client cannot make the server buffer forever.
MAX_BODY = 64 * 1024 * 1024


def generate(template_b64: str, csv_b64: str) -> bytes:
    """Run the exact CLI pipeline on two in-memory files and return the xlsx bytes."""
    with tempfile.TemporaryDirectory() as d:
        template = os.path.join(d, "template.xlsx")
        csv_path = os.path.join(d, "export.csv")
        out = os.path.join(d, "out.xlsx")
        with open(template, "wb") as f:
            f.write(base64.b64decode(template_b64))
        with open(csv_path, "wb") as f:
            f.write(base64.b64decode(csv_b64))
        proc = subprocess.run(
            [sys.executable, SCRIPT, template, csv_path, "-o", out],
            capture_output=True,
            text=True,
            timeout=300,
        )
        if proc.returncode != 0:
            raise RuntimeError((proc.stdout or "") + (proc.stderr or ""))
        with open(out, "rb") as f:
            return f.read()


class Handler(BaseHTTPRequestHandler):
    def _send(self, code: int, obj: dict) -> None:
        body = json.dumps(obj).encode("utf-8")
        self.send_response(code)
        self.send_header("Content-Type", "application/json; charset=utf-8")
        self.send_header("Content-Length", str(len(body)))
        self.end_headers()
        self.wfile.write(body)

    def do_GET(self) -> None:
        if urlparse(self.path).path == "/ready":
            self._send(200, {"ok": True})
        else:
            self._send(404, {"error": "not found"})

    def do_POST(self) -> None:
        if urlparse(self.path).path != "/generate":
            return self._send(404, {"error": "not found"})
        length = int(self.headers.get("Content-Length", "0"))
        if length <= 0 or length > MAX_BODY:
            return self._send(400, {"error": "missing or too large body"})
        try:
            payload = json.loads(self.rfile.read(length).decode("utf-8"))
            xlsx = generate(payload["template"], payload["csv"])
        except (KeyError, ValueError) as err:
            return self._send(400, {"error": f"invalid request: {err}"})
        except Exception as err:  # the pipeline failed (bad files, missing placeholders)
            return self._send(500, {"error": repr(err)})
        self._send(200, {"xlsx": base64.b64encode(xlsx).decode("ascii")})

    def log_message(self, format, *args) -> None:  # keep the /ready poll quiet
        pass


if __name__ == "__main__":
    port = int(os.environ.get("PORT", "8080"))
    print(f"timesheet-updater sidecar listening on :{port}", flush=True)
    ThreadingHTTPServer(("0.0.0.0", port), Handler).serve_forever()