#!/usr/bin/env python3
"""Serve the Travel-Planner static app over HTTP (needed for the service worker)."""

from __future__ import annotations

import argparse
import contextlib
import http.server
import os
import re
import socketserver
import subprocess
import sys
import webbrowser
from pathlib import Path
from urllib.parse import urlparse


_FRONT_MATTER = re.compile(r"\A---\r?\n.*?\r?\n---\r?\n", re.DOTALL)
_CACHE_NAME = re.compile(r"const CACHE_NAME = 'travel-planner-[^']*';")


def git_sha(root: Path) -> str:
    """Short HEAD sha, with a stable -dirty suffix when the worktree has changes."""
    try:
        sha = subprocess.check_output(
            ["git", "rev-parse", "--short=12", "HEAD"],
            cwd=root,
            text=True,
            stderr=subprocess.DEVNULL,
        ).strip()
    except (OSError, subprocess.CalledProcessError):
        return "dev"
    if not sha:
        return "dev"
    try:
        dirty = subprocess.check_output(
            ["git", "status", "--porcelain"],
            cwd=root,
            text=True,
            stderr=subprocess.DEVNULL,
        )
    except (OSError, subprocess.CalledProcessError):
        dirty = ""
    if dirty.strip():
        return f"{sha}-dirty"
    return sha


class Handler(http.server.SimpleHTTPRequestHandler):
    def do_GET(self) -> None:
        if urlparse(self.path).path == "/sw.js":
            self._send_service_worker()
            return
        super().do_GET()

    def _send_service_worker(self) -> None:
        root = Path(".").resolve()
        text = _FRONT_MATTER.sub("", (root / "sw.js").read_text(encoding="utf-8"), count=1)
        text = _CACHE_NAME.sub(
            f"const CACHE_NAME = 'travel-planner-{git_sha(root)}';",
            text,
            count=1,
        )
        body = text.encode("utf-8")
        self.send_response(200)
        self.send_header("Content-Type", "application/javascript; charset=utf-8")
        self.send_header("Content-Length", str(len(body)))
        self.send_header("Cache-Control", "no-cache")
        self.end_headers()
        self.wfile.write(body)


def main() -> int:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument(
        "--port",
        "-p",
        type=int,
        default=8000,
        help="Port to listen on (default: 8000)",
    )
    parser.add_argument(
        "--host",
        default="localhost",
        help="Bind address (default: localhost)",
    )
    parser.add_argument(
        "--no-browser",
        action="store_true",
        help="Do not open the default browser",
    )
    args = parser.parse_args()

    root = Path(__file__).resolve().parent
    os.chdir(root)

    handler = Handler
    url = f"http://{args.host}:{args.port}/"

    try:
        with socketserver.TCPServer((args.host, args.port), handler) as httpd:
            print(f"Serving {root}")
            print(f"Open {url}")
            if not args.no_browser:
                with contextlib.suppress(OSError):
                    webbrowser.open(url)
            try:
                httpd.serve_forever()
            except KeyboardInterrupt:
                print("\nStopped.")
    except OSError as e:
        print(f"Could not start server: {e}", file=sys.stderr)
        return 1
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
