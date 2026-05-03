#!/usr/bin/env python3
"""Serve the Travel-Planner static app over HTTP (needed for the service worker)."""

from __future__ import annotations

import argparse
import contextlib
import http.server
import os
import socketserver
import sys
import webbrowser
from pathlib import Path


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

    handler = http.server.SimpleHTTPRequestHandler
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
