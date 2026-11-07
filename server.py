#!/usr/bin/env python3
"""
OfficeLite Local Server
A lightweight, zero-dependency local server for OfficeLite.
Provides optional high-fidelity Pandoc conversion for .docx files,
serves the application assets, and automatically launches your browser.
"""

import os
import sys
import json
import shutil
import tempfile
import subprocess
import webbrowser
import socket
from http.server import HTTPServer, SimpleHTTPRequestHandler

PORT = 8765
BASE_DIR = os.path.dirname(os.path.abspath(__file__))

def check_pandoc():
    pandoc_path = shutil.which("pandoc") or "/opt/homebrew/bin/pandoc"
    if os.path.exists(pandoc_path) and os.access(pandoc_path, os.X_OK):
        try:
            res = subprocess.run([pandoc_path, "--version"], capture_output=True, text=True, timeout=2)
            version_line = res.stdout.splitlines()[0] if res.stdout else "unknown"
            return pandoc_path, version_line
        except Exception:
            return None, None
    return None, None

PANDOC_BIN, PANDOC_VERSION = check_pandoc()

class OfficeLiteRequestHandler(SimpleHTTPRequestHandler):
    def __init__(self, *args, **kwargs):
        super().__init__(*args, directory=BASE_DIR, **kwargs)

    def do_GET(self):
        if self.path == "/api/status":
            self.send_response(200)
            self.send_header("Content-Type", "application/json")
            self.end_headers()
            data = {
                "status": "online",
                "pandoc_available": PANDOC_BIN is not None,
                "pandoc_path": PANDOC_BIN,
                "pandoc_version": PANDOC_VERSION
            }
            self.wfile.write(json.dumps(data).encode("utf-8"))
            return
        return super().do_GET()

    def do_POST(self):
        if self.path == "/api/export-docx":
            content_length = int(self.headers.get("Content-Length", 0))
            post_data = self.rfile.read(content_length)
            try:
                payload = json.loads(post_data.decode("utf-8"))
                html_content = payload.get("html", "")
                filename = payload.get("filename", "document.docx")
                if not filename.endswith(".docx"):
                    filename += ".docx"

                if not PANDOC_BIN:
                    self.send_error(501, "Pandoc not installed on server.")
                    return

                with tempfile.NamedTemporaryFile("w", suffix=".html", delete=False, encoding="utf-8") as f_html:
                    f_html.write(html_content)
                    html_path = f_html.name

                docx_path = html_path.replace(".html", ".docx")
                try:
                    subprocess.run(
                        [PANDOC_BIN, "-f", "html", "-t", "docx", "-o", docx_path, html_path],
                        check=True,
                        timeout=10
                    )
                    with open(docx_path, "rb") as f_docx:
                        docx_bytes = f_docx.read()

                    self.send_response(200)
                    self.send_header("Content-Type", "application/vnd.openxmlformats-officedocument.wordprocessingml.document")
                    self.send_header("Content-Disposition", f'attachment; filename="{filename}"')
                    self.send_header("Content-Length", str(len(docx_bytes)))
                    self.end_headers()
                    self.wfile.write(docx_bytes)
                finally:
                    if os.path.exists(html_path):
                        os.remove(html_path)
                    if os.path.exists(docx_path):
                        os.remove(docx_path)
            except Exception as e:
                self.send_error(500, f"Error generating docx: {str(e)}")
            return

        elif self.path == "/api/import-docx":
            content_length = int(self.headers.get("Content-Length", 0))
            docx_bytes = self.rfile.read(content_length)
            if not PANDOC_BIN:
                self.send_error(501, "Pandoc not installed on server.")
                return

            with tempfile.NamedTemporaryFile("wb", suffix=".docx", delete=False) as f_docx:
                f_docx.write(docx_bytes)
                docx_path = f_docx.name

            try:
                res = subprocess.run(
                    [PANDOC_BIN, "-f", "docx", "-t", "html", docx_path],
                    capture_output=True,
                    text=True,
                    timeout=10,
                    check=True
                )
                html_output = res.stdout
                self.send_response(200)
                self.send_header("Content-Type", "application/json")
                self.end_headers()
                self.wfile.write(json.dumps({"html": html_output}).encode("utf-8"))
            except Exception as e:
                self.send_error(500, f"Error importing docx: {str(e)}")
            finally:
                if os.path.exists(docx_path):
                    os.remove(docx_path)
            return

        self.send_error(404, "Endpoint not found")

    def do_OPTIONS(self):
        self.send_response(200)
        self.end_headers()

def is_port_in_use(port):
    with socket.socket(socket.AF_INET, socket.SOCK_STREAM) as s:
        return s.connect_ex(('127.0.0.1', port)) == 0

def run_server():
    port = PORT
    while is_port_in_use(port) and port < PORT + 20:
        port += 1

    server_address = ("127.0.0.1", port)
    httpd = HTTPServer(server_address, OfficeLiteRequestHandler)
    url = f"http://127.0.0.1:{port}/"
    print(f"============================================================")
    print(f"  OfficeLite Server Running at {url}")
    if PANDOC_BIN:
        print(f"  Pandoc detected: {PANDOC_VERSION} ({PANDOC_BIN})")
    else:
        print(f"  Pandoc: Not found (client-side fallback mode active)")
    print(f"  Press Ctrl+C to stop the server")
    print(f"============================================================")

    try:
        webbrowser.open(url)
    except Exception:
        pass

    try:
        httpd.serve_forever()
    except KeyboardInterrupt:
        print("\nOfficeLite server stopped.")
        httpd.server_close()

if __name__ == "__main__":
    run_server()
