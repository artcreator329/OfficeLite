#!/usr/bin/env python3
"""
OfficeLite Local Server
A lightweight, zero-dependency local server for OfficeLite.
Serves the application assets and automatically launches your browser.
"""

import os
import sys
import webbrowser
import socket
from http.server import HTTPServer, SimpleHTTPRequestHandler

PORT = 8765
BASE_DIR = os.path.dirname(os.path.abspath(__file__))


class OfficeLiteRequestHandler(SimpleHTTPRequestHandler):
    def __init__(self, *args, **kwargs):
        super().__init__(*args, directory=BASE_DIR, **kwargs)

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
