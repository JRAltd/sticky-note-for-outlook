#!/usr/bin/env python3
"""
Local HTTPS Server for Outlook Email Sticky Notes Add-in
Serves ./dist over https://localhost:3000 (or http://localhost:3000 if --http is passed).
"""
import http.server
import os
import ssl
import sys
from pathlib import Path

PORT = int(os.environ.get("PORT", "3000"))
ROOT_DIR = Path(__file__).resolve().parent
DIST_DIR = ROOT_DIR / "dist"
CERTS_DIR = ROOT_DIR / "certs"
CERT_FILE = CERTS_DIR / "localhost.crt"
KEY_FILE = CERTS_DIR / "localhost.key"


class AddinRequestHandler(http.server.SimpleHTTPRequestHandler):
    def __init__(self, *args, **kwargs):
        super().__init__(*args, directory=str(DIST_DIR), **kwargs)

    def end_headers(self):
        self.send_header("Access-Control-Allow-Origin", "*")
        self.send_header("Access-Control-Allow-Methods", "GET, POST, OPTIONS")
        self.send_header("Access-Control-Allow-Headers", "Content-Type, Authorization")
        self.send_header("Cache-Control", "no-store")
        super().end_headers()

    def do_OPTIONS(self):
        self.send_response(200)
        self.end_headers()

    def do_GET(self):
        if self.path in ("/", ""):
            self.path = "/taskpane.html"
        return super().do_GET()


def main():
    use_https = "--http" not in sys.argv and CERT_FILE.exists() and KEY_FILE.exists()
    server_address = ("localhost", PORT)
    httpd = http.server.ThreadingHTTPServer(server_address, AddinRequestHandler)

    if use_https:
        ctx = ssl.SSLContext(ssl.PROTOCOL_TLS_SERVER)
        ctx.load_cert_chain(certfile=str(CERT_FILE), keyfile=str(KEY_FILE))
        httpd.socket = ctx.wrap_socket(httpd.socket, server_side=True)
        protocol = "https"
    else:
        protocol = "http"

    print("=" * 68)
    print(f" Outlook Email Sticky Notes Add-in Server Running")
    print(f" Taskpane URL : {protocol}://localhost:{PORT}/taskpane.html")
    print(f" Commands URL : {protocol}://localhost:{PORT}/commands.html")
    print(f" Manifest XML : {ROOT_DIR / 'manifest.xml'}")
    print("=" * 68)
    try:
        httpd.serve_forever()
    except KeyboardInterrupt:
        print("\nShutting down server...")
        httpd.server_close()


if __name__ == "__main__":
    main()
