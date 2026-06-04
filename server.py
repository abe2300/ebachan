"""pythonw で起動するローカル HTTP サーバー。
pythonw は stdout/stderr を持たないため、http.server のログ出力で
BrokenPipe / AttributeError が発生し接続が切れる。
それを避けるためログを破棄し、SimpleHTTPRequestHandler のログも無効化する。
"""
import os
import sys
import http.server
import socketserver

# pythonw 配下では sys.stdout / sys.stderr が None になり得るので差し替え
sys.stdout = open(os.devnull, "w")
sys.stderr = open(os.devnull, "w")

PORT = 8767
HOST = "127.0.0.1"


class QuietHandler(http.server.SimpleHTTPRequestHandler):
    def log_message(self, format, *args):
        pass


os.chdir(os.path.dirname(os.path.abspath(__file__)))

socketserver.TCPServer.allow_reuse_address = True
with socketserver.TCPServer((HOST, PORT), QuietHandler) as httpd:
    httpd.serve_forever()
