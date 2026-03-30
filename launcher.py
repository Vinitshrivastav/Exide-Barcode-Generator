import http.server, webbrowser, threading, os, time
from pathlib import Path

PORT    = 7735
BASE    = Path(__file__).parent

class H(http.server.SimpleHTTPRequestHandler):
    def __init__(self,*a,**k): super().__init__(*a,directory=str(BASE),**k)
    def log_message(self,*a): pass
    def do_GET(self):
        if self.path in ("/",""):
            self.send_response(302); self.send_header("Location","/index.html"); self.end_headers()
        else: super().do_GET()

threading.Timer(0.9, lambda: webbrowser.open(f"http://localhost:{PORT}/index.html")).start()
with http.server.HTTPServer(("localhost",PORT),H) as s: s.serve_forever()
