"""Local launcher for ADR Cards.

Serves the tool at http://127.0.0.1:8770/ and gives it direct read/write access
to the ADR folder named in adr-folder.txt, so the page never asks for a folder.
Started by the desktop shortcut (pythonw, no console window). If it is already
running, it just opens the browser.
"""
import json
import os
import sys
import urllib.parse
import webbrowser
from http.server import SimpleHTTPRequestHandler, ThreadingHTTPServer

HOST, PORT = '127.0.0.1', 8770
URL = f'http://{HOST}:{PORT}/'
APP_DIR = os.path.dirname(os.path.abspath(__file__))
CONFIG = os.path.join(APP_DIR, 'adr-folder.txt')


def read_adr_dir():
    with open(CONFIG, encoding='utf-8-sig') as f:
        return os.path.realpath(f.read().strip())


class Handler(SimpleHTTPRequestHandler):
    def __init__(self, *args, **kwargs):
        super().__init__(*args, directory=APP_DIR, **kwargs)

    def end_headers(self):
        # Always serve the current files; no stale copies after updates.
        self.send_header('Cache-Control', 'no-store')
        super().end_headers()

    def log_message(self, *args):
        pass

    # ---- helpers ----

    def reply(self, status, body=b'', ctype='text/plain; charset=utf-8'):
        if isinstance(body, str):
            body = body.encode('utf-8')
        self.send_response(status)
        self.send_header('Content-Type', ctype)
        self.send_header('Content-Length', str(len(body)))
        self.end_headers()
        self.wfile.write(body)

    def allowed(self):
        # Only this page may use the API: the Host check blocks DNS rebinding,
        # and the custom header forces a CORS preflight (which we never
        # answer) for any request coming from another website.
        return (self.headers.get('Host') == f'{HOST}:{PORT}'
                and self.headers.get('X-ADR-Cards') == '1')

    def target(self, query):
        rel = urllib.parse.parse_qs(query).get('path', [''])[0]
        base = read_adr_dir()
        path = os.path.realpath(os.path.join(base, rel))
        if path != base and not path.startswith(base + os.sep):
            raise PermissionError('path outside ADR folder')
        return path

    # ---- API ----

    def do_GET(self):
        url = urllib.parse.urlsplit(self.path)
        if not url.path.startswith('/api/'):
            return super().do_GET()
        if not self.allowed():
            return self.reply(403, 'forbidden')
        try:
            if url.path == '/api/info':
                base = read_adr_dir()
                if not os.path.isdir(base):
                    return self.reply(404, f'ADR folder not found: {base}')
                return self.reply(200, json.dumps({'name': os.path.basename(base)}), 'application/json')
            path = self.target(url.query)
            if url.path == '/api/stat':
                return self.reply(200 if os.path.isfile(path) else 404)
            if url.path == '/api/dir':
                if not os.path.isdir(path):
                    if 'create=1' not in url.query:
                        return self.reply(404)
                    os.makedirs(path, exist_ok=True)
                return self.reply(200)
            if url.path == '/api/list':
                names = sorted(os.listdir(path)) if os.path.isdir(path) else []
                return self.reply(200, json.dumps(names), 'application/json')
            if url.path == '/api/file':
                if not os.path.isfile(path):
                    return self.reply(404)
                with open(path, 'rb') as f:
                    return self.reply(200, f.read(), 'application/octet-stream')
            return self.reply(404)
        except Exception as e:
            return self.reply(500, f'{type(e).__name__}: {e}')

    def do_PUT(self):
        url = urllib.parse.urlsplit(self.path)
        if url.path != '/api/file' or not self.allowed():
            return self.reply(403, 'forbidden')
        try:
            path = self.target(url.query)
            data = self.rfile.read(int(self.headers.get('Content-Length', 0)))
            os.makedirs(os.path.dirname(path), exist_ok=True)
            tmp = path + '.adrtmp'
            with open(tmp, 'wb') as f:
                f.write(data)
            try:
                os.replace(tmp, path)
            except PermissionError:
                os.remove(tmp)
                return self.reply(409, '檔案被其他程式開啟中（請先關閉 Excel）')
            return self.reply(200)
        except Exception as e:
            return self.reply(500, f'{type(e).__name__}: {e}')


def main(open_browser=True):
    try:
        server = ThreadingHTTPServer((HOST, PORT), Handler)
    except OSError:
        # Already running: just show it.
        if open_browser:
            webbrowser.open(URL)
        return
    if open_browser:
        webbrowser.open(URL)
    server.serve_forever()


if __name__ == '__main__':
    main(open_browser='--no-browser' not in sys.argv)
