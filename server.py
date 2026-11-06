"""Local launcher for ADR Cards.

Serves the tool at http://127.0.0.1:8770/ and gives it direct read/write access
to the ADR folder named in adr-folder.txt, so the page never asks for a folder.
Started by the desktop shortcut (pythonw, no console window), and quietly at
Windows sign-in. If it is already running, it just opens the app window.
Each start first updates this folder from GitHub (see update()).
"""
import json
import os
import shutil
import subprocess
import sys
import threading
import time
import urllib.error
import urllib.parse
import urllib.request
import webbrowser
from http.server import SimpleHTTPRequestHandler, ThreadingHTTPServer

HOST, PORT = '127.0.0.1', 8770
URL = f'http://{HOST}:{PORT}/'
APP_DIR = os.path.dirname(os.path.abspath(__file__))
CONFIG = os.path.join(APP_DIR, 'adr-folder.txt')
# Copies of the workbook taken before each overwrite; kept out of the ADR folder.
BACKUP_DIR = os.path.join(APP_DIR, 'backups')
BACKUPS_KEPT = 30
# The hospital's drug information system (intranet). The TFDA page looks up
# drug codes, product names and 健保碼 there; only its two search pages are
# reachable through this helper.
HOSPITAL_DRUG_DB = 'http://10.10.4.25:8080/Key_Search/'
HOSPITAL_PAGES = {'P3', 'P4'}
# Result of the last self-update, for when an update does not show up.
UPDATE_LOG = os.path.join(APP_DIR, 'update.log')
# No console window flashing up for git under pythonw.
NO_WINDOW = getattr(subprocess, 'CREATE_NO_WINDOW', 0)


def backup(path):
    os.makedirs(BACKUP_DIR, exist_ok=True)
    stem, ext = os.path.splitext(os.path.basename(path))
    shutil.copy2(path, os.path.join(BACKUP_DIR, f'{stem}_{time.strftime("%Y%m%d-%H%M%S")}{ext}'))
    old = sorted(f for f in os.listdir(BACKUP_DIR) if f.startswith(stem + '_'))
    for f in old[:-BACKUPS_KEPT]:
        os.remove(os.path.join(BACKUP_DIR, f))


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
            if url.path == '/api/hospital':
                return self.hospital(url.query)
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

    def hospital(self, query):
        params = urllib.parse.parse_qs(query)
        page = params.get('page', [''])[0]
        if page not in HOSPITAL_PAGES:
            return self.reply(403, 'forbidden')
        q = {'query': params.get('q', [''])[0]}
        if page == 'P3':
            # Search by drug code, generic name and product name only.
            q.update(Code='true', PharNm='true', TrdNm='true', Pharm_e='false', Indications='false', ATC='false')
        try:
            with urllib.request.urlopen(HOSPITAL_DRUG_DB + page + '?' + urllib.parse.urlencode(q), timeout=15) as res:
                return self.reply(200, res.read(), 'text/html; charset=utf-8')
        except (urllib.error.URLError, OSError) as e:
            return self.reply(502, f'院內藥品資料庫連不上：{e}')

    def do_POST(self):
        # A newer copy of this helper asks the running one to make way.
        if urllib.parse.urlsplit(self.path).path != '/api/quit' or not self.allowed():
            return self.reply(403, 'forbidden')
        self.reply(200)
        threading.Thread(target=self.server.shutdown, daemon=True).start()

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
                if path.lower().endswith('.xlsx') and os.path.isfile(path):
                    backup(path)
                os.replace(tmp, path)
            except PermissionError:
                os.remove(tmp)
                return self.reply(409, '檔案被其他程式開啟中（請先關閉 Excel）')
            return self.reply(200)
        except Exception as e:
            return self.reply(500, f'{type(e).__name__}: {e}')


class Server(ThreadingHTTPServer):
    # On Windows, SO_REUSEADDR lets a second copy bind the same port, so an
    # old copy could keep answering after an update. Without it the second
    # copy fails to bind and just opens the browser.
    allow_reuse_address = False


# Browsers that can show the tool in its own app window (--app), first found wins.
APP_BROWSERS = [
    os.path.join(os.environ.get(root, ''), *path)
    for root in ('PROGRAMFILES', 'PROGRAMFILES(X86)', 'LOCALAPPDATA')
    for path in (('Google', 'Chrome', 'Application', 'chrome.exe'),
                 ('Microsoft', 'Edge', 'Application', 'msedge.exe'))
]


def open_app():
    """Shows the tool in its own window (no tabs or address bar), or in a
    normal browser tab when neither Chrome nor Edge is installed."""
    for exe in APP_BROWSERS:
        if os.path.isfile(exe):
            subprocess.Popen([exe, f'--app={URL}'])
            return
    webbrowser.open(URL)


def git(*args, timeout=10):
    env = dict(os.environ, GIT_TERMINAL_PROMPT='0', GCM_INTERACTIVE='never')
    return subprocess.run(['git', *args], cwd=APP_DIR, capture_output=True, text=True,
                          timeout=timeout, env=env, creationflags=NO_WINDOW)


def log_update(message):
    try:
        with open(UPDATE_LOG, 'w', encoding='utf-8') as f:
            f.write(f'{time.strftime("%Y-%m-%d %H:%M:%S")} {message}\n')
    except OSError:
        pass


def finish(message, server_changed=False):
    log_update(message)
    return server_changed


def update():
    """Brings this folder up to date with main on GitHub (fast-forward only),
    so changes merged there show up the next time the app is opened.
    Never stops the app from starting: skipped without git, off main, with
    local edits, or offline. Returns True when server.py itself changed."""
    if not os.path.isdir(os.path.join(APP_DIR, '.git')) or not shutil.which('git'):
        return finish('skipped: not a git checkout, or git is not installed')
    try:
        if git('rev-parse', '--abbrev-ref', 'HEAD').stdout.strip() != 'main':
            return finish('skipped: not on the main branch')
        if git('status', '--porcelain', '--untracked-files=no').stdout.strip():
            return finish('skipped: files were changed in this folder')
        before = git('rev-parse', 'HEAD').stdout.strip()
        pulled = git('pull', '--ff-only', '--quiet', 'origin', 'main', timeout=20)
        if pulled.returncode != 0:
            return finish('failed: ' + (pulled.stderr.strip() or 'git pull failed'))
        after = git('rev-parse', 'HEAD').stdout.strip()
        if before == after:
            return finish(f'up to date ({after[:7]})')
        changed = git('diff', '--name-only', before, after).stdout.split()
        return finish(f'updated {before[:7]} -> {after[:7]}: {", ".join(changed)}', 'server.py' in changed)
    except (OSError, subprocess.SubprocessError) as e:
        return finish(f'failed: {e}')


def replace_running():
    """Asks an older running copy to stop, so the updated one can serve."""
    try:
        req = urllib.request.Request(URL + 'api/quit', method='POST', headers={'X-ADR-Cards': '1'})
        urllib.request.urlopen(req, timeout=3).close()
    except (urllib.error.URLError, OSError):
        return None  # an old copy without /api/quit keeps running
    for _ in range(20):
        time.sleep(0.25)
        try:
            return Server((HOST, PORT), Handler)
        except OSError:
            pass
    return None


def main(open_browser=True, replace=False):
    if not replace and update():
        # server.py changed: run the new version (it replaces a running copy).
        args = [a for a in sys.argv[1:] if a != '--replace'] + ['--replace']
        subprocess.Popen([sys.executable, os.path.abspath(__file__), *args], cwd=APP_DIR, creationflags=NO_WINDOW)
        return
    try:
        server = Server((HOST, PORT), Handler)
    except OSError:
        server = replace_running() if replace else None
        if not server:
            # Already running: just show it. Pages are read from disk on
            # every request, so it shows the updated pages too.
            if open_browser:
                open_app()
            return
    if open_browser:
        open_app()
    server.serve_forever()
    server.server_close()


if __name__ == '__main__':
    # --no-browser: start quietly (used by the Windows startup shortcut).
    # --replace: started by an older copy after an update; take over from it.
    main(open_browser='--no-browser' not in sys.argv, replace='--replace' in sys.argv)
