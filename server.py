"""Local launcher for ADR Cards.

Serves the tool at http://127.0.0.1:8770/ and gives it direct read/write access
to the ADR folder named in adr-folder.txt, so the page never asks for a folder.
Started by the desktop shortcut (pythonw, no console window), and quietly at
Windows sign-in. If it is already running, it just opens the app window.
Each start first updates this folder from GitHub (see update()).
"""
import glob
import json
import os
import re
import shutil
import subprocess
import sys
import tempfile
import threading
import time
import urllib.error
import urllib.parse
import urllib.request
import webbrowser
import zipfile
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
# Repairs, refused saves and backups that doubled: kept (last 200 lines),
# not overwritten at each start like update.log, so they can be noticed.
NOTES_LOG = os.path.join(APP_DIR, 'notes.log')
NOTES_KEPT = 200
# No console window flashing up for git under pythonw.
NO_WINDOW = getattr(subprocess, 'CREATE_NO_WINDOW', 0)


def backup(path):
    os.makedirs(BACKUP_DIR, exist_ok=True)
    stem, ext = os.path.splitext(os.path.basename(path))
    previous = sorted(f for f in os.listdir(BACKUP_DIR) if f.startswith(stem + '_'))
    shutil.copy2(path, os.path.join(BACKUP_DIR, f'{stem}_{time.strftime("%Y%m%d-%H%M%S")}{ext}'))
    # The two bloat bugs showed in the backups for weeks: say so when one doubles.
    if previous:
        last = os.path.getsize(os.path.join(BACKUP_DIR, previous[-1]))
        now = os.path.getsize(path)
        if now > 100_000 and now > 2 * last:
            log_repair(f'注意：{os.path.basename(path)} 由 {last // 1024} KB 變成 {now // 1024} KB（比上一份備份大一倍以上）')
    old = sorted(f for f in os.listdir(BACKUP_DIR) if f.startswith(stem + '_'))
    for f in old[:-BACKUPS_KEPT]:
        os.remove(os.path.join(BACKUP_DIR, f))


WORKBOOK = 'ADR通報.xlsx'
# Column widths Excel can hold per sheet; a sheet listing far more than its
# used columns was written by an old version of the page (one entry for all
# 16,384 columns per sheet, uncompressed): ~1 MB per month sheet, enough to
# make every save in the browser crash.
COLS = re.compile(rb'<cols>(.*?)</cols>', re.S)
COL = re.compile(rb'<col\b[^>]*?\bmin="(\d+)"[^>]*?\bmax="(\d+)"[^>]*/>')
KEEP_COLS = 26


def slim_cols(xml):
    """Keeps the width entries of columns A-Z only; None if nothing to cut."""
    m = COLS.search(xml)
    if not m or m.group(1).count(b'<col') <= KEEP_COLS * 2:
        return None
    kept = []
    for col in COL.finditer(m.group(1)):
        lo, hi = int(col.group(1)), int(col.group(2))
        if lo > KEEP_COLS:
            continue
        entry = col.group(0)
        if hi > KEEP_COLS:
            entry = entry.replace(b'max="%d"' % hi, b'max="%d"' % KEEP_COLS, 1)
        kept.append(entry)
    body = b'<cols>' + b''.join(kept) + b'</cols>' if kept else b''
    return xml[:m.start()] + body + xml[m.end():]


# Saves of ADR通報.xlsx that grow it suddenly are refused: no real batch of
# cases adds this much, so it means a bug like the two that bloated the file
# save after save (column widths, then the theme) until the page crashed.
WORKBOOK_CAP = 5_000_000


def too_big(path, new_size):
    """Why a save of ADR通報.xlsx should be refused, or ''."""
    if os.path.basename(path) != WORKBOOK:
        return ''
    if new_size > WORKBOOK_CAP:
        return f'{WORKBOOK} 會變成 {new_size // 1048576} MB，超過上限'
    if os.path.isfile(path):
        old = os.path.getsize(path)
        if new_size > old * 1.5 + 200_000:
            return f'{WORKBOOK} 會由 {old // 1024} KB 突然變成 {new_size // 1024} KB'
    return ''


def unmangle(xml):
    """Undoes the theme garbling (UTF-8 bytes read as latin-1 and encoded
    again, once per save), one run of non-ASCII text at a time."""
    try:
        text = xml.decode('utf-8')
    except UnicodeDecodeError:
        return xml

    def fix(m):
        run = m.group(0)
        for _ in range(64):
            try:
                run = run.encode('latin-1').decode('utf-8')
            except (UnicodeEncodeError, UnicodeDecodeError):
                break
        return run
    return re.sub('[\u0080-\uffff]+', fix, text).encode('utf-8')


def repair_workbook():
    """Cuts the bloat out of ADR通報.xlsx (see COLS), after a backup, and
    stores it compressed. Only column widths change; every cell stays."""
    try:
        path = os.path.join(read_adr_dir(), WORKBOOK)
        if not os.path.isfile(path):
            return
        before = os.path.getsize(path)
        changed = False
        with zipfile.ZipFile(path) as src:
            parts = [(info, src.read(info.filename)) for info in src.infolist()]
        out = []
        for info, data in parts:
            if re.match(r'xl/worksheets/sheet\d+\.xml$', info.filename):
                slim = slim_cols(data)
                if slim is not None:
                    data, changed = slim, True
            elif re.match(r'xl/theme/theme\d+\.xml$', info.filename):
                fixed = unmangle(data)
                if fixed != data:
                    data, changed = fixed, True
            out.append((info.filename, data))
        stored = before > 2_000_000 and any(info.compress_type == zipfile.ZIP_STORED for info, _ in parts)
        if not changed and not stored:
            return
        with WRITE_LOCK:
            fd, tmp = tempfile.mkstemp(dir=os.path.dirname(path), suffix='.adrtmp')
            os.close(fd)
            with zipfile.ZipFile(tmp, 'w', zipfile.ZIP_DEFLATED) as dst:
                for name, data in out:
                    dst.writestr(name, data)
            backup(path)
            os.replace(tmp, path)
        return log_repair(f'{WORKBOOK} 已瘦身：{before // 1024} KB → {os.path.getsize(path) // 1024} KB（備份在 backups）')
    except PermissionError:
        return log_repair(f'{WORKBOOK} 需要修復但檔案正被其他程式使用（Excel 或雲端硬碟同步）：稍後重新開啟 ADR 通報卡會再試')
    except Exception as e:  # never stop the app from starting
        return log_repair(f'{WORKBOOK} 瘦身失敗：{type(e).__name__}: {e}')


# One write to the ADR folder at a time (pages save in parallel threads).
WRITE_LOCK = threading.RLock()


def mtime_ms(path):
    return int(os.path.getmtime(path) * 1000)


def log_repair(message):
    try:
        try:
            with open(NOTES_LOG, encoding='utf-8') as f:
                lines = f.read().splitlines()
        except OSError:
            lines = []
        lines.append(f'{time.strftime("%Y-%m-%d %H:%M:%S")} {message}')
        with open(NOTES_LOG, 'w', encoding='utf-8') as f:
            f.write('\n'.join(lines[-NOTES_KEPT:]) + '\n')
    except OSError:
        pass
    return message


def recent_warning(days=7):
    """The latest warning (refused save, doubled backup, failed repair) of
    the last few days, for the page header; '' if none."""
    try:
        with open(NOTES_LOG, encoding='utf-8') as f:
            lines = f.read().splitlines()
    except OSError:
        return ''
    since = time.strftime('%Y-%m-%d', time.localtime(time.time() - days * 86400))
    for line in reversed(lines):
        if line[:10] < since:
            break
        if re.search(r'注意|存檔異常|突然變成|超過上限|失敗|沒辦法|正被其他程式使用', line):
            return line
    return ''


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

    # Never served as static files: workbook backups (patient data), the
    # git folder, logs and this helper's own settings.
    PRIVATE = re.compile(r'^/(backups|\.git|\.claude|__pycache__|tests|tools)(/|$)|\.(log|txt|py|adrtmp)$', re.I)

    def do_GET(self):
        url = urllib.parse.urlsplit(self.path)
        if not url.path.startswith('/api/'):
            # Only this app, under its own address: a page on another site that
            # points its name at 127.0.0.1 (DNS rebinding) gets nothing.
            path = urllib.parse.unquote(url.path)
            if self.headers.get('Host') != f'{HOST}:{PORT}' or self.PRIVATE.search(path):
                return self.reply(403, 'forbidden')
            return super().do_GET()
        if not self.allowed():
            return self.reply(403, 'forbidden')
        try:
            if url.path == '/api/info':
                base = read_adr_dir()
                if not os.path.isdir(base):
                    return self.reply(404, f'ADR folder not found: {base}')
                return self.reply(200, json.dumps({'name': os.path.basename(base)}), 'application/json')
            if url.path == '/api/version':
                return self.reply(200, json.dumps(version(), ensure_ascii=False), 'application/json')
            if url.path == '/api/hospital':
                return self.hospital(url.query)
            path = self.target(url.query)
            if url.path == '/api/stat':
                if not os.path.isfile(path):
                    # Asking whether a file exists yet is not an error.
                    return self.reply(200, '{"exists": false}', 'application/json')
                # When it was last saved, without sending the file.
                return self.reply(200, json.dumps({'exists': True, 'modified': mtime_ms(path)}), 'application/json')
            if url.path == '/api/dir':
                if not os.path.isdir(path):
                    if 'create=1' not in url.query:
                        # Pages probe optional folders (ADR_cards): not an
                        # error, so no 404 in the browser's console.
                        return self.reply(200, '{"exists": false}', 'application/json')
                    os.makedirs(path, exist_ok=True)
                return self.reply(200, '{"exists": true}', 'application/json')
            if url.path == '/api/list':
                names = sorted(os.listdir(path)) if os.path.isdir(path) else []
                return self.reply(200, json.dumps(names), 'application/json')
            if url.path == '/api/file':
                if not os.path.isfile(path):
                    return self.reply(404)
                with open(path, 'rb') as f:
                    data = f.read()
                # The page reads the date a file was saved (e.g. ADR_<識別碼>_通報.pdf).
                self.send_response(200)
                self.send_header('Content-Type', 'application/octet-stream')
                self.send_header('Content-Length', str(len(data)))
                self.send_header('X-File-Modified', str(mtime_ms(path)))
                self.end_headers()
                self.wfile.write(data)
                return
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
        path = urllib.parse.urlsplit(self.path).path
        if not self.allowed():
            return self.reply(403, 'forbidden')
        if path == '/api/repair':
            # The page found ADR通報.xlsx too big to handle: shrink it now.
            return self.reply(200, repair_workbook() or '', 'text/plain; charset=utf-8')
        if path != '/api/quit':
            return self.reply(403, 'forbidden')
        # A newer copy of this helper asks the running one to make way.
        self.reply(200)
        threading.Thread(target=self.server.shutdown, daemon=True).start()

    def do_DELETE(self):
        # Only card images: a case that moved to another 識別碼 leaves its old one.
        url = urllib.parse.urlsplit(self.path)
        if url.path != '/api/file' or not self.allowed():
            return self.reply(403, 'forbidden')
        try:
            path = self.target(url.query)
            if not re.fullmatch(r'ADR_\d{7}\.png', os.path.basename(path), re.I):
                return self.reply(403, 'only card images can be deleted')
            if not os.path.isfile(path):
                return self.reply(404)
            os.remove(path)
            return self.reply(200)
        except PermissionError:
            return self.reply(409, '檔案被其他程式開啟中')
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
            with WRITE_LOCK:
                # The page says which version of the file its change is based
                # on; if another page saved in between, refuse rather than
                # silently drop that save.
                base = self.headers.get('X-Based-On')
                if base and os.path.isfile(path) and str(mtime_ms(path)) != base:
                    return self.reply(409, '檔案剛被另一個頁面存過，這次沒有存入：請再按一次（會用最新的檔案）。')
                refused = too_big(path, len(data))
                if refused:
                    log_repair(refused)
                    return self.reply(409, '存檔異常：' + refused + '。已取消存檔並保留原檔，請聯絡管理者。')
                fd, tmp = tempfile.mkstemp(dir=os.path.dirname(path), suffix='.adrtmp')
                with os.fdopen(fd, 'wb') as f:
                    f.write(data)
                try:
                    if path.lower().endswith('.xlsx') and os.path.isfile(path):
                        backup(path)
                    os.replace(tmp, path)
                except PermissionError:
                    os.remove(tmp)
                    return self.reply(409, '檔案被其他程式開啟中（請先關閉 Excel）')
            return self.reply(200, json.dumps({'modified': mtime_ms(path)}), 'application/json')
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


# Where Git usually is on Windows when it is not on the PATH the shortcut
# sees (e.g. Git that came with Cmder or GitHub Desktop).
GIT_CANDIDATES = [
    os.path.join(os.environ.get(root, ''), *path)
    for root in ('PROGRAMFILES', 'PROGRAMFILES(X86)', 'LOCALAPPDATA', 'CMDER_ROOT', 'USERPROFILE')
    for path in (('Git', 'cmd', 'git.exe'), ('Programs', 'Git', 'cmd', 'git.exe'),
                 ('vendor', 'git-for-windows', 'cmd', 'git.exe'),
                 ('cmder', 'vendor', 'git-for-windows', 'cmd', 'git.exe'))
] + [os.path.join(drive, *path) for drive in ('C:\\', 'D:\\')
     for path in (('cmder', 'vendor', 'git-for-windows', 'cmd', 'git.exe'),
                  ('tools', 'cmder', 'vendor', 'git-for-windows', 'cmd', 'git.exe'))]


GIT = None  # path to git.exe, found by update()


def find_git():
    found = shutil.which('git')
    if found:
        return found
    desktop = glob.glob(os.path.join(os.environ.get('LOCALAPPDATA', ''), 'GitHubDesktop', 'app-*',
                                     'resources', 'app', 'git', 'cmd', 'git.exe'))
    return next((p for p in GIT_CANDIDATES + sorted(desktop, reverse=True) if os.path.isfile(p)), None)


def git(*args, timeout=10):
    env = dict(os.environ, GIT_TERMINAL_PROMPT='0', GCM_INTERACTIVE='never')
    return subprocess.run([GIT or 'git', *args], cwd=APP_DIR, capture_output=True, text=True,
                          encoding='utf-8', errors='replace', timeout=timeout, env=env, creationflags=NO_WINDOW)



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
    Never stops the app from starting: skipped without git, with local
    edits, or offline. Returns True when server.py itself changed.
    The result is in update.log and shown on the pages (/api/version)."""
    global GIT
    if not os.path.isdir(os.path.join(APP_DIR, '.git')):
        return finish('略過更新：這個資料夾不是 git 下載的（沒有 .git）')
    GIT = find_git()
    if not GIT:
        return finish('略過更新：找不到 git（請安裝 Git for Windows）')
    try:
        if git('status', '--porcelain', '--untracked-files=no').stdout.strip():
            return finish('略過更新：資料夾裡有檔案被修改過，請執行 git status 查看')
        before = git('rev-parse', 'HEAD').stdout.strip()
        pulled = git('pull', '--ff-only', '--quiet', 'origin', 'main', timeout=30)
        if pulled.returncode != 0:
            return finish('更新失敗：' + (pulled.stderr.strip() or 'git pull 失敗').replace('\n', ' '))
        after = git('rev-parse', 'HEAD').stdout.strip()
        if before == after:
            return finish(f'已是最新版（{after[:7]}）')
        changed = git('diff', '--name-only', before, after).stdout.split()
        return finish(f'已更新 {before[:7]} → {after[:7]}：{"、".join(changed)}', 'server.py' in changed)
    except (OSError, subprocess.SubprocessError) as e:
        return finish(f'更新失敗：{e}')


def version():
    """The commit being served and the last update result, for the pages."""
    head = ''
    try:
        with open(os.path.join(APP_DIR, '.git', 'HEAD'), encoding='utf-8') as f:
            ref = f.read().strip()
        if ref.startswith('ref: '):
            ref_path = os.path.join(APP_DIR, '.git', *ref[5:].split('/'))
            if os.path.isfile(ref_path):
                with open(ref_path, encoding='utf-8') as f:
                    head = f.read().strip()
            else:
                with open(os.path.join(APP_DIR, '.git', 'packed-refs'), encoding='utf-8') as f:
                    head = next((l.split()[0] for l in f if l.strip().endswith(' ' + ref[5:])), '')
        else:
            head = ref
    except OSError:
        pass
    try:
        with open(UPDATE_LOG, encoding='utf-8') as f:
            last = f.read().strip()
    except OSError:
        last = ''
    return {'commit': head[:7], 'update': last, 'warning': recent_warning()}


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
    repair_workbook()
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
