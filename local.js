// When the page is served by server.py (desktop shortcut), the ADR folder is
// reached through its small HTTP API instead of the browser's folder picker.
// This exposes the same subset of the FileSystemDirectoryHandle interface
// that app.js uses, so the rest of the app does not care which one it gets.
(function (root) {
  'use strict';

  const HEADERS = { 'X-ADR-Cards': '1' };
  const q = path => '?path=' + encodeURIComponent(path);
  const join = (dir, name) => (dir ? dir + '/' + name : name);

  // The version (modified time) of each file as last read, so a save can
  // say what it is based on and the helper can refuse it if another page
  // saved in between (instead of silently dropping that save).
  const readVersion = new Map();

  // Saves on this page run one after another, never interleaved: each reads
  // the workbook, changes it and writes it back.
  let queue = Promise.resolve();
  function serial(fn) {
    const run = queue.then(fn, fn);
    queue = run.catch(() => {});
    return run;
  }

  // { exists, modified } of a file; an older helper answers 404 when missing.
  async function stat(path) {
    const res = await call('api/stat' + q(path));
    if (!res) return { exists: false };
    const info = res.headers.get('Content-Type') === 'application/json' ? await res.json() : {};
    return Object.assign({ exists: true }, info);
  }

  function notFound(path) {
    const e = new Error(path + ' not found');
    e.name = 'NotFoundError';
    return e;
  }

  async function call(url, opts = {}) {
    let res;
    try {
      res = await fetch(url, Object.assign({ cache: 'no-store', headers: HEADERS }, opts));
    } catch (e) {
      // The helper is not running (closed, or the PC restarted).
      const err = new Error('本機程式沒有在執行：請雙擊桌面上的「ADR 通報卡」，再按一次存入。');
      err.name = 'LocalServerDown';
      throw err;
    }
    if (res.status === 404) return null;
    if (!res.ok) {
      const e = new Error(await res.text());
      e.name = 'LocalServerError';
      throw e;
    }
    return res;
  }

  function fileHandle(path, name) {
    return {
      kind: 'file',
      name,
      async getFile() {
        const res = await call('api/file' + q(path));
        if (!res) throw notFound(path);
        const lastModified = Number(res.headers.get('X-File-Modified')) || Date.now();
        if (res.headers.get('X-File-Modified')) readVersion.set(path, res.headers.get('X-File-Modified'));
        return new File([await res.blob()], name, { lastModified });
      },
      // The day it was last saved, without downloading it.
      async modified() {
        const info = await stat(path);
        if (!info.exists) throw notFound(path);
        return info.modified;
      },
      async createWritable() {
        const parts = [];
        return {
          async write(data) { parts.push(data); },
          async close() {
            const headers = Object.assign({}, HEADERS, readVersion.has(path) ? { 'X-Based-On': readVersion.get(path) } : {});
            const res = await call('api/file' + q(path), { method: 'PUT', body: new Blob(parts), headers });
            const saved = res && res.headers.get('Content-Type') === 'application/json' ? await res.json() : null;
            if (saved && saved.modified) readVersion.set(path, String(saved.modified));
            else readVersion.delete(path);
          },
        };
      },
    };
  }

  function dirHandle(path, name) {
    return {
      kind: 'directory',
      name,
      local: true,
      async queryPermission() { return 'granted'; },
      async requestPermission() { return 'granted'; },
      async getDirectoryHandle(child, { create = false } = {}) {
        const p = join(path, child);
        const res = await call('api/dir' + q(p) + (create ? '&create=1' : ''));
        // An older helper answers 404; this one says {"exists": false}.
        const info = res && res.headers.get('Content-Type') === 'application/json' ? await res.json() : { exists: !!res };
        if (!info.exists) throw notFound(p);
        return dirHandle(p, child);
      },
      async getFileHandle(child, { create = false } = {}) {
        const p = join(path, child);
        if (!create && !(await stat(p)).exists) throw notFound(p);
        return fileHandle(p, child);
      },
      async removeEntry(child) {
        const p = join(path, child);
        if (!await call('api/file' + q(p), { method: 'DELETE' })) throw notFound(p);
      },
      async *keys() {
        const res = await call('api/list' + q(path));
        for (const n of res ? await res.json() : []) yield n;
      },
    };
  }

  // Resolves to the ADR folder handle, or null when not served by server.py.
  async function detect() {
    try {
      const res = await call('api/info');
      if (!res) return null;
      return dirHandle('', (await res.json()).name);
    } catch (e) {
      return null;
    }
  }

  // Shows which version the desktop helper serves, and whether its last
  // self-update worked, at the end of the page links.
  async function showVersion() {
    let info;
    try {
      const res = await call('api/version');
      if (!res) return;
      info = await res.json();
    } catch (e) {
      return;  // not served by server.py
    }
    const nav = document.querySelector('.nav');
    if (!nav || !info.commit) return;
    const span = document.createElement('span');
    const failed = /略過|失敗/.test(info.update);
    span.className = 'version' + (failed ? ' warn' : '');
    span.textContent = `版本 ${info.commit}` + (failed ? `　${info.update.replace(/^\S+ \S+ /, '')}` : '');
    span.title = info.update || '';
    nav.appendChild(span);
  }
  if (location.protocol.startsWith('http')) {
    if (document.readyState === 'loading') document.addEventListener('DOMContentLoaded', showVersion);
    else showVersion();
  }

  // ADR通報.xlsx as a File, never one too big to handle in the browser: an
  // old version could leave it over 30 MB, and reading or saving that
  // crashed the page (STATUS_BREAKPOINT). A big one is shrunk by the desktop
  // helper first (/api/repair); if that is not possible, a clear error.
  const BIG = 4 * 1024 * 1024;
  async function workbookFile(handle) {
    let file = await handle.getFile();
    if (file.size <= BIG) return file;
    try { await call('api/repair', { method: 'POST' }); } catch (e) { /* older helper: no repair */ }
    file = await handle.getFile();
    if (file.size <= BIG) return file;
    const err = new Error(`ADR通報.xlsx 有 ${(file.size / 1048576).toFixed(0)} MB，太大，在瀏覽器裡存取會當掉。` +
      '請先關閉 Excel，再關掉 ADR 通報卡視窗、從桌面捷徑重新開啟（會自動把檔案縮小，原檔備份在 backups）。');
    err.name = 'WorkbookTooBig';
    throw err;
  }

  root.ADRLocal = { detect, workbookFile, serial };
})(window);
