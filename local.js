// When the page is served by server.py (desktop shortcut), the ADR folder is
// reached through its small HTTP API instead of the browser's folder picker.
// This exposes the same subset of the FileSystemDirectoryHandle interface
// that app.js uses, so the rest of the app does not care which one it gets.
(function (root) {
  'use strict';

  const HEADERS = { 'X-ADR-Cards': '1' };
  const q = path => '?path=' + encodeURIComponent(path);
  const join = (dir, name) => (dir ? dir + '/' + name : name);

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
        return res.blob();
      },
      async createWritable() {
        const parts = [];
        return {
          async write(data) { parts.push(data); },
          async close() { await call('api/file' + q(path), { method: 'PUT', body: new Blob(parts) }); },
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
        if (!await call('api/dir' + q(p) + (create ? '&create=1' : ''))) throw notFound(p);
        return dirHandle(p, child);
      },
      async getFileHandle(child, { create = false } = {}) {
        const p = join(path, child);
        if (!create && !await call('api/stat' + q(p))) throw notFound(p);
        return fileHandle(p, child);
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

  root.ADRLocal = { detect };
})(window);
