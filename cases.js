// Phone page for checking ADR cases. The cases come from the link made on
// the PC's 案例總覽 (手機查看): they are after "#" and never reach the website.
// The last list is kept on this phone, so the page opens with it next time.
// Without a link, ADR通報.xlsx (and 通報表 exports) can be picked instead.
(function () {
  'use strict';

  const R = window.ADRReview;
  const $ = id => document.getElementById(id);
  const KEEP_KEY = 'adr-phone-cases';
  const LEVEL_RANK = { info: 1, mid: 2, high: 3 };

  let data = null;        // { from, to, at, cases }
  let filter = 'all';
  const open = new Set(); // 識別碼 shown with details

  function el(tag, attrs = {}, ...children) {
    const e = document.createElement(tag);
    for (const [k, v] of Object.entries(attrs)) {
      if (k === 'class') e.className = v;
      else if (k.startsWith('on')) e.addEventListener(k.slice(2), v);
      else e.setAttribute(k, v);
    }
    e.append(...children.filter(c => c != null && c !== false && c !== ''));
    return e;
  }

  const rocMonth = key => `${key.slice(0, 3)} 年 ${Number(key.slice(3))} 月`;
  const topLevel = c => c.hints.reduce((best, h) => (LEVEL_RANK[h.level] > LEVEL_RANK[best] ? h.level : best), 'info');

  // ---- rendering ----

  function matches(c, q) {
    if (filter === 'marked' && !c.marked) return false;
    if (filter === 'hints' && !c.hints.length) return false;
    if (!q) return true;
    return [c.id, c.drug, c.symptom, c.writer, c.date].join(' ').toLowerCase().includes(q);
  }

  function render() {
    const { from, to, at, cases } = data;
    $('period').textContent = from === to ? rocMonth(from) : `${rocMonth(from)} – ${rocMonth(to)}`;
    $('updated').textContent = at ? `資料時間：${at}（在電腦重新按「手機查看」可更新）` : '';
    $('controls').hidden = false;
    $('start').hidden = true;
    $('nAll').textContent = cases.length;
    $('nMarked').textContent = cases.filter(c => c.marked).length;
    $('nHints').textContent = cases.filter(c => c.hints.length).length;
    document.querySelectorAll('.cs-filter').forEach(b => b.classList.toggle('on', b.dataset.filter === filter));

    const q = $('search').value.trim().toLowerCase();
    const shown = cases.filter(c => matches(c, q));
    $('empty').hidden = shown.length > 0;
    $('empty').textContent = cases.length ? '沒有符合的案例。' : '這段期間沒有案例。';
    // Newest month first, then by 識別碼; a heading per month.
    const items = [];
    let month = null;
    shown.slice().sort((a, b) => b.id.slice(0, 5).localeCompare(a.id.slice(0, 5)) || a.id.localeCompare(b.id)).forEach(c => {
      if (c.id.slice(0, 5) !== month) {
        month = c.id.slice(0, 5);
        items.push(el('li', { class: 'cs-month' }, rocMonth(month)));
      }
      items.push(row(c));
    });
    $('list').replaceChildren(...items);
  }

  function row(c) {
    const isOpen = open.has(c.id);
    const head = el('button', {
      type: 'button', class: 'cs-row-head', 'aria-expanded': isOpen ? 'true' : 'false',
      onclick: () => { if (isOpen) open.delete(c.id); else open.add(c.id); render(); },
    },
    el('span', { class: 'cs-line1' },
      el('b', { class: 'cs-id' }, c.id),
      el('span', { class: 'cs-date' }, c.date),
      c.marked ? el('span', { class: 'cs-star' }, '★ 值得注意') : null,
      c.hints.length ? el('span', { class: `cs-dot ${topLevel(c)}`, title: '有自動提示' }) : null),
    el('span', { class: 'cs-sub' }, c.drug, el('span', { class: 'cs-symptom' }, ' · ' + c.symptom)));
    const li = el('li', { class: 'cs-row' + (c.marked ? ' marked' : '') + (isOpen ? ' open' : '') }, head);
    if (isOpen) {
      const line = (label, text) => (text ? el('div', { class: 'cs-fact' }, el('span', {}, label), el('span', {}, text)) : null);
      li.append(el('div', { class: 'cs-details' },
        line('藥品', c.drug),
        line('症狀', c.symptom),
        line('嚴重度', c.severity),
        line('嚴重性', c.serious),
        line('後續結果', c.outcome || '（無通報表）'),
        line('撰寫人', c.writer || '未指定'),
        c.hints.length ? el('div', { class: 'cs-hints' }, ...c.hints.map(h => el('span', { class: 'flag ' + h.level }, h.label))) : null));
    }
    return li;
  }

  function show(next) {
    data = next;
    open.clear();
    render();
  }

  // ---- loading ----

  function keep(text) {
    try { localStorage.setItem(KEEP_KEY, text); } catch (e) { /* only a convenience */ }
  }
  function kept() {
    try { return localStorage.getItem(KEEP_KEY); } catch (e) { return null; }
  }

  async function fromLink() {
    const text = location.hash.slice(1);
    if (text) {
      try {
        show(R.unpackCases(await window.ADRPhoneLink.decode(text)));
        keep(text);
        // Keep the data out of the address bar (and so out of a shared link).
        history.replaceState(null, '', location.pathname);
        return;
      } catch (e) {
        $('pickStatus').textContent = '連結內容無法讀取：' + e.message;
      }
    }
    const last = kept();
    if (last) {
      try { show(R.unpackCases(await window.ADRPhoneLink.decode(last))); return; } catch (e) { /* fall through */ }
    }
    $('start').hidden = false;
  }

  function loadScript(src) {
    return new Promise((resolve, reject) => {
      const s = document.createElement('script');
      s.src = src;
      s.onload = resolve;
      s.onerror = () => reject(new Error('無法載入 ' + src));
      document.head.appendChild(s);
    });
  }

  // ADR通報.xlsx and/or 通報表 exports picked on the phone: the latest
  // month's cases, flagged the same way as on the PC.
  async function fromFiles(files) {
    $('pickStatus').textContent = '讀取中…';
    if (!window.XLSX) await loadScript('lib/xlsx.full.min.js?v=20261030');
    let cases = [];
    const forms = [];
    for (const f of files) {
      const bytes = new Uint8Array(await f.arrayBuffer());
      if (/通報表/.test(f.name)) {
        const wb = XLSX.read(bytes, { type: 'array' });
        forms.push(wb.SheetNames.flatMap(n => window.ADRForms.readCases(XLSX.utils.sheet_to_json(wb.Sheets[n], { header: 1, raw: false, defval: '' }))));
      } else {
        cases = cases.concat(R.workbookCases(XLSX, bytes));
      }
    }
    const all = R.joinForms(cases, window.ADRForms.mergeCases(forms));
    const months = [...new Set(all.map(c => c.month).filter(Boolean))].sort();
    if (!months.length) { $('pickStatus').textContent = '檔案裡找不到案例。'; return; }
    const to = months[months.length - 1];
    const from = months[Math.max(0, months.length - 3)];
    const flagged = R.flagCases(all.filter(c => c.month >= from && c.month <= to));
    const now = new Date();
    const packed = R.packCases(flagged, { from, to, at: `${window.ADRParse.formatDate(now)} ${String(now.getHours()).padStart(2, '0')}:${String(now.getMinutes()).padStart(2, '0')}` });
    show(R.unpackCases(packed));
    keep(await window.ADRPhoneLink.encode(packed));
  }

  // ---- wiring ----

  document.querySelectorAll('.cs-filter').forEach(b => b.addEventListener('click', () => {
    filter = b.dataset.filter;
    render();
    window.scrollTo({ top: 0 });
  }));
  $('search').addEventListener('input', render);
  $('fileInput').addEventListener('change', e => {
    if (e.target.files.length) fromFiles([...e.target.files]).catch(err => { $('pickStatus').textContent = '讀取失敗：' + err.message; });
    e.target.value = '';
  });
  $('toTop').addEventListener('click', () => window.scrollTo({ top: 0, behavior: 'smooth' }));
  window.addEventListener('scroll', () => { $('toTop').hidden = window.scrollY < 600; }, { passive: true });
  window.addEventListener('hashchange', fromLink);
  fromLink();
})();
