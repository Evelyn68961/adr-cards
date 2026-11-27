// Case overview page: reads ADR通報.xlsx and the report-form exports from the
// ADR folder (or picked files) and lists the cases of a chosen period. The
// pharmacist ticks the cases worth noting for the routine ADR meeting (saved
// to column K of ADR通報.xlsx); automatic hints (see review.js) suggest which.
(function () {
  'use strict';

  const R = window.ADRReview;
  const F = window.ADRForms;
  const P = window.ADRParse;
  const W = window.ADRWorkbook;
  const $ = id => document.getElementById(id);
  const WORKBOOK_NAME = 'ADR通報.xlsx';
  // Not Excel's ~$ lock files, which exist while a form is open.
  const FORM_FILE = /^(?!~\$).*通報表.*\.xlsx?$/i;
  const LEVEL_RANK = { none: 0, info: 1, mid: 2, high: 3 };

  let folder = null;
  let picked = null;       // { name, bytes } of a picked workbook (no folder)
  let all = [];            // joined cases (review.joinForms)
  let shown = [];
  let periodCases = [];    // flagged cases of the period, for the phone link          // flagged cases of the period, in table order
  const open = new Set();  // 識別碼 whose details row is expanded
  // Marks not yet in Excel: 識別碼 -> true (marked) / false (unmarked).
  const pendingMarks = new Map();
  let order = null;        // 識別碼 in table order; kept while ticking so rows don't jump
  let lastTicked = null;   // row index of the last tick, for Shift-click ranges
  let markTimer = null;
  let markSaving = false;

  function el(tag, attrs = {}, ...children) {
    const e = document.createElement(tag);
    for (const [k, v] of Object.entries(attrs)) {
      if (k === 'class') e.className = v;
      else if (k.startsWith('on')) e.addEventListener(k.slice(2), v);
      else if (v === true) e.setAttribute(k, '');
      else if (v !== false && v != null) e.setAttribute(k, v);
    }
    e.append(...children.filter(c => c != null && c !== false));
    return e;
  }

  function toast(text) {
    const t = $('toast');
    t.textContent = text;
    t.hidden = false;
    clearTimeout(toast.timer);
    toast.timer = setTimeout(() => { t.hidden = true; }, 1600);
  }

  const rocMonth = key => `${key.slice(0, 3)} 年 ${Number(key.slice(3))} 月`;
  // "11510" shifted by n months.
  function addMonths(key, n) {
    const total = Number(key.slice(0, 3)) * 12 + Number(key.slice(3)) - 1 + n;
    return String(Math.floor(total / 12)) + String(total % 12 + 1).padStart(2, '0');
  }

  // ---- loading ----

  const bytesOf = async file => new Uint8Array(await file.arrayBuffer());
  const formsFrom = bytes => {
    const wb = XLSX.read(bytes, { type: 'array' });
    return wb.SheetNames.flatMap(n => F.readCases(XLSX.utils.sheet_to_json(wb.Sheets[n], { header: 1, raw: false, defval: '' })));
  };

  // ADR通報.xlsx, then every report-form export (found by its columns, not
  // its name). PDFs are returned unread: they are read after the first
  // render (readPdfs), from the results kept on this PC when possible.
  async function loadFromFolder() {
    let cases = [];
    const names = [];
    try {
      const book = await (await folder.getFileHandle(WORKBOOK_NAME)).getFile();
      cases = R.workbookCases(XLSX, await bytesOf(book));
      names.push(WORKBOOK_NAME);
    } catch (e) {
      if (e.name !== 'NotFoundError') throw e;
    }
    const forms = [];
    const pdfs = [];
    const scan = async (dir, prefix) => {
      for await (const name of dir.keys()) {
        if (!window.ADRPdfForm.isFormFile(name)) continue;
        if (/\.pdf$/i.test(name)) { pdfs.push({ dir, name, path: prefix + name }); continue; }
        try {
          const found = formsFrom(await bytesOf(await (await dir.getFileHandle(name)).getFile()));
          if (found.length) { forms.push(found); names.push(prefix + name); }
        } catch (e) { /* not a readable Excel file: not a form */ }
      }
    };
    await scan(folder, '');
    for await (const name of folder.keys()) {
      if (/^ADR_\d{5}$/.test(name)) await scan(await folder.getDirectoryHandle(name), name + '\\');
    }
    return { cases, forms, names, pdfs };
  }

  async function readPdfs(pdfs, forms, names) {
    for (const f of pdfs) {
      try {
        const c = await window.ADRPdfForm.readCached(f.name, await bytesOf(await (await f.dir.getFileHandle(f.name)).getFile()));
        if (c && /^\d{7}$/.test(c.id)) { forms.push([c]); names.push(f.path); }
      } catch (e) { /* the PDF reader could not load: Excel exports still count */ }
    }
  }

  async function loadPicked(files) {
    const forms = [];
    const names = [];
    let book = null;
    for (const f of files) {
      const bytes = await bytesOf(f);
      if (/\.pdf$/i.test(f.name)) { const c = await window.ADRPdfForm.readCached(f.name, bytes).catch(() => null); if (c) forms.push([c]); }
      else if (FORM_FILE.test(f.name) || formsFrom(bytes).length) forms.push(formsFrom(bytes));
      else book = { name: f.name, bytes };
      names.push(f.name);
    }
    picked = book;
    pendingMarks.clear();
    setData({ cases: book ? R.workbookCases(XLSX, book.bytes) : [], forms: F.mergeCases(forms), names });
  }

  async function reload(resort = true) {
    if (picked) return;
    $('sourceStatus').textContent = '讀取中…';
    try {
      if (!folder) folder = await window.ADRLocal.detect();
      if (!folder) {
        $('sourceStatus').textContent = `請按「選擇檔案」，選取 ${WORKBOOK_NAME}（可同時選取「通報表」Excel）。`;
        return;
      }
      const data = await loadFromFolder();
      const show = keepOrder => setData({ cases: data.cases, forms: F.mergeCases(data.forms), names: data.names }, keepOrder);
      show(resort);
      if (data.pdfs.length) {
        await readPdfs(data.pdfs, data.forms, data.names);
        show(false);
      }
    } catch (e) {
      $('sourceStatus').textContent = '讀取失敗：' + e.message;
    }
  }

  function setData({ cases, forms, names }, resort = true) {
    all = R.joinForms(cases, forms);
    if (!names.length) {
      $('sourceStatus').textContent = `ADR 資料夾裡找不到 ${WORKBOOK_NAME} 或「通報表」Excel。`;
      return;
    }
    const withForm = all.filter(c => c.form).length;
    $('sourceStatus').textContent = `讀取自：${names.join('、')}（共 ${all.length} 案，${withForm} 案有通報表）`;
    const months = [...new Set(all.map(c => c.month).filter(Boolean))].sort().reverse();
    const keepFrom = $('fromMonth').value, keepTo = $('toMonth').value;
    ['fromMonth', 'toMonth'].forEach(id => {
      $(id).replaceChildren(...months.map(m => el('option', { value: m }, rocMonth(m))));
    });
    if (months.includes(keepFrom) && months.includes(keepTo)) {
      $('fromMonth').value = keepFrom;
      $('toMonth').value = keepTo;
    } else if (months.length) {
      $('fromMonth').value = $('toMonth').value = months[0];
    }
    $('overviewPanel').hidden = false;
    render(resort);
  }

  // ---- marking (column K of ADR通報.xlsx) ----

  function setMarks(ids, marked) {
    ids.forEach(id => pendingMarks.set(id, marked));
    clearTimeout(markTimer);
    markTimer = setTimeout(saveMarks, 600);
    $('markStatus').textContent = '儲存中…';
    render(false);
  }

  // Writes the queued marks into the workbook in one go. With picked files
  // there is no folder to write to: the marked copy is offered as a download.
  async function saveMarks() {
    if (!pendingMarks.size) return;
    if (markSaving) { markTimer = setTimeout(saveMarks, 600); return; }
    if (!folder && !picked) {
      $('markStatus').textContent = `沒有讀入 ${WORKBOOK_NAME}，標記無法儲存。`;
      return;
    }
    markSaving = true;
    const batch = new Map(pendingMarks);
    const cells = new Map([...batch].map(([id, marked]) => [id, marked ? W.MARKED : '']));
    try {
      const handle = folder && await folder.getFileHandle(WORKBOOK_NAME);
      const bytes = folder ? await bytesOf(await handle.getFile()) : picked.bytes;
      const result = W.setNotes(XLSX, bytes, cells);
      if (result.changed.length) {
        if (folder) {
          const w = await handle.createWritable();
          await w.write(result.bytes);
          await w.close();
        } else {
          picked.bytes = result.bytes;
        }
      }
      batch.forEach((v, id) => { if (pendingMarks.get(id) === v) pendingMarks.delete(id); });
      result.missing.forEach(id => pendingMarks.delete(id));
      const missing = result.missing.length ? `（${result.missing.join('、')} 不在 ${WORKBOOK_NAME} 中，無法標記）` : '';
      if (folder) {
        $('markStatus').textContent = `已存入 ${WORKBOOK_NAME} 的「值得注意」欄${missing}`;
        if (!pendingMarks.size) await keepStatus(() => reload(false));
      } else {
        showDownload(missing);
        if (!pendingMarks.size) {
          all = mergeMarks();
          render(false);
        }
      }
    } catch (e) {
      $('markStatus').textContent = `標記無法存入 ${WORKBOOK_NAME}（若檔案在 Excel 中開著，請先關閉）：${e.message}。回到此頁時會自動再試。`;
    } finally {
      markSaving = false;
    }
  }

  // After saving into a picked workbook, refresh each case's mark from it.
  function mergeMarks() {
    const byId = new Map(R.workbookCases(XLSX, picked.bytes).map(c => [c.id + '|' + c.chartNo, c]));
    return all.map(c => {
      const fresh = byId.get(c.id + '|' + c.chartNo);
      return fresh ? Object.assign({}, c, { marked: fresh.marked }) : c;
    });
  }

  function showDownload(missing) {
    const url = URL.createObjectURL(new Blob([picked.bytes], {
      type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
    }));
    if (showDownload.url) URL.revokeObjectURL(showDownload.url);
    showDownload.url = url;
    $('markStatus').replaceChildren(
      `標記已加入 ${picked.name}，但這是選取的檔案，請下載後覆蓋原檔${missing}：`,
      el('a', { href: url, download: picked.name }, `下載 ${picked.name}`),
      '。用桌面「ADR 通報卡」開啟本頁，標記會直接存入。');
  }

  async function keepStatus(fn) {
    const message = $('markStatus').textContent;
    await fn();
    $('markStatus').textContent = message;
  }

  const canMark = c => c.id && !c.formOnly;

  // A tick on row i; with Shift, every row from the last tick to this one
  // gets the same state.
  function tick(i, shift) {
    const c = shown[i];
    if (!canMark(c)) return;
    const marked = !c.notable;
    const [a, b] = shift && lastTicked != null ? [Math.min(lastTicked, i), Math.max(lastTicked, i)] : [i, i];
    lastTicked = i;
    setMarks(shown.slice(a, b + 1).filter(canMark).map(x => x.id), marked);
  }

  // ---- rendering ----

  function period() {
    let from = $('fromMonth').value, to = $('toMonth').value;
    if (from > to) [from, to] = [to, from];
    return { from, to };
  }

  const hintLevel = c => c.hints.reduce((best, f) => (LEVEL_RANK[f.level] > LEVEL_RANK[best] ? f.level : best), 'none');
  const rank = c => (c.notable ? 10 : 0) + LEVEL_RANK[hintLevel(c)];
  const rowKey = c => c.id + '|' + c.chartNo;

  // resort: re-apply the sort and the 只看已標記 filter. Ticking and saving
  // keep the current rows and order, so nothing moves under the mouse.
  function render(resort = true) {
    const { from, to } = period();
    const inPeriod = all.filter(c => c.month >= from && c.month <= to);
    const flagged = R.flagCases(inPeriod, { marks: pendingMarks });
    periodCases = flagged;
    showMeetingCount();
    if (!$('phonePanel').hidden) showPhoneLink();
    const sum = R.summarize(flagged);
    $('printTitle').textContent = `期間：${rocMonth(from)}${from === to ? '' : ' 至 ' + rocMonth(to)}`;
    $('statTotal').textContent = `${sum.total} 件`;
    $('statNotable').textContent = `${sum.notable} 件`;
    $('statSuggested').textContent = `${sum.suggested} 件`;
    $('statSerious').textContent = `${sum.serious} 件`;
    $('statNoForm').textContent = `${sum.noForm} 件`;
    const chips = list => list.length
      ? list.slice(0, 6).map(d => el('span', { class: 'ov-chip' + (d.n > 1 ? ' many' : '') }, `${d.name} ×${d.n}`))
      : [el('span', { class: 'hint' }, '—')];
    $('topDrugs').replaceChildren(...chips(sum.drugs));
    $('topSymptoms').replaceChildren(...chips(sum.symptoms));

    if (resort || !order) {
      const list = flagged.filter(c => !$('notableOnly').checked || c.notable);
      if ($('notableFirst').checked) list.sort((a, b) => rank(b) - rank(a) || a.id.localeCompare(b.id));
      order = list.map(rowKey);
      lastTicked = null;
    }
    const byKey = new Map(flagged.map(c => [rowKey(c), c]));
    shown = order.map(k => byKey.get(k)).filter(Boolean);
    $('emptyHint').hidden = shown.length > 0;
    $('emptyHint').textContent = flagged.length ? '這段期間還沒有標記為值得注意的案例。' : '這段期間沒有案例。';
    $('rows').replaceChildren(...shown.flatMap(rowsFor));
  }

  function outcomeText(c) {
    if (!c.form) return el('span', { class: 'hint' }, '無通報表');
    return c.form.outcome || c.form.outcomeRaw || '—';
  }

  function severityCell(c) {
    const serious = c.form && R.SERIOUS_LABEL[c.form.seriousness];
    return el('td', { 'data-label': '嚴重度' }, c.severity || '—',
      c.form && c.form.seriousness ? el('div', { class: 'ov-sub' + (serious ? ' serious' : '') }, serious || '非嚴重') : null);
  }

  function rowsFor(c, i) {
    const isOpen = open.has(c.id);
    const box = el('input', {
      type: 'checkbox', class: 'ov-check', 'aria-label': `${c.id} 值得注意`,
      disabled: !canMark(c),
      title: canMark(c) ? '' : `${WORKBOOK_NAME} 中沒有這個案例，無法標記`,
    });
    box.checked = c.notable;
    const details = el('button', {
      type: 'button', class: 'ov-more no-print', 'aria-expanded': isOpen ? 'true' : 'false',
      onclick: e => {
        e.stopPropagation();
        if (isOpen) open.delete(c.id); else open.add(c.id);
        render(false);
      },
    }, isOpen ? '收合' : '詳情');
    // Clicking anywhere on the row ticks it (Shift-click ticks a range).
    const tr = el('tr', {
      class: 'ov-row' + (c.notable ? ' marked' : '') + ` hint-${hintLevel(c)}` + (isOpen ? ' open' : '') + (canMark(c) ? '' : ' locked'),
      onmousedown: e => { if (e.shiftKey) e.preventDefault(); },  // no text selection on Shift-click
      onclick: e => tick(i, e.shiftKey),
    },
    el('td', { class: 'ov-check-cell', 'data-label': '值得注意' }, box, el('span', { class: 'print-only' }, c.notable ? '✓' : '')),
    el('td', { class: 'ov-id', 'data-label': '識別碼' }, c.id || '—'),
    el('td', { class: 'ov-date', 'data-label': '得知日期' }, c.date ? P.formatDate(c.date) : '—'),
    el('td', { class: 'ov-chart', 'data-label': '病歷號' }, c.chartNo),
    el('td', { class: 'ov-drug', 'data-label': '藥品' }, c.drug),
    el('td', { 'data-label': '症狀' }, c.symptom),
    severityCell(c),
    el('td', { 'data-label': '後續結果' }, outcomeText(c)),
    el('td', { 'data-label': '撰寫人' }, c.writer || el('span', { class: 'hint' }, '—')),
    el('td', { class: 'ov-hints', 'data-label': '自動提示' }, ...c.hints.map(f => el('span', { class: 'flag ' + f.level }, f.label))),
    el('td', { class: 'ov-more-cell no-print' }, details));
    if (!isOpen) return [tr];
    return [tr, detailRow(c)];
  }

  function detailRow(c) {
    const f = c.form;
    const item = (label, text) => (text ? el('div', { class: 'ov-detail-item' }, el('span', { class: 'tf-label' }, label), el('div', {}, text)) : null);
    const body = f
      ? [
        item('案件描述', f.description),
        item('處置', f.treatment),
        item('不良反應嚴重性', f.seriousnessRaw),
        item('後續結果', f.outcomeRaw),
        item('Naranjo', [f.naranjoScore, f.naranjo].filter(x => x != null && x !== '').join('　')),
        item('性別／年齡', [f.sex, f.age].filter(Boolean).join('　')),
        item('通報來源', c.source),
        item('通報者', f.reporter),
      ]
      : [item('通報來源', c.source), el('p', { class: 'hint' }, '尚無通報表，只有 HIS 資料。')];
    if (c.formOnly) body.push(el('p', { class: 'hint' }, `${WORKBOOK_NAME} 中沒有這個識別碼，資料取自通報表。`));
    return el('tr', { class: 'ov-detail' }, el('td', { colspan: '11' }, el('div', { class: 'ov-detail-grid' }, ...body)));
  }

  // ---- phone link (cases.html) ----

  // The phone opens the published site; the desktop helper is only on this PC.
  const PHONE_PAGE = location.hostname.endsWith('github.io')
    ? new URL('cases.html', location.href).href
    : 'https://evelyn68961.github.io/adr-cards/cases.html';
  // Most a QR code holds (version 40, low error correction, bytes).
  const QR_MAX = 2900;
  const QR_LIB = 'https://cdn.jsdelivr.net/npm/qrcode-generator@1.4.4/qrcode.js';
  let phoneUrl = '';
  let phoneToken = 0;

  function loadQrLib() {
    if (window.qrcode) return Promise.resolve();
    if (!loadQrLib.promise) {
      loadQrLib.promise = new Promise((resolve, reject) => {
        const s = document.createElement('script');
        s.src = QR_LIB;
        s.onload = resolve;
        s.onerror = () => { loadQrLib.promise = null; reject(new Error('QR code 元件無法載入（沒有網路？）')); };
        document.head.appendChild(s);
      });
    }
    return loadQrLib.promise;
  }

  async function showPhoneLink() {
    const token = ++phoneToken;
    const { from, to } = period();
    const now = new Date();
    const at = `${P.formatDate(now)} ${String(now.getHours()).padStart(2, '0')}:${String(now.getMinutes()).padStart(2, '0')}`;
    const url = PHONE_PAGE + '#' + await window.ADRPhoneLink.encode(R.packCases(periodCases, { from, to, at }));
    if (token !== phoneToken) return;
    phoneUrl = url;
    const qr = $('phoneQr');
    $('phoneNote').textContent = `${periodCases.length} 件案例（${rocMonth(from)}${from === to ? '' : ' 至 ' + rocMonth(to)}）。之後有新案例或勾選，請再按一次「手機查看」。`;
    if (url.length > QR_MAX) {
      qr.hidden = true;
      $('phoneNote').textContent = `案例太多，放不進 QR code：請把期間縮短（例如只選一個月），或按「複製連結」傳到手機。`;
      return;
    }
    try {
      await loadQrLib();
      if (token !== phoneToken) return;
      const code = window.qrcode(0, 'L');
      code.addData(url);
      code.make();
      qr.innerHTML = code.createSvgTag({ cellSize: 4, margin: 2, scalable: true });
      qr.hidden = false;
    } catch (e) {
      qr.hidden = true;
      $('phoneNote').textContent = `${e.message}。請按「複製連結」傳到手機。`;
    }
  }

  // ---- meeting list (例會 Excel) ----

  const REVIEWER_KEY = 'adr-meeting-reviewer';
  const gregorian = key => String(Number(key.slice(0, 3)) + 1911) + key.slice(3);   // "11508" -> "202608"

  function openMeeting() {
    try {
      const saved = JSON.parse(localStorage.getItem(REVIEWER_KEY) || '{}');
      if (!$('reviewer').value) $('reviewer').value = saved.name || '';
      if (!$('reviewerId').value) $('reviewerId').value = saved.id || '';
    } catch (e) { /* only a convenience */ }
    $('meetingPanel').hidden = false;
    showMeetingCount();
  }

  function showMeetingCount() {
    if ($('meetingPanel').hidden) return;
    const n = periodCases.filter(c => c.notable).length;
    $('meetingStatus').replaceChildren(n
      ? `這段期間已勾選 ${n} 件，會依識別碼排列。ICU 和通報日期請在 Excel 中確認。`
      : '這段期間還沒有勾選的案例：先按「勾選 Hartwig ≥4」或自己勾選。');
  }

  async function saveMeeting() {
    const picked = periodCases.filter(c => c.notable);
    if (!picked.length) { showMeetingCount(); return; }
    const reviewer = $('reviewer').value.trim();
    const reviewerId = $('reviewerId').value.trim().toUpperCase();
    try { localStorage.setItem(REVIEWER_KEY, JSON.stringify({ name: reviewer, id: reviewerId })); } catch (e) { /* convenience */ }
    const { from, to } = period();
    const name = `ADR例會名單_${gregorian(from)}-${gregorian(to)}.xlsx`;
    const ws = XLSX.utils.aoa_to_sheet(R.meetingSheet(picked, { allCases: all, reviewer, reviewerId }));
    ws['!cols'] = R.MEETING_WIDTHS.map(wch => ({ wch }));
    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, ws, '工作表1');
    const bytes = new Uint8Array(XLSX.write(wb, { type: 'array', bookType: 'xlsx' }));
    const url = URL.createObjectURL(new Blob([bytes], { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' }));
    const link = el('a', { href: url, download: name }, '下載');
    let where = '';
    if (folder) {
      try {
        const w = await (await folder.getFileHandle(name, { create: true })).createWritable();
        await w.write(bytes);
        await w.close();
        where = '（已存進 ADR 資料夾）';
      } catch (e) {
        where = `（沒有存進 ADR 資料夾：${e.message}）`;
      }
    }
    $('meetingStatus').replaceChildren(`✓ ${name}：${picked.length} 件${where}　`, link);
  }

  // ---- wiring ----

  ['fromMonth', 'toMonth', 'notableOnly', 'notableFirst'].forEach(id => $(id).addEventListener('change', () => render()));
  $('markHartwig').addEventListener('click', () => {
    const ids = periodCases.filter(c => canMark(c) && !c.notable && c.hints.some(h => h.code === 'hartwig')).map(c => c.id);
    if (ids.length) setMarks(ids, true);
    toast(ids.length ? `已勾選 ${ids.length} 件 Hartwig ≥4 的案例` : '這段期間沒有其他 Hartwig ≥4 的案例');
  });
  $('meetingBtn').addEventListener('click', openMeeting);
  $('meetingSave').addEventListener('click', () => saveMeeting().catch(e => $('meetingStatus').replaceChildren('存檔失敗：' + e.message)));
  $('markAll').addEventListener('click', () => setMarks(shown.filter(c => canMark(c) && !c.notable).map(c => c.id), true));
  $('markNone').addEventListener('click', () => {
    const marked = shown.filter(c => canMark(c) && c.notable);
    if (marked.length > 1 && !confirm(`取消這 ${marked.length} 件的標記？`)) return;
    setMarks(marked.map(c => c.id), false);
  });
  document.querySelectorAll('[data-range]').forEach(b => b.addEventListener('click', () => {
    const latest = $('toMonth').options[0] && $('toMonth').options[0].value;
    if (!latest) return;
    const options = [...$('fromMonth').options].map(o => o.value);
    const want = addMonths(latest, 1 - Number(b.dataset.range));
    // Earliest month on file that is not before the wanted start.
    $('fromMonth').value = options.filter(m => m >= want).pop() || latest;
    $('toMonth').value = latest;
    render();
  }));
  $('copyBtn').addEventListener('click', async () => {
    try {
      await navigator.clipboard.writeText(R.toTsv(shown));
      toast(`已複製 ${shown.length} 列，可貼到 Excel 或 Word`);
    } catch (e) {
      toast('無法複製：' + e.message);
    }
  });
  $('printBtn').addEventListener('click', () => window.print());
  $('phoneBtn').addEventListener('click', () => { $('phonePanel').hidden = false; showPhoneLink(); });
  $('phoneClose').addEventListener('click', () => { $('phonePanel').hidden = true; });
  $('phoneCopy').addEventListener('click', async () => {
    try {
      await navigator.clipboard.writeText(phoneUrl);
      toast('已複製連結，可以傳到手機（例如 LINE 傳給自己）');
    } catch (e) {
      toast('無法複製：' + e.message);
    }
  });
  $('reloadBtn').addEventListener('click', () => { picked = null; reload(); });
  $('fileInput').addEventListener('change', e => {
    if (e.target.files.length) loadPicked([...e.target.files]).catch(err => { $('sourceStatus').textContent = '讀取失敗：' + err.message; });
    e.target.value = '';
  });
  // Pick up changes made in Excel or by the other pages, and retry marks
  // that could not be saved (e.g. the workbook was open in Excel).
  window.addEventListener('focus', () => {
    if (!folder || picked) return;
    if (pendingMarks.size) saveMarks();
    else reload(false);
  });
  reload();
})();
