// Case overview page: reads ADR通報.xlsx and the report-form exports from the
// ADR folder (or picked files), lists the cases of a chosen period and flags
// the ones worth reporting at the routine ADR meeting (see review.js).
(function () {
  'use strict';

  const R = window.ADRReview;
  const F = window.ADRForms;
  const P = window.ADRParse;
  const $ = id => document.getElementById(id);
  const WORKBOOK_NAME = 'ADR通報.xlsx';
  // Not Excel's ~$ lock files, which exist while a form is open.
  const FORM_FILE = /^(?!~\$).*通報表.*\.xlsx?$/i;
  const STAR_KEY = 'adr-overview-starred';
  const LEVEL_RANK = { none: 0, info: 1, mid: 2, high: 3 };

  let folder = null;
  let all = [];            // joined cases (review.joinForms)
  let shown = [];          // flagged cases of the period, in table order
  let starred = loadStarred();
  const open = new Set();  // 識別碼 whose details row is expanded

  function loadStarred() {
    try { return new Set(JSON.parse(localStorage.getItem(STAR_KEY) || '[]')); } catch (e) { return new Set(); }
  }
  function saveStarred() {
    try { localStorage.setItem(STAR_KEY, JSON.stringify([...starred])); } catch (e) { /* per-browser convenience only */ }
  }

  function el(tag, attrs = {}, ...children) {
    const e = document.createElement(tag);
    for (const [k, v] of Object.entries(attrs)) {
      if (k === 'class') e.className = v;
      else if (k.startsWith('on')) e.addEventListener(k.slice(2), v);
      else e.setAttribute(k, v);
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
    const scan = async (dir, prefix) => {
      for await (const name of dir.keys()) {
        if (!FORM_FILE.test(name)) continue;
        forms.push(formsFrom(await bytesOf(await (await dir.getFileHandle(name)).getFile())));
        names.push(prefix + name);
      }
    };
    await scan(folder, '');
    for await (const name of folder.keys()) {
      if (/^ADR_\d{5}$/.test(name)) await scan(await folder.getDirectoryHandle(name), name + '\\');
    }
    return { cases, forms: F.mergeCases(forms), names };
  }

  async function loadPicked(files) {
    let cases = [];
    const forms = [];
    const names = [];
    for (const f of files) {
      const bytes = await bytesOf(f);
      if (FORM_FILE.test(f.name)) forms.push(formsFrom(bytes));
      else cases = cases.concat(R.workbookCases(XLSX, bytes));
      names.push(f.name);
    }
    setData({ cases, forms: F.mergeCases(forms), names });
  }

  async function reload() {
    $('sourceStatus').textContent = '讀取中…';
    try {
      if (!folder) folder = await window.ADRLocal.detect();
      if (!folder) {
        $('sourceStatus').textContent = `請按「選擇檔案」，選取 ${WORKBOOK_NAME}（可同時選取「通報表」Excel）。`;
        return;
      }
      setData(await loadFromFolder());
    } catch (e) {
      $('sourceStatus').textContent = '讀取失敗：' + e.message;
    }
  }

  function setData({ cases, forms, names }) {
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
    render();
  }

  // ---- rendering ----

  function period() {
    let from = $('fromMonth').value, to = $('toMonth').value;
    if (from > to) [from, to] = [to, from];
    return { from, to };
  }

  const topLevel = c => Math.max(0, ...c.flags.map(f => LEVEL_RANK[f.level]));

  function render() {
    const { from, to } = period();
    const inPeriod = all.filter(c => c.month >= from && c.month <= to);
    const flagged = R.flagCases(inPeriod, { starred });
    const sum = R.summarize(flagged);
    $('printTitle').textContent = `期間：${rocMonth(from)}${from === to ? '' : ' 至 ' + rocMonth(to)}`;
    $('statTotal').textContent = `${sum.total} 件`;
    $('statNotable').textContent = `${sum.notable} 件`;
    $('statSerious').textContent = `${sum.serious} 件`;
    $('statNoForm').textContent = `${sum.noForm} 件`;
    const chips = list => list.length
      ? list.slice(0, 6).map(d => el('span', { class: 'ov-chip' + (d.n > 1 ? ' many' : '') }, `${d.name} ×${d.n}`))
      : [el('span', { class: 'hint' }, '—')];
    $('topDrugs').replaceChildren(...chips(sum.drugs));
    $('topSymptoms').replaceChildren(...chips(sum.symptoms));

    shown = flagged.filter(c => !$('notableOnly').checked || c.notable);
    if ($('notableFirst').checked) shown.sort((a, b) => topLevel(b) - topLevel(a) || a.id.localeCompare(b.id));
    $('emptyHint').hidden = shown.length > 0;
    $('emptyHint').textContent = flagged.length ? '這段期間沒有值得注意的案例。' : '這段期間沒有案例。';
    $('rows').replaceChildren(...shown.flatMap(rowsFor));
  }

  function outcomeText(c) {
    if (!c.form) return el('span', { class: 'hint' }, '無通報表');
    return c.form.outcome || c.form.outcomeRaw || '—';
  }

  function severityCell(c) {
    const serious = c.form && R.SERIOUS_LABEL[c.form.seriousness];
    return el('td', {}, c.severity || '—',
      c.form && c.form.seriousness ? el('div', { class: 'ov-sub' + (serious ? ' serious' : '') }, serious || '非嚴重') : null);
  }

  function rowsFor(c) {
    const level = c.flags.reduce((best, f) => (LEVEL_RANK[f.level] > LEVEL_RANK[best] ? f.level : best), 'none');
    const isOpen = open.has(c.id);
    const star = el('button', {
      type: 'button', class: 'ov-star' + (starred.has(c.id) ? ' on' : ''),
      title: starred.has(c.id) ? '取消標記' : '標記為要報告',
      onclick: e => {
        e.stopPropagation();
        if (starred.has(c.id)) starred.delete(c.id); else starred.add(c.id);
        saveStarred();
        render();
      },
    }, starred.has(c.id) ? '★' : '☆');
    const tr = el('tr', {
      class: `ov-row level-${level}` + (isOpen ? ' open' : ''),
      onclick: () => { if (isOpen) open.delete(c.id); else open.add(c.id); render(); },
    },
    el('td', { class: 'no-print' }, star),
    el('td', { class: 'ov-id' }, c.id || '—'),
    el('td', {}, c.date ? P.formatDate(c.date) : '—'),
    el('td', {}, c.chartNo),
    el('td', {}, c.drug),
    el('td', {}, c.symptom),
    severityCell(c),
    el('td', {}, outcomeText(c)),
    el('td', {}, c.writer || el('span', { class: 'hint' }, '—')),
    el('td', {}, ...c.flags.map(f => el('span', { class: 'flag ' + f.level }, f.label))));
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
    return el('tr', { class: 'ov-detail' }, el('td', { colspan: '10' }, el('div', { class: 'ov-detail-grid' }, ...body)));
  }

  // ---- wiring ----

  $('fromMonth').addEventListener('change', render);
  $('toMonth').addEventListener('change', render);
  $('notableOnly').addEventListener('change', render);
  $('notableFirst').addEventListener('change', render);
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
  $('reloadBtn').addEventListener('click', reload);
  $('fileInput').addEventListener('change', e => {
    if (e.target.files.length) loadPicked([...e.target.files]).catch(err => { $('sourceStatus').textContent = '讀取失敗：' + err.message; });
    e.target.value = '';
  });
  // Pick up changes made in Excel or by the other pages.
  window.addEventListener('focus', () => { if (folder) reload(); });
  reload();
})();
