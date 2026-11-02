// TFDA helper page: reads the report-form export from the ADR folder (or a
// picked file), shows one case at a time laid out like the TFDA online form,
// copies any value on click, and makes the lab-results PDF to upload.
(function () {
  'use strict';

  const F = window.ADRForms;
  const $ = id => document.getElementById(id);
  const FORM_FILE = /通報表.*\.xlsx?$/i;
  const DONE_KEY = 'adr-tfda-done';

  let folder = null;      // ADR folder (desktop helper only)
  let cases = [];
  let current = null;     // selected case
  let done = loadDone();

  function loadDone() {
    try { return new Set(JSON.parse(localStorage.getItem(DONE_KEY) || '[]')); } catch (e) { return new Set(); }
  }
  function saveDone() {
    try { localStorage.setItem(DONE_KEY, JSON.stringify([...done])); } catch (e) { /* per-browser convenience only */ }
  }
  const caseKey = c => c.id + '|' + c.chartNo;

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

  // ---- loading ----

  function casesFromBytes(bytes) {
    const wb = XLSX.read(bytes, { type: 'array' });
    return wb.SheetNames.map(n => F.readCases(XLSX.utils.sheet_to_json(wb.Sheets[n], { header: 1, raw: false, defval: '' })));
  }

  async function loadFromFolder() {
    const found = [];
    const scan = async (dir, prefix) => {
      for await (const name of dir.keys()) {
        if (FORM_FILE.test(name) && !name.startsWith('~$')) found.push({ dir, name, path: prefix + name });
      }
    };
    await scan(folder, '');
    for await (const name of folder.keys()) {
      if (/^ADR_\d{5}$/.test(name)) await scan(await folder.getDirectoryHandle(name), name + '\\');
    }
    const lists = [];
    for (const f of found) {
      const file = await (await f.dir.getFileHandle(f.name)).getFile();
      lists.push(...casesFromBytes(new Uint8Array(await file.arrayBuffer())));
    }
    return { lists, names: found.map(f => f.path) };
  }

  async function reload() {
    $('sourceStatus').textContent = '讀取中…';
    try {
      if (!folder) folder = await window.ADRLocal.detect();
      if (!folder) {
        $('sourceStatus').textContent = '請按「選擇檔案」，選取「2.5.6－藥品不良反應 (ADR) 通報表.xls」。';
        return;
      }
      const { lists, names } = await loadFromFolder();
      if (!names.length) {
        $('sourceStatus').textContent = 'ADR 資料夾和月份資料夾裡找不到「通報表」Excel；請把匯出檔放進去，或按「選擇檔案」。';
        return;
      }
      setCases(F.mergeCases(lists), names);
    } catch (e) {
      $('sourceStatus').textContent = '讀取失敗：' + e.message;
    }
  }

  async function loadPicked(files) {
    const lists = [];
    for (const f of files) lists.push(...casesFromBytes(new Uint8Array(await f.arrayBuffer())));
    setCases(F.mergeCases(lists), [...files].map(f => f.name));
  }

  function setCases(list, names) {
    cases = list;
    $('sourceStatus').textContent = `讀取自：${names.join('、')}（共 ${list.length} 案）`;
    const months = [...new Set(list.map(F.monthOf))].sort().reverse();
    const sel = $('month');
    const keep = sel.value;
    sel.replaceChildren(...months.map(m => el('option', { value: m }, `${m.slice(0, 3)} 年 ${Number(m.slice(3))} 月`)));
    if (months.includes(keep)) sel.value = keep;
    renderList();
  }

  // ---- case list ----

  function monthCases() {
    return cases.filter(c => F.monthOf(c) === $('month').value);
  }

  function renderList() {
    const list = monthCases();
    const dupIds = new Set(list.map(c => c.id).filter((id, i, a) => a.indexOf(id) !== i));
    $('caseList').replaceChildren(...list.map(c => {
      const drug = (F.parseDrugs(c.suspect)[0] || { items: [] }).items.find(i => /英文商品名/.test(i.label));
      return el('button', {
        type: 'button',
        class: 'case-btn' + (current && caseKey(current) === caseKey(c) ? ' active' : '') + (done.has(caseKey(c)) ? ' done' : ''),
        onclick: () => select(c),
      }, el('strong', {}, c.id), el('span', {}, drug ? drug.value : c.symptoms),
      dupIds.has(c.id) ? el('span', { class: 'case-flag' }, '識別碼重複') : null,
      done.has(caseKey(c)) ? el('span', { class: 'case-flag ok' }, '已通報') : null);
    }));
    $('allPdfBtn').disabled = !list.some(F.hasLabs);
    if (current && !list.some(c => caseKey(c) === caseKey(current))) current = null;
    if (!current && list.length) select(list[0]);
    else if (!list.length) $('caseView').hidden = true;
  }

  // ---- case view ----

  function copyable(text, opts = {}) {
    const value = String(text || '');
    const box = el('button', {
      type: 'button',
      class: 'copy' + (opts.long ? ' long' : '') + (value ? '' : ' empty'),
      title: value ? '點一下複製' : '',
      onclick: () => value && copy(value, box),
    }, value || '（空白）');
    return box;
  }

  async function copy(text, box) {
    try {
      await navigator.clipboard.writeText(text);
    } catch (e) {
      const ta = el('textarea', {}, text);
      document.body.append(ta);
      ta.select();
      document.execCommand('copy');
      ta.remove();
    }
    document.querySelectorAll('.copy.copied').forEach(b => b.classList.remove('copied'));
    box.classList.add('copied', 'used');
    toast('已複製');
  }

  let toastTimer = null;
  function toast(text) {
    const t = $('toast');
    t.textContent = text;
    t.hidden = false;
    clearTimeout(toastTimer);
    toastTimer = setTimeout(() => { t.hidden = true; }, 1200);
  }

  function field(f) {
    const row = el('div', { class: 'tf-field' + (f.long ? ' wide' : '') }, el('div', { class: 'tf-label' }, f.label));
    if (f.options) {
      row.append(el('div', { class: 'tf-options' }, ...f.options.map(o =>
        el('span', { class: 'tf-option' + (o === f.pick ? ' picked' : '') }, o))));
      if (!f.pick && !f.hint) row.append(el('div', { class: 'tf-hint warn' }, `無法判斷，原始內容：${f.raw || '（空白）'}`));
    } else if (f.list) {
      row.append(f.list.length ? el('div', { class: 'tf-chips' }, ...f.list.map(s => copyable(s))) : el('div', { class: 'tf-hint' }, '（無）'));
    } else if (f.text !== '' || !f.hint) {
      row.append(copyable(f.text, f));
      if (f.max) row.append(el('div', { class: 'tf-hint' + (f.text.length > f.max ? ' warn' : '') }, `${f.text.length}/${f.max}`));
    }
    if (f.hint) row.append(el('div', { class: 'tf-hint' }, f.hint));
    return row;
  }

  function drugBlock(group) {
    if (!group.list.length) return null;
    return el('div', { class: 'tf-drugs' }, ...group.list.map(d =>
      el('div', { class: 'tf-drug' }, el('div', { class: 'tf-drug-title' }, `${group.kind}・${d.title}`),
        ...d.items.map(i => el('div', { class: 'tf-field' }, el('div', { class: 'tf-label' }, i.label), copyable(i.value))))));
  }

  function select(c) {
    current = c;
    const { sections, warnings } = F.tfdaSections(c);
    const key = caseKey(c);
    const doneBox = el('input', { type: 'checkbox' });
    doneBox.checked = done.has(key);
    doneBox.addEventListener('change', () => {
      if (doneBox.checked) done.add(key); else done.delete(key);
      saveDone();
      renderList();
    });
    const view = $('caseView');
    view.replaceChildren(
      el('div', { class: 'panel case-head' },
        el('h2', {}, `${c.id}　${c.chartNo}`),
        el('span', { class: 'hint' }, `撰寫：${c.trainee || c.reporter || '—'}`),
        el('span', { class: 'actions' },
          F.hasLabs(c) ? el('button', { type: 'button', onclick: () => makeAttachments([c]) }, '檢驗資料 PDF') : null,
          el('label', { class: 'done-toggle' }, doneBox, '已完成 TFDA 通報')),
        warnings.length ? el('ul', { class: 'messages' }, ...warnings.map(w => el('li', { class: 'warn' }, w))) : null),
      ...sections.map(s => el('section', { class: 'panel tf-section' },
        el('h3', { class: 'tf-tab' }, s.tab),
        ...(s.drugs || []).map(drugBlock),
        s.drugs && !s.drugs[0].list.length ? el('div', { class: 'tf-hint warn' }, '沒有可疑藥品資料') : null,
        el('div', { class: 'tf-grid' }, ...s.fields.map(field)))));
    view.hidden = false;
    renderListMarks();
  }

  function renderListMarks() {
    document.querySelectorAll('.case-btn').forEach((b, i) => {
      const c = monthCases()[i];
      b.classList.toggle('active', !!c && caseKey(c) === caseKey(current));
    });
  }

  // ---- lab attachment PDF ----

  const PAGE = { w: 1240, h: 1754, margin: 110 };  // A4 at 150 dpi

  function wrap(ctx, text, width) {
    const lines = [];
    for (const para of String(text).split('\n')) {
      let line = '';
      for (const ch of para) {
        if (ctx.measureText(line + ch).width > width && line) { lines.push(line); line = ''; }
        line += ch;
      }
      lines.push(line);
    }
    return lines;
  }

  async function attachmentBlob(c) {
    const content = F.attachmentContent(c);
    const font = '"Microsoft JhengHei", "Noto Sans TC", sans-serif';
    const width = PAGE.w - PAGE.margin * 2;
    const items = [];   // [font, text, gapBefore]
    const measure = document.createElement('canvas').getContext('2d');
    const add = (f, text, gap = 0) => {
      measure.font = f;
      wrap(measure, text, width).forEach((l, i) => items.push([f, l, i ? 0 : gap]));
    };
    add(`bold 40px ${font}`, content.title);
    content.meta.forEach(m => add(`26px ${font}`, m, 12));
    for (const b of content.blocks) {
      add(`bold 30px ${font}`, b.heading, 40);
      add(`26px ${font}`, b.text, 12);
    }
    const lineH = f => Math.round(parseInt(f.match(/(\d+)px/)[1], 10) * 1.5);
    const pages = [];
    let ctx = null;
    let y = 0;
    const newPage = () => {
      const canvas = document.createElement('canvas');
      canvas.width = PAGE.w;
      canvas.height = PAGE.h;
      ctx = canvas.getContext('2d');
      ctx.fillStyle = '#fff';
      ctx.fillRect(0, 0, PAGE.w, PAGE.h);
      ctx.fillStyle = '#000';
      ctx.textBaseline = 'top';
      pages.push(canvas);
      y = PAGE.margin;
    };
    newPage();
    for (const [f, text, gap] of items) {
      if (y + gap + lineH(f) > PAGE.h - PAGE.margin) newPage();
      else y += gap;
      ctx.font = f;
      ctx.fillText(text, PAGE.margin, y);
      y += lineH(f);
    }
    const out = [];
    for (const canvas of pages) {
      const blob = await new Promise(r => canvas.toBlob(r, 'image/jpeg', 0.92));
      out.push({ jpeg: new Uint8Array(await blob.arrayBuffer()), width: canvas.width, height: canvas.height });
    }
    return new Blob([window.ADRPdf.build(out, { width: 595.28, height: 841.89 })], { type: 'application/pdf' });
  }

  async function makeAttachments(list) {
    list = list.filter(F.hasLabs);
    const status = $('pdfStatus');
    status.textContent = '產生中…';
    try {
      const saved = [];
      for (const c of list) {
        const blob = await attachmentBlob(c);
        const name = F.attachmentName(c);
        if (folder) {
          const dir = await folder.getDirectoryHandle('ADR_' + F.monthOf(c), { create: true });
          const w = await (await dir.getFileHandle(name, { create: true })).createWritable();
          await w.write(blob);
          await w.close();
          saved.push(`ADR_${F.monthOf(c)}\\${name}`);
        } else {
          const a = el('a', { href: URL.createObjectURL(blob), download: name });
          a.click();
          URL.revokeObjectURL(a.href);
          saved.push(name);
          await new Promise(r => setTimeout(r, 250));
        }
        if (list.length === 1) window.open(URL.createObjectURL(blob), '_blank');
      }
      status.textContent = folder ? `已儲存 ${saved.length} 個：${saved.join('、')}` : `已下載 ${saved.length} 個`;
      toast(`檢驗資料 PDF：${saved.length} 個`);
    } catch (e) {
      status.textContent = '產生失敗：' + e.message;
    }
  }

  // ---- wiring ----

  $('month').addEventListener('change', () => { current = null; renderList(); });
  $('reloadBtn').addEventListener('click', reload);
  $('fileInput').addEventListener('change', e => e.target.files.length && loadPicked(e.target.files));
  $('allPdfBtn').addEventListener('click', () => makeAttachments(monthCases()));
  reload();
})();
