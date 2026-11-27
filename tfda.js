// TFDA helper page: reads the report-form export from the ADR folder (or a
// picked file), shows one case at a time laid out like the TFDA online form,
// copies any value on click, and makes the lab-results PDF to upload.
(function () {
  'use strict';

  const F = window.ADRForms;
  const $ = id => document.getElementById(id);
  // Not Excel's ~$ lock files, which exist while a form is open.
  const FORM_FILE = /^(?!~\$).*通報表.*\.xlsx?$/i;
  const EXCEL = /\.xlsx?$/i;
  // The report form printed to PDF; not the PDFs this app makes (ADR_<識別碼>…).
  const PDF = /\.pdf$/i;
  // Examples and blank forms (e.g. ADR通報表_填寫範例集.pdf) are not cases.
  const isFormFile = name => !name.startsWith('~$') && name !== WORKBOOK_NAME && !/範例|空白|樣本|template/i.test(name) &&
    (EXCEL.test(name) || (PDF.test(name) && !/^ADR_\d/i.test(name)));
  const WORKBOOK_NAME = 'ADR通報.xlsx';
  const DONE_KEY = 'adr-tfda-done';

  let folder = null;      // ADR folder (desktop helper only)
  let cases = [];
  let current = null;     // selected case
  let done = loadDone();
  let lookup = null;      // licence finder, once data/tfda-licenses.json is loaded
  let licenseIndex = null;
  let reported = new Map();  // 識別碼 -> 藥碼 recorded in ADR通報.xlsx (HIS)
  const hospitalCache = new Map();

  async function loadLicenses() {
    try {
      const res = await fetch('data/tfda-licenses.json', { cache: 'no-store' });
      const index = licenseIndex = window.ADRLicenses.buildIndex((await res.json()).rows);
      const cache = new Map();
      lookup = q => {
        const key = q.en + '|' + q.zh;
        if (!cache.has(key)) cache.set(key, window.ADRLicenses.findLicenses(index, q));
        return cache.get(key);
      };
      if (current) select(current);
    } catch (e) {
      lookup = null;  // page still works; 快速搜尋 falls back to the product name
    }
  }

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

  // Cases in one Excel file, and what was found, to explain an empty result.
  // A printed form gives one case; reading draws every page, so each file is
  // read once per visit.
  const pdfCache = new Map();
  async function readPdfFile(key, bytes) {
    if (!pdfCache.has(key)) {
      const c = await window.ADRPdfForm.readPdf(bytes);
      pdfCache.set(key, { cases: c && /^\d{7}$/.test(c.id) ? [c] : [], stats: { headers: c ? 1 : 0, rows: c ? 1 : 0, noId: c && !/^\d{7}$/.test(c.id) ? 1 : 0 } });
    }
    return pdfCache.get(key);
  }

  async function readAny(name, bytes) {
    return PDF.test(name) ? readPdfFile(name + '|' + bytes.length, bytes) : readFile(bytes);
  }

  function readFile(bytes) {
    const wb = XLSX.read(bytes, { type: 'array' });
    const stats = {};
    const cases = wb.SheetNames.flatMap(n => F.readCases(XLSX.utils.sheet_to_json(wb.Sheets[n], { header: 1, raw: false, defval: '' }), stats));
    return { cases, stats };
  }

  // One line per file: { text, ok }. Files that gave nothing are shown
  // up front; the rest only in the folded 檔案明細.
  function describe(path, { cases, stats }) {
    if (cases.length) return { ok: true, text: `${path}：${cases.length} 案` + (stats.noId ? `（${stats.noId} 列沒有識別代碼）` : '') };
    if (!stats.headers) return { ok: false, text: `${path}：不是通報表` };
    if (!stats.rows) return { ok: false, text: `${path}：沒有資料` };
    return { ok: false, text: `${path}：沒有 7 碼識別代碼（例如 1151103），未讀入` };
  }

  // Every Excel file in the ADR folder and its ADR_<yyyMM> folders except the
  // workbook itself; a file counts as a report form by its columns, not its name.
  async function loadFromFolder() {
    const found = [];
    const scan = async (dir, prefix) => {
      for await (const name of dir.keys()) {
        if (isFormFile(name)) found.push({ dir, name, path: prefix + name });
      }
    };
    await scan(folder, '');
    for await (const name of folder.keys()) {
      if (/^ADR_\d{5}$/.test(name)) await scan(await folder.getDirectoryHandle(name), name + '\\');
    }
    const lists = [];
    const lines = [];
    for (const f of found) {
      try {
        const file = await (await f.dir.getFileHandle(f.name)).getFile();
        const result = await readAny(f.name, new Uint8Array(await file.arrayBuffer()));
        if (result.cases.length) lists.push(result.cases);
        if (result.cases.length || FORM_FILE.test(f.name) || result.stats.headers) lines.push(describe(f.path, result));
      } catch (e) {
        lines.push({ ok: false, text: `${f.path}：無法讀取（${e.message}）` });
      }
    }
    reported = await readReported();
    return { lists, lines };
  }

  // 藥碼 per 識別碼 from the monthly sheets of ADR通報.xlsx (識別碼 | 病歷號 | 姓名 | 藥碼 | …).
  async function readReported() {
    const map = new Map();
    try {
      const file = await (await folder.getFileHandle('ADR通報.xlsx')).getFile();
      const wb = XLSX.read(new Uint8Array(await file.arrayBuffer()), { type: 'array' });
      for (const n of wb.SheetNames) {
        for (const r of XLSX.utils.sheet_to_json(wb.Sheets[n], { header: 1, raw: false, defval: '' })) {
          const id = String(r[0]).trim();
          if (/^\d{7}$/.test(id) && r[3]) map.set(id, String(r[3]).trim().toUpperCase());
        }
      }
    } catch (e) { /* no workbook: the hospital list is still searched by name */ }
    return map;
  }

  // ---- hospital drug information system (through server.py) ----

  function hospital(page, q) {
    const key = page + '|' + q;
    if (!hospitalCache.has(key)) {
      hospitalCache.set(key, fetch(`api/hospital?page=${page}&q=${encodeURIComponent(q)}`,
        { cache: 'no-store', headers: { 'X-ADR-Cards': '1' } })
        .then(res => {
          if (!res.ok) throw new Error('院內藥品資料庫連不上');
          return res.text();
        })
        .catch(e => { hospitalCache.delete(key); throw e; }));
    }
    return hospitalCache.get(key);
  }

  const H = () => window.ADRHospital;
  const hospitalDetail = code => hospital('P4', code).then(html => H().parseDetail(html));

  // Hospital products for one drug on the form: those found by name, plus the
  // 藥碼 HIS recorded for this case when the case has a single suspect drug and
  // a single code. Otherwise a HIS code is only tagged where the name search
  // finds it, so it never lands under another drug.
  async function hospitalProducts(query, hisCodes, single) {
    const found = query ? H().parseSearch(await hospital('P3', query)).slice(0, 8) : [];
    const extraCode = single && hisCodes.length === 1 ? hisCodes[0] : null;
    if (extraCode && !found.some(p => p.code === extraCode)) found.unshift({ code: extraCode });
    const details = await Promise.all(found.map(p => hospitalDetail(p.code).catch(() => null)));
    return details.filter(Boolean).map(d => Object.assign(d,
      licenseIndex ? { tfda: window.ADRLicenses.licenseFromNhiCode(licenseIndex, d.nhi, d.zh) } : {}));
  }

  function drugNotice(text, warn) {
    const n = $('drugNotice');
    if (!n || (n.classList.contains('warn') && !warn)) return;
    n.textContent = ' ' + text;
    n.classList.toggle('warn', warn);
  }

  function hospitalBlock(f, his) {
    const box = el('div', { class: 'tf-hospital' }, el('div', { class: 'tf-hint' }, '查詢院內藥品資料庫…'));
    const formZh = window.ADRLicenses.normZh(f.hospital.zh);
    const isHis = p => his.codes.includes(p.code);
    const sameZh = p => {
      const zh = window.ADRLicenses.normZh(p.zh);
      return !!(formZh && zh && (formZh.includes(zh) || zh.includes(formZh)));
    };
    const rank = p => 2 * isHis(p) + sameZh(p);
    hospitalProducts(f.hospital.query, his.codes, his.single).then(products => {
      if (!products.length) throw new Error('院內藥品資料庫沒有符合的品項');
      products.sort((a, b) => rank(b) - rank(a));
      box.replaceChildren(...products.map(p => {
        const tags = [];
        if (isHis(p)) tags.push('HIS 通報紀錄的藥碼');
        if (sameZh(p)) tags.push('與表單中文品名相同');
        const lic = p.tfda;
        return el('div', { class: 'tf-candidate' + (tags.length ? ' best' : '') },
          lic ? copyable(lic.license) : el('div', { class: 'tf-hint warn' }, `健保碼 ${p.nhi || '—'} 無法對應許可證`),
          el('div', { class: 'tf-hint' },
            el('a', { href: H().PAGE_URL + encodeURIComponent(p.code), target: '_blank', rel: 'noopener' }, p.code),
            `｜${p.en}｜${p.zh}`, lic && lic.company ? `｜${lic.company}` : '', lic && lic.guessed ? '｜（由健保碼推算）' : '',
            ...tags.map(t => el('span', { class: 'tf-tag' }, t))));
      }));
      drugNotice('快速搜尋以病人實際使用的院內品項為準（點藥碼可開啟院內藥品資料）。', false);
    }).catch(e => {
      box.replaceChildren(...nameCandidates(f));
      drugNotice(`${e.message}：快速搜尋改用品名比對 TFDA 許可證的結果，請核對品名與廠牌。`, true);
    });
    return box;
  }

  function nameCandidates(f) {
    if (!f.candidates) return [copyable(f.text)];
    return f.candidates.map((h, i) => el('div', { class: 'tf-candidate' + (i ? '' : ' best') },
      copyable(h.license), el('div', { class: 'tf-hint' }, `${h.zh}｜${h.en}｜${h.company}`)));
  }

  async function reload() {
    $('sourceStatus').textContent = '讀取中…';
    try {
      if (!folder) folder = await window.ADRLocal.detect();
      if (!folder) {
        showSource(['請按「上傳通報表」，選取「2.5.6－藥品不良反應 (ADR) 通報表」Excel。']);
        return;
      }
      const { lists, lines } = await loadFromFolder();
      if (!lines.length) {
        showSource(['ADR 資料夾和 ADR_yyyMM 月份資料夾裡找不到通報表 Excel。請按「上傳通報表」選取匯出的檔案。']);
        return;
      }
      setCases(F.mergeCases(lists), lines);
    } catch (e) {
      showSource(['讀取失敗：' + e.message]);
    }
  }

  // Picked files are read at once and, with the desktop helper, saved into
  // the ADR folder (replacing a file of the same name), so the page keeps
  // reading the latest export from there.
  async function loadPicked(files) {
    showSource(['讀取中…']);
    const lists = [];
    const lines = [];
    const saved = [];
    for (const f of files) {
      const bytes = new Uint8Array(await f.arrayBuffer());
      const result = await readAny(f.name, bytes);
      lines.push(describe(f.name, result));
      if (!result.cases.length) continue;
      lists.push(result.cases);
      if (folder) {
        try {
          const w = await (await folder.getFileHandle(f.name, { create: true })).createWritable();
          await w.write(bytes);
          await w.close();
          saved.push(f.name);
        } catch (e) {
          lines.push({ ok: false, text: `${f.name} 沒有存進 ADR 資料夾：${e.message}` });
        }
      }
    }
    if (saved.length) {
      // Read the whole folder again, so older exports there are merged too.
      const { lists: all, lines: folderLines } = await loadFromFolder();
      setCases(F.mergeCases(all), folderLines, `已存進 ADR 資料夾：${saved.join('、')}`);
    } else {
      setCases(F.mergeCases(lists), lines);
    }
  }

  function showSource(lines) {
    $('sourceStatus').replaceChildren(...lines.map(t => el('div', {}, t)));
  }

  // One summary line; files that gave nothing are listed, the rest folded.
  function setCases(list, lines, note = '') {
    cases = list;
    const bad = lines.filter(l => !l.ok);
    const good = lines.filter(l => l.ok);
    $('sourceStatus').replaceChildren(
      el('div', {}, `共 ${list.length} 案，讀自 ${good.length} 個檔案` + (note ? `　${note}` : '')),
      ...bad.map(l => el('div', { class: 'source-warn' }, l.text)),
      good.length ? el('details', { class: 'source-files' }, el('summary', {}, '檔案明細'), ...good.map(l => el('div', {}, l.text))) : null);
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
    // One click: a 檢驗資料 PDF for every case of the month.
    const empty = list.filter(c => !F.hasLabs(c));
    $('allPdfBtn').disabled = !list.length;
    $('allPdfBtn').textContent = `一次產生本月全部檢驗資料 PDF（${list.length} 件）`;
    $('allPdfBtn').title = empty.length ? `沒有檢驗資料的案例（PDF 內容為「無」）：${empty.map(c => c.id).join('、')}` : '';
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
    }, value || '空白');
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

  function field(f, his = { codes: [], single: false }) {
    const row = el('div', { class: 'tf-field' + (f.long ? ' wide' : '') + (f.hospital ? ' wide' : '') },
      el('div', { class: 'tf-label' }, f.label));
    if (f.hospital && folder) {
      row.append(hospitalBlock(f, his));
      return row;
    }
    if (f.attachment) {
      row.append(attachmentBlock(current));
      return row;
    }
    if (f.options) {
      row.append(el('div', { class: 'tf-options' }, ...f.options.map(o =>
        el('span', { class: 'tf-option' + (o === f.pick ? ' picked' : '') }, o))));
      if (!f.pick && !f.hint && f.raw) row.append(el('div', { class: 'tf-hint warn' }, `無法判斷，原始內容：${f.raw}`));
    } else if (f.list) {
      row.append(f.list.length ? el('div', { class: 'tf-chips' }, ...f.list.map(s => copyable(s))) : el('div', { class: 'tf-hint' }, '（無）'));
    } else if (f.candidates) {
      row.append(...nameCandidates(f));
    } else if (f.text !== '' || !f.hint) {
      row.append(copyable(f.text, f));
      // The length only matters near TFDA's limit.
      if (f.max && f.text.length > f.max * 0.8) row.append(el('div', { class: 'tf-hint' + (f.text.length > f.max ? ' warn' : '') }, `${f.text.length}/${f.max}`));
    }
    if (f.hint) row.append(el('div', { class: 'tf-hint' }, f.hint));
    return row;
  }

  function drugBlock(group, i, groups) {
    if (!group.list.length) return null;
    const suspects = group.kind === '可疑藥品' ? [] : (groups.find(g => g.kind === '可疑藥品') || { list: [] }).list;
    const his = group.kind === '可疑藥品'
      ? { codes: window.ADRParse.splitCodes(reported.get(current.id)), single: group.list.length === 1 }
      : { codes: [], single: false };
    const norm = s => String(s || '').replace(/\s+/g, '').toLowerCase();
    return el('div', { class: 'tf-drugs' }, ...group.list.map(d => {
      // Forms sometimes list the suspect drugs again as interacting drugs.
      const same = suspects.find(s => norm(s.name) === norm(d.name));
      return el('div', { class: 'tf-drug' },
        el('div', { class: 'tf-drug-title' }, `${group.kind}・${d.title}　${d.name}`),
        same ? el('div', { class: 'tf-hint warn' }, `與可疑藥品・${same.title}相同：表單把同一個藥品重複列為${group.kind}，請確認是否真的要在 TFDA 重複新增。`) : null,
        ...d.fields.map(f => field(f, his)));
    }));
  }

  function select(c) {
    current = c;
    const { sections, warnings } = F.tfdaSections(c, lookup);
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
          el('button', { type: 'button', onclick: async () => { await makeAttachments([c]); select(c); } }, '檢驗資料 PDF'),
          el('label', { class: 'done-toggle' }, doneBox, '已完成 TFDA 通報')),
        warnings.length ? el('ul', { class: 'messages' }, ...warnings.map(w => el('li', { class: 'warn' }, w))) : null),
      ...sections.map(s => el('section', { class: 'panel tf-section' },
        el('h3', { class: 'tf-tab' }, s.tab),
        // Said once here instead of in every drug block.
        s.drugs ? el('p', { class: 'tf-hint tf-drug-note' }, '每個藥品按 TFDA「新增藥品」依序填入。',
          el('span', { id: 'drugNotice' })) : null,
        ...(s.drugs || []).map(drugBlock),
        s.drugs && !s.drugs[0].list.length ? el('div', { class: 'tf-hint warn' }, '沒有可疑藥品資料') : null,
        el('div', { class: 'tf-grid' }, ...s.fields.map(f => field(f))))));
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

  // The case's lab PDF: where it is, whether it exists yet, and buttons to
  // make or open it.
  function attachmentBlock(c) {
    const name = F.attachmentName(c);
    const dirName = 'ADR_' + F.monthOf(c);
    const box = el('div', { class: 'tf-attachment' });
    const show = async () => {
      let file = null;
      if (folder) {
        try { file = await (await (await folder.getDirectoryHandle(dirName)).getFileHandle(name)).getFile(); } catch (e) { file = null; }
      }
      const make = el('button', { type: 'button', class: file ? 'secondary' : '' }, file ? '重新產生' : '產生檢驗資料 PDF');
      make.addEventListener('click', async () => {
        make.disabled = true;
        await makeAttachments([c], { open: true });
        show();
      });
      box.replaceChildren(
        copyable(name),
        el('div', { class: 'tf-hint' + (file ? '' : ' warn') },
          file ? `已產生，位置：${folder.name}\\${dirName}\\${name}。在 TFDA 按「新增附件」，選這個檔案上傳。`
            : folder ? `尚未產生。按下方按鈕，會存到 ${folder.name}\\${dirName}\\ 並開啟。` : '尚未產生。按下方按鈕下載。'),
        el('div', { class: 'actions' }, make,
          file ? el('button', { type: 'button', onclick: () => window.open(URL.createObjectURL(file), '_blank') }, '開啟 PDF') : null));
    };
    show();
    return box;
  }

  async function makeAttachments(list, { open = list.length === 1 } = {}) {
    if (!list.length) return;
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
        if (open) window.open(URL.createObjectURL(blob), '_blank');
      }
      status.textContent = folder ? `已儲存 ${saved.length} 個到 ${[...new Set(saved.map(n => n.split('\\')[0]))].join('、')}` : `已下載 ${saved.length} 個`;
      status.title = saved.join('\n');
      toast(`檢驗資料 PDF：${saved.length} 個`);
    } catch (e) {
      status.textContent = '產生失敗：' + e.message;
    }
  }

  // ---- wiring ----

  $('month').addEventListener('change', () => { current = null; renderList(); });
  $('reloadBtn').addEventListener('click', reload);
  $('fileInput').addEventListener('change', e => {
    const files = [...e.target.files];
    e.target.value = '';
    if (files.length) loadPicked(files).catch(err => showSource(['讀取失敗：' + err.message]));
  });
  $('allPdfBtn').addEventListener('click', async () => {
    await makeAttachments(monthCases());
    if (current) select(current);
  });
  reload();
  loadLicenses();
})();
