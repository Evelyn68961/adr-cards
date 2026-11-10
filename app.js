(function () {
  'use strict';
  const P = window.ADRParse;
  const W = window.ADRWorkbook;

  const FIELDS = [
    { key: 'id', label: '識別碼' },
    { key: 'chartNo', label: '病歷號' },
    { key: 'dateText', label: '得知日期' },
    { key: 'source', label: '通報人員' },
    { key: 'drug', label: '藥品' },
    { key: 'symptom', label: '症狀' },
  ];
  const CARDS_PER_PAGE = 4;
  const WORKBOOK_NAME = 'ADR通報.xlsx';

  const $ = id => document.getElementById(id);
  const pasteEl = $('paste');
  const startEl = $('startSeq');
  const sheetEl = $('sheet');
  const messagesEl = $('messages');

  let cards = [];
  let folder = null;          // FileSystemDirectoryHandle of the ADR folder
  let startTouched = false;   // user typed a start number; stop following Excel
  let ledger = null;          // what ADR通報.xlsx holds (see ADRWorkbook.readLedger)
  let pngs = {};              // month key -> Set of card file names on disk
  let shownMonth = null;      // month key picked in the status panel
  // 識別碼 typed on a card (rowKey -> 識別碼): that case takes the number, e.g. to
  // replace a case in Excel that turned out not to be reportable.
  const fixedIds = new Map();
  let saving = false;

  // ---- Parsing and cards ----------------------------------------------------

  // Keeps ticks and hand edits for cases that are still in the list, and only
  // redraws when numbering changed, so a background sync never disturbs typing.
  function rebuild(force = false) {
    const { rows, skipped } = P.parseText(pasteEl.value);
    const before = new Map(cards.map(c => [c.key, c]));
    const start = startTouched ? Number(startEl.value) || 1 : null;
    const keys = new Set(rows.map(r => P.rowKey(r.chartNo, r.drugCode, r.date)));
    [...fixedIds.keys()].filter(k => !keys.has(k)).forEach(k => fixedIds.delete(k));
    const next = P.assignIds(rows, start, ledgerView(), fixedIds).map(c => {
      const key = P.rowKey(c.chartNo, c.drugCode, c.date);
      const old = before.get(key);
      const kept = old ? {
        include: old.include, chartNo: old.chartNo, dateText: old.dateText,
        source: old.source, drug: old.drug, symptom: old.symptom,
      } : { include: true, dateText: P.formatDate(c.date) };
      return Object.assign(c, { key }, kept);
    });
    const signature = list => list.map(c => `${c.key}=${c.id}=${!!c.recorded}=${c.previousId}`).join();
    if (force || signature(next) !== signature(cards)) {
      cards = next;
      renderCards();
    }
    if (!startTouched) {
      const first = cards.find(c => !c.recorded && !c.fixed);
      startEl.value = first ? Number(first.id.slice(5)) : ledger ? nextSeq(currentMonthKey()) : 1;
    }
    renderMessages(skipped);
    updateButtons();
  }

  function renderList(el, msgs) {
    el.replaceChildren(...msgs.map(m => {
      const li = document.createElement('li');
      li.className = m.cls;
      li.textContent = m.text;
      if (m.href) {
        const a = document.createElement('a');
        a.href = m.href;
        a.download = m.download;
        a.textContent = '下載';
        li.append(' ', a);
      }
      return li;
    }));
  }

  function renderMessages(skipped) {
    const msgs = skipped.map(s => ({ cls: 'warn', text: `第 ${s.line} 行略過：${s.reason}` }));
    const months = [...new Set(cards.map(c => c.id.slice(0, 5)))];
    if (months.length > 1 && !ledger) {
      msgs.push({ cls: 'warn', text: `包含多個月份（${months.join('、')}）：起始流水號只套用於第一個月，其後月份從 01 開始。` });
    }
    const fresh = cards.filter(c => !c.recorded && !c.fixed);
    const recorded = cards.filter(c => c.recorded).length;
    if (cards.length) {
      msgs.unshift({ cls: 'ok', text: `共 ${cards.length} 張卡片，已依得知日期時間排序。` +
        (recorded ? `其中 ${recorded} 張已在 Excel，沿用原識別碼。` : '') });
    }
    if (ledger) {
      // A typed number already in Excel: this case replaces what holds it.
      const isSelf = (r, c) => r.chartNo === c.chartNo && r.date && r.date.getTime() === c.date.getTime();
      cards.filter(c => !c.recorded || c.fixed).forEach(c => {
        const rows = Object.values(ledger.months).flatMap(m => m.rows).filter(r => r.id === c.id && !isSelf(r, c));
        if (rows.length) {
          msgs.push({ cls: 'warn', text: `${c.id} 已在 Excel：儲存後將以本卡（${c.name} ${c.drug}）取代 ` +
            rows.map(r => `${r.name} ${r.drug}`).join('、') + `，卡片圖檔 ADR_${c.id}.png 也會換成本卡。` });
        }
      });
    }
    cards.filter(c => c.previousId && c.previousId !== c.id).forEach(c => {
      msgs.push({ cls: 'warn', text: `${c.name} ${c.drug} 原為 ${c.previousId}，儲存後改為 ${c.id}；` +
        `ADR_${c.previousId}.png 會刪除，${c.previousId} 空出來。` });
    });
    const dup = [...new Set(cards.map(c => c.id).filter((id, i, all) => all.indexOf(id) !== i))];
    if (dup.length) msgs.push({ cls: 'warn', text: `識別碼重複：${dup.join('、')}。請修改其中一張卡片的識別碼，否則無法存入。` });
    if (ledger && startTouched && fresh.length) {
      const expected = nextSeq(fresh[0].id.slice(0, 5));
      if (Number(fresh[0].id.slice(5)) !== expected) {
        msgs.push({ cls: 'warn', text: `Excel 的下一號是 ${fresh[0].id.slice(0, 5)}${String(expected).padStart(2, '0')}，您輸入的起始流水號與它不同。` });
      }
    }
    renderList(messagesEl, msgs);
    $('idHint').textContent = fresh.length
      ? `新案例識別碼：${fresh[0].id} – ${fresh[fresh.length - 1].id}${ledger && !startTouched ? '（已與 Excel 同步）' : ''}`
      : '';
  }

  function renderCards() {
    sheetEl.replaceChildren(...cards.map((card, i) => {
      const el = document.createElement('article');
      el.className = 'card' + (card.include ? '' : ' excluded');

      const toggle = document.createElement('input');
      toggle.type = 'checkbox';
      toggle.className = 'include no-print';
      toggle.checked = card.include;
      toggle.title = '列印此卡';
      toggle.addEventListener('change', () => {
        card.include = toggle.checked;
        el.classList.toggle('excluded', !card.include);
        updateButtons();
      });
      el.appendChild(toggle);

      if (card.recorded || card.fixed) {
        const badge = document.createElement('span');
        badge.className = 'recorded no-print' + (card.fixed ? ' fixed' : '');
        badge.textContent = card.fixed ? '指定識別碼' : '已在 Excel';
        el.appendChild(badge);
      }

      FIELDS.forEach(f => {
        const row = document.createElement('div');
        row.className = 'field';
        const label = document.createElement('span');
        label.className = 'label';
        label.textContent = f.label + ': ';
        const value = document.createElement('span');
        value.className = 'value';
        value.contentEditable = 'plaintext-only';
        value.spellcheck = false;
        value.textContent = card[f.key];
        if (f.key === 'id') {
          // A typed 識別碼 is kept through re-syncs; empty goes back to automatic.
          value.title = '改成 Excel 中某案例的識別碼，可用本卡取代它（例如不可通報的案例）；清空恢復自動編號';
          value.addEventListener('keydown', e => { if (e.key === 'Enter') { e.preventDefault(); value.blur(); } });
          value.addEventListener('blur', () => setFixedId(card, value.textContent));
        } else {
          value.addEventListener('input', () => { cards[i][f.key] = value.textContent.trim(); });
        }
        row.append(label, value);
        el.appendChild(row);
      });
      return el;
    }));
    updateButtons();
  }

  function setFixedId(card, text) {
    const id = text.trim();
    if (id && !/^\d{7}$/.test(id)) {
      alert('識別碼要是 7 位數字，例如 1151103。');
      renderCards();
      return;
    }
    if (!id) {
      if (!fixedIds.delete(card.key)) { renderCards(); return; }
    } else {
      if (id === card.id) return;
      fixedIds.set(card.key, id);
    }
    rebuild(true);
  }

  function selected() {
    return cards.filter(c => c.include);
  }

  function setAll(include) {
    cards.forEach(c => { c.include = include; });
    renderCards();
  }

  function updateButtons() {
    const n = selected().length;
    const none = n === 0;
    $('selCount').textContent = cards.length ? `已選 ${n} / ${cards.length} 張` : '';
    $('pdfBtn').disabled = none;
    // Ticks only choose what to print; every case is saved (Excel is the
    // continuous record, and every 識別碼 needs its PNG for numbering).
    $('pngBtn').disabled = !cards.length;
    $('saveBtn').disabled = !cards.length || (!folder && !window.showDirectoryPicker);
    $('pngBtn').hidden = !!folder;
  }

  // ---- Card image -----------------------------------------------------------

  // Card layout of the existing ADR_*.png files (1128x825 px), drawn into a
  // w x h box at (left, top); text scales to fit the box.
  const CARD = { W: 1128, H: 825, X: 66, TOP: 116, LINE: 91, SIZE: 44, BORDER: 6 };

  function drawCard(ctx, card, left, top, w, h) {
    const s = Math.min(w / CARD.W, h / CARD.H);
    const border = CARD.BORDER * s;
    ctx.fillStyle = '#fff';
    ctx.fillRect(left, top, w, h);
    ctx.strokeStyle = '#000';
    ctx.lineWidth = border;
    ctx.strokeRect(left + border / 2, top + border / 2, w - border, h - border);
    ctx.fillStyle = '#000';
    ctx.font = `${CARD.SIZE * s}px "Noto Sans TC", "Microsoft JhengHei", sans-serif`;
    ctx.textBaseline = 'alphabetic';

    const x = left + CARD.X * s;
    const line = CARD.LINE * s;
    let y = top + CARD.TOP * s;
    FIELDS.forEach(f => {
      const label = f.label + ': ';
      ctx.fillText(label, x, y);
      const indent = x + ctx.measureText(label).width;
      const lines = wrap(ctx, card[f.key], left + w - CARD.X * s - indent);
      lines.forEach((text, n) => ctx.fillText(text, indent, y + n * 0.75 * line));
      y += line + (lines.length - 1) * 0.75 * line;
    });
  }

  function cardToPng(card) {
    const canvas = document.createElement('canvas');
    canvas.width = CARD.W;
    canvas.height = CARD.H;
    drawCard(canvas.getContext('2d'), card, 0, 0, CARD.W, CARD.H);
    return canvas;
  }

  // Breaks at spaces and 、 where possible, otherwise at any character.
  function wrap(ctx, text, width) {
    const tokens = String(text).match(/[^\s、]+、?|\s+|、/g) || [''];
    const lines = [];
    let line = '';
    tokens.forEach(tok => {
      const next = line + tok;
      if (ctx.measureText(next).width <= width || !line.trim()) {
        line = next;
      } else {
        lines.push(line.trimEnd());
        line = tok.trimStart();
      }
      while (ctx.measureText(line).width > width && line.length > 1) {
        let cut = line.length - 1;
        while (cut > 1 && ctx.measureText(line.slice(0, cut)).width > width) cut--;
        lines.push(line.slice(0, cut));
        line = line.slice(cut);
      }
    });
    lines.push(line.trimEnd());
    return lines;
  }

  async function cardBlob(card) {
    if (document.fonts && document.fonts.ready) await document.fonts.ready;
    const canvas = cardToPng(card);
    return new Promise(resolve => canvas.toBlob(resolve, 'image/png'));
  }

  // Fallback when no folder is linked.
  async function downloadPngs() {
    for (const card of cards) {
      const a = document.createElement('a');
      a.href = URL.createObjectURL(await cardBlob(card));
      a.download = `ADR_${card.id}.png`;
      a.click();
      URL.revokeObjectURL(a.href);
      // Browsers drop rapid back-to-back downloads; space them out a little.
      await new Promise(r => setTimeout(r, 250));
    }
  }

  // ---- PDF ------------------------------------------------------------------

  // Landscape A4 at 250 dpi, 2x2 cards of 138x94 mm, centred like the print CSS.
  const PAGE_MM = { w: 297, h: 210, cardW: 138, cardH: 94, gap: 1 };
  const DPI = 250;

  async function cardsToPdf(list) {
    if (document.fonts && document.fonts.ready) await document.fonts.ready;
    const px = mm => Math.round(mm / 25.4 * DPI);
    const canvas = document.createElement('canvas');
    canvas.width = px(PAGE_MM.w);
    canvas.height = px(PAGE_MM.h);
    const ctx = canvas.getContext('2d');
    const gridW = PAGE_MM.cardW * 2 + PAGE_MM.gap;
    const gridH = PAGE_MM.cardH * 2 + PAGE_MM.gap;
    const left = (PAGE_MM.w - gridW) / 2;
    const top = (PAGE_MM.h - gridH) / 2;

    const pages = [];
    for (let i = 0; i < list.length; i += CARDS_PER_PAGE) {
      ctx.fillStyle = '#fff';
      ctx.fillRect(0, 0, canvas.width, canvas.height);
      list.slice(i, i + CARDS_PER_PAGE).forEach((card, n) => {
        const col = n % 2, row = Math.floor(n / 2);
        drawCard(ctx, card,
          px(left + col * (PAGE_MM.cardW + PAGE_MM.gap)), px(top + row * (PAGE_MM.cardH + PAGE_MM.gap)),
          px(PAGE_MM.cardW), px(PAGE_MM.cardH));
      });
      const blob = await new Promise(r => canvas.toBlob(r, 'image/jpeg', 0.92));
      pages.push({ jpeg: new Uint8Array(await blob.arrayBuffer()), width: canvas.width, height: canvas.height });
    }
    return new Blob([window.ADRPdf.build(pages)], { type: 'application/pdf' });
  }

  async function makePdf() {
    const list = selected();
    // Open the tab now: browsers block pop-ups opened after an await.
    const tab = window.open('', '_blank');
    const msgs = [];
    try {
      const blob = await cardsToPdf(list);
      const url = URL.createObjectURL(blob);
      const name = `ADR_${list[0].id}-${list[list.length - 1].id}.pdf`;
      if (tab) tab.location = url;
      else msgs.push({ cls: 'warn', text: '瀏覽器擋住了新分頁，請改用下方連結開啟 PDF。' });
      msgs.push({ cls: 'ok', text: `PDF：${list.length} 張卡片，共 ${Math.ceil(list.length / CARDS_PER_PAGE)} 頁 A4`, href: url, download: name });
      if (folder && await folder.requestPermission({ mode: 'readwrite' }) === 'granted') {
        await writeFile(await subfolder('ADR_' + list[0].id.slice(0, 5), true), name, blob);
        msgs.push({ cls: 'ok', text: `已儲存到 ADR_${list[0].id.slice(0, 5)}\\${name}` });
      }
    } catch (e) {
      if (tab) tab.close();
      msgs.push({ cls: 'warn', text: '無法產生 PDF：' + e.message });
    }
    renderList($('saveMessages'), msgs);
  }

  // ---- Printing -------------------------------------------------------------

  // Grid page breaks are unreliable, so printing uses explicit 4-card pages.
  function buildPrintArea() {
    const printed = [...sheetEl.querySelectorAll('.card:not(.excluded)')];
    const pages = [];
    for (let i = 0; i < printed.length; i += CARDS_PER_PAGE) {
      const page = document.createElement('div');
      page.className = 'page';
      printed.slice(i, i + CARDS_PER_PAGE).forEach(card => {
        const copy = card.cloneNode(true);
        copy.querySelector('.include').remove();
        page.appendChild(copy);
      });
      pages.push(page);
    }
    $('printArea').replaceChildren(...pages);
  }

  // ---- ADR folder (File System Access API) ----------------------------------

  // The folder handle is kept in IndexedDB so it survives reloads; the
  // browser still asks for permission again once per session.
  function idb(mode, fn) {
    return new Promise((resolve, reject) => {
      const open = indexedDB.open('adr-cards', 1);
      open.onupgradeneeded = () => open.result.createObjectStore('kv');
      open.onerror = () => reject(open.error);
      open.onsuccess = () => {
        const tx = open.result.transaction('kv', mode);
        const req = fn(tx.objectStore('kv'));
        tx.oncomplete = () => resolve(req.result);
        tx.onerror = () => reject(tx.error);
      };
    });
  }
  const savedFolder = () => idb('readonly', s => s.get('folder')).catch(() => null);
  const saveFolder = h => idb('readwrite', s => s.put(h, 'folder')).catch(() => {});

  async function useFolder(handle) {
    folder = handle;
    $('folderStatus').textContent = `已連結：${handle.name}`;
    $('folderStatus').classList.add('linked');
    $('linkBtn').textContent = '更換資料夾';
    updateButtons();
    await refreshLedger();
  }

  // A month folder like ADR_11510 is an easy mis-pick; the workbook lives one level up.
  const isMonthFolder = handle => /^ADR_\d{5}$/i.test(handle.name);

  async function checkFolder(handle) {
    if (await existingFile(handle, WORKBOOK_NAME)) return true;
    if (isMonthFolder(handle)) {
      alert(`「${handle.name}」是月份資料夾。請選擇上一層、含有 ${WORKBOOK_NAME} 的 ADR 資料夾。`);
      return false;
    }
    return confirm(`「${handle.name}」中沒有 ${WORKBOOK_NAME}。要在這個資料夾新建嗎？`);
  }

  // Resolves true once a folder is linked.
  async function linkFolder() {
    const stored = await savedFolder();
    // Reconnect to the remembered folder if that is all that's needed.
    if (stored && !folder && !isMonthFolder(stored) &&
        await stored.requestPermission({ mode: 'readwrite' }) === 'granted') {
      await useFolder(stored);
      return true;
    }
    try {
      const handle = await window.showDirectoryPicker({ id: 'adr', mode: 'readwrite' });
      if (!await checkFolder(handle)) return false;
      await saveFolder(handle);
      await useFolder(handle);
      return true;
    } catch (e) {
      if (e.name !== 'AbortError') alert('無法連結資料夾：' + e.message);
      return false;
    }
  }

  async function initFolder() {
    // Desktop shortcut (server.py): the ADR folder is built in, nothing to link.
    const local = await window.ADRLocal.detect();
    if (local) {
      $('linkBtn').hidden = true;
      await useFolder(local);
      $('folderStatus').textContent = `已連結：${local.name}（本機自動連結）`;
      return;
    }
    if (!window.showDirectoryPicker) {
      $('linkBtn').hidden = true;
      $('folderUnsupported').hidden = false;
      return;
    }
    const stored = await savedFolder();
    if (!stored || isMonthFolder(stored)) return;
    if (await stored.queryPermission({ mode: 'readwrite' }) === 'granted') return useFolder(stored);
    $('folderStatus').textContent = `上次使用：${stored.name}`;
    $('linkBtn').textContent = '重新連結';
  }

  async function subfolder(name, create) {
    try { return await folder.getDirectoryHandle(name, { create }); }
    catch (e) { if (e.name === 'NotFoundError') return null; throw e; }
  }

  async function existingFile(dir, name) {
    try { return await dir.getFileHandle(name); }
    catch (e) { if (e.name === 'NotFoundError') return null; throw e; }
  }

  // ---- Sync with ADR通報.xlsx ----------------------------------------------------

  const currentMonthKey = () => P.monthKey(new Date());
  const pngFor = id => `ADR_${id}.png`;

  // Next number for a month = 1 + the larger of the workbook's highest number
  // and the highest ADR_<yyyMM>NN.png in the month folder.
  function nextSeq(key) {
    const month = ledger && Object.values(ledger.months).find(m => m.key === key);
    let max = month ? month.maxSeq : 0;
    (pngs[key] || []).forEach(name => {
      const m = name.match(new RegExp(`^ADR_${key}(\\d{2})\\.png$`, 'i'));
      if (m) max = Math.max(max, Number(m[1]));
    });
    return max + 1;
  }

  function ledgerView() {
    if (!ledger) return null;
    return {
      idFor: r => ledger.ids.get(P.rowKey(r.chartNo, r.drugCode, r.date)) || null,
      nextSeq,
    };
  }

  // Re-reads the workbook and month folders. Runs on open, after each save and
  // whenever the page regains focus, so manual Excel edits show up too.
  let ledgerToken = 0;
  async function refreshLedger() {
    if (!folder || saving) return;
    const token = ++ledgerToken;
    try {
      const book = await existingFile(folder, WORKBOOK_NAME);
      const bytes = book ? new Uint8Array(await (await book.getFile()).arrayBuffer()) : null;
      const next = bytes ? W.readLedger(window.XLSX, P, bytes) : { months: {}, ids: new Map(), owners: new Map() };
      const keys = new Set([
        ...Object.values(next.months).map(m => m.key),
        ...cards.map(c => c.id.slice(0, 5)),
        currentMonthKey(),
      ]);
      const found = {};
      for (const key of keys) {
        const dir = await subfolder('ADR_' + key, false);
        const names = new Set();
        if (dir) for await (const name of dir.keys()) names.add(name);
        found[key] = names;
      }
      if (token !== ledgerToken || saving) return;
      ledger = next;
      pngs = found;
      $('statusError').hidden = true;
      renderStatus();
      rebuild();
    } catch (e) {
      if (token !== ledgerToken) return;
      $('statusPanel').hidden = false;
      $('statusError').hidden = false;
      $('statusError').textContent = `無法讀取 ${WORKBOOK_NAME}：${e.message}`;
    }
  }

  function renderStatus() {
    $('statusPanel').hidden = false;
    const months = [...new Set([currentMonthKey(), ...Object.values(ledger.months).map(m => m.key)])]
      .sort().reverse();
    if (!months.includes(shownMonth)) shownMonth = currentMonthKey();
    const select = $('statusMonth');
    select.replaceChildren(...months.map(key => {
      const option = document.createElement('option');
      option.value = key;
      option.textContent = `${key.slice(0, 3)} 年 ${Number(key.slice(3))} 月`;
      return option;
    }));
    select.value = shownMonth;

    const month = Object.values(ledger.months).find(m => m.key === shownMonth);
    const rows = month ? month.rows.slice().sort((a, b) => a.id.localeCompare(b.id) || a.date - b.date) : [];
    const ids = rows.map(r => r.id).filter(Boolean);
    const onDisk = pngs[shownMonth] || new Set();
    const noCard = rows.filter(r => !r.id || !onDisk.has(pngFor(r.id)));
    $('statCount').textContent = `${rows.length} 件`;
    $('statLatest').textContent = ids.length ? ids[ids.length - 1] : '—';
    $('statNext').textContent = shownMonth + String(nextSeq(shownMonth)).padStart(2, '0');
    $('statNoCard').textContent = `${noCard.length} 件`;

    const unassigned = rows.filter(r => r.id && !writerOf(r)).length;
    if (!writerSaving && !pendingWriters.size && !/saved|failed/.test($('writerStatus').className)) {
      $('writerStatus').textContent = rows.length ? (unassigned ? `尚未指定撰寫人：${unassigned} 件` : '每個案例都已指定撰寫人') : '';
    }
    showWriterState(/failed/.test($('writerStatus').className));
    $('statusEmpty').hidden = rows.length > 0;
    // Don't redraw the rows under an open name dropdown.
    if ($('statusRows').contains(document.activeElement)) return;

    $('statusRows').replaceChildren(...rows.map(r => {
      const tr = document.createElement('tr');
      [r.id || '—', r.chartNo, r.drug, r.symptom, r.date ? P.formatDate(r.date) : '—'].forEach(text => {
        const td = document.createElement('td');
        td.textContent = text;
        tr.appendChild(td);
      });
      const td = document.createElement('td');
      const done = r.id && onDisk.has(pngFor(r.id));
      td.textContent = done ? '已製卡' : '未製卡';
      td.className = done ? 'card-done' : 'card-missing';
      tr.appendChild(td);
      tr.appendChild(writerCell(r));
      return tr;
    }));
  }

  // ---- 撰寫人 (who writes up each case) ----------------------------------------

  const pendingWriters = new Map();   // 識別碼 -> name typed but not saved yet
  let writerSaving = false;

  const writerOf = r => (pendingWriters.has(r.id) ? pendingWriters.get(r.id) : r.writer);

  // A plain text box: type or paste the name, then press 儲存撰寫人.
  function writerCell(r) {
    const td = document.createElement('td');
    if (!r.id) return td;
    const input = document.createElement('input');
    input.type = 'text';
    input.className = 'writer' + (pendingWriters.has(r.id) ? ' unsaved' : '');
    input.autocomplete = 'off';
    input.setAttribute('aria-label', `${r.id} 撰寫人`);
    input.placeholder = '未指定';
    input.value = writerOf(r);
    input.addEventListener('input', () => {
      const name = input.value.trim();
      if (name === r.writer) pendingWriters.delete(r.id);
      else pendingWriters.set(r.id, name);
      input.classList.toggle('unsaved', pendingWriters.has(r.id));
      showWriterState();
    });
    // Enter (not the one that picks a character in a Chinese input method) saves.
    input.addEventListener('keydown', e => {
      if (e.key === 'Enter' && !e.isComposing && e.keyCode !== 229) saveWriters();
    });
    td.appendChild(input);
    return td;
  }

  // Enables the button, and (unless a save result is showing) says how many
  // typed names are not saved yet.
  function showWriterState(keepMessage = false) {
    $('saveWritersBtn').disabled = !pendingWriters.size || writerSaving;
    if (pendingWriters.size && !keepMessage) {
      $('writerStatus').className = 'hint writer-status';
      $('writerStatus').textContent = `${pendingWriters.size} 位撰寫人尚未儲存，請按「儲存撰寫人」。`;
    }
  }

  // Writes every typed name into ADR通報.xlsx in one go and says plainly
  // whether it worked.
  async function saveWriters() {
    if (!pendingWriters.size || writerSaving) return;
    if (saving) {
      $('writerStatus').textContent = '正在存入卡片，請稍後再按一次「儲存撰寫人」。';
      return;
    }
    writerSaving = true;
    showWriterState(true);
    $('writerStatus').className = 'hint writer-status';
    $('writerStatus').textContent = '儲存中…';
    const batch = new Map(pendingWriters);
    let ok = false;
    try {
      const handle = await existingFile(folder, WORKBOOK_NAME);
      if (!handle) throw new Error(`找不到 ${WORKBOOK_NAME}`);
      const bytes = new Uint8Array(await (await handle.getFile()).arrayBuffer());
      const result = W.setWriters(window.XLSX, bytes, batch);
      if (result.changed.length) await writeFile(folder, WORKBOOK_NAME, result.bytes);
      batch.forEach((name, id) => { if (pendingWriters.get(id) === name) pendingWriters.delete(id); });
      result.missing.forEach(id => pendingWriters.delete(id));
      const saved = [...batch].filter(([id]) => !result.missing.includes(id));
      $('writerStatus').className = 'hint writer-status saved';
      $('writerStatus').textContent = `✓ 已存入 ${WORKBOOK_NAME}：` +
        saved.map(([id, n]) => `${id} ${n || '（清除）'}`).join('、') +
        (result.missing.length ? `。Excel 中找不到 ${result.missing.join('、')}，這些沒有存入。` : '');
      ok = true;
    } catch (e) {
      const locked = /開啟中|Excel|NoModification|InvalidState|lock|Permission/i.test(e.name + ' ' + e.message);
      $('writerStatus').className = 'hint writer-status failed';
      $('writerStatus').textContent = locked
        ? `沒有存入：${WORKBOOK_NAME} 正在 Excel 中開著（或雲端硬碟正在同步）。請關閉 Excel，等幾秒後再按一次「儲存撰寫人」。名字都還留在格子裡。`
        : `沒有存入：${e.message}。名字都還留在格子裡，可以再按一次「儲存撰寫人」。`;
    } finally {
      writerSaving = false;
      showWriterState(true);
    }
    if (ok) {
      const message = $('writerStatus').textContent;
      await refreshLedger();
      $('writerStatus').textContent = message;
    }
  }

  async function writeFile(dir, name, data) {
    const handle = await dir.getFileHandle(name, { create: true });
    const w = await handle.createWritable();
    await w.write(data);
    await w.close();
  }

  // The workbook keeps HIS's own text; a value fixed by hand on the card
  // replaces it. A changed 識別碼 on the card is written as well.
  function excelRow(card) {
    const edited = (shown, raw, clean) => (shown !== clean(raw) ? shown : raw);
    return Object.assign({}, card, {
      drugRaw: edited(card.drug, card.drugRaw, P.cleanDrug),
      symptomRaw: edited(card.symptom, card.symptomRaw, P.cleanSymptom),
    });
  }

  async function saveToFolder() {
    // One click: link if needed, then re-sync with Excel so the numbers are
    // current (unless the user typed a start number) before writing anything.
    if (!folder && !await linkFolder()) return;
    await refreshLedger();
    const btn = $('saveBtn');
    const msgs = [];
    saving = true;
    btn.disabled = true;
    btn.textContent = '儲存中…';
    try {
      if (await folder.requestPermission({ mode: 'readwrite' }) !== 'granted') {
        throw new Error('沒有資料夾寫入權限');
      }
      const ids = cards.map(c => c.id);
      const dup = [...new Set(ids.filter((id, i) => ids.indexOf(id) !== i))];
      if (dup.length) throw new Error(`識別碼重複：${dup.join('、')}。請先修改卡片上的識別碼。`);

      // 1. Excel: append new cases, overwrite the ones already there.
      const fileHandle = await existingFile(folder, WORKBOOK_NAME);
      const original = fileHandle ? new Uint8Array(await (await fileHandle.getFile()).arrayBuffer()) : null;
      const result = W.appendRows(window.XLSX, P, original, cards.map(excelRow));
      if (result.added.length || result.updated.length || result.removed.length) {
        try {
          await writeFile(folder, WORKBOOK_NAME, result.bytes);
        } catch (e) {
          throw new Error(`無法寫入 ${WORKBOOK_NAME}（若檔案在 Excel 中開著，請先關閉）：${e.name} ${e.message}`);
        }
        if (result.added.length) {
          const sheets = [...new Set(result.added.map(a => a.sheet))].join('、');
          msgs.push({ cls: 'ok', text: `${WORKBOOK_NAME}：新增 ${result.added.length} 列到工作表 ${sheets}${original ? '' : '（新建檔案）'}` });
        }
        if (result.updated.length) {
          msgs.push({ cls: 'ok', text: `${WORKBOOK_NAME}：覆寫 ${result.updated.length} 列：${result.updated.map(u => u.row.id).join('、')}` });
        }
        if (result.removed.length) {
          msgs.push({ cls: 'ok', text: `${WORKBOOK_NAME}：刪除被取代的 ${result.removed.length} 列：` +
            result.removed.map(d => `${d.values[0]} ${d.values[2]} ${P.cleanDrug(d.values[4])}`).join('、') });
        }
      }
      if (result.duplicates.length) {
        msgs.push({ cls: 'ok', text: `${result.duplicates.length} 列在 Excel 中內容相同，不需更動：${result.duplicates.map(r => r.id).join('、')}` });
      }

      // 2. PNGs: always write the card as it is now, replacing an older image.
      const saved = [];
      const replaced = [];
      for (const card of cards) {
        const dir = await subfolder('ADR_' + card.id.slice(0, 5), true);
        const name = `ADR_${card.id}.png`;
        (await existingFile(dir, name) ? replaced : saved).push(name);
        await writeFile(dir, name, await cardBlob(card));
      }
      if (saved.length) msgs.push({ cls: 'ok', text: `已儲存 ${saved.length} 張卡片圖檔：${saved.join('、')}` });
      if (replaced.length) msgs.push({ cls: 'ok', text: `已覆寫 ${replaced.length} 張卡片圖檔：${replaced.join('、')}` });

      // 3. A case moved to another number leaves its old image behind: remove it.
      const gone = [];
      for (const card of cards.filter(c => c.previousId && !ids.includes(c.previousId))) {
        const name = `ADR_${card.previousId}.png`;
        const dir = await subfolder('ADR_' + card.previousId.slice(0, 5), false);
        try {
          if (dir && await existingFile(dir, name)) { await dir.removeEntry(name); gone.push(name); }
        } catch (e) {
          msgs.push({ cls: 'warn', text: `無法刪除 ${name}，請自行刪除：${e.message}` });
        }
      }
      if (gone.length) msgs.push({ cls: 'ok', text: `已刪除改號案例的舊圖檔：${gone.join('、')}` });
      // Saved: the numbers are in Excel now, so the cards follow Excel again.
      fixedIds.clear();
    } catch (e) {
      msgs.push({ cls: 'warn', text: e.message });
    } finally {
      saving = false;
      btn.textContent = '存入 Excel 與 PNG（全部）';
      updateButtons();
      renderList($('saveMessages'), msgs);
    }
    await refreshLedger();
  }

  // ---- Wiring ---------------------------------------------------------------

  pasteEl.addEventListener('input', () => {
    startTouched = false;
    $('saveMessages').replaceChildren();
    rebuild();
  });
  startEl.addEventListener('input', () => { startTouched = true; rebuild(); });
  $('clearBtn').addEventListener('click', () => {
    pasteEl.value = '';
    startTouched = false;
    rebuild();
    pasteEl.focus();
  });
  $('statusMonth').addEventListener('change', e => {
    shownMonth = e.target.value;
    renderStatus();
  });
  window.addEventListener('focus', () => refreshLedger());
  $('saveWritersBtn').addEventListener('click', saveWriters);
  // Typed names not saved yet: ask before the window is closed.
  window.addEventListener('beforeunload', e => { if (pendingWriters.size) { e.preventDefault(); e.returnValue = ''; } });
  $('linkBtn').addEventListener('click', linkFolder);
  $('saveBtn').addEventListener('click', saveToFolder);
  $('pdfBtn').addEventListener('click', makePdf);
  $('selAll').addEventListener('click', () => setAll(true));
  $('selNone').addEventListener('click', () => setAll(false));
  $('pngBtn').addEventListener('click', downloadPngs);
  // Also covers Ctrl+P.
  window.addEventListener('beforeprint', buildPrintArea);

  rebuild();
  initFolder();
})();
