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
  let startTouched = false;   // user typed a start number; stop suggesting

  // ---- Parsing and cards ----------------------------------------------------

  function rebuild() {
    const { rows, skipped } = P.parseText(pasteEl.value);
    cards = P.assignIds(rows, Number(startEl.value) || 1).map(c => Object.assign(c, {
      dateText: P.formatDate(c.date),
      include: true,
    }));
    renderMessages(skipped);
    renderCards();
    $('saveMessages').replaceChildren();
    suggestStart();
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
    if (months.length > 1) {
      msgs.push({ cls: 'warn', text: `包含多個月份（${months.join('、')}）：起始流水號只套用於第一個月，其後月份從 01 開始。` });
    }
    if (cards.length) msgs.unshift({ cls: 'ok', text: `共 ${cards.length} 張卡片，已依得知日期時間排序。` });
    renderList(messagesEl, msgs);
    $('idHint').textContent = cards.length
      ? `識別碼：${cards[0].id} – ${cards[cards.length - 1].id}` : '';
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
        value.addEventListener('input', () => { cards[i][f.key] = value.textContent.trim(); });
        row.append(label, value);
        el.appendChild(row);
      });
      return el;
    }));
    updateButtons();
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
    $('saveBtn').disabled = !cards.length || !folder;
    $('saveBtn').hidden = !folder;
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
    suggestStart();
  }

  async function linkFolder() {
    const stored = await savedFolder();
    // Reconnect to the remembered folder if that is all that's needed.
    if (stored && !folder && await stored.requestPermission({ mode: 'readwrite' }) === 'granted') {
      return useFolder(stored);
    }
    try {
      const handle = await window.showDirectoryPicker({ id: 'adr', mode: 'readwrite' });
      await saveFolder(handle);
      useFolder(handle);
    } catch (e) {
      if (e.name !== 'AbortError') alert('無法連結資料夾：' + e.message);
    }
  }

  async function initFolder() {
    if (!window.showDirectoryPicker) {
      $('linkBtn').hidden = true;
      $('folderUnsupported').hidden = false;
      return;
    }
    const stored = await savedFolder();
    if (!stored) return;
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

  // Next sequence number = highest ADR_<yyyMM>NN.png in ADR_<yyyMM>\ + 1.
  async function nextSeq(key) {
    const dir = await subfolder('ADR_' + key, false);
    if (!dir) return 1;
    const pattern = new RegExp(`^ADR_${key}(\\d{2})\\.png$`, 'i');
    let max = 0;
    for await (const name of dir.keys()) {
      const m = name.match(pattern);
      if (m) max = Math.max(max, Number(m[1]));
    }
    return max + 1;
  }

  let suggestToken = 0;
  async function suggestStart() {
    if (!folder || startTouched || !cards.length) return;
    const token = ++suggestToken;
    const key = cards[0].id.slice(0, 5);
    try {
      const seq = await nextSeq(key);
      if (token !== suggestToken || startTouched) return;
      if (Number(startEl.value) !== seq) {
        startEl.value = seq;
        rebuild();  // runs suggestStart again, which adds the hint
        return;
      }
      $('idHint').textContent += `（依 ADR_${key} 已有卡片自動帶入）`;
    } catch (e) { /* folder unavailable; keep the typed number */ }
  }

  async function writeFile(dir, name, data) {
    const handle = await dir.getFileHandle(name, { create: true });
    const w = await handle.createWritable();
    await w.write(data);
    await w.close();
  }

  async function saveToFolder() {
    const btn = $('saveBtn');
    const msgs = [];
    btn.disabled = true;
    btn.textContent = '儲存中…';
    try {
      if (await folder.requestPermission({ mode: 'readwrite' }) !== 'granted') {
        throw new Error('沒有資料夾寫入權限');
      }

      // 1. Excel: append new rows to the month's sheet.
      const fileHandle = await existingFile(folder, WORKBOOK_NAME);
      const original = fileHandle ? new Uint8Array(await (await fileHandle.getFile()).arrayBuffer()) : null;
      const result = W.appendRows(window.XLSX, P, original, cards);
      if (result.added.length) {
        try {
          if (original) await writeFile(folder, WORKBOOK_NAME + '.bak', original);
          await writeFile(folder, WORKBOOK_NAME, result.bytes);
        } catch (e) {
          throw new Error(`無法寫入 ${WORKBOOK_NAME}，請先關閉 Excel 中的這個檔案再試一次。（${e.name}）`);
        }
        const sheets = [...new Set(result.added.map(a => a.sheet))].join('、');
        msgs.push({ cls: 'ok', text: `${WORKBOOK_NAME}：新增 ${result.added.length} 列到工作表 ${sheets}${original ? '' : '（新建檔案）'}` });
      }
      if (result.duplicates.length) {
        msgs.push({ cls: 'warn', text: `${result.duplicates.length} 列已在 Excel 中，未重複新增：${result.duplicates.map(r => r.chartNo).join('、')}` });
      }

      // 2. PNGs: never overwrite an existing card.
      const saved = [];
      const exists = [];
      for (const card of cards) {
        const dir = await subfolder('ADR_' + card.id.slice(0, 5), true);
        const name = `ADR_${card.id}.png`;
        if (await existingFile(dir, name)) { exists.push(name); continue; }
        await writeFile(dir, name, await cardBlob(card));
        saved.push(name);
      }
      if (saved.length) msgs.push({ cls: 'ok', text: `已儲存 ${saved.length} 張卡片圖檔：${saved.join('、')}` });
      if (exists.length) msgs.push({ cls: 'warn', text: `已存在、未覆寫（請確認起始流水號）：${exists.join('、')}` });
    } catch (e) {
      msgs.push({ cls: 'warn', text: e.message });
    } finally {
      btn.textContent = '存入 Excel 與 PNG（全部）';
      updateButtons();
      renderList($('saveMessages'), msgs);
    }
  }

  // ---- Wiring ---------------------------------------------------------------

  pasteEl.addEventListener('input', () => { startTouched = false; rebuild(); });
  startEl.addEventListener('input', () => { startTouched = true; rebuild(); });
  $('clearBtn').addEventListener('click', () => {
    pasteEl.value = '';
    startTouched = false;
    rebuild();
    pasteEl.focus();
  });
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
