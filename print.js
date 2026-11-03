// Picks saved card images (ADR_<識別碼>.png in the ADR_<yyyMM> folders) and
// merges the chosen ones into an A4 PDF, 4 per page, laid out like app.js.
(function () {
  'use strict';

  const CARDS_PER_PAGE = 4;
  // Landscape A4 at 250 dpi, 2x2 cards of 138x94 mm (same as app.js).
  const PAGE_MM = { w: 297, h: 210, cardW: 138, cardH: 94, gap: 1 };
  const DPI = 250;
  const CARD_FILE = /^ADR_(\d{7})\.png$/i;
  const MONTH_DIR = /^ADR_(\d{5})$/;

  const $ = id => document.getElementById(id);
  let folder = null;
  const cards = new Map();     // 識別碼 -> { id, month, file: () => Promise<Blob>, url }
  const chosen = new Set();    // 識別碼

  const rocMonth = key => `${key.slice(0, 3)} 年 ${Number(key.slice(3))} 月`;

  function message(cls, text, href, download) {
    const li = document.createElement('li');
    li.className = cls;
    li.textContent = text;
    if (href) {
      const a = document.createElement('a');
      a.href = href;
      a.download = download;
      a.textContent = '下載';
      li.append(' ', a);
    }
    $('messages').replaceChildren(li);
  }

  // ---- Loading cards ----------------------------------------------------------

  async function loadFolder() {
    folder = await window.ADRLocal.detect();
    if (!folder) {
      $('folderStatus').textContent = '未連到本機程式：請用桌面「ADR 通報卡」開啟，或選擇卡片圖檔。';
      return;
    }
    $('folderStatus').textContent = `已連結：${folder.name}`;
    $('folderStatus').classList.add('linked');
    for await (const name of folder.keys()) {
      const m = name.match(MONTH_DIR);
      if (!m) continue;
      const dir = await folder.getDirectoryHandle(name);
      for await (const file of dir.keys()) {
        const c = file.match(CARD_FILE);
        if (!c) continue;
        cards.set(c[1], { id: c[1], month: m[1], file: async () => (await dir.getFileHandle(file)).getFile() });
      }
    }
  }

  function addPicked(files) {
    const ids = [];
    [...files].forEach(f => {
      const c = f.name.match(CARD_FILE);
      if (!c) return;
      const old = cards.get(c[1]);
      if (old && old.url) URL.revokeObjectURL(old.url);
      cards.set(c[1], { id: c[1], month: c[1].slice(0, 5), file: async () => f });
      ids.push(c[1]);
    });
    if (ids.length < files.length) message('warn', `${files.length - ids.length} 個檔案不是 ADR_<識別碼>.png，已略過。`);
    renderMonths(ids.length ? ids[ids.length - 1].slice(0, 5) : null);
  }

  // ---- Rendering -------------------------------------------------------------

  function renderMonths(prefer) {
    const months = [...new Set([...cards.values()].map(c => c.month))].sort().reverse();
    const sel = $('month');
    const keep = prefer || sel.value || months[0];
    sel.replaceChildren(...months.map(k => {
      const o = document.createElement('option');
      o.value = k;
      o.textContent = `${rocMonth(k)}（${[...cards.values()].filter(c => c.month === k).length} 張）`;
      return o;
    }));
    if (months.includes(keep)) sel.value = keep;
    renderGrid();
  }

  function renderGrid() {
    const month = $('month').value;
    const list = [...cards.values()].filter(c => c.month === month).sort((a, b) => a.id.localeCompare(b.id));
    if (!list.length) {
      const p = document.createElement('p');
      p.className = 'hint';
      p.textContent = cards.size ? '這個月沒有卡片。' : '找不到卡片圖檔（ADR_<yyyMM> 資料夾中的 ADR_<識別碼>.png）。';
      $('grid').replaceChildren(p);
    } else {
      $('grid').replaceChildren(...list.map(tile));
    }
    updateCount();
  }

  function tile(card) {
    const el = document.createElement('button');
    el.type = 'button';
    el.className = 'pr-card' + (chosen.has(card.id) ? ' chosen' : '');
    el.setAttribute('aria-pressed', chosen.has(card.id));
    const img = document.createElement('img');
    img.alt = card.id;
    img.loading = 'lazy';
    if (card.url) img.src = card.url;
    else card.file().then(b => { card.url = URL.createObjectURL(b); img.src = card.url; })
      .catch(() => { img.alt = `${card.id}（無法讀取）`; });
    const label = document.createElement('span');
    label.className = 'pr-label';
    label.textContent = card.id;
    el.append(img, label);
    el.addEventListener('click', () => {
      if (chosen.has(card.id)) chosen.delete(card.id); else chosen.add(card.id);
      el.classList.toggle('chosen', chosen.has(card.id));
      el.setAttribute('aria-pressed', chosen.has(card.id));
      updateCount();
    });
    return el;
  }

  function updateCount() {
    const n = chosen.size;
    const months = [...new Set([...chosen].map(id => id.slice(0, 5)))];
    $('count').textContent = n
      ? `已選 ${n} 張，${Math.ceil(n / CARDS_PER_PAGE)} 頁` + (months.length > 1 ? `（含 ${months.map(rocMonth).join('、')}）` : '')
      : '尚未選取';
    $('pdfBtn').disabled = !n;
  }

  // ---- PDF ---------------------------------------------------------------------

  async function bitmap(card) {
    return createImageBitmap(await card.file());
  }

  async function buildPdf(list) {
    const px = mm => Math.round(mm / 25.4 * DPI);
    const canvas = document.createElement('canvas');
    canvas.width = px(PAGE_MM.w);
    canvas.height = px(PAGE_MM.h);
    const ctx = canvas.getContext('2d');
    ctx.imageSmoothingQuality = 'high';
    const gridW = PAGE_MM.cardW * 2 + PAGE_MM.gap;
    const gridH = PAGE_MM.cardH * 2 + PAGE_MM.gap;
    const left = (PAGE_MM.w - gridW) / 2;
    const top = (PAGE_MM.h - gridH) / 2;
    const pages = [];
    for (let i = 0; i < list.length; i += CARDS_PER_PAGE) {
      ctx.fillStyle = '#fff';
      ctx.fillRect(0, 0, canvas.width, canvas.height);
      const images = await Promise.all(list.slice(i, i + CARDS_PER_PAGE).map(bitmap));
      images.forEach((img, n) => {
        const col = n % 2, row = Math.floor(n / 2);
        ctx.drawImage(img,
          px(left + col * (PAGE_MM.cardW + PAGE_MM.gap)), px(top + row * (PAGE_MM.cardH + PAGE_MM.gap)),
          px(PAGE_MM.cardW), px(PAGE_MM.cardH));
        img.close();
      });
      const blob = await new Promise(r => canvas.toBlob(r, 'image/jpeg', 0.92));
      pages.push({ jpeg: new Uint8Array(await blob.arrayBuffer()), width: canvas.width, height: canvas.height });
    }
    return new Blob([window.ADRPdf.build(pages)], { type: 'application/pdf' });
  }

  async function makePdf() {
    const list = [...chosen].sort().map(id => cards.get(id));
    // Open the tab now: browsers block pop-ups opened after an await.
    const tab = window.open('', '_blank');
    $('pdfBtn').disabled = true;
    try {
      const url = URL.createObjectURL(await buildPdf(list));
      const name = `ADR_${list[0].id}-${list[list.length - 1].id}.pdf`;
      if (tab) tab.location = url;
      message(tab ? 'ok' : 'warn',
        `PDF：${list.length} 張卡片，共 ${Math.ceil(list.length / CARDS_PER_PAGE)} 頁 A4` +
        (tab ? '，已在新分頁開啟，可直接列印。' : '。瀏覽器擋住了新分頁，請用右側連結。'), url, name);
    } catch (e) {
      if (tab) tab.close();
      message('warn', '無法產生 PDF：' + e.message);
    }
    updateCount();
  }

  // ---- Wiring --------------------------------------------------------------------

  $('month').addEventListener('change', renderGrid);
  $('pick').addEventListener('change', e => { addPicked(e.target.files); e.target.value = ''; });
  $('allBtn').addEventListener('click', () => {
    [...cards.values()].filter(c => c.month === $('month').value).forEach(c => chosen.add(c.id));
    renderGrid();
  });
  $('noneBtn').addEventListener('click', () => { chosen.clear(); renderGrid(); });
  $('pdfBtn').addEventListener('click', makePdf);

  loadFolder()
    .catch(e => message('warn', '無法讀取 ADR 資料夾：' + e.message))
    .then(() => renderMonths());
})();
