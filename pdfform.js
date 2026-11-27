// Reads the report form printed to PDF from the fill-in page
// (ADR通報表_空白表單.html → 列印 → 另存為 PDF), for when the system's Excel
// export does not work. Gives the same case object as ADRForms.readCases, so
// the TFDA helper treats both alike.
//
// The PDF has the typed answers as text, but a ticked choice (性別, 來源,
// 結果, 嚴重性, Naranjo …) only as a drawn blue dot. readPdf (browser only,
// with pdf.js) draws each page and marks every option that has a circle
// left of it (`circle`) and whether the circle is filled (`ticked`).
// parsePages (pure, tested in Node) turns that into the case.
(function (root) {
  'use strict';

  const VALUE_X = 190;          // left of this: labels; right: answers
  const norm = s => String(s).normalize('NFKC');   // ⽣ → 生, ⼀ → 一 …
  const squash = s => norm(s).replace(/\s+/g, '');

  // Left-column labels (start of the text, without spaces) → field.
  // kind: text (typed answer), radio (one ticked option), age, skip.
  const LABELS = [
    ['通報者姓名', 'reporter', 'text'],
    ['不良反應發生之最初日期', 'onset', 'text'],
    ['通報者獲知日期', 'learned', 'text'],
    ['國內案件', 'domestic', 'radio'],
    ['原始藥物不良反應獲知來源', 'source', 'radio'],
    ['本案件是否為文獻通報', 'literature', 'radio'],
    ['識別代碼', 'id', 'text'],
    ['原通報單位識別代碼', 'chartNo', 'text'],
    ['性別', 'sex', 'radio'],
    ['出生日期或年齡', 'age', 'age'],
    ['身高', 'height', 'text'],
    ['體重', 'weight', 'text'],
    ['不良反應症狀', 'symptoms', 'text'],
    ['處置', 'treatment', 'text'],
    ['不良反應後續結果', 'outcome', 'radio'],
    ['嚴重不良反應【院內', 'hartwig', 'radio'],
    ['不良反應嚴重性', 'seriousness', 'radio'],
    ['案件描述', 'description', 'text'],
    ['相關檢查及檢驗數據', 'labs', 'text'],
    ['檢查描述', 'labNote', 'text'],
    ['其他相關資訊', 'other', 'text'],
    ['可疑藥品', 'suspect', 'text'],
    ['併用產品', 'concomitant', 'text'],
    ['交互作用產品', 'interacting', 'text'],
    ['請將上述問題分數加總', 'naranjoScore', 'text'],
    ['Naranjo最終評估', 'naranjo', 'radio'],
  ];
  // Section headings: the answer before them has ended.
  const HEADINGS = ['通報者資料', '一般資訊', '病患資訊', '不良反應資訊', '用藥相關資訊', 'Naranjoscore', '案例標題'];

  // Items of one page → lines (same baseline), each sorted left to right.
  function lines(items) {
    const sorted = items.slice().sort((a, b) => a.y - b.y || a.x - b.x);
    const out = [];
    for (const it of sorted) {
      const last = out[out.length - 1];
      if (last && Math.abs(last.y - it.y) <= 3) last.items.push(it);
      else out.push({ y: it.y, items: [it] });
    }
    out.forEach(l => l.items.sort((a, b) => a.x - b.x));
    return out;
  }

  // Pieces of one line joined, with a space where there is a visible gap.
  function join(items) {
    let text = '';
    let end = null;
    for (const it of items) {
      if (end != null && it.x - end > 1.5 && !/\s$/.test(text)) text += ' ';
      text += norm(it.str);
      end = it.x + (it.w || 0);
    }
    return text.trim();
  }

  // Options on a line: each starts at a piece with a circle left of it. A
  // circle needs room: a piece right after another (e.g. 生 in 已知出生日期,
  // whose left neighbour is text, not a circle) never starts an option.
  const CIRCLE_ROOM = 8;
  function options(items) {
    const out = [];
    let end = -Infinity;
    for (const it of items) {
      const starts = it.circle && it.x - end >= CIRCLE_ROOM;
      if (starts || !out.length) out.push({ ticked: starts && !!it.ticked, items: [it] });
      else out[out.length - 1].items.push(it);
      end = it.x + (it.w || 0);
    }
    return out.filter(o => o.items[0].circle).map(o => ({ ticked: o.ticked, text: join(o.items) }));
  }

  const NARANJO_Q = /^(\d{1,2})\./;

  // pages: [{ height, items: [{ x, y, w, str, circle, ticked }] }] with y
  // measured from the top. Returns a case like ADRForms.readCases, or null
  // when this is not the report form.
  function parsePages(pages) {
    const values = {};
    let field = null;           // [key, kind] being read
    let naranjoSum = null;
    let isForm = false;
    for (const page of pages) {
      const body = page.items.filter(it => it.str.trim() && it.y > 32 && it.y < (page.height || 842) - 20);
      for (const line of lines(body)) {
        const left = line.items.filter(it => it.x < VALUE_X);
        const right = line.items.filter(it => it.x >= VALUE_X);
        const label = squash(join(left));
        if (/藥品不良反應\(ADR\)通報表/.test(squash(join(line.items)))) isForm = true;
        if (label) {
          const hit = LABELS.find(([prefix]) => label.startsWith(prefix));
          if (hit) field = [hit[1], hit[2]];
          else if (NARANJO_Q.test(label)) field = ['naranjoQ', 'naranjoQ'];
          else if (HEADINGS.some(h => label.startsWith(h))) field = null;
          // Anything else is the rest of a label that wraps: keep reading.
        }
        if (!field || !right.length) continue;
        const [key, kind] = field;
        if (kind === 'text') {
          values[key] = (values[key] ? values[key] + '\n' : '') + join(right);
        } else if (kind === 'radio' || kind === 'age') {
          const picked = options(right).find(o => o.ticked);
          if (!picked) continue;
          values[key] = kind === 'age' && /^不知/.test(squash(picked.text)) ? '未知' : picked.text;
        } else if (kind === 'naranjoQ') {
          const picked = options(right).find(o => o.ticked);
          const m = picked && squash(picked.text).replace(/−/g, '-').match(/\(([+-]?\d+)\)/);
          if (m) naranjoSum = (naranjoSum || 0) + Number(m[1]);
        }
      }
    }
    if (!isForm && !values.id) return null;
    const unit = v => String(v || '').replace(/\s*(cm|kg)\s*$/i, '').replace(/^未知$/, '').trim();
    const c = {
      trainee: '', reporter: '', onset: '', learned: '', domestic: '', source: '', literature: '',
      id: '', chartNo: '', sex: '', age: '', height: '', weight: '', symptoms: '', treatment: '',
      outcome: '', seriousness: '', hartwig: '', description: '', labs: '', labNote: '', other: '', suspect: '',
      concomitant: '', interacting: '', naranjoScore: '', naranjo: '', submitted: '',
    };
    for (const k of Object.keys(c)) if (values[k] != null) c[k] = values[k].trim();
    c.id = squash(c.id);
    c.chartNo = squash(c.chartNo).toUpperCase();
    c.height = unit(c.height);
    c.weight = unit(c.weight);
    if (!c.naranjoScore && naranjoSum != null) c.naranjoScore = String(naranjoSum);
    c.naranjo = c.naranjo.replace(/\s*\([^)]*\)\s*$/, '');
    c.filled = Object.values(c).filter(Boolean).length;
    c.fromPdf = true;
    return c;
  }

  // ---- browser: PDF → pages with ticks (needs pdf.js) ----

  const PDFJS = 'https://cdn.jsdelivr.net/npm/pdfjs-dist@3.11.174/build/';
  let ready = null;
  const LOAD_TIMEOUT = 20000;

  function loadPdfJs() {
    if (!ready) {
      const fail = () => new Error('PDF 讀取元件無法載入（網路不通或被擋）');
      ready = (async () => {
        if (!root.pdfjsLib) {
          await new Promise((resolve, reject) => {
            const s = document.createElement('script');
            s.src = PDFJS + 'pdf.min.js';
            // The exact published files only: a changed copy (CDN, proxy) is refused.
            s.integrity = 'sha384-/1qUCSGwTur9vjf/z9lmu/eCUYbpOTgSjmpbMQZ1/CtX2v/WcAIKqRv+U1DUCG6e';
            s.crossOrigin = 'anonymous';
            s.onload = resolve;
            s.onerror = () => reject(fail());
            // A blocked site can leave the request hanging: give up after a while.
            setTimeout(() => reject(fail()), LOAD_TIMEOUT);
            document.head.appendChild(s);
          });
        }
        // A worker from another site must be started from a local copy.
        const res = await fetch(PDFJS + 'pdf.worker.min.js', { signal: AbortSignal.timeout(LOAD_TIMEOUT),
          integrity: 'sha384-SnzOobpRMLXZ52iJvZm/C0fYw0OQemTXzTjIsdsfMcrCtCEe9qgzxTd3RSklO5x2' }).catch(() => { throw fail(); });
        if (!res.ok) throw new Error('PDF 讀取元件無法載入');
        root.pdfjsLib.GlobalWorkerOptions.workerSrc = URL.createObjectURL(new Blob([await res.text()], { type: 'text/javascript' }));
        return root.pdfjsLib;
      })();
      // A failure stands for the rest of this visit (no 20 s wait per PDF);
      // retry() lets 重新讀取 try again.
      ready.catch(() => {});
    }
    return ready;
  }

  const SCALE = 3;
  // Circle left of an answer piece: a ring of grey pixels; ticked: blue fill.
  function circleAt(ctx, x, y) {
    const box = ctx.getImageData(Math.round((x - 14) * SCALE), Math.round((y - 11) * SCALE), 11 * SCALE, 10 * SCALE).data;
    let ink = 0, blue = 0;
    for (let i = 0; i < box.length; i += 4) {
      const [r, g, b] = [box[i], box[i + 1], box[i + 2]];
      if (r < 200 || g < 200 || b < 200) ink++;
      if (b > 150 && r < 100) blue++;
    }
    return { circle: ink > 8, ticked: blue > 12 };
  }

  async function readPdf(bytes) {
    const pdfjs = await loadPdfJs();
    // pdf.js takes over (detaches) the bytes it is given, which emptied the
    // caller's copy: an uploaded PDF was then saved as a 0-byte file.
    const doc = await pdfjs.getDocument({ data: bytes.slice() }).promise;
    const pages = [];
    for (let n = 1; n <= doc.numPages; n++) {
      const page = await doc.getPage(n);
      const vp1 = page.getViewport({ scale: 1 });
      const items = (await page.getTextContent()).items
        .filter(it => it.str && it.str.trim())
        .map(it => ({ x: it.transform[4], y: vp1.height - it.transform[5], w: it.width, str: it.str }));
      // Only the answer column needs looking at for circles.
      if (items.some(it => it.x >= VALUE_X)) {
        const vp = page.getViewport({ scale: SCALE });
        const canvas = document.createElement('canvas');
        canvas.width = Math.ceil(vp.width);
        canvas.height = Math.ceil(vp.height);
        const ctx = canvas.getContext('2d', { willReadFrequently: true });
        await page.render({ canvasContext: ctx, viewport: vp }).promise;
        items.filter(it => it.x >= VALUE_X).forEach(it => Object.assign(it, circleAt(ctx, it.x, it.y)));
      }
      pages.push({ height: vp1.height, items });
    }
    doc.destroy();
    return parsePages(pages);
  }

  const retry = () => { ready = null; };

  // Reading draws every page, so each PDF is read once: the result is kept
  // on this PC by file name and size, shared by the pages that read forms.
  const CACHE_KEY = 'adr-pdf-cases-v2';   // v2: with Hartwig
  let cache = null;
  function cacheMap() {
    if (!cache) {
      try { cache = new Map(Object.entries(JSON.parse(localStorage.getItem(CACHE_KEY) || '{}'))); } catch (e) { cache = new Map(); }
    }
    return cache;
  }
  // `modified` (the file's date) tells a corrected PDF of the same size apart.
  async function readCached(name, bytes, modified = '') {
    const key = name + '|' + bytes.length + '|' + modified;
    const map = cacheMap();
    if (!map.has(key)) {
      // A read that hangs (seen once on the user's PC) must not leave the
      // page on 讀取中 forever with cases silently missing.
      let timer;
      const timeout = new Promise((_, reject) => { timer = setTimeout(() => reject(new Error('讀取逾時（超過 60 秒），請按「重新讀取」')), 60000); });
      try { map.set(key, await Promise.race([readPdf(bytes), timeout])); } finally { clearTimeout(timer); }
      try { localStorage.setItem(CACHE_KEY, JSON.stringify(Object.fromEntries([...map].slice(-300)))); } catch (e) { /* saves time only */ }
    }
    return map.get(key);
  }

  // Files that may hold report forms: Excel exports and printed PDFs, not
  // ADR通報.xlsx, Excel lock files, the PDFs this app makes (ADR_<識別碼>…),
  // the meeting list, or examples and blank forms.
  const isFormFile = name => !name.startsWith('~$') && name !== 'ADR通報.xlsx' &&
    !/範例|空白|樣本|template|例會名單/i.test(name) &&
    (/\.xlsx?$/i.test(name) || (/\.pdf$/i.test(name) && !/^ADR_\d/i.test(name)));

  const api = { parsePages, readPdf, readCached, isFormFile, retry };
  if (typeof module !== 'undefined' && module.exports) module.exports = api;
  else root.ADRPdfForm = api;
})(typeof window !== 'undefined' ? window : globalThis);
