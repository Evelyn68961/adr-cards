(function () {
  'use strict';
  const P = window.ADRParse;

  const FIELDS = [
    { key: 'id', label: '識別碼' },
    { key: 'chartNo', label: '病歷號' },
    { key: 'dateText', label: '得知日期' },
    { key: 'source', label: '通報人員' },
    { key: 'drug', label: '藥品' },
    { key: 'symptom', label: '症狀' },
  ];

  const CARDS_PER_PAGE = 4;

  const $ = id => document.getElementById(id);
  const pasteEl = $('paste');
  const startEl = $('startSeq');
  const sheetEl = $('sheet');
  const messagesEl = $('messages');

  let cards = [];

  function rebuild() {
    const { rows, skipped } = P.parseText(pasteEl.value);
    cards = P.assignIds(rows, Number(startEl.value) || 1).map(c => ({
      id: c.id,
      chartNo: c.chartNo,
      dateText: P.formatDate(c.date),
      source: c.source,
      drug: c.drug,
      symptom: c.symptom,
      include: true,
    }));
    renderMessages(skipped);
    renderCards();
  }

  function renderMessages(skipped) {
    const msgs = skipped.map(s => ({ cls: 'warn', text: `第 ${s.line} 行略過：${s.reason}` }));
    const months = [...new Set(cards.map(c => c.id.slice(0, 5)))];
    if (months.length > 1) {
      msgs.push({ cls: 'warn', text: `包含多個月份（${months.join('、')}）：起始流水號只套用於第一個月，其後月份從 01 開始。` });
    }
    if (cards.length) msgs.unshift({ cls: 'ok', text: `共 ${cards.length} 張卡片，已依得知日期時間排序。` });
    messagesEl.replaceChildren(...msgs.map(m => {
      const li = document.createElement('li');
      li.className = m.cls;
      li.textContent = m.text;
      return li;
    }));
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
      toggle.title = '列印 / 匯出此卡';
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

  function updateButtons() {
    const none = selected().length === 0;
    $('printBtn').disabled = none;
    $('pngBtn').disabled = none;
  }

  // Draws one card at the same size and layout as the existing ADR_*.png files.
  function cardToPng(card) {
    const W = 1128, H = 825, X = 66, LINE = 91, SIZE = 44;
    const canvas = document.createElement('canvas');
    canvas.width = W;
    canvas.height = H;
    const ctx = canvas.getContext('2d');
    ctx.fillStyle = '#fff';
    ctx.fillRect(0, 0, W, H);
    ctx.strokeStyle = '#000';
    ctx.lineWidth = 6;
    ctx.strokeRect(3, 3, W - 6, H - 6);
    ctx.fillStyle = '#000';
    ctx.font = `${SIZE}px "Noto Sans TC", "Microsoft JhengHei", sans-serif`;
    ctx.textBaseline = 'alphabetic';

    let y = 116;
    FIELDS.forEach(f => {
      const label = f.label + ': ';
      ctx.fillText(label, X, y);
      const indent = X + ctx.measureText(label).width;
      const lines = wrap(ctx, card[f.key], W - X - indent);
      lines.forEach((line, n) => {
        ctx.fillText(line, indent, y + n * 0.75 * LINE);
      });
      y += LINE + (lines.length - 1) * 0.75 * LINE;
    });
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

  async function downloadPngs() {
    if (document.fonts && document.fonts.ready) await document.fonts.ready;
    for (const card of selected()) {
      const canvas = cardToPng(card);
      const blob = await new Promise(resolve => canvas.toBlob(resolve, 'image/png'));
      const a = document.createElement('a');
      a.href = URL.createObjectURL(blob);
      a.download = `ADR_${card.id}.png`;
      a.click();
      URL.revokeObjectURL(a.href);
      // Browsers drop rapid back-to-back downloads; space them out a little.
      await new Promise(r => setTimeout(r, 250));
    }
  }

  pasteEl.addEventListener('input', rebuild);
  startEl.addEventListener('input', rebuild);
  $('clearBtn').addEventListener('click', () => {
    pasteEl.value = '';
    rebuild();
    pasteEl.focus();
  });
  $('printBtn').addEventListener('click', () => window.print());
  // Also covers Ctrl+P.
  window.addEventListener('beforeprint', buildPrintArea);
  $('pngBtn').addEventListener('click', downloadPngs);

  rebuild();
})();
