// Where card images (ADR_<識別碼>.png) live in the ADR folder. They can be in
// the month folder itself, or in an ADR_cards folder, either inside each
// month folder or at the top with or without month folders:
//
//   ADR_11511\ADR_1151101.png
//   ADR_11511\ADR_cards\ADR_1151101.png
//   ADR_cards\ADR_11511\ADR_1151101.png
//   ADR_cards\ADR_1151101.png
//
// Cards are found in any of these; new ones go where the cards already are.
(function (root) {
  'use strict';

  const CARDS = 'ADR_cards';
  const MONTH = /^ADR_\d{5}$/;

  async function child(dir, name, create = false) {
    if (!dir) return null;
    try { return await dir.getDirectoryHandle(name, { create }); }
    catch (e) { if (e.name === 'NotFoundError' || e.name === 'TypeMismatchError') return null; throw e; }
  }

  // How this ADR folder keeps its cards: 'month' (in the month folder),
  // 'month-cards' (ADR_yyyMM\ADR_cards), 'cards-month' (ADR_cards\ADR_yyyMM)
  // or 'cards' (all in ADR_cards).
  async function layout(folder) {
    const top = await child(folder, CARDS);
    if (top) {
      for await (const name of top.keys()) if (MONTH.test(name) && await child(top, name)) return 'cards-month';
      return 'cards';
    }
    for await (const name of folder.keys()) {
      if (MONTH.test(name) && await child(await child(folder, name), CARDS)) return 'month-cards';
    }
    return 'month';
  }

  // Every folder that may hold the month's cards (key "11511").
  async function dirsOf(folder, key) {
    const month = await child(folder, 'ADR_' + key);
    const top = await child(folder, CARDS);
    return [await child(month, CARDS), await child(top, 'ADR_' + key), top, month].filter(Boolean);
  }

  // The folder new cards of the month go into, made if needed.
  async function dirFor(folder, key, style) {
    style = style || await layout(folder);
    if (style === 'cards') return child(folder, CARDS, true);
    if (style === 'cards-month') return child(await child(folder, CARDS, true), 'ADR_' + key, true);
    const month = await child(folder, 'ADR_' + key, true);
    return style === 'month-cards' ? child(month, CARDS, true) : month;
  }

  // Names of every file in the month's card folders.
  async function namesOf(folder, key) {
    const names = new Set();
    for (const dir of await dirsOf(folder, key)) for await (const name of dir.keys()) names.add(name);
    return names;
  }

  // Every card image in the ADR folder: [{ id, month, dir, name }].
  async function allCards(folder) {
    const found = new Map();
    const take = async dir => {
      for await (const name of dir.keys()) {
        const m = name.match(/^ADR_(\d{7})\.png$/i);
        if (m && !found.has(m[1])) found.set(m[1], { id: m[1], month: m[1].slice(0, 5), dir, name });
      }
    };
    const months = [];
    for await (const name of folder.keys()) if (MONTH.test(name)) months.push(name.slice(4));
    const top = await child(folder, CARDS);
    if (top) for await (const name of top.keys()) if (MONTH.test(name)) months.push(name.slice(4));
    for (const key of [...new Set(months)]) for (const dir of await dirsOf(folder, key)) await take(dir);
    if (top) await take(top);
    return [...found.values()];
  }

  // The PDF saved after reporting a case to TFDA, ADR_<識別碼>_通報.pdf, in the
  // month folder (or its ADR_cards): Map(識別碼 -> "2026/11/04", the day it
  // was saved). Its being there means the case was reported.
  const REPORT_PDF = /^ADR_(\d{7})_通報\.pdf$/i;
  async function reportDates(folder) {
    const dates = new Map();
    const day = d => `${d.getFullYear()}/${String(d.getMonth() + 1).padStart(2, '0')}/${String(d.getDate()).padStart(2, '0')}`;
    for await (const name of folder.keys()) {
      if (!MONTH.test(name)) continue;
      const month = await child(folder, name);
      for (const dir of [month, await child(month, CARDS)].filter(Boolean)) {
        for await (const file of dir.keys()) {
          const m = file.match(REPORT_PDF);
          if (!m || dates.has(m[1])) continue;
          try {
            const handle = await dir.getFileHandle(file);
            // The helper says when without sending the file; the browser's own handle gives a lazy File.
            const time = handle.modified ? await handle.modified() : (await handle.getFile()).lastModified;
            dates.set(m[1], day(new Date(time)));
          } catch (e) { /* unreadable: skip */ }
        }
      }
    }
    return dates;
  }

  root.ADRCardFolders = { layout, dirsOf, dirFor, namesOf, allCards, reportDates, CARDS };
})(window);
