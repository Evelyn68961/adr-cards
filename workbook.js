// Appends parsed rows to the ADR workbook (one sheet per month, e.g. "202610").
// Pure: takes the file bytes and returns new bytes, so it runs in Node tests too.
(function (root) {
  'use strict';

  // Matches the existing sheet's 得知日期 column.
  const DATE_FORMAT = 'm/d/yy h:mm';

  function load(XLSX, bytes) {
    return XLSX.read(bytes, { type: 'array', cellDates: true, cellStyles: true, cellNF: true });
  }

  // Excel stores whole seconds; SheetJS can return x.999s for them.
  function roundSeconds(date) {
    return new Date(Math.round(date.getTime() / 1000) * 1000);
  }

  function existingKeys(XLSX, ws, P) {
    const keys = new Set();
    XLSX.utils.sheet_to_json(ws, { header: 1, raw: true, defval: '' }).forEach(r => {
      const date = r[8] instanceof Date ? roundSeconds(r[8]) : P.parseDate(r[8]);
      if (date) keys.add(P.rowKey(r[1], r[3], date));
    });
    return keys;
  }

  // Returns { bytes, added: [{sheet, row}], duplicates: [row] }. Rows already
  // in their month's sheet (same chart no + drug code + date-time) are skipped.
  function appendRows(XLSX, P, bytes, rows) {
    const wb = bytes ? load(XLSX, bytes) : XLSX.utils.book_new();
    const added = [];
    const duplicates = [];
    const bySheet = new Map();
    rows.slice().sort((a, b) => a.date - b.date).forEach(r => {
      const name = P.sheetName(r.date);
      if (!bySheet.has(name)) bySheet.set(name, []);
      bySheet.get(name).push(r);
    });

    for (const [name, sheetRows] of bySheet) {
      let ws = wb.Sheets[name];
      if (!ws) {
        ws = XLSX.utils.aoa_to_sheet([P.SHEET_HEADER]);
        ws['!cols'] = [10, 12, 8, 8, 36, 8, 14, 30, 16].map(wch => ({ wch }));
        XLSX.utils.book_append_sheet(wb, ws, name);
        // Keep monthly sheets in chronological order.
        wb.SheetNames.sort();
      }
      const keys = existingKeys(XLSX, ws, P);
      const fresh = [];
      sheetRows.forEach(r => {
        const key = P.rowKey(r.chartNo, r.drugCode, r.date);
        if (keys.has(key)) { duplicates.push(r); return; }
        keys.add(key);
        fresh.push(r);
        added.push({ sheet: name, row: r });
      });
      if (fresh.length) {
        XLSX.utils.sheet_add_aoa(ws, fresh.map(P.sheetRow), { origin: -1, cellDates: true, dateNF: DATE_FORMAT });
      }
    }

    // A "rest of the columns" width rule (e.g. I:XFD) is read back as 16,384
    // separate columns and would bloat the file ~80x; keep only used columns.
    wb.SheetNames.forEach(name => {
      const ws = wb.Sheets[name];
      if (ws['!cols'] && ws['!ref']) {
        ws['!cols'] = ws['!cols'].slice(0, XLSX.utils.decode_range(ws['!ref']).e.c + 1);
      }
    });

    const out = added.length
      ? new Uint8Array(XLSX.write(wb, { type: 'array', bookType: 'xlsx', cellStyles: true }))
      : bytes;
    return { bytes: out, added, duplicates };
  }

  const api = { appendRows };
  if (typeof module !== 'undefined' && module.exports) module.exports = api;
  else root.ADRWorkbook = api;
})(typeof window !== 'undefined' ? window : globalThis);
