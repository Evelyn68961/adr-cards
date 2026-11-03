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

  // Sheet row index (0-based, as sheet_to_json with header:1 counts them) of
  // every case, by rowKey and by 識別碼 (first row carrying it).
  function existingRows(XLSX, ws, P) {
    const byKey = new Map();
    const byId = new Map();
    const aoa = XLSX.utils.sheet_to_json(ws, { header: 1, raw: true, defval: '', blankrows: true });
    aoa.forEach((r, i) => {
      const date = r[8] instanceof Date ? roundSeconds(r[8]) : P.parseDate(r[8]);
      if (date) byKey.set(P.rowKey(r[1], r[3], date), i);
      const id = String(r[0]).trim();
      if (/^\d{7}$/.test(id) && !byId.has(id)) byId.set(id, i);
    });
    return { aoa, byKey, byId };
  }

  function sameCell(a, b) {
    if (a instanceof Date || b instanceof Date) {
      return a instanceof Date && b instanceof Date && roundSeconds(a).getTime() === roundSeconds(b).getTime();
    }
    return String(a == null ? '' : a).trim() === String(b == null ? '' : b).trim();
  }

  // What the workbook holds: { months: { '202610': { key: '11510', rows, maxSeq } },
  // ids: Map(rowKey -> 識別碼), owners: Map(識別碼 -> 病歷號) }. maxSeq is the
  // highest number in use, counting rows without a 識別碼 so a blank one can
  // never lead to a reused number. A 識別碼 that a second case also carries is
  // left off that case (it gets a new number on the next save).
  function readLedger(XLSX, P, bytes) {
    const wb = XLSX.read(bytes, { type: 'array', cellDates: true });
    const months = {};
    const ids = new Map();
    const owners = new Map();
    wb.SheetNames.filter(name => /^\d{6}$/.test(name)).forEach(name => {
      const key = String(Number(name.slice(0, 4)) - 1911) + name.slice(4);
      const rows = [];
      let maxSeq = 0;
      let blank = 0;
      XLSX.utils.sheet_to_json(wb.Sheets[name], { header: 1, raw: true, defval: '' }).forEach(r => {
        const chartNo = String(r[1]).trim().toUpperCase();
        if (!/^\d{9}[A-Z]$/.test(chartNo)) return;
        let id = String(r[0]).trim();
        const date = r[8] instanceof Date ? roundSeconds(r[8]) : P.parseDate(r[8]);
        if (owners.has(id)) id = '';
        else if (id) owners.set(id, chartNo);
        else blank++;
        if (id.length === 7 && id.startsWith(key)) maxSeq = Math.max(maxSeq, Number(id.slice(5)));
        if (id && date) ids.set(P.rowKey(chartNo, r[3], date), id);
        rows.push({ id, chartNo, drug: P.cleanDrug(r[4]), symptom: P.cleanSymptom(r[7]), date });
      });
      months[name] = { key, rows, maxSeq: Math.max(maxSeq, countIn(owners, key) + blank) };
    });
    return { months, ids, owners };
  }

  function countIn(owners, key) {
    let n = 0;
    owners.forEach((_, id) => { if (id.startsWith(key)) n++; });
    return n;
  }

  // Returns { bytes, added: [{sheet, row}], updated: [{sheet, row}], duplicates: [row] }.
  // A case already in its month's sheet (same chart no + drug code + date-time,
  // or else the same 識別碼 on the same patient) is overwritten in place when
  // anything differs; identical ones are left alone (`duplicates`). Others are
  // appended; a 識別碼 is never taken over from another patient's row.
  function appendRows(XLSX, P, bytes, rows) {
    const wb = bytes ? load(XLSX, bytes) : XLSX.utils.book_new();
    const added = [];
    const updated = [];
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
      const { aoa, byKey, byId } = existingRows(XLSX, ws, P);
      // aoa[0] is the sheet's first used row, which need not be row 1.
      const base = ws['!ref'] ? XLSX.utils.decode_range(ws['!ref']).s.r : 0;
      const opts = { cellDates: true, dateNF: DATE_FORMAT };
      sheetRows.forEach(r => {
        const key = P.rowKey(r.chartNo, r.drugCode, r.date);
        const sameId = byId.get(r.id);
        const samePatient = sameId !== undefined &&
          String(aoa[sameId][1]).trim().toUpperCase() === String(r.chartNo).trim().toUpperCase();
        const at = byKey.has(key) ? byKey.get(key) : samePatient ? sameId : undefined;
        const values = P.sheetRow(r);
        if (at === undefined) {
          XLSX.utils.sheet_add_aoa(ws, [values], Object.assign({ origin: -1 }, opts));
          const i = XLSX.utils.decode_range(ws['!ref']).e.r - base;
          aoa[i] = values;
          byKey.set(key, i);
          if (r.id && !byId.has(r.id)) byId.set(r.id, i);
          added.push({ sheet: name, row: r });
          return;
        }
        // Keep the 識別碼 the sheet already has if the row comes without one.
        if (!values[0]) values[0] = aoa[at][0];
        if (values.every((v, c) => sameCell(v, aoa[at][c]))) { duplicates.push(r); return; }
        XLSX.utils.sheet_add_aoa(ws, [values], Object.assign({ origin: { r: at + base, c: 0 } }, opts));
        aoa[at] = values;
        byKey.set(key, at);
        updated.push({ sheet: name, row: r });
      });
    }

    // A "rest of the columns" width rule (e.g. I:XFD) is read back as 16,384
    // separate columns and would bloat the file ~80x; keep only used columns.
    wb.SheetNames.forEach(name => {
      const ws = wb.Sheets[name];
      if (ws['!cols'] && ws['!ref']) {
        ws['!cols'] = ws['!cols'].slice(0, XLSX.utils.decode_range(ws['!ref']).e.c + 1);
      }
    });

    const out = added.length || updated.length
      ? new Uint8Array(XLSX.write(wb, { type: 'array', bookType: 'xlsx', cellStyles: true }))
      : bytes;
    return { bytes: out, added, updated, duplicates };
  }

  const api = { appendRows, readLedger };
  if (typeof module !== 'undefined' && module.exports) module.exports = api;
  else root.ADRWorkbook = api;
})(typeof window !== 'undefined' ? window : globalThis);
