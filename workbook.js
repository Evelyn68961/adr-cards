// Appends parsed rows to the ADR workbook (one sheet per month, e.g. "202610").
// Pure: takes the file bytes and returns new bytes, so it runs in Node tests too.
(function (root) {
  'use strict';

  // Matches the existing sheet's 得知日期 column.
  const DATE_FORMAT = 'm/d/yy h:mm';

  function load(XLSX, bytes) {
    return XLSX.read(bytes, { type: 'array', cellDates: true, cellStyles: true, cellNF: true });
  }

  // Every write goes through here. A "rest of the columns" width rule (e.g.
  // K:XFD) is read back as 16,384 separate columns; written out as such it
  // bloats the file ~80x and can make the page run out of memory and crash
  // (STATUS_BREAKPOINT). Only the columns in use are kept.
  function save(XLSX, wb) {
    wb.SheetNames.forEach(name => {
      const ws = wb.Sheets[name];
      if (ws['!cols'] && ws['!ref']) {
        ws['!cols'] = ws['!cols'].slice(0, XLSX.utils.decode_range(ws['!ref']).e.c + 1);
      }
    });
    return new Uint8Array(XLSX.write(wb, { type: 'array', bookType: 'xlsx', cellStyles: true }));
  }

  // Excel stores whole seconds; SheetJS can return x.999s for them.
  function roundSeconds(date) {
    return new Date(Math.round(date.getTime() / 1000) * 1000);
  }

  // Sheet row index (0-based, as sheet_to_json with header:1 counts them) of
  // every case by rowKey, and the indices of all rows carrying each 識別碼.
  function existingRows(XLSX, ws, P) {
    const byKey = new Map();
    const byId = new Map();
    const aoa = XLSX.utils.sheet_to_json(ws, { header: 1, raw: true, defval: '', blankrows: true });
    aoa.forEach((r, i) => {
      const date = r[8] instanceof Date ? roundSeconds(r[8]) : P.parseDate(r[8]);
      if (date) byKey.set(P.rowKey(r[1], r[3], date), i);
      const id = String(r[0]).trim();
      if (/^\d{7}$/.test(id)) byId.set(id, (byId.get(id) || []).concat(i));
    });
    return { aoa, byKey, byId };
  }

  // Removes one sheet row (absolute index), moving the rows below it up.
  function deleteRow(XLSX, ws, at) {
    const range = XLSX.utils.decode_range(ws['!ref']);
    for (let r = at; r <= range.e.r; r++) {
      for (let c = range.s.c; c <= range.e.c; c++) {
        const to = XLSX.utils.encode_cell({ r, c });
        const from = XLSX.utils.encode_cell({ r: r + 1, c });
        if (r < range.e.r && ws[from]) ws[to] = ws[from];
        else delete ws[to];
      }
    }
    if (ws['!rows']) ws['!rows'].splice(at, 1);
    range.e.r--;
    ws['!ref'] = XLSX.utils.encode_range(range);
  }

  function sameCell(a, b) {
    if (a instanceof Date || b instanceof Date) {
      return a instanceof Date && b instanceof Date && roundSeconds(a).getTime() === roundSeconds(b).getTime();
    }
    return String(a == null ? '' : a).trim() === String(b == null ? '' : b).trim();
  }

  // Who writes up each case is kept in column J (撰寫人), after the HIS columns.
  // Rows are rewritten only up to column I, so the column survives saves.
  const WRITER = '撰寫人';
  const WRITER_COL = 9;
  const writerAt = r => String(r[WRITER_COL] == null ? '' : r[WRITER_COL]).trim();
  // Cases marked as worth noting for the ADR meeting (overview.html): MARKED,
  // empty when not marked.
  const NOTE = '值得注意';
  const NOTE_COL = 10;
  const MARKED = 'V';

  // What the workbook holds: { months: { '202610': { key: '11510', rows, maxSeq } },
  // ids: Map(rowKey -> 識別碼), owners: Map(識別碼 -> its row) }. maxSeq is the
  // highest number in use, counting rows without a 識別碼 so a blank one can
  // never lead to a reused number. A 識別碼 that a second case also carries
  // stays with the first; the second gets a new number on the next save.
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
        const id = String(r[0]).trim();
        const date = r[8] instanceof Date ? roundSeconds(r[8]) : P.parseDate(r[8]);
        const row = { id, chartNo, name: String(r[2]).trim(), drug: P.cleanDrug(r[4]), symptom: P.cleanSymptom(r[7]), date, writer: writerAt(r) };
        rows.push(row);
        if (!id) { blank++; return; }
        if (owners.has(id)) return;
        owners.set(id, row);
        if (id.length === 7 && id.startsWith(key)) maxSeq = Math.max(maxSeq, Number(id.slice(5)));
        if (date) ids.set(P.rowKey(chartNo, r[3], date), id);
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

  // Returns { bytes, added: [{sheet, row}], updated: [{sheet, row}],
  // duplicates: [row], removed: [{sheet, values}] }.
  // A case already in its month's sheet (same chart no + drug code + date-time)
  // is overwritten in place when anything differs; identical ones are left
  // alone (`duplicates`). A new case (not `recorded`) whose 識別碼 is already in
  // the sheet takes that number over: it is written in place of the first row
  // holding it, and any other row holding it is removed. Others are appended.
  // A case given a number by hand (`fixed`) takes it over the same way, also
  // when it is already in the workbook under another number (it keeps its own
  // row, and the row that held the number is removed), and in any month sheet.
  function appendRows(XLSX, P, bytes, rows) {
    const wb = bytes ? load(XLSX, bytes) : XLSX.utils.book_new();
    const added = [];
    const updated = [];
    const duplicates = [];
    const removed = [];
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
      const written = new Set();
      const drop = new Set();
      sheetRows.forEach(r => {
        const key = P.rowKey(r.chartNo, r.drugCode, r.date);
        const holders = (r.recorded && !r.fixed) || !r.id ? [] : (byId.get(r.id) || []);
        const at = byKey.has(key) ? byKey.get(key) : holders[0];
        holders.forEach(i => { if (i !== at) drop.add(i); });
        // A row taken over from another case does not keep that case's 撰寫人 or mark.
        if (!byKey.has(key) && at !== undefined) {
          [WRITER_COL, NOTE_COL].forEach(c => delete ws[XLSX.utils.encode_cell({ r: at + base, c })]);
        }
        const values = P.sheetRow(r);
        if (at === undefined) {
          XLSX.utils.sheet_add_aoa(ws, [values], Object.assign({ origin: -1 }, opts));
          const i = XLSX.utils.decode_range(ws['!ref']).e.r - base;
          aoa[i] = values;
          byKey.set(key, i);
          written.add(i);
          added.push({ sheet: name, row: r });
          return;
        }
        written.add(at);
        // Keep the 識別碼 the sheet already has if the row comes without one.
        if (!values[0]) values[0] = aoa[at][0];
        if (values.every((v, c) => sameCell(v, aoa[at][c]))) { duplicates.push(r); return; }
        XLSX.utils.sheet_add_aoa(ws, [values], Object.assign({ origin: { r: at + base, c: 0 } }, opts));
        aoa[at] = values;
        byKey.set(key, at);
        updated.push({ sheet: name, row: r });
      });
      [...drop].filter(i => !written.has(i)).sort((a, b) => b - a).forEach(i => {
        removed.push({ sheet: name, values: aoa[i] });
        deleteRow(XLSX, ws, i + base);
      });
    }

    // A number typed by hand takes over rows holding it in other months' sheets too.
    rows.filter(r => r.fixed && r.id).forEach(r => {
      const own = P.sheetName(r.date);
      wb.SheetNames.filter(name => /^\d{6}$/.test(name) && name !== own).forEach(name => {
        const ws = wb.Sheets[name];
        const { aoa } = existingRows(XLSX, ws, P);
        const base = ws['!ref'] ? XLSX.utils.decode_range(ws['!ref']).s.r : 0;
        aoa.map((v, i) => i).filter(i => String(aoa[i][0]).trim() === r.id).reverse().forEach(i => {
          removed.push({ sheet: name, values: aoa[i] });
          deleteRow(XLSX, ws, i + base);
        });
      });
    });

    const out = added.length || updated.length || removed.length ? save(XLSX, wb) : bytes;
    return { bytes: out, added, updated, duplicates, removed };
  }

  // Sets one extra column (after the A–I case columns) for cases by 識別碼:
  // values is a Map(識別碼 -> text, '' clears). Labels the column in the
  // header row (the 識別碼 row, or an empty row above the data) when it is
  // not labelled yet. Returns { bytes, changed: [識別碼], missing: [識別碼] }.
  function setColumn(XLSX, bytes, col, label, writers) {
    const cellAt = r => String(r[col] == null ? '' : r[col]).trim();
    const wb = load(XLSX, bytes);
    const changed = [];
    const left = new Map(writers);
    wb.SheetNames.filter(name => /^\d{6}$/.test(name)).forEach(name => {
      const ws = wb.Sheets[name];
      if (!ws['!ref']) return;
      const base = XLSX.utils.decode_range(ws['!ref']).s.r;
      const aoa = XLSX.utils.sheet_to_json(ws, { header: 1, raw: true, defval: '', blankrows: true });
      let touched = false;
      aoa.forEach((r, i) => {
        const id = String(r[0]).trim();
        if (!left.has(id)) return;
        const value = String(left.get(id)).trim();
        left.delete(id);
        if (cellAt(r) === value) return;
        const addr = XLSX.utils.encode_cell({ r: base + i, c: col });
        if (value) XLSX.utils.sheet_add_aoa(ws, [[value]], { origin: addr });
        else delete ws[addr];
        changed.push(id);
        touched = true;
      });
      if (!touched) return;
      const head = aoa.findIndex(r => String(r[0]).trim() === '識別碼');
      const firstData = aoa.findIndex(r => /^\d{9}[A-Z]$/i.test(String(r[1]).trim()));
      const emptyAbove = firstData === 0 || (firstData > 0 && aoa[firstData - 1].every(v => v === ''));
      const headRow = head >= 0 ? base + head : emptyAbove ? base + firstData - 1 : -1;
      const headAddr = XLSX.utils.encode_cell({ r: headRow, c: col });
      if (headRow >= 0 && !ws[headAddr]) {
        XLSX.utils.sheet_add_aoa(ws, [[label]], { origin: headAddr });
        if (ws['!cols'] && !ws['!cols'][col]) ws['!cols'][col] = { wch: 10 };
      }
    });
    const out = changed.length ? save(XLSX, wb) : bytes;
    return { bytes: out, changed, missing: [...left.keys()] };
  }

  // 撰寫人 (column J): Map(識別碼 -> name).
  const setWriters = (XLSX, bytes, writers) => setColumn(XLSX, bytes, WRITER_COL, WRITER, writers);
  // 值得注意 (column K): Map(識別碼 -> MARKED, or '' to unmark).
  const setNotes = (XLSX, bytes, notes) => setColumn(XLSX, bytes, NOTE_COL, NOTE, notes);

  const api = { appendRows, readLedger, setWriters, setNotes, NOTE_COL, MARKED };
  if (typeof module !== 'undefined' && module.exports) module.exports = api;
  else root.ADRWorkbook = api;
})(typeof window !== 'undefined' ? window : globalThis);
