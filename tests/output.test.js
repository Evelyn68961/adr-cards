// Run with: npm test
// PDF structure and workbook append, without a browser. All data is made up.
const test = require('node:test');
const assert = require('node:assert/strict');
const XLSX = require('../lib/xlsx.full.min.js');
const P = require('../parse.js');
const W = require('../workbook.js');
const PDF = require('../pdf.js');

test('PDF xref offsets point at their objects', () => {
  const fakeJpeg = new Uint8Array([0xff, 0xd8, 0xff, 0xd9]);
  const bytes = PDF.build([{ jpeg: fakeJpeg, width: 10, height: 7 }, { jpeg: fakeJpeg, width: 10, height: 7 }]);
  const text = Buffer.from(bytes).toString('latin1');
  assert.match(text, /^%PDF-1\.4/);
  assert.match(text, /\/Count 2/);
  const xrefAt = Number(text.match(/startxref\n(\d+)/)[1]);
  assert.ok(text.startsWith('xref', xrefAt));
  const entries = text.slice(xrefAt).match(/^\d{10} 00000 n $/gm);
  assert.equal(entries.length, 8);  // catalog, pages, 3 per page
  entries.forEach((e, i) => assert.ok(text.startsWith(`${i + 1} 0 obj`, Number(e.slice(0, 10)))));
});

const row = (n, d) => `00000000${n}Z\tX\tCEL01\tCelebrex 膠囊 200 mg\t2_中等\t病人/家屬自述\tRash\t${d}`;

test('appendRows creates monthly sheets, appends, and skips duplicates', () => {
  const first = P.parseText([row(1, '2026/10/17 10:52:08'), row(2, '2026/11/02 09:00:00')].join('\n')).rows;
  const r1 = W.appendRows(XLSX, P, null, first);
  assert.equal(r1.added.length, 2);
  const wb1 = XLSX.read(r1.bytes, { type: 'array' });
  assert.deepEqual(wb1.SheetNames, ['202610', '202611']);

  const second = P.parseText([row(1, '2026/10/17 10:52:08'), row(3, '2026/10/20 14:26:51')].join('\n')).rows;
  const r2 = W.appendRows(XLSX, P, r1.bytes, second);
  assert.deepEqual(r2.added.map(a => a.row.chartNo), ['000000003Z']);
  assert.deepEqual(r2.duplicates.map(d => d.chartNo), ['000000001Z']);
  const sheet = XLSX.utils.sheet_to_json(XLSX.read(r2.bytes, { type: 'array' }).Sheets['202610'], { header: 1 });
  assert.deepEqual(sheet.map(r => r[1]), ['病歷號', '000000001Z', '000000003Z']);
  assert.equal(sheet[0][0], '識別碼');
});

test('appendRows with nothing new returns the original bytes', () => {
  const rows = P.parseText(row(1, '2026/10/17 10:52:08')).rows;
  const { bytes } = W.appendRows(XLSX, P, null, rows);
  const again = W.appendRows(XLSX, P, bytes, rows);
  assert.equal(again.added.length, 0);
  assert.equal(again.bytes, bytes);
});

test('appendRows overwrites a changed case in place, by key or by 識別碼', () => {
  const cards = P.assignIds(P.parseText([row(1, '2026/10/17 10:52:08'), row(2, '2026/10/20 14:26:51')].join('\n')).rows, 1);
  const { bytes } = W.appendRows(XLSX, P, null, cards);
  // Same case, symptom fixed by hand.
  const fixed = Object.assign({}, cards[0], { symptomRaw: 'Rash、Itching' });
  // A corrected HIS row (other drug code) reusing 1151002.
  const redo = Object.assign({}, cards[1], { drugCode: 'PRO23', drugRaw: 'PromERAN 針' });
  const r = W.appendRows(XLSX, P, bytes, [fixed, redo]);
  assert.deepEqual(r.updated.map(u => u.row.id), ['1151001', '1151002']);
  assert.equal(r.added.length, 0);
  const sheet = XLSX.utils.sheet_to_json(XLSX.read(r.bytes, { type: 'array' }).Sheets['202610'], { header: 1 });
  assert.equal(sheet.length, 3);
  assert.equal(sheet[1][7], 'Rash、Itching');
  assert.deepEqual([sheet[2][0], sheet[2][3], sheet[2][4]], ['1151002', 'PRO23', 'PromERAN 針']);
});

// A workbook whose 202611 sheet has two cases both numbered 1151101.
function clashedBook() {
  const cards = P.assignIds(P.parseText([row(1, '2026/11/02 11:50:00'), row(2, '2026/11/03 14:57:00')].join('\n')).rows, 1);
  const wb = XLSX.read(W.appendRows(XLSX, P, null, cards).bytes, { type: 'array', cellDates: true });
  wb.Sheets['202611'].A3 = { t: 's', v: '1151101' };
  return new Uint8Array(XLSX.write(wb, { type: 'array', bookType: 'xlsx' }));
}
const viewOf = ledger => ({
  idFor: r => ledger.ids.get(P.rowKey(r.chartNo, r.drugCode, r.date)) || null,
  nextSeq: () => ledger.months['202611'].maxSeq + 1,
});
const sheet11 = bytes => XLSX.utils.sheet_to_json(XLSX.read(bytes, { type: 'array' }).Sheets['202611'], { header: 1 })
  .slice(1).map(s => [s[0], s[1]]);

test('a typed 識別碼 already in Excel replaces every row holding it', () => {
  const bytes = clashedBook();
  const ledger = W.readLedger(XLSX, P, bytes);
  // The second case is pasted again with 起始流水號 typed as 1.
  const cards = P.assignIds(P.parseText(row(2, '2026/11/03 14:57:00')).rows, 1, viewOf(ledger));
  assert.deepEqual(cards.map(c => [c.id, !!c.recorded]), [['1151101', false]]);
  const r = W.appendRows(XLSX, P, bytes, cards);
  assert.deepEqual(r.removed.map(d => d.values[1]), ['000000001Z']);
  assert.deepEqual(sheet11(r.bytes), [['1151101', '000000002Z']]);
});

test('a 識別碼 carried by two cases is renumbered on the next plain save', () => {
  const bytes = clashedBook();
  const ledger = W.readLedger(XLSX, P, bytes);
  assert.equal(ledger.months['202611'].maxSeq, 1);
  const again = P.assignIds(P.parseText([row(1, '2026/11/02 11:50:00'), row(2, '2026/11/03 14:57:00')].join('\n')).rows, null, viewOf(ledger));
  assert.deepEqual(again.map(c => c.id), ['1151101', '1151102']);
  const r = W.appendRows(XLSX, P, bytes, again);
  assert.equal(r.removed.length, 0);
  assert.deepEqual(sheet11(r.bytes), [['1151101', '000000001Z'], ['1151102', '000000002Z']]);
});

test('setWriters adds a 撰寫人 column that survives later saves', () => {
  const cards = P.assignIds(P.parseText([row(1, '2026/10/17 10:52:08'), row(2, '2026/10/20 14:26:51')].join('\n')).rows, 1);
  const { bytes } = W.appendRows(XLSX, P, null, cards);
  const set = W.setWriters(XLSX, bytes, new Map([['1151002', '王藥師'], ['1159999', 'X']]));
  assert.deepEqual(set.changed, ['1151002']);
  assert.deepEqual(set.missing, ['1159999']);
  const ledger = W.readLedger(XLSX, P, set.bytes);
  assert.deepEqual(ledger.months['202610'].rows.map(r => r.writer), ['', '王藥師']);

  // Overwriting the case with a hand fix keeps its 撰寫人.
  const fixed = Object.assign({}, cards[1], { symptomRaw: 'Rash、Itching' });
  const after = W.appendRows(XLSX, P, set.bytes, [fixed]);
  assert.equal(after.updated.length, 1);
  const sheet = XLSX.utils.sheet_to_json(XLSX.read(after.bytes, { type: 'array' }).Sheets['202610'], { header: 1 });
  assert.equal(sheet[0][9], '撰寫人');
  assert.deepEqual([sheet[2][7], sheet[2][9]], ['Rash、Itching', '王藥師']);

  // Clearing removes the name; an unchanged name writes nothing.
  assert.equal(W.setWriters(XLSX, after.bytes, new Map([['1151002', '王藥師']])).changed.length, 0);
  const cleared = W.setWriters(XLSX, after.bytes, new Map([['1151002', '']]));
  assert.deepEqual(W.readLedger(XLSX, P, cleared.bytes).months['202610'].rows.map(r => r.writer), ['', '']);
});

test('setWriters on a sheet without a header row labels the empty row above', () => {
  const cards = P.assignIds(P.parseText(row(1, '2026/10/05 09:57:01')).rows, 1);
  const ws = XLSX.utils.aoa_to_sheet([P.sheetRow(cards[0])], { origin: 'A2', cellDates: true });
  const wb = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(wb, ws, '202610');
  const bytes = new Uint8Array(XLSX.write(wb, { type: 'array', bookType: 'xlsx' }));
  const out = W.setWriters(XLSX, bytes, new Map([['1151001', '王藥師']])).bytes;
  const sheet = XLSX.read(out, { type: 'array' }).Sheets['202610'];
  assert.deepEqual([sheet.J1.v, sheet.A2.v, sheet.J2.v], ['撰寫人', '1151001', '王藥師']);
  // The rest of the empty title row gets its titles too.
  assert.deepEqual([sheet.A1.v, sheet.I1.v], ['識別碼', '得知日期']);
  assert.equal(W.readLedger(XLSX, P, out).months['202610'].rows[0].writer, '王藥師');
});

test('a case that takes over a 識別碼 does not inherit its 撰寫人', () => {
  const [a] = P.assignIds(P.parseText(row(1, '2026/11/02 11:50:00')).rows, 1);
  const named = W.setWriters(XLSX, W.appendRows(XLSX, P, null, [a]).bytes, new Map([['1151101', '王藥師']])).bytes;
  const [b] = P.assignIds(P.parseText(row(2, '2026/11/03 14:57:00')).rows, 1);
  const r = W.appendRows(XLSX, P, named, [b]);
  const rows = W.readLedger(XLSX, P, r.bytes).months['202611'].rows;
  assert.deepEqual(rows.map(x => [x.chartNo, x.writer]), [['000000002Z', '']]);
});

test('appendRows does not expand a whole-row column width rule', () => {
  const ws = XLSX.utils.aoa_to_sheet([P.SHEET_HEADER]);
  ws['!cols'] = Array.from({ length: 16384 }, () => ({ wch: 9 }));
  const wb = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(wb, ws, '202610');
  const bytes = new Uint8Array(XLSX.write(wb, { type: 'array', bookType: 'xlsx', cellStyles: true }));
  const cards = P.assignIds(P.parseText(row(1, '2026/10/17 10:52:08')).rows, 11);
  const { bytes: out } = W.appendRows(XLSX, P, bytes, cards);
  const sheet = XLSX.read(out, { type: 'array', cellStyles: true }).Sheets['202610'];
  assert.equal(sheet.A2.v, '1151011');
  const cols = sheet['!cols'];
  assert.ok(cols.length <= 9, `kept ${cols.length} column entries`);
  assert.ok(out.length < 20000, `file is ${out.length} bytes`);
});

test('readLedger reports numbers per month and maps cases to their 識別碼', () => {
  const cards = P.assignIds(P.parseText([row(1, '2026/10/17 10:52:08'), row(2, '2026/10/20 14:26:51')].join('\n')).rows, 11);
  const { bytes } = W.appendRows(XLSX, P, null, cards);
  const ledger = W.readLedger(XLSX, P, bytes);
  const month = ledger.months['202610'];
  assert.equal(month.key, '11510');
  assert.equal(month.rows.length, 2);
  assert.equal(month.maxSeq, 12);
  assert.equal(month.rows[1].drug, 'Celebrex 膠囊 200 mg');
  assert.equal(ledger.ids.get(P.rowKey('000000001Z', 'CEL01', cards[0].date)), '1151011');
});

test('assignIds keeps recorded numbers and continues from the ledger', () => {
  const rows = P.parseText([row(1, '2026/10/17 10:52:08'), row(2, '2026/10/20 14:26:51'), row(3, '2026/11/02 09:00:00')].join('\n')).rows;
  const ledger = {
    idFor: r => (r.chartNo === '000000001Z' ? '1151006' : null),
    nextSeq: key => (key === '11510' ? 11 : 1),
  };
  const cards = P.assignIds(rows, null, ledger);
  assert.deepEqual(cards.map(c => c.id), ['1151006', '1151011', '1151101']);
  assert.deepEqual(cards.map(c => !!c.recorded), [true, false, false]);
  // A typed start number applies to the first month with new cases.
  assert.deepEqual(P.assignIds(rows, 20, ledger).map(c => c.id), ['1151006', '1151020', '1151101']);
});

test('a 識別碼 typed on a card replaces an unreportable case, also for a case already in Excel', () => {
  // 1151101 (case 1) turned out unreportable; case 2 is 1151102.
  const cards = P.assignIds(P.parseText([row(1, '2026/11/02 11:50:00'), row(2, '2026/11/03 14:57:00')].join('\n')).rows, 1);
  let bytes = W.appendRows(XLSX, P, null, cards).bytes;
  bytes = W.setWriters(XLSX, bytes, new Map([['1151102', '王藥師']])).bytes;
  const ledger = W.readLedger(XLSX, P, bytes);
  const pasted = P.parseText([row(2, '2026/11/03 14:57:00'), row(3, '2026/11/04 09:00:00')].join('\n')).rows;
  // Case 2 (in Excel as 1151102) is given 1151101 by hand; new case 3 is numbered automatically.
  const fixed = new Map([[P.rowKey('000000002Z', 'CEL01', pasted[0].date), '1151101']]);
  const out = P.assignIds(pasted, null, viewOf(ledger), fixed);
  assert.deepEqual(out.map(c => [c.id, !!c.recorded, !!c.fixed, c.previousId || null]),
    [['1151101', true, true, '1151102'], ['1151103', false, false, null]]);
  const r = W.appendRows(XLSX, P, bytes, out);
  assert.deepEqual(r.removed.map(d => d.values[1]), ['000000001Z']);
  assert.deepEqual(sheet11(r.bytes), [['1151101', '000000002Z'], ['1151103', '000000003Z']]);
  // The moved case keeps its 撰寫人.
  const kept = XLSX.utils.sheet_to_json(XLSX.read(r.bytes, { type: 'array' }).Sheets['202611'], { header: 1 });
  assert.equal(kept[1][9], '王藥師');
});

test('a typed 識別碼 skips automatic numbers and replaces a case in another month', () => {
  // A new case is free to take 1151002 (unreportable, in October) and the
  // October case pasted with it continues numbering without reusing 1151002.
  const oct = P.assignIds(P.parseText([row(1, '2026/10/02 11:50:00'), row(2, '2026/10/03 14:57:00')].join('\n')).rows, 1);
  const bytes = W.appendRows(XLSX, P, null, oct).bytes;
  const ledger = W.readLedger(XLSX, P, bytes);
  const view = { idFor: r => ledger.ids.get(P.rowKey(r.chartNo, r.drugCode, r.date)) || null, nextSeq: k => (k === '11510' ? 3 : 1) };
  const pasted = P.parseText([row(3, '2026/10/05 08:00:00'), row(4, '2026/11/01 08:00:00')].join('\n')).rows;
  const fixed = new Map([[P.rowKey('000000004Z', 'CEL01', pasted[1].date), '1151002']]);
  const out = P.assignIds(pasted, 3, view, fixed);
  assert.deepEqual(out.map(c => c.id), ['1151003', '1151002']);
  const r = W.appendRows(XLSX, P, bytes, out);
  const book = XLSX.read(r.bytes, { type: 'array' });
  const ids = name => XLSX.utils.sheet_to_json(book.Sheets[name], { header: 1 }).slice(1).map(s => s[0]);
  assert.deepEqual(ids('202610'), ['1151001', '1151003']);
  assert.deepEqual(ids('202611'), ['1151002']);
});

test('setWriters and setNotes do not expand a whole-row column width rule', () => {
  // Excel's "rest of the columns" width rule (e.g. K:XFD) is read back as
  // 16,384 columns; writing them all made the page run out of memory.
  const cards = P.assignIds(P.parseText(row(1, '2026/10/17 10:52:08')).rows, 1);
  const wb = XLSX.read(W.appendRows(XLSX, P, null, cards).bytes, { type: 'array', cellStyles: true });
  wb.Sheets['202610']['!cols'] = Array.from({ length: 16384 }, () => ({ wch: 9 }));
  const other = XLSX.utils.aoa_to_sheet([['notes']]);
  other['!cols'] = Array.from({ length: 16384 }, () => ({ wch: 9 }));
  XLSX.utils.book_append_sheet(wb, other, '說明');
  const bytes = new Uint8Array(XLSX.write(wb, { type: 'array', bookType: 'xlsx', cellStyles: true }));
  for (const out of [W.setWriters(XLSX, bytes, new Map([['1151001', '王藥師']])).bytes,
    W.setNotes(XLSX, bytes, new Map([['1151001', W.MARKED]])).bytes]) {
    const book = XLSX.read(out, { type: 'array', cellStyles: true });
    assert.ok(book.Sheets['202610']['!cols'].length <= 11, `kept ${book.Sheets['202610']['!cols'].length} column entries`);
    assert.ok(out.length < 30000, `file is ${out.length} bytes`);
  }
  assert.equal(XLSX.read(W.setWriters(XLSX, bytes, new Map([['1151001', '王藥師']])).bytes, { type: 'array' }).Sheets['202610'].J2.v, '王藥師');
});

test('a case saved before under a stale 登記日 is moved to its real month', () => {
  const OLD = '000000777J\t王小美\tCOD02\tCodeine 針\tR05DA04\t480800\t鎮咳劑\t0_不明\t病人/家屬自述\t0_藥物不良反應\t2016/09/21\tUrticaria\t\t\tB00000\t陳藥師\t2026/11/17 21:31:20';
  const [r] = P.parseText(OLD).rows;
  // What the old version saved: dated 2016/09/21, numbered 1050901, on sheet 201609.
  const before = Object.assign({}, r, { date: r.oldDate, id: '1050901' });
  let bytes = W.appendRows(XLSX, P, null, P.assignIds(P.parseText(row(1, '2026/11/02 11:50:00')).rows, 1)).bytes;
  bytes = W.appendRows(XLSX, P, bytes, [before]).bytes;
  assert.ok(XLSX.read(bytes, { type: 'array' }).SheetNames.includes('201609'));
  const ledger = W.readLedger(XLSX, P, bytes);
  const [card] = P.assignIds([r], 10, viewOf(ledger));
  card.staleId = ledger.ids.get(P.rowKey(r.chartNo, r.drugCode, r.oldDate));
  assert.deepEqual([card.id, card.staleId], ['1151110', '1050901']);
  const out = W.appendRows(XLSX, P, bytes, [card]);
  assert.deepEqual(out.removedSheets, ['201609']);
  assert.deepEqual(out.removed.map(d => d.values[0]), ['1050901']);
  const book = XLSX.read(out.bytes, { type: 'array' });
  assert.deepEqual(book.SheetNames, ['202611']);
  assert.deepEqual(sheet11(out.bytes), [['1151101', '000000001Z'], ['1151110', '000000777J']]);
});

test('a sheet whose used area Excel stretched to row 1,048,576 saves quickly', () => {
  // Formatting on whole rows makes Excel report e.g. A1:N1048576; walking
  // that area cell by cell used ~500 MB per save and crashed the page.
  const cards = P.assignIds(P.parseText(row(1, '2026/10/17 10:52:08')).rows, 1);
  const wb = XLSX.read(W.appendRows(XLSX, P, null, cards).bytes, { type: 'array' });
  const old = XLSX.utils.aoa_to_sheet([P.SHEET_HEADER, ['1150101', '000000009Z', 'Y', 'A', 'B', '1', 's', 'x', '1/1/26 9:00']]);
  old['!ref'] = 'A1:N1048576';
  XLSX.utils.book_append_sheet(wb, old, '202601');
  const bytes = new Uint8Array(XLSX.write(wb, { type: 'array', bookType: 'xlsx' }));
  const start = Date.now();
  const out = W.setTfdaDates(XLSX, bytes, new Map([['1150101', '2026/11/27']]));
  assert.deepEqual(out.changed, ['1150101']);
  assert.ok(Date.now() - start < 2000, `took ${Date.now() - start} ms`);
  assert.equal(XLSX.read(out.bytes, { type: 'array' }).Sheets['202601']['!ref'], 'A1:N2');
  assert.ok(W.readLedger(XLSX, P, bytes).months['202601']);
});

// A workbook whose theme names East Asian fonts, as Excel writes it.
function withTheme(bytes, fonts) {
  const cfb = XLSX.CFB.read(bytes, { type: 'array' });
  const entry = XLSX.CFB.find(cfb, '/xl/theme/theme1.xml');
  entry.content = Buffer.from(Buffer.from(entry.content).toString('utf8')
    .replace(/<a:latin typeface="[^"]*"\/>/, m => m + fonts), 'utf8');
  return new Uint8Array(XLSX.CFB.write(cfb, { type: 'array', fileType: 'zip' }));
}
const themeOf = bytes => Buffer.from(XLSX.CFB.find(XLSX.CFB.read(bytes, { type: 'array' }), '/xl/theme/theme1.xml').content);

test('saving keeps a theme with East Asian font names the same size, and undoes earlier garbling', () => {
  const cards = P.assignIds(P.parseText(row(1, '2026/10/17 10:52:08')).rows, 1);
  const fonts = '<a:font script="Jpan" typeface="ＭＳ Ｐゴシック"/><a:font script="Hant" typeface="新細明體"/>';
  let bytes = withTheme(W.appendRows(XLSX, P, null, cards).bytes, fonts);
  const size = themeOf(bytes).length;
  // SheetJS used to UTF-8-encode the theme's raw bytes again on every save,
  // doubling each non-ASCII byte: 100+ MB after ~20 saves.
  for (let i = 0; i < 6; i++) bytes = W.setTfdaDates(XLSX, bytes, new Map([['1151001', `2026/11/1${i}`]])).bytes;
  assert.equal(themeOf(bytes).length, size);
  assert.ok(themeOf(bytes).toString('utf8').includes('typeface="新細明體"'));
  // A theme already garbled three times by the old saves comes back whole.
  let garbled = fonts;
  for (let i = 0; i < 3; i++) garbled = Buffer.from(garbled, 'utf8').toString('latin1');
  const fixed = W.setWriters(XLSX, withTheme(W.appendRows(XLSX, P, null, cards).bytes, garbled), new Map([['1151001', '王']])).bytes;
  assert.ok(themeOf(fixed).toString('utf8').includes('typeface="新細明體"'));
  assert.equal(themeOf(fixed).length, size);
});

test('a workbook with every known trap saved 10 times keeps every part the same size', () => {
  // East Asian theme fonts (theme garbling), widths for all 16,384 columns
  // (cols bloat) and an area stretched to row 1,048,576 (slow range walks).
  const cards = P.assignIds(P.parseText([row(1, '2026/10/17 10:52:08'), row(2, '2026/11/02 09:00:00')].join('\n')).rows, 1);
  const wb = XLSX.read(W.appendRows(XLSX, P, null, cards).bytes, { type: 'array', cellStyles: true });
  wb.Sheets['202610']['!cols'] = Array.from({ length: 16384 }, () => ({ wch: 9 }));
  wb.Sheets['202611']['!ref'] = 'A1:N200000';
  let bytes = withTheme(new Uint8Array(XLSX.write(wb, { type: 'array', bookType: 'xlsx', cellStyles: true })),
    '<a:font script="Jpan" typeface="ＭＳ Ｐゴシック"/><a:font script="Hant" typeface="新細明體"/>');
  const parts = b => Object.fromEntries(XLSX.CFB.read(b, { type: 'array' }).FullPaths
    .map((p, i, all) => [p.replace(/^Root Entry/, ''), XLSX.CFB.read(b, { type: 'array' }).FileIndex[i].size]));
  // The first saves clean up and add the cells the later saves only change.
  bytes = W.setWriters(XLSX, bytes, new Map([['1151001', '王藥師0']])).bytes;
  bytes = W.setTfdaDates(XLSX, bytes, new Map([['1151101', '2026/11/10']])).bytes;
  const first = parts(bytes);
  for (let i = 0; i < 10; i++) {
    bytes = (i % 2 ? W.setWriters(XLSX, bytes, new Map([['1151001', '王藥師' + i]]))
      : W.setTfdaDates(XLSX, bytes, new Map([['1151101', `2026/11/1${i}`]]))).bytes;
  }
  const last = parts(bytes);
  for (const [name, size] of Object.entries(first)) {
    if (/theme|styles|worksheets/.test(name)) assert.ok(Math.abs(last[name] - size) <= 4, `${name}: ${size} → ${last[name]}`);
  }
  assert.ok(bytes.length < 40000, `file is ${bytes.length} bytes`);
  assert.ok(themeOf(bytes).toString('utf8').includes('typeface="ＭＳ Ｐゴシック"'));
});

test('blank titles in a month sheet are filled in on save, existing ones kept', () => {
  // Like the user's 202610: row 1 only had J1 = 撰寫人.
  const cards = P.assignIds(P.parseText(row(1, '2026/10/17 10:52:08')).rows, 1);
  const wb = XLSX.read(W.appendRows(XLSX, P, null, cards).bytes, { type: 'array' });
  const ws = wb.Sheets['202610'];
  for (const c of 'ABCDEFGHI') delete ws[c + '1'];
  ws.J1 = { t: 's', v: '撰寫人' };
  ws.F1 = { t: 's', v: '我的嚴重度' };   // a title the user typed stays
  ws['!ref'] = 'A1:J2';
  const bytes = new Uint8Array(XLSX.write(wb, { type: 'array', bookType: 'xlsx' }));
  const out = W.setTfdaDates(XLSX, bytes, new Map([['1151001', '2026/11/02']])).bytes;
  const head = XLSX.utils.sheet_to_json(XLSX.read(out, { type: 'array' }).Sheets['202610'], { header: 1, defval: '' })[0];
  assert.deepEqual(head.slice(0, 9), ['識別碼', '病歷號', '姓名', '藥碼', '藥名', '我的嚴重度', '通報來源', '症狀', '得知日期']);
  assert.equal(head[9], '撰寫人');
  assert.equal(head[13], 'TFDA通報日期');
  assert.equal(head[10] || '', '');   // 值得注意 unused: no title added
});
