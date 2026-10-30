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
