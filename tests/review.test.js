// Run with: npm test
// Case overview: workbook + report forms -> flagged cases. All data is made up.
const test = require('node:test');
const assert = require('node:assert/strict');
const XLSX = require('../lib/xlsx.full.min.js');
const P = require('../parse.js');
const W = require('../workbook.js');
const F = require('../forms.js');
const R = require('../review.js');

const his = (n, code, drug, severity, symptom, date) =>
  `00000000${n}Z\tX\t${code}\t${drug}\t${severity}\t醫療人員提供\t${symptom}\t${date}`;

function workbook(lines) {
  const rows = P.assignIds(P.parseText(lines.join('\n')).rows, 1, null);
  return W.appendRows(XLSX, P, null, rows).bytes;
}

const FORM_HEADER = ['臨床訓練計畫', '學員', '識別代碼\r', '原通報單位識別代碼\r', '通報者獲知日期\r', '不良反應症狀\r',
  '不良反應嚴重性\r', '不良反應後續結果\r', '案件描述\r', '可疑藥品\r', '請將上述問題分數加總\r'];
const form = (id, chartNo, seriousness, outcome, suspect = '', score = '') =>
  ['計畫', '甲', id, chartNo, '2026-10-05', 'Rash', seriousness, outcome, '描述', suspect, score];

const bytes = workbook([
  his(1, 'CEL01', 'Celebrex 膠囊 200 mg', '1_輕微', 'Rash', '2026/10/03 09:00:00'),
  his(2, 'CEL01', 'Celebrex 膠囊 200 mg', '2_中等', 'Pruritus', '2026/10/09 10:00:00'),
  his(3, 'VAN02', 'Vancomycin 針 1 g', '3_嚴重', 'Red man syndrome', '2026/10/12 11:00:00'),
  his(4, 'AMO01', 'Amoxicillin 膠囊', '1_輕微', 'Rash', '2026/11/02 08:00:00'),
  his(4, 'PAR01', 'Panadol 錠', '1_輕微', 'Nausea', '2026/11/04 08:00:00'),
]);

test('reads every workbook case with its month', () => {
  const cases = R.workbookCases(XLSX, bytes);
  assert.deepEqual(cases.map(c => c.id), ['1151001', '1151002', '1151003', '1151101', '1151102']);
  assert.equal(cases[2].severity, '3_嚴重');
  assert.equal(R.monthOf(cases[3]), '11511');
  assert.equal(R.severityLevel('1_輕微、3_嚴重'), 3);
});

test('flags serious, not recovered, positive rechallenge and repeats', () => {
  const rechallenged = '[藥品一]\n1. 英文商品名：Celebrex\n9. 再投予：有再投予且不良反應再發生';
  const forms = F.readCases([FORM_HEADER,
    form('1151001', '000000001Z', '非嚴重', '已恢復/已解決', rechallenged, '9'),
    form('1151002', '000000002Z', '導致病人住院或延長病人住院時間', '尚未恢復/尚未解決'),
    form('1151199', '000000009Z', '危及生命', '恢復中/解決中', '[藥品一]\n1. 英文商品名：Fakedrug'),
  ]);
  const all = R.joinForms(R.workbookCases(XLSX, bytes), forms);
  assert.equal(all.length, 6);
  const formOnly = all.find(c => c.id === '1151199');
  assert.ok(formOnly.formOnly);
  assert.equal(formOnly.drug, 'Fakedrug');

  const oct = R.flagCases(all.filter(c => c.month === '11510'), { starred: new Set(['1151003']) });
  const codes = id => oct.find(c => c.id === id).flags.map(f => f.code);
  assert.deepEqual(codes('1151001'), ['rechallenge', 'naranjo', 'repeatDrug']);
  assert.deepEqual(codes('1151002'), ['serious', 'outcome', 'repeatDrug']);
  // No form: the HIS 嚴重度 decides.
  assert.deepEqual(codes('1151003'), ['serious', 'starred']);
  assert.equal(oct.find(c => c.id === '1151001').flags[2].label, '同藥品多件：CEL01×2');

  const nov = R.flagCases(all.filter(c => c.month === '11511'));
  assert.deepEqual(nov.find(c => c.id === '1151101').flags.map(f => f.label), ['同病人 2 件']);
  assert.equal(nov.find(c => c.id === '1151199').flags[0].label, '危及生命');

  const sum = R.summarize(oct);
  assert.deepEqual([sum.total, sum.notable, sum.serious, sum.noForm], [3, 3, 2, 1]);
  assert.deepEqual(sum.drugs[0], { key: 'CEL01', name: 'Celebrex 膠囊 200 mg', n: 2 });

  const tsv = R.toTsv(oct).split('\r\n');
  assert.equal(tsv.length, 4);
  assert.match(tsv[2], /導致住院或延長住院\t尚未恢復/);
});

test('a quiet case has no flags', () => {
  const [c] = R.flagCases(R.joinForms(R.workbookCases(XLSX, bytes), []).slice(0, 1));
  assert.deepEqual(c.flags, []);
  assert.equal(c.notable, false);
});
