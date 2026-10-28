// Run with: npm test
// All chart numbers and names below are made up.
const test = require('node:test');
const assert = require('node:assert/strict');
const P = require('../parse.js');

const SAMPLE = [
  '病歷號\t姓名\t藥碼\t藥名\t嚴重度\t通報來源\t症狀\t得知日期',
  '000000001A\t王小明\tMOP03\tMopride 錠劑 5 mg\t2_中等\t醫療人員提供\tLiver dysfunction\t10/5/26 9:57',
  '000000002B\t李小華\tPRO23\tPromERAN 針 9.08 mg/2 mL/Amp\t1_輕微\t病人/家屬自述\tEPS\t10/9/26 13:28',
  '000000003C\t張小美\tCOM05\t【RT給】Combivent 吸入 2.5 mL\t2_中等\t病人/家屬自述\tface swelling\t10/10/26 15:28',
  '000000004D\t陳大文\tVAN02\t★Vancomycin SANDOZ 針 1 g\t2_中等\t醫療人員提供\tanaphylaxis shock\t10/14/26 18:59',
  '000000005E\t林小玉\tCEF10\tCefin 針【禁用含鈣溶液】2 g/Vial\t2_中等\t醫療人員提供\tRash,Swelling of the skin,Eye swelling,Mouth swelling\t10/14/26 10:12',
].join('\n');

test('parseDate handles Excel, Gregorian and ROC formats', () => {
  const expect = new Date(2026, 9, 5, 9, 57).getTime();
  assert.equal(P.parseDate('10/5/26 9:57').getTime(), expect);
  assert.equal(P.parseDate('2026/10/5 09:57').getTime(), expect);
  assert.equal(P.parseDate('2026-10-05 09:57:00').getTime(), expect);
  assert.equal(P.parseDate('115/10/05 09:57').getTime(), expect);
  assert.equal(P.parseDate('2026/10/5').getTime(), new Date(2026, 9, 5).getTime());
  assert.equal(P.parseDate('2026/13/40'), null);
  assert.equal(P.parseDate('得知日期'), null);
});

test('cleanDrug strips HIS tags and markers', () => {
  assert.equal(P.cleanDrug('【RT給】Combivent 吸入 2.5 mL'), 'Combivent 吸入 2.5 mL');
  assert.equal(P.cleanDrug('★Vancomycin SANDOZ 針 1 g'), 'Vancomycin SANDOZ 針 1 g');
  assert.equal(P.cleanDrug('Cefin 針【禁用含鈣溶液】2 g/Vial'), 'Cefin 針 2 g/Vial');
  assert.equal(P.cleanDrug('Mopride 錠劑 5 mg'), 'Mopride 錠劑 5 mg');
});

test('cleanSymptom joins with 、', () => {
  assert.equal(P.cleanSymptom('Rash,Swelling of the skin, Eye swelling'), 'Rash、Swelling of the skin、Eye swelling');
  assert.equal(P.cleanSymptom('Rash'), 'Rash');
});

test('parseText reads the 8-column sheet layout with header', () => {
  const { rows, skipped } = P.parseText(SAMPLE);
  assert.equal(rows.length, 5);
  assert.equal(skipped.length, 0);
  assert.equal(rows[2].drug, 'Combivent 吸入 2.5 mL');
  assert.equal(rows[2].drugRaw, '【RT給】Combivent 吸入 2.5 mL');
  assert.equal(rows[1].severity, '1_輕微');
});

test('parseText reports rows with a chart number but bad date', () => {
  const { rows, skipped } = P.parseText('000000009Z\tX\tA\tB\tC\tD\tE\tnot a date');
  assert.equal(rows.length, 0);
  assert.deepEqual(skipped, [{ line: 1, reason: '無法辨識得知日期' }]);
});

test('assignIds numbers by date-time, not paste order', () => {
  const cards = P.assignIds(P.parseText(SAMPLE).rows, 1);
  assert.deepEqual(cards.map(c => c.id), ['1151001', '1151002', '1151003', '1151004', '1151005']);
  // 10:12 Cefin comes before 18:59 Vancomycin on the same day
  assert.equal(cards[3].chartNo, '000000005E');
  assert.equal(cards[4].chartNo, '000000004D');
  assert.equal(P.formatDate(cards[4].date), '2026/10/14');
});

test('assignIds honours start number; next month restarts at 01', () => {
  const text = [
    '000000006F\tX\tCEL01\tCelebrex 膠囊 200 mg\t2_中等\t病人/家屬自述\tRash\t10/17/26 10:52',
    '000000007G\tX\tPEN04\tPentop SR 錠劑 400 mg\t2_中等\t病人/家屬自述\tPalpitation\t10/20/26 14:26',
    '000000008H\tX\tCEL01\tCelebrex 膠囊 200 mg\t1_輕微\t醫療人員提供\tRash\t11/2/26 08:00',
  ].join('\n');
  const cards = P.assignIds(P.parseText(text).rows, 6);
  assert.deepEqual(cards.map(c => c.id), ['1151006', '1151007', '1151101']);
});

// Same shape as a real HIS copy: line breaks inside 記錄來源 swallow a tab,
// 登記日 is empty, and 得知日期 comes from 最後異動時間.
const HIS_HEADER = '病歷號\t姓名\t藥碼\tMDPNAME\tATC\t藥理分類\t藥理分類名稱\t嚴重程度\t記錄來源\t是否過敏\t登記日\t症狀\t疫苗接種日期\t疫苗接種院所\t最後異動者員編\t最後異動者\t最後異動時間';
const HIS_ROW = '000000007G\t王小明\tPEN04\tPentop SR 錠劑 400 mg\tC04AD03\t202400\t血液流變性劑\t2_中等\t\n病人/家屬自述\n0_藥物不良反應\t\tPalpitation\t\t\tB00000\t李藥師\t2026/10/20 14:26:51';
const HIS_ROW2 = '000000008H\t陳小美\tCEF10\tCefin 針【禁用含鈣溶液】2 g/Vial\tJ01DD04\t081204\t頭孢子菌素\t2_中等\t醫療人員提供\t0_藥物不良反應\t\tRash,Eye swelling\t\t\tB00000\t李藥師\t2026/10/14 10:12:15';

function checkPentop(r) {
  assert.equal(r.chartNo, '000000007G');
  assert.equal(r.name, '王小明');
  assert.equal(r.drugCode, 'PEN04');
  assert.equal(r.drugRaw, 'Pentop SR 錠劑 400 mg');
  assert.equal(r.severity, '2_中等');
  assert.equal(r.source, '病人/家屬自述');
  assert.equal(r.symptom, 'Palpitation');
  assert.equal(P.formatDateTime(r.date), '2026/10/20 14:26:51');
}

test('HIS export with header: multi-line records and 最後異動時間', () => {
  const { rows, skipped } = P.parseText([HIS_HEADER, HIS_ROW, HIS_ROW2].join('\n'));
  assert.equal(skipped.length, 0);
  assert.equal(rows.length, 2);
  checkPentop(rows[0]);
  assert.equal(rows[1].source, '醫療人員提供');
  assert.equal(rows[1].symptom, 'Rash、Eye swelling');
  assert.equal(rows[1].drug, 'Cefin 針 2 g/Vial');
});

test('HIS export without header is detected by shape', () => {
  const { rows } = P.parseText([HIS_ROW, HIS_ROW2].join('\r\n'));
  assert.equal(rows.length, 2);
  checkPentop(rows[0]);
  const cards = P.assignIds(rows, 4);
  assert.deepEqual(cards.map(c => c.id), ['1151004', '1151005']);
});

test('sheetRow and rowKey match the monthly sheet', () => {
  const r = P.parseText(HIS_ROW).rows[0];
  const row = P.sheetRow(r);
  assert.deepEqual(row.slice(1, 8), ['000000007G', '王小明', 'PEN04', 'Pentop SR 錠劑 400 mg', '2_中等', '病人/家屬自述', 'Palpitation']);
  assert.equal(row[0], '');
  assert.equal(P.sheetRow(P.assignIds([r], 9)[0])[0], '1151009');
  assert.equal(P.sheetName(r.date), '202610');
  assert.equal(P.rowKey(r.chartNo, r.drugCode, r.date), '000000007G|PEN04|2026/10/20 14:26:51');
});

test('rows copied back from the sheet (識別碼 first) parse', () => {
  const { rows } = P.parseText('1151006\t000000006F\tX\tCEL01\tCelebrex 膠囊 200 mg\t2_中等\t病人/家屬自述\tRash\t10/17/26 10:52');
  assert.equal(rows.length, 1);
  assert.equal(rows[0].chartNo, '000000006F');
  assert.equal(rows[0].drug, 'Celebrex 膠囊 200 mg');
});

test('sheet header starting with 識別碼 is recognised', () => {
  const text = P.SHEET_HEADER.join('\t') + '\n' + '1151006\t000000006F\tX\tCEL01\tCelebrex 膠囊 200 mg\t2_中等\t病人/家屬自述\tRash\t10/17/26 10:52';
  const { rows, skipped } = P.parseText(text);
  assert.equal(skipped.length, 0);
  assert.equal(rows.length, 1);
  assert.equal(rows[0].symptom, 'Rash');
});
