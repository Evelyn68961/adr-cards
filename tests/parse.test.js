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

test('parseText skips header and keeps card fields only', () => {
  const { rows, skipped } = P.parseText(SAMPLE);
  assert.equal(rows.length, 5);
  assert.equal(skipped.length, 0);
  assert.deepEqual(Object.keys(rows[0]).sort(), ['chartNo', 'date', 'drug', 'source', 'symptom']);
  assert.equal(rows[2].drug, 'Combivent 吸入 2.5 mL');
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
