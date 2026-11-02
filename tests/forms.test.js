// Run with: npm test
// Report-form export -> TFDA fields. All data is made up.
const test = require('node:test');
const assert = require('node:assert/strict');
const F = require('../forms.js');

const HEADER_A = ['臨床訓練計畫', '學員', '不良反應發生之最初日期\r', '通報者獲知日期\r', '國內案件\r',
  '原始藥物不良反應獲知來源\r', '識別代碼 (請先詢問再填答)\r', '原通報單位識別代碼 (病歷號)\r', '性別\r',
  '出生日期或年齡\r', '身高\r', '體重\r', '不良反應症狀\r\r(以中/英醫學詞彙填寫)\r', '不良反應嚴重性【ADR通報中心要求】\r',
  '案件描述\r', '不良反應後續結果\r', '相關檢查及檢驗數據\r', '檢查描述\r', '可疑藥品\r', '併用產品\r'];
// A later form version with columns in another order.
const HEADER_B = ['臨床訓練計畫', '學員', '識別代碼 (請先詢問再填答)\r', '原通報單位識別代碼 (病歷號)\r',
  '通報者獲知日期\r', '不良反應症狀\r', '不良反應嚴重性【ADR通報中心要求】\r', '不良反應後續結果\r', '案件描述\r'];

const DRUG = '[藥品一]\n1. 英文商品名：Fakedrug 5 mg\n2. 中文商品名：假藥錠\n5. 起始日期：2026-9-7\n7. 用藥原因：上腹痛\n延續的說明';

const rowA = ['計畫', '甲', '2026-09-24', '2026-10-05', '是', '民眾', '1151001', '000000001a\t', '男',
  '已知首次發生不良反應年齡 70歲', '160', '0', 'Rash, Pruritus、Fever', '其他嚴重不良反應 (例如: 不需要住院處理的癲癇發作)',
  '病人為七十歲男性，服藥後第三天出現全身紅疹與搔癢。', '未知 (病人未回診)', '1. 2026/9/24。AST 895 U/L', '無', DRUG, '無'];

test('reads every header version and maps columns by name', () => {
  const rows = [HEADER_A, rowA, HEADER_B,
    ['計畫', '乙', '1151002', '000000002B', '2026-10-09', 'EPS', '非嚴重', '已恢復/已解決', '短'],
    ['計畫', '乙', '0', '', '', '', '', '', '']];
  const cases = F.readCases(rows);
  assert.equal(cases.length, 2);
  assert.equal(cases[0].chartNo, '000000001A');
  assert.equal(cases[1].id, '1151002');
  assert.equal(cases[1].symptoms, 'EPS');
});

test('keeps the most complete copy of a duplicated case', () => {
  const full = F.readCases([HEADER_A, rowA])[0];
  const draft = Object.assign({}, full, { suspect: '', filled: 5 });
  const merged = F.mergeCases([[draft], [full]]);
  assert.equal(merged.length, 1);
  assert.equal(merged[0].suspect, DRUG);
});

test('maps choices onto the TFDA options', () => {
  assert.equal(F.pickSource('民眾'), '民眾或其他非醫療人員');
  assert.equal(F.pickSource('藥師'), '藥師');
  assert.equal(F.pickSeriousness('其他嚴重不良反應 (例如: 不需要住院處理的癲癇發作)'), 'Other Medically Important Condition');
  assert.equal(F.pickSeriousness('導致病人住院或延長病人住院時間'), 'Caused/Prolonged Hospitalisation');
  assert.equal(F.pickSeriousness('非嚴重'), 'Non-serious reaction');
  assert.equal(F.pickOutcome('未知 (病人未回診、或已出院無法追蹤)'), '未知');
  assert.equal(F.pickOutcome('已恢復/解決但有後遺症'), '已恢復/解決但有後遺症');
  assert.equal(F.pickOutcome('恢復中/解決中'), '恢復中/解決中');
  assert.deepEqual(F.splitAge('已知出生日期 2018-7-13'), { kind: '已知出生日期', value: '2018-07-13' });
  assert.deepEqual(F.splitAge('僅知大約年齡 50'), { kind: '大約年齡', value: '50' });
});

test('splits symptoms and drug blocks', () => {
  assert.deepEqual(F.splitSymptoms('Rash, Pruritus、Fever'), ['Rash', 'Pruritus', 'Fever']);
  assert.deepEqual(F.splitSymptoms('無'), []);
  const [drug] = F.parseDrugs(DRUG);
  assert.equal(drug.title, '藥品一');
  assert.deepEqual(drug.items.map(i => i.label), ['英文商品名', '中文商品名', '起始日期', '用藥原因']);
  assert.equal(drug.items[2].value, '2026-09-07');
  assert.equal(drug.items[3].value, '上腹痛\n延續的說明');
  assert.deepEqual(F.parseDrugs('無'), []);
});

test('builds TFDA sections and points lab results at the attachment', () => {
  const c = F.readCases([HEADER_A, rowA])[0];
  const { sections, warnings } = F.tfdaSections(c);
  assert.deepEqual(sections.map(s => s.tab), ['General Information', 'Patient Information',
    'Reaction Information(1)', 'Reaction Information(2)', 'Drug Information', 'Upload Attachment']);
  const byLabel = Object.fromEntries(sections.flatMap(s => s.fields).map(f => [f.label, f]));
  assert.equal(byLabel['獲知藥品不良反應日期'].text, '2026-10-05');
  assert.equal(byLabel['體重(公斤)'].text, '');
  assert.match(byLabel['檢驗說明'].text, /ADR_1151001_檢驗資料\.pdf/);
  assert.ok(warnings.includes('身高或體重未填'));
  assert.ok(F.hasLabs(c));
  assert.equal(F.attachmentContent(c).blocks[0].text, '1. 2026/9/24。AST 895 U/L');
});
