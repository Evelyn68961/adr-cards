// Run with: npm test
// Report form printed to PDF -> case. Positions mimic Chrome's print of the
// fill-in page; all data is made up.
const test = require('node:test');
const assert = require('node:assert/strict');
const PF = require('../pdfform.js');
const F = require('../forms.js');

// A piece of text: x, y (from the top), text; ring = has a circle left of
// it, on = that circle is ticked.
const t = (x, y, str, ring, on) => ({ x, y, w: Math.min(str.length * 9, 28), str, circle: !!ring, ticked: !!on });

const page1 = { height: 842, items: [
  t(24, 23, '2026/11/10 下午3:46'),
  t(170, 54, '2.5.6 - 藥品不良反應 (ADR) 通報表'),
  t(34, 112, '通報者資料'),
  t(31, 135, '通報者姓名（你的姓名）'), t(196, 139, '王藥師'),
  t(34, 178, '⼀'), t(44, 178, '般資訊'),
  t(31, 237, '通報者獲知'), t(78, 237, '⽇'), t(88, 237, '期'), t(196, 241, '2026/11/04'),
  t(31, 272, '國內案件'), t(209, 273, '是', 1, 1), t(247, 273, '否', 1),
  t(31, 299, '原始藥物不良反應獲知來源'), t(209, 300, '醫師', 1), t(256, 300, '藥師', 1, 1), t(303, 300, '其他醫療', 1), t(341, 300, '⼈'),
  t(31, 386, '識別代碼（請先詢問再填答）'), t(196, 390, '1151103'),
  t(31, 420, '原通報單位識別代碼（病歷號）'), t(196, 424, '000000001a'),
  t(31, 454, '性別'), t(209, 454, '男', 1), t(247, 454, '女', 1, 1), t(285, 454, '不知', 1),
  t(31, 481, '出'), t(41, 481, '⽣'), t(50, 481, '⽇'), t(60, 481, '期或年齡'),
  t(209, 485, '已知出', 1), t(237, 485, '⽣', 1), t(246, 485, '⽇', 1), t(255, 485, '期：', 1), t(280, 485, '1950/04/17'),
  t(209, 508, '已知首次發生不良反應年齡：', 1, 1), t(380, 508, '70'), t(400, 508, '歲'),
  t(209, 550, '不知', 1),
  t(31, 577, '⾝'), t(41, 577, '⾼'), t(196, 580, '未知'), t(260, 580, 'cm'),
  t(31, 610, '體重'), t(196, 614, '62'), t(260, 614, 'kg'),
  t(34, 654, '不良反應資訊（⼀）'),
  t(31, 677, '不良反應症狀（以中/英醫學詞彙填寫）'), t(196, 681, 'Rash, Pruritus'),
  t(24, 825, 'file:///C:/Users/x/Desktop/ADR通報表_空白表單.html'),
] };
const page2 = { height: 842, items: [
  t(31, 45, '不良反應後續結果'), t(209, 46, '已恢復/已解決', 1), t(297, 46, '恢復中/解決中', 1, 1),
  t(209, 61, '已恢復/解決但有後遺症', 1), t(335, 61, '死亡', 1),
  t(31, 88, '嚴重不良反應【院內通報獎勵⾦要'), t(31, 103, '求】'), t(209, 170, '3級(含)以下', 1, 1),
  t(31, 193, '不良反應嚴重性【ADR通報中⼼要'), t(31, 208, '求】'),
  t(209, 214, '導致死亡', 1), t(209, 290, '其他嚴重不良反應（具重要臨床意義）', 1, 1), t(209, 310, '非嚴重', 1),
  t(31, 314, '案件描述'), t(196, 318, '病人服藥後出現紅疹。'), t(196, 336, '停藥後改善。'),
  t(34, 455, '不良反應資訊（⼆）'),
  t(31, 478, '相關檢查及檢驗數據'), t(196, 482, 'AST 95 U/L'),
  t(34, 720, '⽤藥相關資訊：可疑藥品'),
] };
const page3 = { height: 842, items: [
  t(31, 45, '可疑藥品'), t(196, 49, '[藥品一]'), t(196, 66, '1. 英文商品名：Fakedrug'), t(196, 83, '12. 再投藥是否出現同樣反應(擇一)：沒有再投予'),
  t(31, 307, '併⽤產品（沒有併⽤產品請寫「無」）'), t(196, 311, '無'),
  t(34, 507, 'Naranjo score 不良反應評估問題列表'),
  t(31, 530, '1. 是否有文獻報告確定此項不良反'), t(209, 531, '是 (+1)', 1, 1), t(270, 531, '否 (+0)', 1),
  t(31, 545, '應'),
  t(31, 571, '2. 不良反應是否發⽣於給藥之後'), t(209, 572, '是 (+2)', 1, 1), t(270, 572, '否 (', 1), t(284, 572, '−'), t(290, 572, '1)'),
  t(31, 681, '5. 是否有其他因素（此藥以外）可'), t(209, 682, '是 (', 1, 1), t(223, 682, '−'), t(229, 682, '1)'), t(268, 682, '否 (+2)', 1),
  t(31, 188 + 600, 'Naranjo 最終評估'), t(209, 789, '確定 (', 1), t(233, 789, '≥'), t(238, 789, '9)'), t(275, 789, '極可能的 (5-8)', 1, 1),
] };

test('reads the printed report form, ticked choices included', () => {
  const c = PF.parsePages([page1, page2, page3]);
  assert.equal(c.reporter, '王藥師');
  assert.equal(c.learned, '2026/11/04');
  assert.deepEqual([c.domestic, c.source, c.sex], ['是', '藥師', '女']);
  assert.deepEqual([c.id, c.chartNo], ['1151103', '000000001A']);
  assert.match(c.age, /^已知首次發生不良反應年齡/);
  assert.deepEqual(F.splitAge(c.age), { kind: '已知首次發生不良反應年齡', value: '70 歲' });
  assert.deepEqual([c.height, c.weight], ['', '62']);
  assert.equal(c.symptoms, 'Rash, Pruritus');
  assert.equal(c.outcome, '恢復中/解決中');
  assert.equal(F.pickSeriousness(c.seriousness), 'Other Medically Important Condition');
  assert.equal(c.description, '病人服藥後出現紅疹。\n停藥後改善。');
  assert.equal(c.labs, 'AST 95 U/L');
  assert.equal(F.parseDrugs(c.suspect)[0].items[0].value, 'Fakedrug');
  assert.equal(c.concomitant, '無');
  // 1 + 2 - 1 = 2, and the final assessment without its range.
  assert.deepEqual([c.naranjoScore, c.naranjo], ['2', '極可能的']);
  // Usable for the TFDA sections like an Excel case.
  assert.equal(F.tfdaSections(c).sections[0].fields[0].text, '2026-11-04');
});

test('a known birth date is read with its date', () => {
  const p = { height: 842, items: page1.items.map(it => (it.y === 485 && it.x === 209 ? Object.assign({}, it, { ticked: true })
    : it.y === 508 && it.x === 209 ? Object.assign({}, it, { ticked: false }) : it)) };
  const c = PF.parsePages([p]);
  assert.deepEqual(F.splitAge(c.age), { kind: '已知出生日期', value: '1950-04-17' });
});

test('a PDF that is not the report form gives nothing', () => {
  assert.equal(PF.parsePages([{ height: 842, items: [t(40, 100, '其他文件'), t(200, 100, '內容')] }]), null);
});
