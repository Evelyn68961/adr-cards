// Reads the hospital's ADR report-form export (2.5.6－藥品不良反應 (ADR) 通報表)
// and maps each case onto the fields of the TFDA online reporting form.
// Pure functions only: window.ADRForms in the browser, module.exports in Node.
(function (root) {
  'use strict';

  // The export repeats a header row whenever the form version changes, and
  // the column order differs between versions, so columns are found by the
  // first line of their header text.
  const HEADER_START = '臨床訓練計畫';
  const COLUMNS = {
    trainee: '學員',
    reporter: '通報者姓名',
    onset: '不良反應發生之最初日期',
    learned: '通報者獲知日期',
    domestic: '國內案件',
    source: '原始藥物不良反應獲知來源',
    literature: '本案件是否為文獻通報',
    id: '識別代碼',
    chartNo: '原通報單位識別代碼',
    sex: '性別',
    age: '出生日期或年齡',
    height: '身高',
    weight: '體重',
    symptoms: '不良反應症狀',
    treatment: '處置',
    outcome: '不良反應後續結果',
    seriousness: '不良反應嚴重性',
    description: '案件描述',
    labs: '相關檢查及檢驗數據',
    labNote: '檢查描述',
    other: '其他相關資訊',
    suspect: '可疑藥品',
    concomitant: '併用產品',
    interacting: '交互作用產品',
    naranjoScore: '請將上述問題分數加總',
    naranjo: 'Naranjo score',
    submitted: '表單提交日期',
  };

  const firstLine = h => String(h).split(/[\r\n]/)[0].trim();
  const clean = v => String(v == null ? '' : v).replace(/\r\n?/g, '\n').replace(/[ \t]+\n/g, '\n').trim();
  const isNone = v => /^(無|nil|none|n\/a|-)?$/i.test(clean(v).replace(/[。.]$/, ''));

  function headerIndex(row) {
    const names = row.map(firstLine);
    const index = {};
    for (const [key, prefix] of Object.entries(COLUMNS)) {
      const i = names.findIndex(n => n.startsWith(prefix));
      if (i >= 0) index[key] = i;
    }
    return index;
  }

  // rows: array of arrays (SheetJS sheet_to_json with header:1, raw:false).
  function readCases(rows) {
    const cases = [];
    let index = null;
    for (const row of rows) {
      if (clean(row[0]) === HEADER_START) { index = headerIndex(row); continue; }
      if (!index) continue;
      const c = {};
      for (const key of Object.keys(COLUMNS)) c[key] = key in index ? clean(row[index[key]]) : '';
      c.chartNo = c.chartNo.toUpperCase();
      if (!/^\d{7}$/.test(c.id)) continue;
      c.filled = row.filter(v => clean(v)).length;
      cases.push(c);
    }
    return cases;
  }

  // The same case can be exported more than once (resubmitted, or an empty
  // draft); keep the most complete copy per 識別代碼 + 病歷號.
  function mergeCases(lists) {
    const best = new Map();
    for (const c of [].concat(...lists)) {
      const key = c.id + '|' + c.chartNo;
      const prev = best.get(key);
      if (!prev || c.filled >= prev.filled) best.set(key, c);
    }
    return [...best.values()].sort((a, b) => a.id.localeCompare(b.id) || a.chartNo.localeCompare(b.chartNo));
  }

  // "2026/9/24", "2026-09-24", "9/24/26" -> "2026-09-24"
  function isoDate(text) {
    const s = clean(text);
    let m = s.match(/(\d{4})[\/\-.](\d{1,2})[\/\-.](\d{1,2})/);
    if (m) return `${m[1]}-${m[2].padStart(2, '0')}-${m[3].padStart(2, '0')}`;
    m = s.match(/^(\d{1,2})\/(\d{1,2})\/(\d{2})$/);
    if (m) return `20${m[3]}-${m[1].padStart(2, '0')}-${m[2].padStart(2, '0')}`;
    return s;
  }

  const SOURCES = ['醫師', '藥師', '其他醫療人員', '民眾或其他非醫療人員'];
  const AGE_KINDS = ['已知出生日期', '已知首次發生不良反應年齡', '大約年齡', '已知懷孕週數', '未知'];
  const SERIOUSNESS = ['Results in Death', 'Life Threatening', 'Disabling/Incapacitating',
    'Congenital Anomaly/Birth Defect', 'Caused/Prolonged Hospitalisation',
    'Other Medically Important Condition', 'Non-serious reaction'];
  const OUTCOMES = ['已恢復/已解決', '恢復中/解決中', '尚未恢復/尚未解決/不良反應持續中',
    '已恢復/解決但有後遺症', '死亡', '未知'];

  function pickSource(v) {
    const i = SOURCES.slice(0, 3).indexOf(v);
    return i >= 0 ? SOURCES[i] : v ? SOURCES[3] : '';
  }

  function pickSeriousness(v) {
    // 其他嚴重 first: its option text gives examples that mention 住院.
    const rules = [[/非嚴重/, 6], [/其他嚴重|重要臨床意義/, 5], [/死亡/, 0], [/危及生命/, 1],
      [/殘|失能/, 2], [/先天|畸形/, 3], [/住院/, 4]];
    const hit = rules.find(([re]) => re.test(v));
    return hit ? SERIOUSNESS[hit[1]] : '';
  }

  function pickOutcome(v) {
    const rules = [[/後遺症/, 3], [/死亡/, 4], [/尚未|持續/, 2], [/恢復中|解決中/, 1], [/已恢復|已解決/, 0], [/未知/, 5]];
    const hit = rules.find(([re]) => re.test(v));
    return hit ? OUTCOMES[hit[1]] : '';
  }

  // "已知出生日期 2018-07-13" / "已知首次發生不良反應年齡 70歲" / "僅知大約年齡 50"
  function splitAge(v) {
    const s = clean(v);
    const kinds = [[/出生日期/, 0], [/首次發生/, 1], [/大約/, 2], [/懷孕/, 3], [/未知/, 4]];
    const hit = kinds.find(([re]) => re.test(s));
    const value = (s.match(/\d[\d\-\/.]*\s*(歲|週|周)?/) || [''])[0].trim();
    return { kind: hit ? AGE_KINDS[hit[1]] : '', value: hit && hit[1] === 0 ? isoDate(value) : value };
  }

  // 0 means "not recorded" on the internal form.
  const measure = v => (/^0*(\.0*)?$/.test(clean(v)) ? '' : clean(v));

  const splitSymptoms = v => (isNone(v) ? [] : clean(v).split(/\s*[,，、;；\n]\s*/).filter(Boolean));

  // "[藥品一]\n1. 英文商品名：…\n2. 中文商品名：…" -> [{ title, items: [{ label, value }] }]
  function parseDrugs(text) {
    if (isNone(text)) return [];
    const drugs = [];
    let cur = null;
    for (const raw of clean(text).split('\n')) {
      const line = raw.trim();
      if (!line) continue;
      const head = line.match(/^\[(.+?)\]\s*(.*)$/);
      if (head || !cur) {
        cur = { title: head ? head[1] : '藥品', items: [] };
        drugs.push(cur);
        if (!head) { addItem(cur, line); continue; }
        if (head[2]) addItem(cur, head[2]);
        continue;
      }
      addItem(cur, line);
    }
    return drugs;
  }

  function addItem(drug, line) {
    const m = line.match(/^(?:\d+\s*[.、]\s*)?([^：:]{1,20})[：:]\s*(.*)$/);
    if (m) {
      const label = m[1].trim();
      const value = /日期/.test(label) ? isoDate(m[2]) : m[2].trim();
      drug.items.push({ label, value });
    } else if (drug.items.length) {
      drug.items[drug.items.length - 1].value += '\n' + line;
    } else {
      drug.items.push({ label: '', value: line });
    }
  }

  const DISPOSITIONS = ['停藥', '降低劑量', '增加劑量', '未改變劑量', '未知', '不適用'];
  const RECHALLENGE = ['有再投予且不良反應再發生', '有再投予但不良反應未發生', '有再投予但結果未知', '沒有再投予'];

  const itemValue = (drug, re) => {
    const v = ((drug.items.find(i => re.test(i.label)) || {}).value || '').trim();
    return /^(未知|不詳|無|n\/a)$/i.test(v) ? '' : v;
  };

  // "1.2g" -> ["1.2", "g"]; "160 mg (1 tab)" -> ["160", "mg"]; free text ("適量") -> ["", ""]
  function splitDose(v) {
    const m = String(v).replace(/^.*[：:]/, '').replace(/\s*[(（][^)）]*[)）]\s*$/, '').match(/^\s*([\d.]+)\s*([^\d\s/&]*)\s*$/);
    return m ? [m[1], m[2]] : ['', ''];
  }

  // Days from start to end, both counted.
  function useDays(start, end) {
    const a = Date.parse(start), b = Date.parse(end);
    if (isNaN(a) || isNaN(b) || b < a) return '';
    return String(Math.round((b - a) / 86400000) + 1);
  }

  // TFDA's quick search matches a product name; the first word of the English
  // brand name ("CEFIN FOR I.V. INJECTION …" -> "CEFIN") finds it most often.
  const searchKey = name => (String(name).match(/[A-Za-z][A-Za-z-]*/) || [''])[0];

  function pickFrom(options, v) {
    // Longest option first, so "未改變劑量" is not taken for "未知" etc.
    return [...options].sort((a, b) => b.length - a.length).find(o => v.includes(o)) || '';
  }

  // Older forms say "有再投予且不良反應未發生", so match on meaning, not exact text.
  function pickRechallenge(v) {
    const rules = [[/沒有再投|未再投/, 3], [/未發生/, 1], [/結果未知/, 2], [/再發生|再度發生/, 0]];
    const hit = rules.find(([re]) => re.test(v));
    return hit ? RECHALLENGE[hit[1]] : '';
  }

  // A licence number in 快速搜尋 brings up exactly one product, which then
  // fills 成分名, 許可證字號, 商品名, 廠牌 and 劑型 by itself.
  // With the desktop helper the page replaces this field with the hospital's
  // own drug list (`hospital`), and the name-based TFDA matches are the fallback.
  function searchField(en, zh, lookup) {
    const hospital = { query: searchKey(en) || zh, en, zh };
    return Object.assign(nameSearchField(en, zh, lookup), { hospital });
  }

  function nameSearchField(en, zh, lookup) {
    const hits = lookup ? lookup({ en, zh }) : [];
    if (!hits.length) {
      return { label: '快速搜尋', text: searchKey(en) || zh, hint: '找不到許可證字號，改用名稱搜尋，再選對品項' };
    }
    return {
      label: '快速搜尋（許可證字號）', text: hits[0].license, candidates: hits,
      hint: hits[0].score < 0.9 ? '名稱不完全相符，請核對品名與廠牌' : '',
    };
  }

  // One drug block -> fields in the order of TFDA's 新增藥品 dialog.
  // `lookup({ en, zh })` returns licence candidates (licenses.js), if loaded.
  function tfdaDrug(drug, lookup) {
    const en = itemValue(drug, /英文/);
    const zh = itemValue(drug, /中文/);
    const doseText = itemValue(drug, /劑量/);
    const [dose, doseUnit] = splitDose(doseText);
    const start = itemValue(drug, /起始/), end = itemValue(drug, /結束/);
    const days = useDays(start, end);
    const disposition = itemValue(drug, /處置/);
    const rechallenge = itemValue(drug, /再投/);
    return {
      title: drug.title,
      name: en || zh,
      fields: [
        searchField(en, zh, lookup),
        { label: '英文商品名（核對用）', text: en },
        { label: '中文商品名（核對用）', text: zh },
        { label: '劑量', text: dose, hint: !dose && doseText ? `無法拆分，原始內容：${doseText}` : '' },
        { label: '劑量單位', text: doseUnit },
        { label: '給藥頻率（下拉選）', text: itemValue(drug, /頻率/) },
        { label: '給藥途徑（下拉選）', text: itemValue(drug, /途徑/) },
        { label: '用藥起始日期', text: start },
        { label: '用藥結束日期', text: end },
        { label: '藥品使用時間', text: days, hint: days ? '單位選「天」' : '' },
        { label: '用藥原因', text: itemValue(drug, /原因/), long: true },
        { label: '批號', text: itemValue(drug, /批號/) },
        { label: '效期', text: isoDate(itemValue(drug, /效期/)) },
        { label: '處置情形', options: DISPOSITIONS, pick: pickFrom(DISPOSITIONS, disposition), raw: disposition },
        { label: '再投藥是否出現同樣反應', options: RECHALLENGE, pick: pickRechallenge(rechallenge), raw: rechallenge },
      ],
    };
  }

  const DESCRIPTION_MIN = 20;
  const LAB_NOTE_MAX = 250;

  // Text for TFDA's 檢驗說明 when the lab results go in an attachment.
  function labNoteText(c) {
    const parts = [];
    if (!isNone(c.labs)) parts.push(`檢驗數據詳見附件 ${attachmentName(c)}。`);
    if (!isNone(c.labNote)) parts.push(c.labNote);
    return parts.join('\n');
  }

  const attachmentName = c => `ADR_${c.id}_檢驗資料.pdf`;
  // 檢查描述 alone fits in 檢驗說明, so only lab data needs an attachment.
  const hasLabs = c => !isNone(c.labs);

  // Everything the TFDA form asks for, grouped by its tabs. A field is either
  // text to paste ({ text }) or a choice to click ({ options, pick }).
  function tfdaSections(c, lookup) {
    const drugs = text => parseDrugs(text).map(d => tfdaDrug(d, lookup));
    const age = splitAge(c.age);
    const warn = [];
    if (clean(c.description).length < DESCRIPTION_MIN) warn.push(`不良反應描述少於 ${DESCRIPTION_MIN} 字`);
    const note = labNoteText(c);
    if (note.length > LAB_NOTE_MAX) warn.push(`檢驗說明超過 ${LAB_NOTE_MAX} 字（${note.length}）`);
    if (!measure(c.height) || !measure(c.weight)) warn.push('身高或體重未填');
    const seriousness = pickSeriousness(c.seriousness);
    return {
      warnings: warn,
      sections: [
        { tab: 'General Information', fields: [
          { label: '獲知藥品不良反應日期', text: isoDate(c.learned) },
          { label: '獲知不良反應最新資訊日期', text: isoDate(c.learned), hint: '同獲知日期；如有更新資訊請改填' },
          { label: '獲知藥品不良反應來源', options: SOURCES, pick: pickSource(c.source), raw: c.source },
          { label: '是否為國內案件', options: ['是', '否'], pick: c.domestic === '否' ? '否' : '是' },
          { label: '文獻通報', options: ['是', '否'], pick: c.literature === '是' ? '是' : '否' },
        ] },
        { tab: 'Patient Information', fields: [
          { label: '識別代碼', text: c.id },
          { label: '原通報單位識別代碼', text: c.chartNo },
          { label: '性別', options: ['男', '女', '未知'], pick: ['男', '女'].includes(c.sex) ? c.sex : '未知' },
          { label: '體重(公斤)', text: measure(c.weight) },
          { label: '身高(公分)', text: measure(c.height) },
          { label: '出生日期或年齡', options: AGE_KINDS, pick: age.kind, raw: c.age },
          { label: age.kind === '已知出生日期' ? '出生日期' : '年齡', text: age.value },
        ] },
        { tab: 'Reaction Information(1)', fields: [
          { label: '發生日期', text: isoDate(c.onset) },
          { label: '不良反應嚴重性', options: SERIOUSNESS, pick: seriousness, raw: c.seriousness },
          ...(seriousness === SERIOUSNESS[0]
            ? [{ label: '死亡日期 / 死亡原因', text: '', hint: '內部表單沒有此欄，請自行填寫' }] : []),
          { label: '是否同意提供通報者姓名等資料給藥商', options: ['是', '否'], pick: '', hint: '內部表單沒有此欄，請自行選擇' },
          { label: '不良反應症狀（逐一新增）', list: splitSymptoms(c.symptoms) },
          { label: '不良反應描述', text: c.description, long: true },
          ...(isNone(c.treatment) ? [] : [{ label: '處置（TFDA 無此欄，可併入描述）', text: c.treatment, long: true }]),
          { label: '不良反應後續結果', options: OUTCOMES, pick: pickOutcome(c.outcome), raw: c.outcome },
        ] },
        { tab: 'Reaction Information(2)', fields: [
          { label: '檢驗結果', text: '', hint: isNone(c.labs) ? '無檢驗數據，不需新增' : `不用逐筆新增，改上傳附件 ${attachmentName(c)}` },
          { label: '檢驗說明', text: note, long: true, max: LAB_NOTE_MAX },
          { label: '其他相關資料', text: isNone(c.other) ? '' : c.other, long: true },
        ] },
        { tab: 'Drug Information', drugs: [
          { kind: '可疑藥品', list: drugs(c.suspect) },
          { kind: '併用藥品', list: drugs(c.concomitant) },
          { kind: '相互作用藥品', list: drugs(c.interacting) },
        ], fields: [
          { label: 'Naranjo score（參考）', text: [c.naranjoScore, c.naranjo].filter(Boolean).join('　') },
        ] },
        { tab: 'Upload Attachment', fields: [
          hasLabs(c) ? { label: '附件', attachment: true } : { label: '附件', text: '', hint: '無檢驗資料，不需附件' },
        ] },
      ],
    };
  }

  // Lines for the lab attachment, in reading order.
  function attachmentContent(c) {
    return {
      title: '藥品不良反應通報　檢驗資料',
      meta: [`識別代碼：${c.id}`, `原通報單位識別代碼：${c.chartNo}`, `不良反應：${splitSymptoms(c.symptoms).join('、')}`],
      blocks: [
        { heading: '相關檢查及檢驗數據', text: isNone(c.labs) ? '無' : c.labs },
        { heading: '檢查描述', text: isNone(c.labNote) ? '無' : c.labNote },
      ],
    };
  }

  // "1151001" -> "11510"; ROC year-month the 識別碼 belongs to.
  const monthOf = c => c.id.slice(0, 5);

  const api = {
    readCases, mergeCases, isoDate, splitAge, splitSymptoms, parseDrugs, tfdaDrug, pickSource, pickSeriousness,
    pickOutcome, tfdaSections, attachmentContent, attachmentName, hasLabs, monthOf, isNone,
    SOURCES, SERIOUSNESS, OUTCOMES, AGE_KINDS,
  };
  if (typeof module !== 'undefined' && module.exports) module.exports = api;
  else root.ADRForms = api;
})(typeof window !== 'undefined' ? window : globalThis);
