// Case overview for the routine ADR meeting: every case in ADR通報.xlsx,
// joined with its report form (通報表) when there is one, and flagged when it
// is worth bringing up. Pure: window.ADRReview in the browser, module.exports in Node.
(function (root) {
  'use strict';

  const inNode = typeof module !== 'undefined' && module.exports;
  const P = inNode ? require('./parse.js') : root.ADRParse;
  const F = inNode ? require('./forms.js') : root.ADRForms;

  const CHART_NO = /^\d{9}[A-Z]$/;
  const ID = /^\d{7}$/;
  const SEP = '、';
  const roundSeconds = d => new Date(Math.round(d.getTime() / 1000) * 1000);
  const str = v => String(v == null ? '' : v).trim();

  // Every case row of the monthly sheets
  // (識別碼 | 病歷號 | 姓名 | 藥碼 | 藥名 | 嚴重度 | 通報來源 | 症狀 | 得知日期 | 撰寫人 | 值得注意).
  // 值得注意 is ticked on the overview page ("V"); anything typed there counts as marked.
  function workbookCases(XLSX, bytes) {
    const wb = XLSX.read(bytes, { type: 'array', cellDates: true });
    const out = [];
    wb.SheetNames.filter(n => /^\d{6}$/.test(n)).forEach(name => {
      XLSX.utils.sheet_to_json(wb.Sheets[name], { header: 1, raw: true, defval: '' }).forEach(r => {
        const chartNo = str(r[1]).toUpperCase();
        if (!CHART_NO.test(chartNo)) return;
        const date = r[8] instanceof Date ? roundSeconds(r[8]) : P.parseDate(r[8]);
        out.push({
          id: str(r[0]), chartNo, drugCode: str(r[3]), drug: P.cleanDrug(r[4]), severity: str(r[5]),
          source: str(r[6]), symptom: P.cleanSymptom(r[7]), date, writer: str(r[9]), ...markOf(r[10]),
          hisUser: str(r[11]), hisUserId: str(r[12]).toUpperCase(),
          // Typed into Excel by hand it may be a real date.
          tfdaDate: r[13] instanceof Date ? P.formatDate(r[13]) : str(r[13]),
        });
      });
    });
    return out;
  }

  const markOf = value => ({ marked: str(value) !== '' });

  // ROC year-month ("11510") a case belongs to.
  function monthOf(c) {
    if (ID.test(c.id)) return c.id.slice(0, 5);
    return c.date ? P.monthKey(c.date) : '';
  }

  // What the report form says, in the words the overview shows.
  function formFacts(form) {
    if (!form) return null;
    const suspects = F.parseDrugs(form.suspect).map(d => F.tfdaDrug(d, null));
    const rechallenge = suspects.map(d => d.fields.find(f => f.label.startsWith('再投藥')).pick);
    const score = Number((String(form.naranjoScore).match(/-?\d+/) || [])[0]);
    return {
      seriousness: F.pickSeriousness(form.seriousness),
      seriousnessRaw: form.seriousness,
      outcome: F.pickOutcome(form.outcome),
      outcomeRaw: form.outcome,
      naranjoScore: isNaN(score) ? null : score,
      naranjo: form.naranjo,
      rechallengePositive: rechallenge.includes(F.RECHALLENGE[0]),
      suspects: suspects.map(d => d.name).filter(Boolean),
      sex: form.sex,
      age: form.age,
      description: form.description,
      treatment: F.isNone(form.treatment) ? '' : form.treatment,
      reporter: form.reporter,
      writer: form.trainee || form.reporter,
      hartwig: F.hartwigLevel(form.hartwig),
      onset: form.onset,
      submitted: form.submitted,
      symptoms: form.symptoms,
    };
  }

  // Workbook cases with their report form attached (by 識別碼, and 病歷號 when
  // two forms share a number). Forms with no workbook row are listed too.
  function joinForms(cases, forms) {
    const byId = new Map();
    forms.forEach(f => byId.set(f.id, (byId.get(f.id) || []).concat(f)));
    const used = new Set();
    const out = cases.map(c => {
      const list = byId.get(c.id) || [];
      const form = list.find(f => f.chartNo === c.chartNo) || (list.length === 1 ? list[0] : null);
      if (form) used.add(form);
      return Object.assign({}, c, { month: monthOf(c), form: formFacts(form) });
    });
    forms.filter(f => !used.has(f)).forEach(f => {
      const facts = formFacts(f);
      const date = P.parseDate(F.isoDate(f.learned).replace(/-/g, '/'));
      out.push({
        id: f.id, chartNo: f.chartNo, drugCode: '', drug: facts.suspects.join(SEP), severity: '', source: f.source,
        symptom: F.splitSymptoms(f.symptoms).join(SEP), date, writer: '', month: f.id.slice(0, 5),
        marked: false, form: facts, formOnly: true,
      });
    });
    return out.sort((a, b) => a.id.localeCompare(b.id) || (a.date || 0) - (b.date || 0));
  }

  // HIS 嚴重度 "1_輕微" / "2_中等" / "3_嚴重" -> 1..3 (0 when unknown).
  function severityLevel(text) {
    return Math.max(0, ...String(text || '').split(SEP).map(s => {
      const n = s.match(/^\s*(\d)/);
      if (n) return Number(n[1]);
      return /嚴重|重度/.test(s) ? 3 : /中/.test(s) ? 2 : /輕/.test(s) ? 1 : 0;
    }));
  }

  const SERIOUS_LABEL = {
    'Results in Death': '死亡',
    'Life Threatening': '危及生命',
    'Disabling/Incapacitating': '永久性殘疾',
    'Congenital Anomaly/Birth Defect': '先天性畸形',
    'Caused/Prolonged Hospitalisation': '導致住院或延長住院',
    'Other Medically Important Condition': '其他嚴重不良反應',
  };

  // The drugs of a case as [{ key, name }]: by 藥碼 when the workbook has
  // them, else by the form's suspect drug names.
  function drugsOf(c) {
    const codes = P.splitCodes(c.drugCode);
    const names = String(c.drug || '').split(SEP).map(s => s.trim());
    if (codes.length) return codes.map((code, i) => ({ key: code, name: names[i] || code }));
    return names.filter(Boolean).map(n => ({ key: n.toUpperCase(), name: n }));
  }

  // Which cases are worth noting is the pharmacist's call: a case is
  // `notable` when it is marked (column K, or `marks` not yet saved:
  // Map(識別碼 -> true / false)).
  // `hints` only suggest what to look at: { code, label, level }, level 'high'
  // (serious), 'mid' (strong causality or not recovered) or 'info' (patterns
  // across the period). `cases` is the period being reviewed: repeats are
  // counted within it.
  function flagCases(cases, { marks = new Map() } = {}) {
    const byDrug = new Map();
    const byPatient = new Map();
    cases.forEach(c => {
      new Set(drugsOf(c).map(d => d.key)).forEach(k => byDrug.set(k, (byDrug.get(k) || 0) + 1));
      byPatient.set(c.chartNo, (byPatient.get(c.chartNo) || 0) + 1);
    });
    return cases.map(c => {
      const flags = [];
      const add = (code, label, level) => flags.push({ code, label, level });
      const f = c.form;
      const seriousness = f ? f.seriousness : '';
      if (SERIOUS_LABEL[seriousness]) add('serious', SERIOUS_LABEL[seriousness], 'high');
      else if (f && f.outcome === '死亡') add('serious', '死亡', 'high');
      else if (!seriousness && severityLevel(c.severity) >= 3) add('serious', 'HIS 嚴重度：嚴重', 'high');
      if (f && f.outcome === '已恢復/解決但有後遺症') add('outcome', '有後遺症', 'mid');
      if (f && f.outcome === '尚未恢復/尚未解決/不良反應持續中') add('outcome', '尚未恢復', 'mid');
      if (f && f.rechallengePositive) add('rechallenge', '再投予後再發生', 'mid');
      if (f && (f.naranjoScore >= 9 || /definite|確定/i.test(f.naranjo))) add('naranjo', 'Naranjo 確定相關', 'mid');
      const repeated = drugsOf(c).filter(d => byDrug.get(d.key) > 1);
      if (repeated.length) {
        add('repeatDrug', '同藥品多件：' + repeated.map(d => `${d.key}×${byDrug.get(d.key)}`).join(SEP), 'info');
      }
      if (byPatient.get(c.chartNo) > 1) add('repeatPatient', `同病人 ${byPatient.get(c.chartNo)} 件`, 'info');
      if (f && f.hartwig >= 4) add('hartwig', `Hartwig ${f.hartwig} 級（獎勵）`, 'high');
      const marked = marks.has(c.id) ? marks.get(c.id) : c.marked;
      return Object.assign({}, c, { hints: flags, marked, notable: marked, serious: flags.some(x => x.level === 'high') });
    });
  }

  // Numbers for the top of the overview.
  function summarize(flagged) {
    const count = (list, keyOf) => {
      const m = new Map();
      list.forEach(x => keyOf(x).forEach(({ key, name }) => {
        const e = m.get(key) || { key, name, n: 0 };
        e.n++;
        m.set(key, e);
      }));
      return [...m.values()].sort((a, b) => b.n - a.n || a.name.localeCompare(b.name));
    };
    return {
      total: flagged.length,
      notable: flagged.filter(c => c.notable).length,
      suggested: flagged.filter(c => !c.notable && c.hints.length).length,
      serious: flagged.filter(c => c.serious).length,
      noForm: flagged.filter(c => !c.form).length,
      noWriter: flagged.filter(c => !c.formOnly && !c.writer).length,
      drugs: count(flagged, c => [...new Map(drugsOf(c).map(d => [d.key, d])).values()]),
      symptoms: count(flagged, c => [...new Set(String(c.symptom).split(SEP).map(s => s.trim()).filter(Boolean))]
        .map(s => ({ key: s.toLowerCase(), name: s }))),
    };
  }

  // Tab-separated table (with header) for pasting into the meeting report.
  const TSV_HEADER = ['識別碼', '得知日期', '病歷號', '藥品', '症狀', '嚴重度', '嚴重性（通報表）', '後續結果', 'Naranjo', '撰寫人', '值得注意', '自動提示'];
  function toTsv(flagged) {
    const cell = v => String(v == null ? '' : v).replace(/[\t\r\n]+/g, ' ');
    return [TSV_HEADER].concat(flagged.map(c => [
      c.id, c.date ? P.formatDate(c.date) : '', c.chartNo, c.drug, c.symptom, c.severity,
      c.form ? SERIOUS_LABEL[c.form.seriousness] || (c.form.seriousness ? '非嚴重' : '') : '',
      c.form ? c.form.outcome : '', c.form && c.form.naranjoScore != null ? c.form.naranjoScore : '',
      c.writer, c.notable ? 'V' : '', c.hints.map(x => x.label).join('；'),
    ])).map(r => r.map(cell).join('\t')).join('\r\n');
  }

  // A period's cases, packed small for the phone link (cases.html). No names
  // or 病歷號: only what is needed to check a case.
  const LEVEL_CODE = { high: 'h', mid: 'm', info: 'i' };
  const LEVEL_NAME = { h: 'high', m: 'mid', i: 'info' };
  function packCases(flagged, { from, to, at }) {
    return {
      v: 1, f: from, t: to, at,
      c: flagged.map(c => [
        c.id, c.date ? `${c.date.getMonth() + 1}/${c.date.getDate()}` : '', c.drug, c.symptom, c.severity,
        c.form ? SERIOUS_LABEL[c.form.seriousness] || (c.form.seriousness ? '非嚴重' : '') : '',
        c.form ? c.form.outcome || '' : '', c.writer || '', c.notable ? 1 : 0,
        c.hints.map(h => [h.label, LEVEL_CODE[h.level]]),
      ]),
    };
  }
  function unpackCases(data) {
    if (!data || data.v !== 1 || !Array.isArray(data.c)) throw new Error('連結內容無法辨識');
    return {
      from: data.f, to: data.t, at: data.at,
      cases: data.c.map(([id, date, drug, symptom, severity, serious, outcome, writer, marked, hints]) => ({
        id, date, drug, symptom, severity, serious, outcome, writer, marked: !!marked,
        hints: (hints || []).map(([label, level]) => ({ label, level: LEVEL_NAME[level] || 'info' })),
      })),
    };
  }

  // ---- meeting list (例會名單), one row per ticked case ----

  const MEETING_HEADER = ['識別代碼', 'ICU', '病歷號', '通報者(200)', '員編', '發生日期', '通報日期', '通報藥品', '通報症狀', '撰寫人', '員編', '審核者', '員編'];
  const MEETING_WIDTHS = [9, 5, 11, 12, 7, 10, 10, 9, 21, 8, 7, 8, 7];
  // 員編 starts with the staff group: A 醫, B 護, C 藥.
  const ROLE = { A: '醫', B: '護', C: '藥' };
  const roleOf = id => ROLE[String(id || '').trim().charAt(0).toUpperCase()] || '';

  // "2026-04-16" / "2026/4/16" -> "2026/4/16" (pad: "2026/04/16").
  function slashDate(text, pad) {
    const m = String(text || '').match(/(\d{4})[\/\-.](\d{1,2})[\/\-.](\d{1,2})/);
    if (!m) return String(text || '').trim();
    const p = n => (pad ? n.padStart(2, '0') : String(Number(n)));
    return `${m[1]}/${p(m[2])}/${p(m[3])}`;
  }

  // 員編 by name, from the 通報者 HIS recorded (pharmacists key in most cases).
  function staffIds(cases) {
    const ids = new Map();
    cases.forEach(c => { if (c.hisUser && c.hisUserId && !ids.has(c.hisUser)) ids.set(c.hisUser, c.hisUserId); });
    return ids;
  }

  // Rows for the meeting list (array of arrays, header first, 單位主管 last).
  // ICU is not recorded anywhere and is left blank; 通報日期 is the day it was
  // reported to TFDA (TFDA helper), else the form's 表單提交日期. `allCases`
  // (every workbook case) supplies the 員編 of names.
  function meetingSheet(cases, { allCases = cases, reviewer = '', reviewerId = '' } = {}) {
    const ids = staffIds(allCases);
    const rows = cases.slice().sort((a, b) => a.id.localeCompare(b.id)).map(c => {
      const f = c.form || {};
      const writer = c.writer || f.writer || '';
      const reporter = c.hisUser ? `${c.hisUser}${roleOf(c.hisUserId) ? `(${roleOf(c.hisUserId)})` : ''}` : '';
      return [
        c.id, '', c.chartNo, reporter, c.hisUserId || '',
        slashDate(f.onset, false), slashDate(c.tfdaDate || f.submitted, true),
        P.splitCodes(c.drugCode).join('、'), f.symptoms || c.symptom || '',
        writer, ids.get(writer) || '', reviewer, reviewerId,
      ];
    });
    return [MEETING_HEADER, ...rows, [], ['單位主管：']];
  }

  const api = { meetingSheet, MEETING_WIDTHS, staffIds, packCases, unpackCases, workbookCases, joinForms, monthOf, severityLevel, flagCases, summarize, toTsv, SERIOUS_LABEL };
  if (inNode) module.exports = api;
  else root.ADRReview = api;
})(typeof window !== 'undefined' ? window : globalThis);
