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

  const api = { workbookCases, joinForms, monthOf, severityLevel, flagCases, summarize, toTsv, SERIOUS_LABEL };
  if (inNode) module.exports = api;
  else root.ADRReview = api;
})(typeof window !== 'undefined' ? window : globalThis);
