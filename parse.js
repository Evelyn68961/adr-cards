// Parsing and formatting for ADR cards. Pure functions only, so the same file
// runs in the browser (window.ADRParse) and in Node tests (module.exports).
(function (root) {
  'use strict';

  // HIS ADR list column order (tab-separated when copied from HIS or Excel).
  const COL = { chartNo: 0, name: 1, drugCode: 2, drug: 3, severity: 4, source: 5, symptom: 6, date: 7 };
  const CHART_NO = /^\d{9}[A-Z]$/i;

  // Accepts the formats the date column shows up in:
  //   2026/10/5 9:57[:01], 2026-10-05 09:57   (Gregorian Y/M/D)
  //   115/10/05 09:57                          (ROC Y/M/D)
  //   10/5/26 9:57                             (Excel "m/d/yy h:mm" display)
  function parseDate(text) {
    const m = String(text || '').trim().match(
      /^(\d{1,4})[\/\-.](\d{1,2})[\/\-.](\d{1,4})(?:[ T]+(\d{1,2}):(\d{2})(?::(\d{2}))?)?/);
    if (!m) return null;
    let [a, b, c] = [m[1], m[2], m[3]].map(Number);
    let y, mo, d;
    if (m[1].length === 4) { y = a; mo = b; d = c; }
    else if (m[1].length === 3) { y = a + 1911; mo = b; d = c; }
    else { mo = a; d = b; y = m[3].length === 4 ? c : 2000 + c; }
    const date = new Date(y, mo - 1, d, Number(m[4] || 0), Number(m[5] || 0), Number(m[6] || 0));
    if (date.getMonth() !== mo - 1 || date.getDate() !== d) return null;
    return date;
  }

  // "【RT給】Combivent 吸入 2.5 mL" -> "Combivent 吸入 2.5 mL"
  // "★Vancomycin SANDOZ 針 1 g"   -> "Vancomycin SANDOZ 針 1 g"
  // "Cefin 針【禁用含鈣溶液】2 g/Vial" -> "Cefin 針 2 g/Vial"
  function cleanDrug(text) {
    return String(text || '')
      .replace(/【[^】]*】/g, ' ')
      .replace(/[★☆◎●○▲△■□◆◇※＊*]/g, ' ')
      .replace(/\s+/g, ' ')
      .trim();
  }

  // "Rash,Swelling of the skin" -> "Rash、Swelling of the skin"
  function cleanSymptom(text) {
    return String(text || '')
      .split(/\s*[,，;；、]\s*/)
      .map(s => s.trim())
      .filter(Boolean)
      .join('、');
  }

  function formatDate(date) {
    return `${date.getFullYear()}/${date.getMonth() + 1}/${date.getDate()}`;
  }

  // 識別碼 prefix: ROC year + 2-digit month, e.g. 2026/10 -> "11510".
  function monthKey(date) {
    return String(date.getFullYear() - 1911) + String(date.getMonth() + 1).padStart(2, '0');
  }

  function splitRow(line) {
    if (line.includes('\t')) return line.split('\t').map(s => s.trim());
    return line.trim().split(/\s{2,}/);
  }

  // Returns { rows, skipped } where each row is a card-ready object.
  // Lines without a chart number (headers, blank lines) are skipped silently;
  // lines that have one but fail to parse are reported in `skipped`.
  function parseText(text) {
    const rows = [];
    const skipped = [];
    String(text || '').split(/\r?\n/).forEach((line, i) => {
      if (!line.trim()) return;
      const cells = splitRow(line);
      const chartNo = (cells[COL.chartNo] || '').toUpperCase();
      if (!CHART_NO.test(chartNo)) {
        if (cells.some(c => CHART_NO.test(c))) skipped.push({ line: i + 1, reason: '欄位順序不符' });
        return;
      }
      const date = parseDate(cells[COL.date]);
      if (!date) { skipped.push({ line: i + 1, reason: '無法辨識得知日期' }); return; }
      rows.push({
        chartNo,
        date,
        source: (cells[COL.source] || '').trim(),
        drug: cleanDrug(cells[COL.drug]),
        symptom: cleanSymptom(cells[COL.symptom]),
      });
    });
    return { rows, skipped };
  }

  // Sorts by 得知日期 (date + time) and assigns 識別碼. `startSeq` applies to
  // the earliest month in the batch; any later month starts again at 1.
  function assignIds(rows, startSeq) {
    const sorted = rows.slice().sort((a, b) => a.date - b.date);
    const counters = {};
    const firstKey = sorted.length ? monthKey(sorted[0].date) : null;
    return sorted.map(r => {
      const key = monthKey(r.date);
      if (!(key in counters)) counters[key] = key === firstKey ? Math.max(1, startSeq | 0) : 1;
      const seq = counters[key]++;
      return Object.assign({}, r, { id: key + String(seq).padStart(2, '0') });
    });
  }

  const api = { parseDate, cleanDrug, cleanSymptom, formatDate, monthKey, parseText, assignIds };
  if (typeof module !== 'undefined' && module.exports) module.exports = api;
  else root.ADRParse = api;
})(typeof window !== 'undefined' ? window : globalThis);
