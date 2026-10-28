// Parsing and formatting for ADR cards. Pure functions only, so the same file
// runs in the browser (window.ADRParse) and in Node tests (module.exports).
(function (root) {
  'use strict';

  // Column layouts. With a header row, columns are found by name (aliases
  // cover the HIS export and the monthly Excel sheet); without one, the
  // layout is picked by column count. `date` lists candidates in priority
  // order: HIS leaves 登記日 empty, and the sheet's 得知日期 is 最後異動時間.
  const HEADER_ALIASES = {
    chartNo: ['病歷號'],
    name: ['姓名'],
    drugCode: ['藥碼'],
    drug: ['MDPNAME', '藥名', '藥品'],
    severity: ['嚴重程度', '嚴重度'],
    source: ['記錄來源', '通報來源', '通報人員'],
    symptom: ['症狀'],
    date: ['得知日期', '最後異動時間', '登記日'],
  };
  const LAYOUTS = {
    // HIS export: 病歷號 姓名 藥碼 MDPNAME ATC 藥理分類 藥理分類名稱 嚴重程度 記錄來源
    //             是否過敏 登記日 症狀 疫苗接種日期 疫苗接種院所 最後異動者員編 最後異動者 最後異動時間
    17: { chartNo: 0, name: 1, drugCode: 2, drug: 3, severity: 7, source: 8, symptom: 11, date: [16, 10] },
    // Monthly Excel sheet: 病歷號 姓名 藥碼 藥名 嚴重度 通報來源 症狀 得知日期
    8: { chartNo: 0, name: 1, drugCode: 2, drug: 3, severity: 4, source: 5, symptom: 6, date: [7] },
  };
  const CHART_NO = /^\d{9}[A-Z]$/i;

  function headerLayout(cells) {
    const norm = cells.map(c => c.replace(/\s/g, '').toUpperCase());
    const find = name => norm.indexOf(name.toUpperCase());
    const layout = {};
    for (const [key, names] of Object.entries(HEADER_ALIASES)) {
      const found = names.map(find).filter(i => i >= 0);
      if (found.length) layout[key] = key === 'date' ? found : found[0];
    }
    return 'chartNo' in layout && 'date' in layout ? { cols: layout, width: cells.length } : null;
  }

  // Accepts the formats the date column shows up in:
  //   2026/10/5 9:57[:01], 2026-10-05 09:57   (Gregorian Y/M/D)
  //   115/10/05 09:57                          (ROC Y/M/D)
  //   10/5/26 9:57                             (Excel "m/d/yy h:mm" display)
  function parseDate(text) {
    const m = String(text || '').trim().match(
      /^(\d{1,4})[\/\-.](\d{1,2})[\/\-.](\d{1,4})(?:[ T]+(\d{1,2}):(\d{2})(?::(\d{2}))?)?/);
    if (!m) return null;
    const [a, b, c] = [m[1], m[2], m[3]].map(Number);
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

  function formatDateTime(date) {
    const pad = n => String(n).padStart(2, '0');
    return `${formatDate(date)} ${pad(date.getHours())}:${pad(date.getMinutes())}:${pad(date.getSeconds())}`;
  }

  // 識別碼 prefix: ROC year + 2-digit month, e.g. 2026/10 -> "11510".
  function monthKey(date) {
    return String(date.getFullYear() - 1911) + String(date.getMonth() + 1).padStart(2, '0');
  }

  // Monthly sheet name in the Excel workbook, e.g. 2026/10 -> "202610".
  function sheetName(date) {
    return String(date.getFullYear()) + String(date.getMonth() + 1).padStart(2, '0');
  }

  // HIS puts line breaks inside cells (e.g. 記錄來源), so one record can span
  // several lines. A record starts at a line whose first cell is a chart
  // number or a header; any other line continues the previous record.
  function splitRecords(text) {
    const records = [];
    String(text || '').split(/\r?\n/).forEach((line, i) => {
      const first = line.split('\t')[0].trim();
      if (CHART_NO.test(first) || first === '病歷號') {
        records.push({ line: i + 1, text: line });
      } else if (records.length && line.trim()) {
        records[records.length - 1].text += '\n' + line;
      }
    });
    return records;
  }

  // Splits a record into cells. A line break inside a record sometimes stands
  // in for a lost tab, so while there are too few cells, the first cell with
  // an inner line break is split at its last break.
  function splitCells(text, width) {
    const cells = text.split('\t');
    while (width && cells.length < width) {
      const i = cells.findIndex(c => c.trim().includes('\n'));
      if (i < 0) break;
      const c = cells[i].trim();
      const cut = c.lastIndexOf('\n');
      cells.splice(i, 1, c.slice(0, cut), c.slice(cut + 1));
    }
    return cells.map(c => c.replace(/\s*\n\s*/g, ' ').trim());
  }

  // Returns { rows, skipped }. Rows keep every field the Excel sheet needs;
  // the card only shows a subset. Records that have a chart number but fail
  // to parse are reported in `skipped`.
  function parseText(text) {
    const rows = [];
    const skipped = [];
    let header = null;
    splitRecords(text).forEach(rec => {
      if (rec.text.split('\t')[0].trim() === '病歷號') {
        header = headerLayout(splitCells(rec.text));
        return;
      }
      const tabs = rec.text.split('\t').length;
      const width = header ? header.width : tabs >= 12 || rec.text.includes('\n') ? 17 : 8;
      const cells = splitCells(rec.text, width);
      const cols = header ? header.cols : LAYOUTS[width];
      const get = key => (key in cols ? cells[cols[key]] || '' : '');
      const date = cols.date.map(i => parseDate(cells[i])).find(Boolean);
      if (!date) { skipped.push({ line: rec.line, reason: '無法辨識得知日期' }); return; }
      rows.push({
        chartNo: get('chartNo').toUpperCase(),
        name: get('name'),
        drugCode: get('drugCode'),
        drugRaw: get('drug'),
        severity: get('severity'),
        source: get('source'),
        symptomRaw: get('symptom'),
        date,
        drug: cleanDrug(get('drug')),
        symptom: cleanSymptom(get('symptom')),
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

  // One row of the monthly Excel sheet (raw HIS text, as the sheet stores it).
  // 識別碼 is the last column so the original A-H layout is unchanged.
  const SHEET_HEADER = ['病歷號', '姓名', '藥碼', '藥名', '嚴重度', '通報來源', '症狀', '得知日期', '識別碼'];
  function sheetRow(r) {
    return [r.chartNo, r.name, r.drugCode, r.drugRaw, r.severity, r.source, r.symptomRaw, r.date, r.id || ''];
  }

  // Identifies a report across the paste and the workbook, to avoid
  // appending the same row twice.
  function rowKey(chartNo, drugCode, date) {
    return [String(chartNo).trim().toUpperCase(), String(drugCode).trim().toUpperCase(),
      date instanceof Date ? formatDateTime(date) : ''].join('|');
  }

  // Tab-separated rows (no header) for pasting under an existing month's sheet.
  function toTsv(rows) {
    return rows.map(r => sheetRow(r)
      .map(v => (v instanceof Date ? formatDateTime(v) : String(v).replace(/[\t\r\n]+/g, ' ')))
      .join('\t')).join('\r\n');
  }

  const api = {
    parseDate, cleanDrug, cleanSymptom, formatDate, formatDateTime, monthKey, sheetName,
    parseText, assignIds, SHEET_HEADER, sheetRow, rowKey, toTsv,
  };
  if (typeof module !== 'undefined' && module.exports) module.exports = api;
  else root.ADRParse = api;
})(typeof window !== 'undefined' ? window : globalThis);
