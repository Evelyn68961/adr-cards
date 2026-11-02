// Reads pages of the hospital's drug information system (藥品關鍵字查詢,
// reached through server.py's /api/hospital). Pure string parsing, so it is
// testable in Node: window.ADRHospital / module.exports.
(function (root) {
  'use strict';

  // Address of a drug's own page, for opening it from the TFDA helper.
  const PAGE_URL = 'http://10.10.4.25:8080/Key_Search/P4?query=';

  const text = html => String(html || '')
    .replace(/<!--[\s\S]*?-->/g, '')
    .replace(/<[^>]+>/g, ' ')
    .replace(/&nbsp;/g, ' ').replace(/&amp;/g, '&').replace(/&quot;/g, '"').replace(/&#39;/g, "'")
    .replace(/&lt;/g, '<').replace(/&gt;/g, '>')
    .replace(/\s+/g, ' ').trim();

  // Search results (P3): one link per drug.
  function parseSearch(html) {
    const out = [];
    const re = /<a href="\/Key_Search\/P4\?query=([^"]+)">([\s\S]*?)<\/a>/g;
    let m;
    while ((m = re.exec(String(html)))) {
      const labels = {};
      const lr = /<div class="ImgLabelP3">\s*<div class="label"><span>([^<]+)<\/span><\/div>([\s\S]*?)<\/div>\s*(?=<div class="ImgLabelP3"|<\/div>)/g;
      let l;
      while ((l = lr.exec(m[2]))) labels[l[1].trim()] = text(l[2]);
      out.push({ code: decodeURIComponent(m[1]), generic: labels['藥品學名'] || '', en: labels['英文商品名'] || '', zh: labels['中文商品名'] || '' });
    }
    return out;
  }

  // Drug page (P4): label/value pairs. Returns null when the code is unknown.
  function parseDetail(html) {
    const f = {};
    const re = /<div class="MedLabelText">([^<]+)<\/div>[\s\S]*?<div class="MedInputText">([\s\S]*?)<\/div>/g;
    let m;
    while ((m = re.exec(String(html)))) {
      const k = m[1].trim();
      if (!(k in f)) f[k] = text(m[2]);
    }
    if (!f['藥碼']) return null;
    return { code: f['藥碼'], generic: f['藥名'] || '', en: f['英文商品名'] || '', zh: f['中文商品名'] || '',
      form: f['劑型'] || '', spec: f['規格'] || '', nhi: f['健保碼'] || '' };
  }

  const api = { parseSearch, parseDetail, PAGE_URL };
  if (typeof module !== 'undefined' && module.exports) module.exports = api;
  else root.ADRHospital = api;
})(typeof window !== 'undefined' ? window : globalThis);
