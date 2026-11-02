// Finds TFDA drug licence numbers (許可證字號) for the drugs named on a report
// form, so TFDA's 快速搜尋 brings up the exact product. The licence list is a
// trimmed copy of TFDA's public open data (data/tfda-licenses.json, rebuilt
// with tools/build-licenses.js). Pure: window.ADRLicenses / module.exports.
(function (root) {
  'use strict';

  // Words that say nothing about which product it is.
  const STOP = new Set(('for tablet tablets tab tabs capsule capsules cap caps injection inj injections ' +
    'film coated fc f c solution sol oral powder lyophilized iv i v im i m sr er xr cr ml mg g gm mcg ' +
    'ug iu u unit units of and the in with w premixed infusion concentrate vial amp syrup').split(' '));

  const nfkc = s => String(s || '').normalize('NFKC');
  // Chinese names: drop quotes, brackets' contents kept, spacing and punctuation removed.
  const normZh = s => nfkc(s).toLowerCase().replace(/["'“”「」『』\s,，.。:：;；()（）\[\]【】\-－_/]/g, '');
  const tokensEn = s => new Set((nfkc(s).toLowerCase().match(/[a-z][a-z0-9-]*|\d+(?:\.\d+)?/g) || [])
    .map(t => t.replace(/-+$/, '')).filter(t => t && !STOP.has(t)));

  function bigrams(s) {
    const out = new Map();
    for (let i = 0; i < s.length - 1; i++) {
      const b = s.slice(i, i + 2);
      out.set(b, (out.get(b) || 0) + 1);
    }
    return out;
  }

  function dice(a, b) {
    if (!a || !b) return 0;
    if (a === b) return 1;
    const A = bigrams(a), B = bigrams(b);
    let common = 0;
    for (const [k, n] of A) common += Math.min(n, B.get(k) || 0);
    const total = Math.max(1, a.length - 1) + Math.max(1, b.length - 1);
    return (2 * common) / total;
  }

  function zhScore(a, b) {
    if (!a || !b) return 0;
    if (a === b) return 1;
    // The form often leaves out the strength: "摩舒胃清膜衣錠" vs "摩舒胃清膜衣錠5毫克".
    if (b.includes(a) || a.includes(b)) return 0.75 + 0.2 * Math.min(a.length, b.length) / Math.max(a.length, b.length);
    return dice(a, b) * 0.8;
  }

  function enScore(A, B) {
    if (!A.size || !B.size) return 0;
    let common = 0;
    for (const t of A) if (B.has(t)) common++;
    return (2 * common) / (A.size + B.size);
  }

  // rows: [[licence, 中文品名, 英文品名, 主成分, 劑型, 申請商], ...]
  function buildIndex(rows) {
    return rows.map(([license, zh, en, ingredient, form, company]) => ({
      license, zh, en, ingredient, form, company, nz: normZh(zh), te: tokensEn(en),
    }));
  }

  // Best matches for a drug named by its English and/or Chinese product name.
  function findLicenses(index, { en = '', zh = '' }, limit = 3) {
    const nz = normZh(zh), te = tokensEn(en);
    if (!nz && !te.size) return [];
    // The brand (first English word) must appear, unless the Chinese name matches.
    const brand = [...te][0];
    const scored = [];
    for (const r of index) {
      const z = zhScore(nz, r.nz);
      if (z < 0.5 && brand && !r.te.has(brand)) continue;
      const e = enScore(te, r.te);
      const score = nz && te.size ? Math.max(z, e) * 0.7 + Math.min(z, e) * 0.3 : Math.max(z, e);
      if (score >= 0.33) scored.push({ score, r });
    }
    scored.sort((a, b) => b.score - a.score);
    return scored.slice(0, limit).map(({ score, r }) => ({
      license: r.license, zh: r.zh, en: r.en, ingredient: r.ingredient, form: r.form, company: r.company,
      score: Math.round(score * 100) / 100,
    }));
  }

  // 健保碼 -> 許可證字號. Its first letter is A for a domestic product (藥製),
  // B for an imported one (藥輸), and characters 3-7 are the licence number:
  // AC57254… -> 衛署藥製字第057254號, BC26476… -> 衛部藥輸字第026476號. Whether it is
  // 衛署 or 衛部 comes from the licence list; ties go to the closer Chinese name.
  function licenseFromNhiCode(index, nhi, zh = '') {
    const m = String(nhi || '').trim().toUpperCase().match(/^([AB])[A-Z](\d{5})/);
    if (!m) return null;
    const kind = m[1] === 'A' ? '製' : '輸';
    const re = new RegExp(`藥${kind}字第0*${Number(m[2])}號$`);
    const hits = index.filter(r => re.test(r.license));
    if (!hits.length) return { license: `衛署藥${kind}字第0${m[2]}號`, guessed: true };
    const nz = normZh(zh);
    hits.sort((a, b) => zhScore(nz, b.nz) - zhScore(nz, a.nz));
    const h = hits[0];
    return { license: h.license, zh: h.zh, en: h.en, company: h.company, guessed: false };
  }

  const api = { buildIndex, findLicenses, licenseFromNhiCode, normZh, tokensEn };
  if (typeof module !== 'undefined' && module.exports) module.exports = api;
  else root.ADRLicenses = api;
})(typeof window !== 'undefined' ? window : globalThis);
