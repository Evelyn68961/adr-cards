// Rebuilds data/tfda-licenses.json from TFDA's public drug-licence open data
// (全部藥品許可證資料集, InfoId 36):
//   https://data.fda.gov.tw/opendata/exportDataList.do?method=ExportData&InfoId=36&logType=3
// Unzip it, then: node tools/build-licenses.js path/to/36_3.json
// Keeps licences that are still valid, for finished products and biologics only.
const fs = require('fs');
const path = require('path');

const src = process.argv[2];
if (!src) throw new Error('usage: node tools/build-licenses.js 36_3.json');
const records = JSON.parse(fs.readFileSync(src, 'utf8').replace(/^﻿/, ''))
  .map(fields => Object.assign({}, ...fields));

const seen = new Set();
const rows = [];
for (const r of records) {
  if (r['註銷狀態'] || !/製|菌/.test(r['許可證種類'] || '') || seen.has(r['許可證字號'])) continue;
  seen.add(r['許可證字號']);
  rows.push([r['許可證字號'], r['中文品名'], r['英文品名'], r['主成分略述'], r['劑型'], r['申請商名稱']]
    .map(v => String(v || '').trim()));
}
rows.sort((a, b) => a[0].localeCompare(b[0]));
const out = path.join(__dirname, '..', 'data', 'tfda-licenses.json');
fs.mkdirSync(path.dirname(out), { recursive: true });
fs.writeFileSync(out, JSON.stringify({ source: 'TFDA open data InfoId 36', built: new Date().toISOString().slice(0, 10), rows }));
console.log(`${rows.length} licences -> ${out}`);
