# ADR Cards

Turns the HIS adverse-drug-reaction (ADR) export into ADR case cards for pharmacists to write up,
and records the cases in the monthly ADR workbook.

**Live:** https://evelyn68961.github.io/adr-cards/ (also works by opening `index.html` from disk)

## Usage

1. **連結 ADR 資料夾** (first time only; Chrome or Edge) and pick the `ADR` folder that holds
   `ADR通報.xlsx`. The browser remembers it and only asks for permission again once per session.
2. Copy the rows from the HIS ADR list (header row included is best) and paste.
3. **起始流水號** is filled in automatically from the cards already in `ADR_<yyyMM>\`
   (e.g. 8 cases in sheet `202610` → starts at 09). It can be overridden.
4. Tick the cards to **print** (全選 / 全不選); click any value on a card to fix it.
5. **存入 Excel 與 PNG（全部）** appends **every** pasted case (ticked or not — the workbook is the
   continuous record) to `ADR通報.xlsx` and saves `ADR_<識別碼>.png`
   into `ADR_<yyyMM>\`.
6. **產生 PDF** opens a PDF of the ticked cards, 4 per landscape A4 page, ready to print. With a
   linked folder it is also saved as `ADR_<yyyMM>\ADR_<first>-<last>.pdf`.

Without a linked folder (or in Firefox), **下載 PNG** and **產生 PDF** still work as downloads.

## Writing to the workbook

- Rows go to the sheet for their month (`202610`, `202611`, …), created if missing.
- A row already in the sheet (same 病歷號 + 藥碼 + 得知日期) is never added twice.
- Before each write, the previous file is kept as `ADR通報.xlsx.bak`.
- If `ADR通報.xlsx` is open in Excel, Windows locks it: close it and save again.
- Existing card PNGs are never overwritten; a clash means the start number is wrong.
- Rewriting uses SheetJS, which keeps values, number formats and column widths but may drop
  cell styling such as fonts and fills.

## Input format

The HIS ADR export (17 columns, tab-separated):

`病歷號 姓名 藥碼 MDPNAME ATC 藥理分類 藥理分類名稱 嚴重程度 記錄來源 是否過敏 登記日 症狀 疫苗接種日期 疫苗接種院所 最後異動者員編 最後異動者 最後異動時間`

- With a header row, columns are found by name; without one, the layout is detected by shape.
  Rows copied back out of the monthly sheet (8 columns) also work.
- HIS puts line breaks inside cells (e.g. 記錄來源), so a record can span several lines; a new
  record starts only at a line beginning with a 病歷號.
- **得知日期** = 最後異動時間 (HIS leaves 登記日 empty), falling back to 登記日.

## Card rules

- **識別碼** = ROC year + 2-digit month + 2-digit sequence, e.g. `1151006`, numbered in 得知日期
  date-time order. If the batch spans months, the start number applies to the earliest month and
  later months restart at 01.
- **藥品**: `【…】` tags and markers like `★` removed (`Cefin 針【禁用含鈣溶液】2 g/Vial` →
  `Cefin 針 2 g/Vial`). The workbook keeps the original HIS text.
- **症狀**: commas become `、`.
- **通報人員**: the 記錄來源 column as-is.

## Privacy

Everything runs in the browser; nothing is uploaded. The only files written are the ones in the
ADR folder you link. `.gitignore` blocks spreadsheets, PDFs and card images. Never commit real
case data to this repo.

## Development

No build step. `parse.js` (parsing), `workbook.js` (Excel append), `pdf.js` (PDF writer) are pure
and tested in Node; `app.js` is the page. `lib/xlsx.full.min.js` is SheetJS CE 0.20.3
(Apache-2.0, see `lib/xlsx.LICENSE`).

```
npm test
```
