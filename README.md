# ADR Cards

Turns the HIS adverse-drug-reaction (ADR) export into ADR case cards for pharmacists to write up,
and records the cases in the monthly ADR workbook.

**Live:** https://evelyn68961.github.io/adr-cards/ (also works by opening `index.html` from disk)

## Desktop shortcut (recommended)

`server.py` (standard-library Python) serves the tool at **http://127.0.0.1:8770/** with the ADR
folder built in, so nothing ever has to be linked or allowed.

- `adr-folder.txt` (one line, not committed) holds the ADR folder path, e.g.
  `G:\我的雲端硬碟\#FJUH.Pharm\^藥品資訊\ADR`.
- Desktop and Start-menu shortcut **ADR 通報卡** (icon `icons\adr.ico`): target
  `C:\Windows\pyw.exe "D:\Git\adr-cards\server.py"`, start in `D:\Git\adr-cards`. It starts the
  helper (no console window) and opens the tool in its own app window (Chrome or Edge `--app`,
  no tabs or address bar); if the helper is already running it just opens the window.
- The pages are an installable app (`manifest.json`): Chrome's install icon in the address bar
  adds it to the Start menu and taskbar, with 卡片列印, TFDA 通報小幫手 and 案例總覽 on its right-click
  menu. The installed app still needs the helper running.
- Optional: to have the helper always running, put a shortcut to
  `C:\Windows\pyw.exe "D:\Git\adr-cards\server.py" --no-browser` in the Windows Startup folder
  (`shell:startup`).
- The helper only listens on 127.0.0.1, only answers requests carrying the page's own header,
  and only touches files inside the ADR folder.

## Usage

1. **連結 ADR 資料夾** (first time only; Chrome or Edge) and pick the `ADR` folder that holds
   `ADR通報.xlsx`. The browser remembers it and only asks for permission again once per session.
2. Copy the rows from the HIS ADR list (header row included is best) and paste.
3. The **目前累積** panel reads `ADR通報.xlsx` and the month folder: case count, latest and next
   識別碼, cases without a card, and the month's list. It refreshes on open, after each save, and
   when the page regains focus. Numbers follow it automatically; pasted cases already in Excel
   keep their 識別碼. **起始流水號** can still be overridden.
4. Tick the cards to **print** (全選 / 全不選); click any value on a card to fix it.
5. **存入 Excel 與 PNG（全部）** appends **every** pasted case (ticked or not — the workbook is the
   continuous record) to `ADR通報.xlsx` and saves `ADR_<識別碼>.png`
   into `ADR_<yyyMM>\`.
6. **產生 PDF** opens a PDF of the ticked cards, 4 per landscape A4 page, ready to print. With a
   linked folder it is also saved as `ADR_<yyyMM>\ADR_<first>-<last>.pdf`.

Without a linked folder (or in Firefox), **下載 PNG** and **產生 PDF** still work as downloads.

## Writing to the workbook

- Rows go to the sheet for their month (`202610`, `202611`, …), created if missing.
- A case already in the sheet (same 病歷號 + 藥碼 + 得知日期) is never added twice: saving again **overwrites** that row with the card's current values (hand fixes to
  藥品 / 症狀 / 通報人員 / 病歷號 included) and leaves identical rows untouched.
- Columns: **識別碼** | 病歷號 | 姓名 | 藥碼 | 藥名 | 嚴重度 | 通報來源 | 症狀 | 得知日期 | **撰寫人** |
  **值得注意** (marked on the 案例總覽 page).
- **撰寫人** (who writes up the case) is typed in the 本月案例清單 table (names used before are
  suggested) and saved to column J automatically; clearing the box clears it. Saves only rewrite
  columns A–I, so a case keeps its 撰寫人 when overwritten, except when a different case takes
  over its 識別碼. If the workbook is open in Excel, the name is kept on the page and saved when
  you come back to it after closing Excel.
- With the desktop shortcut, the previous workbook is copied to `backups\` in this folder
  (latest 30) before each write; nothing extra is put in the ADR folder.
- If `ADR通報.xlsx` is open in Excel, Windows locks it: close it and save again.
- Card PNGs are always written as the card is now, replacing an older image of the same 識別碼.
- A typed 起始流水號 means "this case is that number": a new case given a 識別碼 that Excel already
  holds replaces it. Its row takes the place of the first row with that number, every other row
  with that number is removed, and its card image replaces the old one. The page warns which
  case will be replaced before you save.
- If two Excel rows carry the same 識別碼 (left over from older versions), the first keeps it and
  the later one gets a new number the next time it is pasted and saved without a typed number.
- Rewriting uses SheetJS, which keeps values, number formats and column widths but may drop
  cell styling such as fonts and fills.

## Input format

The HIS ADR export (17 columns, tab-separated):

`病歷號 姓名 藥碼 MDPNAME ATC 藥理分類 藥理分類名稱 嚴重程度 記錄來源 是否過敏 登記日 症狀 疫苗接種日期 疫苗接種院所 最後異動者員編 最後異動者 最後異動時間`

- With a header row, columns are found by name; without one, the layout is detected by shape.
  Rows copied back out of the monthly sheet (8 columns) also work.
- HIS puts line breaks inside cells (e.g. 記錄來源), so a record can span several lines; a new
  record starts only at a line beginning with a 病歷號.
- **得知日期** = 登記日 when HIS filled it in, otherwise 最後異動時間 (登記日 is often empty).
- HIS lists one row per suspected drug. Rows with the same 病歷號, 得知日期 and 症狀 are one case:
  paste them all and they become one card, one 識別碼 and one Excel row, with the 藥碼 and 藥名 joined
  by `、` (e.g. `ISO06、RIF04、PYR03`). Pasting the case again in any order is recognised.

## Card rules

- **識別碼** = ROC year + 2-digit month + 2-digit sequence, e.g. `1151006`, numbered in 得知日期
  date-time order. If the batch spans months, the start number applies to the earliest month and
  later months restart at 01.
- **藥品**: `【…】` tags and markers like `★` removed (`Cefin 針【禁用含鈣溶液】2 g/Vial` →
  `Cefin 針 2 g/Vial`). The workbook keeps the original HIS text.
- **症狀**: commas become `、`.
- **通報人員**: the 記錄來源 column as-is.

## Card printing (`print.html`)

For printing cards that are already saved, in any combination.

- Shows the `ADR_<識別碼>.png` images in each `ADR_<yyyMM>` folder, one month at a time (or
  picked files, without the desktop helper).
- Click cards to select them; the selection is kept when you switch months. **本月全選** /
  **全不選**.
- **產生 PDF** opens a PDF of the selected cards in 識別碼 order, 4 per landscape A4 page (same
  layout as the main page), ready to print. Nothing is written to the ADR folder; use the link
  under the button to download it.

## TFDA helper (`tfda.html`)

For entering finished report forms into the TFDA online ADR form.

- Reads every `*通報表*.xls(x)` export in the ADR folder and its `ADR_<yyyMM>` folders (or a picked
  file). Form versions with different column orders are handled; a case exported twice keeps the
  most complete copy.
- Pick a month and a case. The case is laid out in the TFDA tabs (General, Patient, Reaction 1/2,
  Drug, Upload). Click any value to copy it; choices show the option to click on TFDA, mapped from
  the internal wording (e.g. 民眾 → 民眾或其他非醫療人員, 其他嚴重不良反應 →
  Other Medically Important Condition).
- Lab results are not typed in: **檢驗資料 PDF** writes `ADR_<識別碼>_檢驗資料.pdf` (A4, the
  case's 相關檢查及檢驗數據 and 檢查描述) into the month folder to upload under Upload Attachment,
  and 檢驗說明 points to it.
- Drugs follow TFDA's 新增藥品 dialog. 快速搜尋 gets the product's 許可證字號 (best match first,
  plus up to two alternatives), found by product name in `data/tfda-licenses.json`, a trimmed
  copy of TFDA's public licence open data (valid finished-product licences). Refresh it with
  `node tools/build-licenses.js 36_3.json` (instructions in the script).
- **已完成 TFDA 通報** marks a case as done in this browser only.

## Case overview (`overview.html`)

For the routine ADR meeting: every case of a period at a glance, with the cases you choose to
report marked.

- Reads `ADR通報.xlsx` and every `*通報表*.xls(x)` export in the ADR folder and its `ADR_<yyyMM>`
  folders (or picked files), and joins each case to its report form by 識別碼. A form whose
  識別碼 is not in the workbook is listed too (it cannot be marked until it is in the workbook).
- Pick a period (從 / 到, or 本月, 近 3 個月, 近 12 個月). The top shows the case count, marked
  cases, cases with hints that are not marked yet, serious reactions, cases without a report
  form, and the most common drugs and symptoms.
- **值得注意** is your call: **☆ 標記** marks a case and opens a box for the reason (optional,
  e.g. 新藥、教學案例). Marks are saved to column **K (值得注意)** of `ADR通報.xlsx`: the reason, or
  `V` without one, so everyone sees the same marks, and typing into column K in Excel works too.
  Saving again from the main page keeps a case's mark, except when a different case takes over
  its 識別碼. If the workbook is open in Excel, the mark is kept on the page and saved when you
  come back after closing Excel. With picked files (no desktop helper) the marked workbook is
  offered as a download to replace the original.
- **自動提示** only suggests cases to look at; it never marks one by itself:
  - red: the form's 不良反應嚴重性 is serious (death, life-threatening, disability, birth defect,
    hospitalisation, other serious), or the outcome is death; without a form, HIS 嚴重度 3_嚴重.
  - yellow: outcome 有後遺症 or 尚未恢復, a positive rechallenge (有再投予且不良反應再發生), or
    Naranjo ≥ 9 / Definite.
  - green: the same 藥碼 in more than one case of the period, or the same patient more than once.
- Marked cases are highlighted and listed first. Click a row for the form's 案件描述, 處置,
  嚴重性, 結果, Naranjo, sex and age.
- **複製表格** copies the rows shown (tab-separated, no patient names) for Excel or Word;
  **列印** prints the summary and table on landscape A4.

## Privacy

Everything runs in the browser; nothing is uploaded. The only files written are the ones in the
ADR folder you link. `.gitignore` blocks spreadsheets, PDFs and card images. Never commit real
case data to this repo.

## Development

No build step. `parse.js` (parsing), `workbook.js` (Excel append), `pdf.js` (PDF writer), `forms.js`
(report forms) and `review.js` (overview flags) are pure
and tested in Node; `app.js` is the page. `lib/xlsx.full.min.js` is SheetJS CE 0.20.3
(Apache-2.0, see `lib/xlsx.LICENSE`).

```
npm test
```
