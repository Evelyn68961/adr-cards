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
- **Updates itself:** each time the shortcut is opened, the helper first runs `git pull` to get
  the latest `main` from GitHub, so merged changes show up straight away. If `server.py` itself
  changed, the new version takes over from a copy that is already running. The update is
  skipped (and the app starts as before) when this folder has local edits, Git cannot be
  found, or there is no internet. Git is looked for on the PATH and in the usual Git for Windows,
  Cmder and GitHub Desktop folders. The page links show the version being served, and the
  reason in red when the last update was skipped or failed (also in `update.log`).
- **Guards ADR通報.xlsx against growing by itself:** a save that would make it more than 1.5×
  its size plus 200 KB, or over 5 MB, is refused (the old file is kept, the page says
  「存檔異常…」, `update.log` notes it), and a backup that more than doubles is noted in
  `update.log`. Two bugs grew the file save after save until the page crashed: width entries
  for all 16,384 columns, and the theme's East Asian font names being re-encoded on every save
  (each non-ASCII byte doubled; fixed in `workbook.js`, and undone in old files).
- **Repairs an oversized workbook:** an older version saved every month sheet with width entries
  for all 16,384 columns, uncompressed (~1 MB per sheet; a 30 MB `ADR通報.xlsx` crashed every save
  in the browser). On start (and when a page finds the file over 4 MB) the helper cuts those
  entries back to columns A–Z, repairs a garbled theme, and stores the file compressed, after
  copying it to `backups\`; cells are untouched. `update.log` says what it did.
  Saves from the page are compressed too.
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

## Where card images go

Card images (`ADR_<識別碼>.png`) and card PDFs can be kept in the month folder, or in an `ADR_cards`
folder inside each month folder (`ADR_11511\ADR_cards\`), or in a top `ADR_cards` folder (flat
or with month folders). Cards are found in any of these (已製卡, numbering, 卡片列印), and new ones
go where the cards already are, e.g. a new month gets its own `ADR_yyyMM\ADR_cards\`.

## Writing to the workbook

- Rows go to the sheet for their month (`202610`, `202611`, …), created if missing.
- A case already in the sheet (same 病歷號 + 藥碼 + 得知日期) is never added twice: saving again **overwrites** that row with the card's current values (hand fixes to
  藥品 / 症狀 / 通報人員 / 病歷號 included) and leaves identical rows untouched.
- Columns: **識別碼** | 病歷號 | 姓名 | 藥碼 | 藥名 | 嚴重度 | 通報來源 | 症狀 | 得知日期 | **撰寫人** |
  **值得注意** (ticked on the 案例總覽 page) | **通報者** | **通報者員編** (HIS 最後異動者 and 員編, from
  the HIS paste; kept when a case is saved again from rows without them) | **TFDA通報日期** (TFDA
  helper).
- **撰寫人** (who writes up the case) is typed or pasted into the boxes in the 本月案例清單 table
  (unsaved ones turn yellow), then saved to column J with **儲存撰寫人** (or Enter); an empty box
  clears it. The page says whether it worked; if not, the names stay in the boxes to try again. Saves only rewrite
  columns A–I, so a case keeps its 撰寫人 when overwritten, except when a different case takes
  over its 識別碼.
- With the desktop shortcut, the previous workbook is copied to `backups\` in this folder
  (latest 30) before each write; nothing extra is put in the ADR folder.
- If `ADR通報.xlsx` is open in Excel, Windows locks it: close it and save again.
- Card PNGs are always written as the card is now, replacing an older image of the same 識別碼.
- A typed 起始流水號 means "this case is that number": a new case given a 識別碼 that Excel already
  holds replaces it (the orange message names the case it replaces and has a **取代並存入**
  button that saves right away). Its row takes the place of the first row with that number, every other row
  with that number is removed, and its card image replaces the old one. The page warns which
  case will be replaced before you save.
- **Replacing an unreportable case:** click the 識別碼 on a card and type the number of the case
  to replace (Enter or click away). The card shows 指定識別碼 and keeps that number through
  re-syncs; clear it to go back to automatic numbering. On save the card's case takes over the
  number: its row replaces the old case's row (in any month's sheet) and its image replaces
  `ADR_<識別碼>.png`. This also works for a case already in Excel under another number: it keeps
  its own row (and 撰寫人), the old number is left free and its `ADR_<old>.png` is deleted. Other
  new cases pasted with it are numbered around the typed number. The page lists what will be
  replaced before you save.
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
  A 登記日 more than 31 days before 最後異動時間 is an old record edited again (e.g. 2016/09/21
  on a case edited 2026/11/17): it is ignored, 最後異動時間 is used, and the page says so. A copy
  saved earlier under that old date (its row, a month sheet left empty, its card image) is
  removed when the case is saved again.
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

- Reads every report-form export in the ADR folder and its `ADR_<yyyMM>` folders, whatever the
  file is called: an Excel file counts as a report form by its columns (識別代碼 and others), and
  the header row may start with 臨床訓練計畫 or not (e.g. a Google Sheets copy). Form versions with
  different column orders are handled; a case exported twice keeps the most complete copy.
- When the system's Excel export does not work, the report form printed to PDF from the fill-in
  page (ADR通報表_空白表單.html → 列印 → 另存為 PDF, one case per PDF) is read too, in the folders or
  uploaded. Typed answers come from the PDF's text; ticked choices (性別, 來源, 結果, 嚴重性,
  Naranjo …) are read by drawing each page and looking for the filled dot (pdfform.js, with
  pdf.js from cdn.jsdelivr.net, loaded only when there is a PDF to read). PDFs this app makes
  (`ADR_<識別碼>…`) are skipped. If a case is in both an Excel export and a PDF, the more complete
  copy is used.
- **一次產生本月全部檢驗資料 PDF（N 件）** makes the lab-results PDF of every case of the month,
  saved into the month folder; a case with no lab data gets one that says 「無」. A 檢查描述 too
  long for TFDA's 檢驗說明 (250) goes into the attachment, and 檢驗說明 points to it.
- The page shows one summary line for the files read (檔案明細 folds out); files that gave no
  cases are listed. Example and blank forms (範例, 空白) are skipped.
- **上傳通報表** picks exports from anywhere: they are read at once and, with the desktop helper,
  saved into the ADR folder (replacing a file of the same name) so the page keeps reading them.
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
- **已通報** also comes from the PDF saved after reporting: a case with `ADR_<識別碼>_通報.pdf` in its
  month folder (or its `ADR_cards`) counts as reported on the day that file was saved, with no
  tick needed; the meeting list uses that day as 通報日期 unless a date was ticked here. The
  first time the page sees such a PDF, its date is written into column N, so copying or
  restoring the PDFs later (which changes their file dates) doesn't change it.
- **已完成 TFDA 通報** records the day the case was reported to TFDA (today; the date box next to
  it can change it) in column N (TFDA通報日期) of `ADR通報.xlsx`; the case list shows 已通報 11/27,
  and the meeting list uses it as 通報日期. Unticking clears it.

## Case overview (`overview.html`)

For the routine ADR meeting: every case of a period at a glance, with the cases you choose to
report ticked.

- Reads `ADR通報.xlsx` and every `*通報表*.xls(x)` export in the ADR folder and its `ADR_<yyyMM>`
  folders (or picked files), and joins each case to its report form by 識別碼. A form whose
  識別碼 is not in the workbook is listed too (it cannot be ticked until it is in the workbook).
- Pick a period (從 / 到, or 本月, 近 3 個月, 近 12 個月). The top shows the case count, ticked
  cases, cases with hints that are not ticked yet, serious reactions, cases without a report
  form, and the most common drugs and symptoms.
- **值得注意** is your call: click a row (or its checkbox) to tick it, Shift-click to tick a
  whole range, or use **全部勾選** / **全部取消** for every row shown. Rows stay where they are
  while you tick. Ticks are saved to column **K (值得注意)** of `ADR通報.xlsx` as `V`, so everyone
  sees the same ones; anything typed into column K in Excel counts as ticked. Saving again from
  the main page keeps a case's tick, except when a different case takes over its 識別碼. If the
  workbook is open in Excel, ticks are kept on the page and saved when you come back after
  closing Excel. With picked files (no desktop helper) the ticked workbook is offered as a
  download to replace the original. **詳情** shows the form's 案件描述, 處置, 嚴重性, 結果, Naranjo,
  sex and age.
- **自動提示** only suggests cases to look at; it never ticks one by itself:
  - red: the form's 不良反應嚴重性 is serious (death, life-threatening, disability, birth defect,
    hospitalisation, other serious), or the outcome is death; without a form, HIS 嚴重度 3_嚴重.
  - yellow: outcome 有後遺症 or 尚未恢復, a positive rechallenge (有再投予且不良反應再發生), or
    Naranjo ≥ 9 / Definite.
  - green: the same 藥碼 in more than one case of the period, or the same patient more than once.
- Ticked cases are highlighted and, with cases that have hints, listed first.
- **複製表格** copies the rows shown (tab-separated, no patient names) for Excel or Word;
  **列印** prints the summary and table on landscape A4.

### 例會名單 (meeting list)

On 案例總覽, for the meeting every 4 months (e.g. 08–11):

1. Pick the months (**近 4 個月**), press **勾選 Hartwig ≥4** (the cases whose report form says
   嚴重不良反應【院內通報獎勵金要求】 4級 or more; 自動提示 shows them as "Hartwig n 級（獎勵）"),
   and tick or untick any other case.
2. **產生例會 Excel** → fill 審核者 / 員編 once (kept on this PC) → **存成 Excel**: the ticked cases
   go into `ADR例會名單_<yyyyMM>-<yyyyMM>.xlsx` in the ADR folder (and as a download), one row each:
   識別代碼 | ICU | 病歷號 | 通報者(200) | 員編 | 發生日期 | 通報日期 | 通報藥品 | 通報症狀 | 撰寫人 | 員編 |
   審核者 | 員編, then 單位主管：.
   - 通報者 / 員編: from the HIS paste (role from the 員編's letter: A 醫, B 護, C 藥); cases saved
     before this was recorded have them blank.
   - 發生日期, 通報症狀: report form; 通報日期: the day recorded with 已完成 TFDA 通報 (TFDA
     helper), else the form's 表單提交日期; 通報藥品: 藥碼.
   - 撰寫人 員編: looked up from the 通報者 names HIS recorded; blank if that person never keyed a case.
   - ICU is not recorded anywhere: blank, fill it in Excel.

Report forms are found by their columns (not the file name), and printed PDFs are included once
the TFDA helper or this page has read them (kept on this PC).

## Checking cases on a phone (`cases.html`)

A read-only list for the phone: one short row per case (tap for details), filters 全部 /
值得注意 / 有提示, search, newest month first.

- On the PC's 案例總覽, pick the period and press **手機查看**: scan the QR code with the phone's
  camera, or press **複製連結** and send the link to the phone (e.g. LINE).
- The cases travel inside the link after `#` (識別碼, date, drug, symptom, severity, outcome,
  撰寫人, ticks and hints; no names or 病歷號). Browsers never send that part to the website, so
  nothing is uploaded; the phone keeps the last list and shows it when the page is opened again.
  Press 手機查看 again for newer data.
- The QR image comes from the qrcode-generator library on cdn.jsdelivr.net, loaded only when
  手機查看 is pressed; without internet the link can still be copied.
- Without a link, the phone page can also read ADR通報.xlsx picked on the phone (last 3 months).
- Page: https://evelyn68961.github.io/adr-cards/cases.html

## Privacy

Everything runs in the browser; nothing is uploaded. The only files written are the ones in the
ADR folder you link. `.gitignore` blocks spreadsheets, PDFs and card images. Never commit real
case data to this repo.

## Development

No build step. `dropdown.js` draws every page's dropdown lists itself, because Chrome's own
list can crash the page (STATUS_BREAKPOINT) on some Windows PCs. `parse.js` (parsing), `workbook.js` (Excel append), `pdf.js` (PDF writer), `forms.js`
(report forms) and `review.js` (overview flags) are pure
and tested in Node; `app.js` is the page. `lib/xlsx.full.min.js` is SheetJS CE 0.20.3
(Apache-2.0, see `lib/xlsx.LICENSE`).

```
npm test
```
