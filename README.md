# ADR Cards

Turns the HIS adverse-drug-reaction (ADR) list into printable case cards that are handed to
pharmacists to write up the ADR report.

**Live:** https://evelyn68961.github.io/adr-cards/ (also works by opening `index.html` from disk)

## Privacy

Everything runs in the browser. Pasted data is never uploaded, stored, or written to this repo.
Patient name, drug code and severity are dropped during parsing and never reach the card.
`.gitignore` blocks spreadsheets and exported card images. Do not commit real case data.

## Usage

1. Copy the ADR rows from HIS (or the monthly Excel sheet) and paste them. A header row is fine.
2. Set **起始流水號** to the next sequence number for the month (e.g. `6` if 01–05 are done).
3. Check the cards; click any value to fix it, untick a card to leave it out.
4. **列印** prints 4 cards per landscape A4 page. **下載 PNG** saves `ADR_<識別碼>.png` files
   in the same style as the archived cards.

## Input format

Tab-separated, in HIS column order:

| 病歷號 | 姓名 | 藥碼 | 藥名 | 嚴重度 | 通報來源 | 症狀 | 得知日期 |
|---|---|---|---|---|---|---|---|

Date formats accepted: `10/5/26 9:57` (Excel display), `2026/10/5 09:57`, `2026-10-05`,
`115/10/05` (ROC).

## Card rules

- **識別碼** = ROC year + 2-digit month + 2-digit sequence, e.g. `1151006`. Cards are numbered
  in 得知日期 date-time order, not paste order. If the batch spans months, the start number
  applies to the earliest month and later months restart at 01.
- **藥品**: HIS tags in `【…】` and markers like `★` are removed
  (`Cefin 針【禁用含鈣溶液】2 g/Vial` → `Cefin 針 2 g/Vial`).
- **症狀**: commas become `、`.
- **通報人員**: the 通報來源 column as-is.

## Development

No build step: `index.html`, `style.css`, `app.js`, and `parse.js` (pure parsing logic).

```
npm test
```
