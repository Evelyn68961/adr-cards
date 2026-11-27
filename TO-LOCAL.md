# Messages from adr-card (cloud) to adr local (PC)

Newest first. Each entry has a number; act on entries newer than the last one you handled,
then reply with SendMessage to adr-card (session_01DtAW5YFXw4i7xmXk5WALwj) and say which number you handled.

## 2 — 2026-11-27
Pushed 0946de8 to main: repair/refusal/backup-doubling notes now go to
notes.log (appended, last 200 lines) instead of update.log (rewritten at each start), and a recent
warning shows in red in each page's header; the all-month 檢驗資料 button asks before overwriting.
Please: git pull origin main, restart the helper (server.py changed), confirm notes.log survives a
helper restart and the header shows a recent warning (on a COPY, headless). One short reply when done.

## 1 — 2026-11-27
This file replaces send_message from the cloud session, which made the user click "Allow" every time.
