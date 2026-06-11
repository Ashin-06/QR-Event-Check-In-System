# QR Event Check-In System — Upgraded Edition

A production-ready Flask + WebSocket system for scanning attendee QR codes at events,
with real-time log updates, Excel highlighting, and bulk email delivery.

---

## Quick Start

```bash
pip install -r requirements.txt

# Set your Excel file path (or place it next to app.py)
export EXCEL_FILE=/path/to/data_with_qrdemo.xlsx

python app.py
# Open http://localhost:5001
```

---

## File Structure

```
qr_checkin_system/
├── app.py             # Flask app — scan endpoint, real-time socket
├── qradding.py        # Generate QR codes → embed in Excel
├── qrsendemail.py     # Bulk email QR codes to attendees
├── requirements.txt
├── templates/
│   └── index.html     # Dark glassmorphic scanner UI
└── static/
    └── js/
        ├── qr-scanner.min.js
        └── qr-scanner-worker.min.js
```

---

## Environment Variables

| Variable        | Default              | Description                         |
|-----------------|----------------------|-------------------------------------|
| `EXCEL_FILE`    | `data_with_qrdemo.xlsx` | Path to master Excel file         |
| `SECRET_KEY`    | random               | Flask session secret                |
| `CORS_ORIGIN`   | `*`                  | Allowed CORS origin for SocketIO    |
| `PORT`          | `5001`               | Server port                         |
| `FLASK_DEBUG`   | `false`              | Enable Flask debug mode             |

---

## QR Generation

```bash
python qradding.py \
  --csv attendees.csv \
  --output data_with_qrdemo.xlsx \
  --qr-dir qrcodes_samples \
  --id-length 8
```

CSV must have columns: `Name`, `Email Address`, `Registration Number`

---

## Bulk Email

```bash
# Use env vars (recommended — never commit credentials)
export SENDER_EMAIL=you@gmail.com
export APP_PASSWORD=your_app_password
export CSV_PATH=attendees.csv
export QR_FOLDER=qrcodes_samples
export EVENT_NAME="TechConf 2025"
export EMAIL_SUBJECT="Your TechConf QR Code"

python qrsendemail.py
```

Or use flags: `python qrsendemail.py --sender you@gmail.com --password xxx --csv ... --event ...`

---

## Bugs Fixed (Original → Fixed)

| # | File | Bug | Fix |
|---|------|-----|-----|
| 1 | app.py | `scanned_log.loc[len(df)]` index collision when df has gaps | Replaced with `pd.concat()` |
| 2 | app.py | QR match used substring (`name in qr_data`) — false positives | Exact match on `qr` column only |
| 3 | app.py | `/data_csv` read shared DataFrame without lock | Added `with lock` |
| 4 | app.py | `sort_log` mutated shared df without lock | Called inside lock |
| 5 | app.py | No Content-Type check on `/scan` | Returns 415 if not JSON |
| 6 | app.py | No rate limiting — scanner can fire twice per QR | Per-IP 2s cooldown |
| 7 | app.py | Excel highlight rebuild on main thread (blocks) | Moved to daemon thread |
| 8 | app.py | No startup validation of EXCEL_FILE | `FileNotFoundError` with clear message |
| 9 | app.py | No Flask SECRET_KEY | Added from env or random |
| 10 | app.py | `Timestamps` NaN dtype mixed with str | Normalised on load |
| 11 | index.html | `processing = false` only in socket callback — locks if socket drops | Added 8s timeout fallback |
| 12 | index.html | `result` passed as string but new qr-scanner returns `{data}` object | Use `result.data ?? result` |
| 13 | index.html | Old DataTables CDN (1.13.5) | Updated to 1.13.6 |
| 14 | qradding.py | `csv_path = ` and `final_excel = ` were empty (syntax/runtime error) | argparse + env vars |
| 15 | qradding.py | Column letter formula broken for col > 26 | `openpyxl.utils.get_column_letter()` |
| 16 | qradding.py | No row height set — QR images overlapped rows | `ws.row_dimensions[r].height` set |
| 17 | qrsendemail.py | `REG_COL = 'Registration Number '` (trailing space) | Stripped to `'Registration Number'` |
| 18 | qrsendemail.py | Credentials hardcoded as empty strings | argparse + env vars |
| 19 | qrsendemail.py | No retry on send failure | 3-attempt exponential backoff |
| 20 | qrsendemail.py | No send summary | Printed counters at end |

---

## Excel Column Requirements

The master Excel file must have these headers (case-insensitive, trimmed):
- `Name`
- `Email Address`
- `Registration Number`
- `QR` — the raw QR string (added by `qradding.py`)

The app auto-creates `Scanned Status` if missing.
