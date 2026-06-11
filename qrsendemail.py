"""
qrsendemail.py — Send QR code images to attendees via Gmail.

Bug fixes:
1. REG_COL had trailing space 'Registration Number ' — now stripped, consistent with CSV.
2. Credentials hardcoded as empty strings — now use env vars or argparse (never commit creds).
3. No retry logic — added simple 3-attempt retry with backoff.
4. No batch progress / summary — added counters.
5. yagmail SMTP object created once but never closed — wrapped in context or explicit close.
6. Subject line generic — parameterised.
"""

from __future__ import annotations

import sys
if hasattr(sys.stdout, "reconfigure"):
    sys.stdout.reconfigure(encoding="utf-8")
if hasattr(sys.stderr, "reconfigure"):
    sys.stderr.reconfigure(encoding="utf-8")

import argparse
import os
import time

import pandas as pd


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(description="Send QR codes to event attendees.")
    parser.add_argument("--csv",      default=os.environ.get("CSV_PATH"),      help="Input CSV path")
    parser.add_argument("--qr-dir",   default=os.environ.get("QR_FOLDER", "qrcodes"), help="QR images folder")
    parser.add_argument("--sender",   default=os.environ.get("SENDER_EMAIL"),  help="Gmail address")
    parser.add_argument("--password", default=os.environ.get("APP_PASSWORD"),  help="Gmail app password")
    parser.add_argument("--subject",  default=os.environ.get("EMAIL_SUBJECT", "Your Event QR Code"),
                        help="Email subject line")
    parser.add_argument("--event",    default=os.environ.get("EVENT_NAME", "the Event"),
                        help="Event name used in email body")
    parser.add_argument("--retries",  type=int, default=3, help="Send attempts per recipient (default 3)")
    return parser.parse_args()


# ── Column names (consistent, no trailing spaces) ─────────────────────────────
EMAIL_COL = "Email Address"
NAME_COL  = "Name"
REG_COL   = "Registration Number"   # FIX: was "Registration Number " (trailing space)


def send_all(args: argparse.Namespace) -> None:
    import yagmail  # import here so missing dep gives a clear error

    if not args.csv:
        print("❌  No CSV path. Use --csv or set CSV_PATH.")
        sys.exit(1)
    if not args.sender or not args.password:
        print("❌  Sender credentials missing. Use --sender / --password or env vars.")
        sys.exit(1)
    if not os.path.exists(args.csv):
        print(f"❌  CSV not found: {args.csv}")
        sys.exit(1)

    df = pd.read_csv(args.csv)
    df.columns = df.columns.str.strip()

    # FIX: strip column names so trailing-space variants don't cause KeyError
    for col in [EMAIL_COL, NAME_COL, REG_COL]:
        if col not in df.columns:
            # Try stripping trailing spaces from CSV column headers
            stripped = {c.strip(): c for c in df.columns}
            if col in stripped:
                df.rename(columns={stripped[col]: col}, inplace=True)
            else:
                print(f"❌  Missing column: '{col}'")
                sys.exit(1)

    yag = yagmail.SMTP(user=args.sender, password=args.password)

    sent = skipped = failed = 0

    for _, row in df.iterrows():
        recipient = str(row[EMAIL_COL]).strip()
        name      = str(row[NAME_COL]).strip()
        reg       = str(row[REG_COL]).strip().upper()
        qr_file   = os.path.join(args.qr_dir, f"{reg}.png")

        if not os.path.exists(qr_file):
            print(f"⚠️   QR not found for {name} ({reg}) — skipped.")
            skipped += 1
            continue

        body = (
            f"Hi {name},\n\n"
            f"Please find your QR Code attached. Present it at the entrance of {args.event}.\n\n"
            f"Do not share this QR code with others — it is unique to your registration.\n\n"
            f"Regards,\nEvent Team"
        )

        success = False
        for attempt in range(1, args.retries + 1):
            try:
                yag.send(
                    to=recipient,
                    subject=args.subject,
                    contents=body,
                    attachments=qr_file,
                )
                print(f"✅  Sent → {recipient}")
                sent += 1
                success = True
                break
            except Exception as e:
                wait = 2 ** attempt
                print(f"   Attempt {attempt} failed for {recipient}: {e}. Retrying in {wait}s…")
                time.sleep(wait)

        if not success:
            print(f"❌  Failed → {recipient} after {args.retries} attempts.")
            failed += 1

    yag.close()

    print(f"\n── Summary ──────────────────")
    print(f"   Sent:    {sent}")
    print(f"   Skipped: {skipped}  (QR image missing)")
    print(f"   Failed:  {failed}  (send error)")
    print(f"─────────────────────────────")


if __name__ == "__main__":
    send_all(parse_args())
