import pandas as pd
import os
import yagmail

# === Config ===
CSV_PATH = ""# Your CSV Path
QR_FOLDER = "qrcodes"
SENDER_EMAIL = ""            # Your Gmail address
APP_PASSWORD = ""          # Gmail app password


# === Columns ===
EMAIL_COL = 'Email Address'
NAME_COL = 'Name'
REG_COL = 'Registration Number '

# Load CSV
df = pd.read_csv(CSV_PATH)

# Initialize email client
yag = yagmail.SMTP(user=SENDER_EMAIL, password=APP_PASSWORD)

# Send emails
for _, row in df.iterrows():
    recipient = str(row[EMAIL_COL]).strip()
    name = str(row[NAME_COL]).strip()
    reg = str(row[REG_COL]).strip().upper()  # Match file name

    qr_file = os.path.join(QR_FOLDER, f"{reg}.png")
    
    if os.path.exists(qr_file):
        try:
            yag.send(
                to=recipient,
                subject="EVENT: QR Code",
                contents=f"Hi {name},\n\nPlease find your QR Code attached below.\n\nRegards,\nAshin",
                attachments=qr_file
            )
            print(f"✅ Sent to {recipient}")
        except Exception as e:
            print(f"❌ Failed to send to {recipient}: {e}")
    else:
        print(f"❌ QR code not found for {name} ({reg}) - Skipping.")
