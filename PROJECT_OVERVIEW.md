# Upgraded QR Event Check-In System

An advanced, enterprise-grade QR code and barcode check-in system built on Flask, Socket.IO, openpyxl, and pandas. This application features rate limiting, strict phone validation, administrative undo controls, system diagnostics, subgroup notifications, a quarantine queue, and a tamper-evident administrative audit ledger.

---

## ── Quick Links ──
- 📖 [Complete Functions Guide (FUNCTIONS.md)](file:///c:/Users/ashin/Downloads/qr_checkin_system_upgraded/qr_checkin_system/FUNCTIONS.md)
- 💡 [Detailed Troubleshooting & Setup (HELP.md)](file:///c:/Users/ashin/Downloads/qr_checkin_system_upgraded/qr_checkin_system/HELP.md)

---

## ── Core Upgrade Features ──

### 1. Robust Check-In Verification & Error Handling
- **Idempotency Token Caching**: Scan transactions cache idempotency tokens in memory to prevent network retries or scan jitter from causing duplicate registry records.
- **Quarantine Manager**: Failsafe quarantine queue (saved under `quarantine.json`) that captures failed scan attempts. Administrators can inspect quarantines on the dashboard, choose to manually approve them (inserting them into the roster as guest registrations) or discard them.
- **Path Traversal Protection**: Explicit path sanitization controls on `/qrcodes/<path>` and `/barcodes/<path>` routing endpoints to block malicious path traversal.

### 2. Comprehensive Security Controls
- **Event-Scoped Cryptographic QR Code Payloads**: Cryptographically registers attendee registration codes using SHA256 HMAC event secret keys, verified in real-time during scan ingestions.
- **Localhost Auth Toggle**: An optional checkbox configuration `enforce_localhost_auth` that requires terminal-based admin interfaces to supply password credentials.
- **Composite Rate Limiter**: Composite keys linking `f"{ip}:{device_id}"` restrict scan ingestion rates.

### 3. Tamper-Evident Hash-Chained Audit Trail
- Every administrative action (undo check-in, settings update, quarantine approval, etc.) is appended to `audit_log.csv` inside the active event's directory.
- Each row contains a SHA256 signature chaining it to the previous record's hash. A built-in sequence integrity check runs every time the audit log is requested or exported.

### 4. Advanced Administrative Utilities
- **Check-in Revocation (Undo)**: Allows managers to revoke a check-in status from registry tables. The system decrements counters, removes matching timestamp entries, updates Excel files atomically, and syncs all manager screens instantly.
- **Peak Arrivals Timeline**: Interactive HTML/CSS vertical bar chart that maps hourly arrival counts in real-time.
- **Subgroup Broadcaster**: Multi-channel (Email, WhatsApp, SMS, or All) bulk notifications triggered directly from subgroup visual pills.
- **Backup Rotations**: Up to 3 backup versions are kept during atomic spreadsheet commits.
- **System Diagnostics**: A unified `/health` diagnostics endpoint tracking lock statuses, folder write checks, tunnel activities, and message queue depths.

---

## ── Setup and Launch Guide ──

### 1. Requirements
Ensure Python 3.8+ is installed. Install all dependencies:
```powershell
pip install -r requirements.txt
```

### 2. Launching locally
To run the server locally on port 5001:
```powershell
python app.py
```
You can also run the quick launcher:
```powershell
run_system.bat
```

### 3. Running unit tests
To verify all core utilities, endpoint routes, phone normalization, path traversal checks, and security upgrades:
```powershell
python test_upgrades.py
```
