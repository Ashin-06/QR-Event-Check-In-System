# ⬢ QR Event Check-In System — Enterprise Edition

An enterprise-grade, secure, and highly scalable event check-in and attendee management platform. Built on a robust Flask and Socket.IO real-time backend with a premium dark-mesh glassmorphic frontend, the system facilitates seamless roster imports, dynamic barcode/QR code embedding, bulk invitations, and concurrent multi-device scan synchronization.

---

## 🌟 Key Features

### 1. Hierarchical Event & Sub-Event Management
* **Isolated Event Workspaces**: Dynamically manages folders under `events/<Event_Name>/` and sub-events under `events/<Event_Name>/<Sub_Event_Name>/`.
* **Standardized Storage**: Every event workspace encapsulates its own `registrations.xlsx` database, `scanned_log.csv` flat file, custom `config.json` parameters, and image asset subdirectories.

### 2. Advanced Spreadsheet Database Synchronization
* **Dynamic Header Resolution**: Automatically maps and adds missing standard columns (e.g., `Scan Timestamps`, `Scan Devices`, `Email Sent Status`, etc.) to the spreadsheet database on startup or check-in events.
* **Atomic Spreadsheet Updates**: Check-in records are updated in real-time. Timestamps and scanner device labels are appended sequentially, separated by semicolons (`;`) for multi-scan tracking.
* **Background Report Compilation**: Rebuilds highlighted spreadsheet databases containing highlighted scanned rows in a separate background daemon thread, eliminating main-thread locking and UI stuttering.

### 3. QR & Barcode Auto-Embedding
* **Unique ID Generation**: Generates cryptographically collision-free 8-character Unique IDs for guests.
* **Symbology Assets**: Auto-generates high-density QR codes and Code128 barcodes.
* **Automated Excel Formatting**: Fits and embeds QR and Barcode images directly inside [registrations.xlsx](file:///c:/Users/ashin/Downloads/qr_checkin_system_upgraded/qr_checkin_system/events/Default%20Event/registrations.xlsx) cells with optimized cell heights to ensure professional printability.

### 4. Bulk Invitation Campaigns
* **Dynamic Placeholders**: Supports templated emails and Twilio WhatsApp notifications containing dynamic fields like `{Name}`, `{Registration Number}`, `{Event}`, and `{QR_URL}`.
* **Robust Mail Server Delivery**: Features an asynchronous email campaign runner with 3-attempt exponential backoff retries, real-time progression tracking, and automated status logging.

### 5. Multi-Device Scanner Network & Live Monitor
* **Concurrent Scanning**: Supports multiple scanner operators checking in guests simultaneously across different networks (WiFi/LAN or Cellular/Public internet).
* **Live Device Monitor**: Tracks scanner operator statuses (online/offline, active pings, rename operator, processed scan counts, last-scanned attendee).
* **Quiet Real-Time Manager Alerts**: Displays non-intrusive bottom-right dashboard notifications for checking in guests, avoiding loud audio interruptions.

### 6. Secure Passcode Authorization
* **Glassmorphic Security Interface**: Protects the Manager Control Center from unauthorized remote access with a sleek login interface.
* **Smart Local Bypass**: Requests coming from localhost (`127.0.0.1`) are automatically authorized, while remote connections are securely checked.
* **Scan-to-Authorize Access**: Generates a shareable dashboard URL and a scan-to-connect QR code containing authorization parameters to quickly provision secondary organizer laptops and mobiles.

### 7. High-Performance Mobile Scanner Client
* **Dual Verification Modes**:
  * **Detail Verification (Default)**: Pauses the camera feed and presents a detailed guest profile card detailing Name, Registration Number, Phone, check-in status badge, and custom spreadsheet fields.
  * **⚡ Quick Scan Mode**: Bypasses the profile card for high-throughput gates, flashing a large 800ms feedback overlay (✅ checkmark for successful scans, ⚠️ warning for duplicate scans, ❌ cross for unregistered codes) and automatically resuming scanning.
* **Local Session History**: Operators can view their own scanning logs under the "My Scans" tab, backed by `localStorage` persistence.

---

## 📂 Project Architecture

```text
qr_checkin_system/
├── app.py                     # Flask backend, Socket.IO channels, API endpoints, SSH Tunnel loop
├── HELP.md                    # Detailed interactive system user manual
├── qradding.py                # Command-line utility to generate QR/barcodes and compile master Excel
├── qrsendemail.py             # Command-line utility to run bulk email invitations
├── requirements.txt           # Python package dependencies
├── run_system.bat             # Startup batch file for host machine launcher
├── start_public_tunnel.bat    # Standalone script to spin up the public internet tunnel
├── test_checkin.py            # Automated simulation test suite
├── static/
│   ├── css/                   # Datatables and vendor styling assets
│   └── js/                    # Socket.io, jQuery, Datatables, and QR scanner libraries
└── templates/
    ├── dashboard.html         # Manager Control Center HTML template
    ├── dashboard_login.html   # Glassmorphic passcode login HTML template
    └── index.html             # Mobile scanner client HTML template
```

---

## ⚙️ Environment Configuration

The application can be configured dynamically via environment variables:

| Variable | Default Value | Description |
|---|---|---|
| `PORT` | `5001` | Server port number |
| `FLASK_DEBUG` | `false` | Enable or disable Flask debug mode |
| `SECRET_KEY` | Autogenerated | Flask session encryption key |
| `CORS_ORIGIN` | `*` | Allowed CORS origins for WebSocket connections |

---

## 🚀 Installation & Quick Start

### 1. Installation
Install the required packages using pip:
```bash
pip install -r requirements.txt
```

### 2. Launching the System
Double-click **`run_system.bat`** (or run `python app.py` from the terminal).
* **Local Access**: Open `http://localhost:5001/dashboard` to access the Control Center.
* **LAN Access**: Operators on the same Wi-Fi network can connect to the host's IP (e.g. `http://192.168.1.15:5001`).

### 3. Activating the Public Tunnel
For scanners operating on mobile cellular networks:
1. Open the dashboard.
2. Go to the **📡 Public Internet Tunnel** panel and click **"Start Public Tunnel"** (or run `start_public_tunnel.bat`).
3. Share the generated public tunnel URL with your team or have them scan the sharing QR code.
