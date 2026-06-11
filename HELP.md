# ⬢ QR Event Check-In System — User Guide (HELP.md)

Welcome to the **QR Event Check-In System**! This application is a high-speed, secure, and scalable event management tool designed to import registration lists, automatically generate and embed QR/barcodes, distribute invitation passes, and track check-ins in real-time across multiple devices.

---

## 🚀 Getting Started

### 1. Launching the Server
To start the application, double-click **`run_system.bat`** on the host machine. This will:
1. Initialize the Flask backend server on `http://localhost:5001`.
2. Output your **LAN Network URL** (e.g. `http://192.168.1.15:5001`) which allows other devices on the same Wi-Fi network to connect.
3. Automatically launch the Manager Command Center in your default browser.

### 2. Exposing to the Internet (Public Tunneling)
If your check-in operators are using cellular data or are on a different Wi-Fi network:
- Go to the **📡 Public Internet Tunnel** panel on your dashboard and click **"Start Public Tunnel"** (or run `start_public_tunnel.bat` manually).
- The system will spin up a secure connection via `localhost.run` and display a public URL (e.g. `https://your-event.lhrtunnel.link`).
- The tunnel has built-in **SSH TCP Keepalives** and an **automatic reconnect loop** to keep it active and stable indefinitely.

---

## 📂 Event Management

### 1. Folder Structure
All event data is isolated. When you create or switch events, the system reads and writes to:
`qr_checkin_system/events/<Event_Name>/`

Within each event folder, the system automatically creates:
* **`registrations.xlsx`**: The primary Excel database containing attendee details, QR/barcode images, scan statuses, email/whatsapp receipt statuses, and timestamps.
* **`scanned_log.csv`**: A lightweight, flat log of checked-in attendees.
* **`qrcodes/` & `barcodes/`**: Directories storing generated image assets.
* **`config.json`**: Custom email/SMS templates, Twilio credentials, and SMTP details.

### 2. Sub-Events
To create sub-events (e.g., "Day 1", "Workshop A") within a main event:
- In the **📁 Event Explorer** sidebar, select a parent event, click **"Create Event / Sub-Event"**, choose the parent event, and give the sub-event a name.
- It will create a sub-folder under the parent event folder (e.g., `events/Main Event/Workshop A/`) with its own separate registration and check-in database.

---

## 📥 Registration & Roster Uploads

### 1. Importing Attendees (CSV or Excel)
- Click the **"Import Roster"** button on the dashboard.
- Upload any Excel (`.xlsx`) or CSV roster.
- The system will open a **Preview & Duplicate Validation Modal**, analyzing all entries and mapping columns (Name, Email, Registration Number, Phone).
- It checks for:
  - Missing critical columns.
  - Duplicates *within the file*.
  - Duplicates *already present in the active event database*.
- Organizers can opt to **Auto-Resolve Duplicates** (which automatically appends numeric suffixes to duplicate registration numbers) or skip them.

### 2. Auto-Generating QR & Barcodes
- Upon confirming the import (or manually adding an attendee), the system:
  1. Generates a unique 8-character ID.
  2. Generates a QR code containing structural attendee details (Name, Email, ID).
  3. Generates a Code128 barcode matching the Registration Number.
  4. Saves the images in `qrcodes/` and `barcodes/`.
  5. Inserts and displays them directly inside [registrations.xlsx](file:///c:/Users/ashin/Downloads/qr_checkin_system_upgraded/qr_checkin_system/events/Default%20Event/registrations.xlsx) in their respective columns, adjusting row heights for perfect alignment.

---

## ✉️ Distributing Passes (Email & WhatsApp Campaigns)

- Navigate to the **✉️ Bulk Email Invitations** panel.
- Enter your sender email and SMTP App Password.
- Subject lines and email bodies support dynamic placeholders:
  * `{Name}`: Replaced with the guest's name.
  * `{Registration Number}`: Replaced with the guest's registration number.
  * `{Event}`: Replaced with the active event title.
- Click **"Send Emails"** to run the campaign in the background. The system will attach the generated QR code image to each email and update the `Email Sent Status` column in the Excel file automatically.
- Check-in confirmation templates andTwilio integrations for WhatsApp can be set up in the **⚙️ Event Configuration** panel.

---

## 📷 Checking In Attendees

### 1. Local Webcam Scanning
Organizers can scan QR codes using the host computer's webcam directly from the dashboard.

### 2. Multi-Device Mobile Scanning
To connect mobile devices to use their cameras as scanners:
1. Ensure the public tunnel is running (if on cellular data) or the devices are on the same Wi-Fi (if using the LAN link).
2. Go to the **🔒 Dashboard Sharing & Access** card on the dashboard.
3. Open the **LAN URL** or **Public Tunnel URL** on the mobile device (or scan the **Dashboard Access QR Code**).
4. Remote devices will be greeted with a glassmorphic login screen. Enter the **4-digit Passcode** shown on your main dashboard to authorize access.
5. Once authorized, select **"Go to Scanner Client"** to open the camera scanner.

### 3. Check-In Verification Modes
* **Normal Mode (Default)**: After scanning a QR, the operator's screen displays a detailed overlay showing the guest's Name, Email, Reg Number, Phone, and any custom fields, with color-coded alerts (Green: Check-in OK, Orange: Duplicate checked-in, Red: Unregistered). Click **"Verify & Next Scan"** to proceed.
* **⚡ Quick Scan Mode**: Toggle this on for high-throughput entry gates. The details page is bypassed. Scanning flashes a large status indicator directly on the camera view for 800ms (✅ for OK, ⚠️ for duplicate, ❌ for error) and automatically resumes.

---

## 📊 Monitoring & Reports

* **Live Device Monitor**: Tracks all connected scanner devices in real-time. Displays device name (can be renamed from the dashboard), connection status (online/offline), total scans performed, and their last scanned attendee.
* **Scan History Log**: View a unified table of all checked-in attendees, with filtering options (filter by scanning operator device, duplicate status, and scan time frame).
* **Data Synced Excel**: At any time, click **"Download Highlighted Excel"** to export the registrations sheet. Checked-in rows are highlighted in yellow, complete with scan timestamps and scanner device names.

---

## 🔒 Sharing & Permissions
In the **🔒 Dashboard Sharing & Access** card, you can restrict remote users:
* **Passcode Protected**: Only users with the active passcode can view the dashboard and logs.
* **Public Access**: Anyone with the link can view the dashboard (convenient for sharing monitors on big screens).
* **Disabled**: Remote access is shut down; the dashboard can only be accessed from the host machine.
