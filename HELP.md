# ⬢ QR Event Check-In System — Master System Guide & Documentation

Welcome to the **QR Event Check-In System**! This application is an enterprise-grade, high-speed, and secure event management platform. It allows you to manage guest rosters, design stunning custom ID badges, distribute passes across Email, WhatsApp, and SMS, and track attendee check-ins in real-time across multiple devices and entry gates.

---

## 🚀 1. Quick Start Guide (5-Minute Setup)

1. **Launch the System**: Double-click **`START_SYSTEM.bat`** (or `run_system.bat`). This starts the backend server on `http://localhost:5001` and opens the Manager Dashboard in your browser.
2. **Import or Add Attendees**: Click **⚙️ Operations** in the top bar, select **Import Spreadsheet** (or **Add Guest**), and upload your attendee list.
3. **Customize ID Badges**: Click the **🎨 ID Card & Communications Hub** tab to choose a theme, adjust fonts/typography, and preview badges.
4. **Send Passes**: Go to **Operations -> Bulk Send** to email or WhatsApp personalized passes with attached QR codes.
5. **Start Scanning**: Click **📷 Open QR Scanner** on the dashboard or open the scanner URL on any mobile phone/tablet to check in attendees.

---

## 📋 2. Main Navigation Tabs

### Tab A: `📋 Check-In & Logs`
The real-time operational command center for door staff and event managers:
* **Last Checked In Card**: Instantly displays the name, registration number, email, scan time, and scanner device name for the most recently checked-in guest.
* **📈 Peak Arrivals Timeline**: An interactive arrival velocity chart that visualizes incoming attendee volume in customizable 5, 15, 30, or 60-minute intervals.
* **Live Check-In Log Table**: Chronological table recording every scan attempt with timestamp, device name, and duplicate warnings.
* **Device Filter**: Filter logs by specific scanner phones, webcams, or manual dashboard check-ins.
* **Manual Check-In & Revoke**: Check in guests manually without a camera, or revoke/undo a previous scan with audit logging.

### Tab B: `👥 Guest Registry & Groups`
The central database of all event attendees:
* **Master Attendee Table**: Searchable, sortable list of all registered guests with registration numbers, email, phone, check-in status, and pass delivery status.
* **Overseer Hub (Attendee Profile Editor)**: Click on any attendee to open their full profile. You can edit their name, phone, email, assign custom subgroups, upload/change profile photos, download their badge, or send an instant single pass via Email, WhatsApp, or SMS.
* **👥 Custom Named Groups**: Multi-select attendees and group them into custom cohorts (e.g. "VIPs", "Keynote Speakers", "Sponsors", "Exhibitors"). You can assign dedicated badge themes, custom email templates, and broadcast batch passes specifically to that group.
* **Export Options**: Export the full registry to Excel (with check-ins highlighted in yellow and embedded QR codes) or clean CSV at any time.

### Tab C: `🎨 ID Card & Communications Hub`
Visual drag-and-drop badge and ID card designer:
* **11 Curated Themes**: Choose from professional themes including *Cyber Neon, Midnight Executive, Clean Slate, Monolith Modern, Blueprint Industrial, Classic Royal, Nordic Minimal, Solar Flare, Stealth Tech, Ultra Minimalist,* and *Rose Gold Boutique*.
* **Full Typography Controls**: Customize Font Family (Inter, Outfit, Roboto Mono, Georgia, Times, Courier, Segoe UI, Consolas, Playfair Display, Montserrat), Font Size, Bold, Italic, Letter Spacing (Tracking), Case Conversion (Uppercase, Lowercase, Capitalize), and Color for every text layer.
* **Profile Photo Options**:
  * `Show Photo`: Automatically loads attendee photos from Google Drive sharing links or direct image URLs.
  * `Smart Initials Avatar`: Generates high-contrast geometric initial avatars for attendees without photos.
  * `Hide Photo Container`: Hides the photo box for minimalist badge designs.
* **Dual Preview Modes**:
  * `Web Preview`: Ultra-fast live HTML/CSS rendering with real-time responsive updates as you tweak settings.
  * `Final PIL Render`: Generates pixel-perfect rasterized output matching actual print badge files.
* **Print-Ready Downloads**: Download single badges as high-resolution PNGs or vector PDFs, or export batch zip archives.

---

## 🛠️ 3. Operations Drawer (Left Panel Tabs)

Click **⚙️ Operations** in the top navigation bar to open the side drawer:

### 1. `Add Guest`
Register individual attendees on the fly:
* Enter Name, Registration Number, Email, and Phone.
* Add custom fields (e.g. Organization, Seat Number, Subgroup).
* The system automatically generates a unique 8-character ID, QR code, Code128 barcode, and ID badge.

### 2. `Import Spreadsheet`
Batch-import attendee lists from CSV or Excel (`.xlsx`) spreadsheets:
* **Auto-Header Detection**: Automatically detects column names for Name, Email Address, Registration Number, Phone Number, and custom attributes.
* **Duplicate Validation**: Identifies duplicate registration numbers inside the file and existing database duplicates.
* **Destination Options**: Import into the *Current Active Event* or automatically *Create a New Event*.
* **Reset & Overwrite**: Optional checkbox to clear the existing roster before importing.

### 3. `Bulk Send` (Email, WhatsApp, SMS)
Distribute personalized invitation passes with embedded QR codes:
* **📧 Email Subtab**:
  * Send automated HTML/text emails via Gmail SMTP or custom mail servers.
  * Attach generated QR codes and printable ID card badges.
  * Supports placeholders: `{Name}`, `{Registration Number}`, `{Event}`, `{Email Address}`, `{Phone Number}`, `{Unique ID}`.
* **💬 WhatsApp Subtab**:
  * Send WhatsApp messages via Twilio, Meta Cloud API, or 1-click WhatsApp Web links (`wa.me`) with individualized attendee details.
* **📱 SMS Subtab**:
  * Send personalized SMS text passes using a connected local Android SMS Gateway phone or Twilio SMS.

### 4. `Settings & Templates`
Event-level configuration:
* **Event Name Template**: Global event title shown across all passes and ID badges.
* **Check-In Time Restrictions**: Set strict start and end dates/times for valid scanning (manager manual check-ins and quarantine approvals bypass restrictions automatically).
* **Notification Templates**: Customize default Email subjects, Email bodies, and SMS messages.
* **Sound Alerts**: Choose audio themes (Modern Chime, Cyber Beep, Subtle Click, High Contrast) for scanner feedback.

### 5. `Clean & QR Tools`
Maintenance and security diagnostics:
* **Regenerate QR / Barcode Assets**: Re-creates any missing or updated QR code, barcode, and ID card image files for the active event.
* **Clean Database Duplicates**: Scans the Excel database for duplicate registration numbers and deduplicates rows safely.
* **Security Audit Log**: Review the timestamped audit log of all managerial actions (event switches, manual check-ins, revocations, and data modifications).

### 6. `Quarantine`
Security filter for unregistered or anomalous scans:
* When a visitor scans an unregistered QR code, malformed code, or duplicate, the scan is safely logged in the Quarantine Queue.
* Event managers can review quarantined scans and click **Approve** for **1-Click Self-Healing Registration**—the guest is immediately registered into the roster, given QR/barcodes/ID card, and checked in.

---

## 📡 4. Remote Scanners, Network & Internet Tunnels

* **Local LAN Scanning**: Any phone or tablet on the same Wi-Fi network can open the LAN URL (e.g. `http://192.168.1.15:5001`) to scan attendees.
* **📡 Public Internet Tunnel**: If scanning staff are on mobile data (4G/5G) or outside the venue Wi-Fi:
  * Click **"Start Public Tunnel"** in the operations drawer (or run `start_public_tunnel.bat`).
  * The system generates a public HTTPS link (e.g. `https://your-event.lhrtunnel.link`) with automatic reconnection keepalives.
* **🔒 Passcode-Protected Dashboard Sharing**:
  * Protect the dashboard with a 4-digit PIN passcode so only authorized staff can manage the event.
  * Share dedicated scanner links directly to volunteer phones without giving access to settings.

---

## 📷 5. Scanner Features & Verification Modes

1. **Normal Verification Mode**: Displays full attendee details (Name, Reg No, Email, Level, Photo) on scan with color-coded status badges (Green = Success, Orange = Duplicate, Red = Unregistered). Operator taps **"Verify & Next Scan"** to proceed.
2. **⚡ Quick Scan Mode**: Designed for high-throughput gates. Bypasses the details popup, flashes a quick green checkmark on the camera view, beeps, and is ready for the next attendee in 800 milliseconds.
3. **Sound & Haptic Feedback**: Plays audio confirmation chimes and triggers device vibration on mobile scanners for fast eyes-free operation.
4. **Offline Resilience**: If Wi-Fi briefly drops, scans are queued locally on the mobile device and automatically synced once the connection resumes.

---

## ❓ 6. Common Questions & Troubleshooting

* **Q: How do I undo an accidental check-in?**
  * *A*: Go to the **Guest Registry** tab, find the attendee, and click the red **Revoke** button (or click Revoke in the Live Scan Log).
* **Q: How do I use attendee photos from Google Forms or Google Drive?**
  * *A*: In your Excel/CSV roster, paste the Google Drive sharing link or public image URL into the `Profile Photo` or `Photo` column. The system will automatically fetch, thumbnail, and embed the photo onto the attendee's ID card.
* **Q: Can different ticket tiers (VIP vs General) have different badges?**
  * *A*: Yes! Create custom groups under **Guest Registry -> Custom Groups** or define subgroup templates under **Settings & Templates -> Subgroup Rules**.
