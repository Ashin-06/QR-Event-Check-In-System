QR Event Check-In System
A secure and efficient QR code-based attendance tracker for events, ensuring only registered guests gain entry.

Key Features
Individualized QR Codes – Unique QR codes sent via email to verified attendees.
Multi-Device Scanning – Supports simultaneous scanning across multiple devices.
Real-Time Logging – Tracks scans, timestamps, and duplicate attempts in a CSV file.
Prevents Unauthorized Entry – Only pre-registered guests can check in.

How It Works
Admin sends personalized QR codes to registered attendees via email.

Scanner App (mobile/desktop) validates QR codes at the event entrance.

CSV Logs record each scan, including:

Timestamp

Scan count (detects duplicates)

Attendee details

Limitation
⚠ CSV File Locking – Scanning pauses if the CSV log is open (to prevent write conflicts).
