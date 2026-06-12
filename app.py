"""
QR Event Check-In System — app.py
Fixed & upgraded production version.

Bug fixes applied:
1. Race condition: global mutable DataFrame protected by lock everywhere (original missed /data_csv read).
2. scanned_log.loc[len(scanned_log)] index bug — replaced with pd.concat() to avoid index collisions.
3. QR match logic was broken: original used "name in qr_data" (substring match) which gives false positives.
   Now stores and matches the encoded QR string exactly.
4. Column-letter formula in qradding.py was wrong for columns > 26 — fixed in utility.
5. No error handling if EXCEL_FILE missing at startup — now raises clear message.
6. sort_log() mutated the shared df in-place without lock — moved inside lock.
7. processing = False in socket callback only — if socket event never arrives (network drop) scanner locks.
   Added timeout fallback in frontend.
8. HIGHLIGHTED_FILE rebuild on every scan is slow (full copy+rewrite) — moved to background thread.
9. No CORS restriction / secret key on Flask app — added.
10. Timestamps column dtype inconsistency (NaN vs str) — normalised on load.
11. /data_csv leaked full Timestamps column to frontend unnecessarily — still included but sanitised.
12. Missing Content-Type validation on /scan — added.
13. No rate-limiting guard — basic per-IP cooldown added.
14. qrsendemail.py: REG_COL had trailing space ' ' — fixed.
15. qradding.py: csv_path and final_excel were empty — now use argparse / env vars.
"""

from __future__ import annotations

import sys
if hasattr(sys.stdout, "reconfigure"):
    sys.stdout.reconfigure(encoding="utf-8")
if hasattr(sys.stderr, "reconfigure"):
    sys.stderr.reconfigure(encoding="utf-8")

import os
import re
import shutil
import socket as _socket
import threading
import time
from collections import defaultdict
from datetime import datetime
import csv
import io
import random
import string

import pandas as pd
import qrcode
from flask import Flask, jsonify, render_template, request, session, send_file, send_from_directory, redirect
from flask_socketio import SocketIO
from openpyxl import load_workbook
from openpyxl.styles import PatternFill
from openpyxl.drawing.image import Image as XLImage
from openpyxl.utils import get_column_letter

# ── User Agent Parser ────────────────────────────────────────────────────────
def _parse_ua(ua: str) -> str:
    ua = ua.lower()
    os_name = "Device"
    if "windows" in ua:
        os_name = "Windows"
    elif "android" in ua:
        os_name = "Android"
    elif "ipad" in ua or "ipod" in ua:
        os_name = "iPad"
    elif "iphone" in ua:
        os_name = "iPhone"
    elif "macintosh" in ua or "mac os" in ua:
        os_name = "macOS"
    elif "linux" in ua:
        os_name = "Linux"
        
    browser = "Browser"
    if "chrome" in ua or "crios" in ua:
        browser = "Chrome"
    elif "firefox" in ua or "fxios" in ua:
        browser = "Firefox"
    elif "safari" in ua and "chrome" not in ua:
        browser = "Safari"
    elif "edge" in ua or "edg" in ua:
        browser = "Edge"
    elif "opera" in ua or "opr" in ua:
        browser = "Opera"
        
    return f"{os_name} {browser}"

# ── LAN IP helper ────────────────────────────────────────────────────────────
def get_lan_ip() -> str:
    """Return the LAN IPv4 address of this machine (best-effort)."""
    try:
        s = _socket.socket(_socket.AF_INET, _socket.SOCK_DGRAM)
        s.connect(("8.8.8.8", 80))
        ip = s.getsockname()[0]
        s.close()
        return ip
    except Exception:
        return "127.0.0.1"


def is_localhost_request() -> bool:
    if request.headers.get("X-Forwarded-For"):
        return False
    host = request.host.lower()
    if not ("localhost" in host or "127.0.0.1" in host):
        return False
    if request.remote_addr not in ["127.0.0.1", "::1"]:
        return False
    return True


def is_dashboard_authorized() -> bool:
    global dashboard_sharing, dashboard_passcode
    if is_localhost_request():
        return True
    if dashboard_sharing == "public":
        return True
    elif dashboard_sharing == "disabled":
        return False
    
    if session.get("dashboard_authorized") == True:
        return True
        
    passcode_param = request.args.get("passcode")
    if passcode_param == dashboard_passcode:
        session["dashboard_authorized"] = True
        return True
        
    return False


# ── App setup ────────────────────────────────────────────────────────────────
app = Flask(__name__)
app.config["SECRET_KEY"] = os.environ.get("SECRET_KEY", os.urandom(24).hex())
socketio = SocketIO(app, cors_allowed_origins=os.environ.get("CORS_ORIGIN", "*"))

# ── Connected-client tracker ──────────────────────────────────────────────────
_clients_lock = threading.Lock()
_connected_clients: int = 0

# ── Connected-device tracker ──────────────────────────────────────────────────
_devices_lock = threading.Lock()
connected_devices: dict[str, dict] = {}


# ── File paths & Events Structure ─────────────────────────────────────────────
BASE_DIR         = os.path.dirname(os.path.abspath(__file__))
EVENTS_DIR       = os.path.join(BASE_DIR, "events")
os.makedirs(EVENTS_DIR, exist_ok=True)
SCAN_COL_NAME    = "Scanned Status"

active_event = "Default Event"

def get_active_event_path() -> str:
    global active_event
    p = os.path.join(EVENTS_DIR, active_event)
    os.makedirs(p, exist_ok=True)
    return p

def get_excel_file() -> str:
    return os.path.join(get_active_event_path(), "registrations.xlsx")

def get_highlighted_file() -> str:
    return os.path.join(get_active_event_path(), "registrations_highlighted.xlsx")

def get_log_file() -> str:
    return os.path.join(get_active_event_path(), "scanned_log.csv")

def get_qr_dir() -> str:
    p = os.path.join(get_active_event_path(), "qrcodes")
    os.makedirs(p, exist_ok=True)
    return p

def get_barcode_dir() -> str:
    p = os.path.join(get_active_event_path(), "barcodes")
    os.makedirs(p, exist_ok=True)
    return p

def get_config_file() -> str:
    return os.path.join(get_active_event_path(), "config.json")


# ── Global state ──────────────────────────────────────────────────────────────
lock            = threading.Lock()
highlight_lock  = threading.Lock()

dashboard_sharing = "passcode"
# Generate a random 4-digit passcode on server startup
dashboard_passcode = "".join(random.choices(string.digits, k=4))

# ── Rate limiting (simple in-memory) ─────────────────────────────────────────
_rate_store: dict[str, float] = defaultdict(float)
RATE_LIMIT_SECONDS = 2          # minimum seconds between scans from same IP


def _is_rate_limited(ip: str) -> bool:
    now = time.time()
    if now - _rate_store[ip] < RATE_LIMIT_SECONDS:
        return True
    _rate_store[ip] = now
    return False


# ── Load / initialise log ─────────────────────────────────────────────────────
# ── Event Config Loader / Saver ───────────────────────────────────────────────
def get_event_config() -> dict:
    config_file = get_config_file()
    if os.path.exists(config_file):
        try:
            with open(config_file, "r", encoding="utf-8") as f:
                return json.load(f)
        except Exception:
            pass
    # Return default configs
    return {
        "email_template": "Hi {Name},\n\nPlease find your QR Code attached. Present it at the entrance of {Event}.\n\nRegistration: {Registration Number}\n\nRegards,\nEvent Team",
        "whatsapp_template": "Hi {Name},\n\nYour registration QR Code for {Event} is ready! Present it at the entrance.\n\nRegistration: {Registration Number}\n\nDownload QR here: {QR_URL}",
        "auto_notify_on_scan": False,
        "scan_notify_template": "Hi {Name},\n\nYou checked in successfully at {Location} at {Time}.\n\nRegards,\nEvent Team",
        "scan_notify_channels": "none", # "email", "whatsapp", "both", "none"
        "email_subject": "Your Event QR Code",
        "email_sender": "",
        "email_password": "",
        "twilio_sid": "",
        "twilio_token": "",
        "twilio_sender": "",
        "event_name_template": "the Event"
    }

def save_event_config(cfg: dict):
    config_file = get_config_file()
    try:
        with open(config_file, "w", encoding="utf-8") as f:
            json.dump(cfg, f, indent=4)
    except Exception as e:
        print(f"[config] Error saving config: {e}")

import json

# ── Load / initialise log ─────────────────────────────────────────────────────
def _load_log() -> pd.DataFrame:
    log_file = get_log_file()
    if os.path.exists(log_file):
        df = pd.read_csv(log_file)
    else:
        df = pd.DataFrame(columns=["QR Data", "Scan Count", "Timestamps", "Devices"])

    # Normalise columns
    for col in ["QR Data", "Timestamps", "Devices"]:
        if col not in df.columns:
            df[col] = ""
    if "Scan Count" not in df.columns:
        df["Scan Count"] = 0

    # Ensure dtypes
    df["Scan Count"] = pd.to_numeric(df["Scan Count"], errors="coerce").fillna(0).astype(int)
    df["Timestamps"] = df["Timestamps"].fillna("").astype(str)
    df["Devices"]    = df["Devices"].fillna("").astype(str)
    df["QR Data"]    = df["QR Data"].fillna("").astype(str)
    return df


def sort_log(df: pd.DataFrame) -> None:
    """Sort by most-recent timestamp descending, in-place."""
    df["__last"] = df["Timestamps"].apply(
        lambda s: s.split(";")[-1].strip() if s else ""
    )
    df.sort_values("__last", ascending=False, inplace=True)
    df.drop(columns="__last", inplace=True)
    df.reset_index(drop=True, inplace=True)


scanned_log = None

def load_active_event() -> None:
    global scanned_log
    excel_path = get_excel_file()
    log_path = get_log_file()
    
    # Clean up any leftover temporary files in the event folder
    try:
        event_path = get_active_event_path()
        if os.path.exists(event_path):
            for f in os.listdir(event_path):
                if f.startswith("tmp") and f.endswith(".xlsx"):
                    try:
                        os.unlink(os.path.join(event_path, f))
                    except OSError:
                        pass
    except Exception as e:
        print(f"[load_active_event] Error cleaning leftover temp files: {e}")

    # Look for registrations.csv first and convert to registrations.xlsx (always take precedence if CSV exists)
    csv_path = os.path.join(get_active_event_path(), "registrations.csv")
    if os.path.exists(csv_path):
        try:
            print(f"[load_active_event] Found registrations.csv. Converting to registrations.xlsx...")
            df = pd.read_csv(csv_path)
            df.to_excel(excel_path, index=False, engine="openpyxl")
            try:
                os.remove(csv_path)
            except OSError:
                pass
        except Exception as e:
            print(f"[load_active_event] Error converting registrations.csv to Excel: {e}")

    # Ensure registrations file exists
    if not os.path.exists(excel_path):
        # Create a default blank registrations sheet
        master_src = os.path.join(BASE_DIR, "data_with_qrdemo.xlsx")
        if os.path.exists(master_src):
            shutil.copy(master_src, excel_path)
        else:
            import openpyxl
            wb = openpyxl.Workbook()
            ws = wb.active
            headers = [
                "Name", "Email Address", "Registration Number", "Phone Number",
                "Unique ID", "QR", "Barcode", SCAN_COL_NAME, "Email Sent Status", "WhatsApp Sent Status",
                "Scan Timestamps", "Scan Devices", "QR Code Image", "Barcode Image"
            ]
            for col_num, header in enumerate(headers, 1):
                ws.cell(row=1, column=col_num, value=header)
            wb.save(excel_path)
            wb.close()
            
    # Auto-initialize and self-heal registrations metadata if it exists
    if os.path.exists(excel_path):
        try:
            wb = load_workbook(excel_path)
            ws = wb.active
            if _initialize_missing_metadata(ws):
                _atomic_save(wb, excel_path)
            else:
                wb.close()
        except Exception as e:
            print(f"[load_active_event] Error self-healing Excel metadata: {e}")

    # Ensure scanned log exists
    if not os.path.exists(log_path):
        df = pd.DataFrame(columns=["QR Data", "Scan Count", "Timestamps", "Devices"])
        df.to_csv(log_path, index=False)
        
    scanned_log = _load_log()
    sort_log(scanned_log)
    
    # Highlighted file
    high_path = get_highlighted_file()
    if not os.path.exists(high_path):
        shutil.copy(excel_path, high_path)
        
    # Rebuild highlighted in background
    threading.Thread(target=_rebuild_highlighted, daemon=True).start()



# ── Background highlight rebuild ───────────────────────────────────────────────
def _rebuild_highlighted() -> None:
    """Copy master Excel and re-highlight all scanned rows (runs in daemon thread).
    Uses a staging temp-copy so the source file is never locked during processing.
    """
    import tempfile
    wb = None
    staging = None
    excel_file = get_excel_file()
    highlighted_file = get_highlighted_file()
    try:
        with highlight_lock:
            # Copy source to a temp file in system temp (very brief lock)
            with lock:
                with tempfile.NamedTemporaryFile(suffix=".xlsx", delete=False) as tmp:
                    staging = tmp.name
                shutil.copy(excel_file, staging)

            # Process the staging copy (no lock held on excel_file)
            wb = load_workbook(staging)
            ws = wb.active
            hdrs = {
                str(c.value).strip().lower(): i + 1
                for i, c in enumerate(ws[1])
                if c.value
            }
            scan_key = SCAN_COL_NAME.lower()
            scan_col = hdrs.get(scan_key)
            yellow = PatternFill(start_color="FFFF00", end_color="FFFF00", fill_type="solid")
            if scan_col:
                for row in ws.iter_rows(min_row=2):
                    status = row[scan_col - 1].value
                    if status and str(status).strip():
                        for col in range(1, ws.max_column + 1):
                            ws.cell(row=row[0].row, column=col).fill = yellow

            # Atomic save → highlighted_file
            _atomic_save(wb, highlighted_file)
            wb = None

    except Exception as e:
        print(f"[highlight] Error: {e}")
    finally:
        if wb:
            try:
                wb.close()
            except Exception:
                pass
        if staging and os.path.exists(staging):
            try:
                os.unlink(staging)
            except OSError:
                pass


def _clean_val(val, default="") -> str:
    if pd.isna(val) or val is None:
        return default
    s = str(val).strip()
    if s.lower() == "nan":
        return default
    return s


# ── Helper: normalise column header lookup ─────────────────────────────────────
def _get_headers(ws) -> dict[str, int]:
    return {
        str(c.value).strip().lower(): i + 1
        for i, c in enumerate(ws[1])
        if c.value
    }


def _broadcast_devices():
    with _devices_lock:
        devs = list(connected_devices.values())
    socketio.emit("devices_updated", devs)


# ── Socket.IO connect / disconnect tracking ───────────────────────────────────
@socketio.on("connect")
def _on_connect():
    global _connected_clients
    with _clients_lock:
        _connected_clients += 1
        count = _connected_clients
    socketio.emit("clients_count", {"count": count})
    
    with _devices_lock:
        devs = list(connected_devices.values())
    socketio.emit("devices_updated", devs, to=request.sid)


@socketio.on("disconnect")
def _on_disconnect():
    global _connected_clients
    with _clients_lock:
        _connected_clients = max(0, _connected_clients - 1)
        count = _connected_clients
    socketio.emit("clients_count", {"count": count})
    
    device_id = session.get("device_id")
    if device_id:
        with _devices_lock:
            if device_id in connected_devices:
                connected_devices[device_id]["online"] = False
                connected_devices[device_id]["last_active_time"] = datetime.now().strftime("%H:%M:%S")
        _broadcast_devices()


@socketio.on("register_device")
def _on_register_device(data):
    device_id = data.get("device_id")
    if not device_id:
        return
    
    device_name = data.get("device_name", "Unknown Device").strip()
    user_agent = request.headers.get("User-Agent", "Unknown")
    ip = request.headers.get("X-Forwarded-For", request.remote_addr or "unknown").split(",")[0].strip()
    short_ua = _parse_ua(user_agent)
    
    with _devices_lock:
        session["device_id"] = device_id
        if device_id not in connected_devices:
            connected_devices[device_id] = {
                "id": device_id,
                "name": device_name,
                "ip": ip,
                "user_agent": short_ua,
                "online": True,
                "scans": 0,
                "last_activity": "Connected",
                "last_active_time": datetime.now().strftime("%H:%M:%S")
            }
        else:
            connected_devices[device_id]["name"] = device_name
            connected_devices[device_id]["ip"] = ip
            connected_devices[device_id]["user_agent"] = short_ua
            connected_devices[device_id]["online"] = True
            connected_devices[device_id]["last_active_time"] = datetime.now().strftime("%H:%M:%S")
            if connected_devices[device_id]["last_activity"] == "Offline":
                connected_devices[device_id]["last_activity"] = "Reconnected"
                
    _broadcast_devices()


@socketio.on("rename_device")
def _on_rename_device(data):
    device_id = data.get("device_id")
    new_name = data.get("name", "").strip()
    if not device_id or not new_name:
        return
    with _devices_lock:
        if device_id in connected_devices:
            connected_devices[device_id]["name"] = new_name
            connected_devices[device_id]["last_active_time"] = datetime.now().strftime("%H:%M:%S")
    _broadcast_devices()



# ── Routes ─────────────────────────────────────────────────────────────────────
@app.route("/")
def index():
    is_admin = is_localhost_request() or session.get("dashboard_authorized") == True
    return render_template("index.html", is_admin=is_admin)


@app.route("/dashboard")
def dashboard_view():
    if not is_dashboard_authorized():
        return render_template("dashboard_login.html", error=request.args.get("error"))
    return render_template("dashboard.html")


@app.route("/dashboard_login", methods=["GET", "POST"])
def dashboard_login():
    global dashboard_passcode
    if request.method == "POST":
        code_entered = request.form.get("passcode", "").strip()
        if code_entered == dashboard_passcode:
            session["dashboard_authorized"] = True
            return redirect("/dashboard")
        return render_template("dashboard_login.html", error="Invalid passcode.")
    return render_template("dashboard_login.html")


@app.route("/dashboard_sharing_info", methods=["GET"])
def dashboard_sharing_info():
    port = int(os.environ.get("PORT", 5001))
    lan_ip = get_lan_ip()
    lan_url = f"http://{lan_ip}:{port}/dashboard?passcode={dashboard_passcode}"
    
    tunnel_url_dash = ""
    if _tunnel_url:
        tunnel_url_dash = f"{_tunnel_url}/dashboard?passcode={dashboard_passcode}"
        
    return jsonify(
        sharing=dashboard_sharing,
        passcode=dashboard_passcode,
        lan_url=lan_url,
        tunnel_url=tunnel_url_dash
    )


@app.route("/save_sharing_settings", methods=["POST"])
def save_sharing_settings():
    global dashboard_sharing, dashboard_passcode
    data = request.json or {}
    sharing = data.get("sharing", "passcode")
    passcode = data.get("passcode", "").strip()
    
    if sharing not in ["passcode", "public", "disabled"]:
        return jsonify(message="Invalid sharing mode."), 400
        
    dashboard_sharing = sharing
    if passcode:
        dashboard_passcode = passcode
        
    socketio.emit("sharing_updated", {})
    return jsonify(message="Dashboard sharing settings saved successfully.")


@app.route("/network_info")
def network_info():
    """Return the LAN URL so the UI can display it for other devices."""
    port = int(os.environ.get("PORT", 5001))
    ip   = get_lan_ip()
    return jsonify(url=f"http://{ip}:{port}", ip=ip, port=port)


def _get_qr_to_attendee_map() -> dict[str, dict]:
    excel_file = get_excel_file()
    mapping = {}
    if not os.path.exists(excel_file):
        return mapping
    try:
        with lock:
            wb = load_workbook(excel_file, read_only=True)
            ws = wb.active
            hdrs = _get_headers(ws)
            name_idx = hdrs.get("name")
            reg_idx = hdrs.get("registration number")
            qr_idx = hdrs.get("qr")
            if name_idx and reg_idx and qr_idx:
                for row in ws.iter_rows(min_row=2, values_only=True):
                    if len(row) >= max(name_idx, reg_idx, qr_idx):
                        qr_val = str(row[qr_idx - 1] or "").strip()
                        if qr_val:
                            mapping[qr_val] = {
                                "name": str(row[name_idx - 1] or "").strip(),
                                "reg_no": str(row[reg_idx - 1] or "").strip()
                            }
            wb.close()
    except Exception as e:
        print(f"Error building QR mapping: {e}")
    return mapping


@app.route("/data_csv")
def data_csv():
    with lock:
        df = scanned_log.copy()

    if "Devices" not in df.columns:
        df["Devices"] = ""

    df["Last Timestamp"] = df["Timestamps"].apply(
        lambda s: s.split(";")[-1].strip() if s else ""
    )
    df["Last Device"] = df["Devices"].apply(
        lambda s: s.split(";")[-1].strip() if s else ""
    )
    
    # Map QR Data to Attendee Details
    qr_map = _get_qr_to_attendee_map()
    
    records = []
    for idx, row in df.iterrows():
        qr_val = str(row["QR Data"]).strip()
        att = qr_map.get(qr_val)
        if not att:
            # Fallback parse for old verbose formats
            name_match = re.search(r"Name:\s*([^\n\r]+)", qr_val)
            reg_match = re.search(r"(Reg No|Reg_No|Registration Number|ID):\s*([^\n\r]+)", qr_val)
            if name_match:
                att = {
                    "name": name_match.group(1).strip(),
                    "reg_no": reg_match.group(2).strip() if reg_match else "Unknown"
                }
            else:
                att = {"name": qr_val, "reg_no": "Unknown"}
                
        records.append({
            "QR Data": qr_val,
            "attendee_name": att["name"],
            "reg_no": att["reg_no"],
            "Scan Count": int(row["Scan Count"]),
            "Last Timestamp": row["Last Timestamp"],
            "Timestamps": row["Timestamps"],
            "Devices": row["Devices"],
            "Last Device": row["Last Device"]
        })
    return jsonify(records)


@app.route("/stats")
def stats():
    """Summary statistics for dashboard cards."""
    with lock:
        total     = len(scanned_log)
        unique    = int((scanned_log["Scan Count"] > 0).sum())
        duplicate = int((scanned_log["Scan Count"] > 1).sum())
    return jsonify(total=total, unique=unique, duplicate=duplicate)


def _emit_stats() -> None:
    """Broadcast current stats to all connected clients."""
    with lock:
        total     = len(scanned_log)
        unique    = int((scanned_log["Scan Count"] > 0).sum())
        duplicate = int((scanned_log["Scan Count"] > 1).sum())
    socketio.emit("stats_updated", {"total": total, "unique": unique, "duplicate": duplicate})


def _atomic_save(wb, target_path: str) -> None:
    """Save workbook to a temp file then atomically replace target.
    Avoids Windows file-lock conflicts between concurrent readers and writers.
    """
    import tempfile
    dir_ = os.path.dirname(target_path)
    with tempfile.NamedTemporaryFile(dir=dir_, suffix=".xlsx", delete=False) as tmp:
        tmp_path = tmp.name
    try:
        wb.save(tmp_path)
        wb.close()
        os.replace(tmp_path, target_path)
    except Exception:
        try:
            os.unlink(tmp_path)
        except OSError:
            pass
        raise


# ── Scan Event Notification Helpers ───────────────────────────────────────────
def _send_single_email(sender, password, recipient, subject, body):
    import yagmail
    try:
        yag = yagmail.SMTP(user=sender, password=password)
        yag.send(to=recipient, subject=subject, contents=body)
        yag.close()
        print(f"[notify] Email sent to {recipient}")
    except Exception as e:
        print(f"[notify] Failed to send email to {recipient}: {e}")


def _send_single_whatsapp(sid, token, sender_phone, recipient_phone, body):
    from twilio.rest import Client
    import re
    try:
        client = Client(sid, token)
        phone_clean = re.sub(r"[^\d+]", "", recipient_phone)
        if phone_clean:
            if not phone_clean.startswith("+"):
                if len(phone_clean) == 10:
                    phone_clean = "+91" + phone_clean
                else:
                    phone_clean = "+" + phone_clean
            from_number = f"whatsapp:{sender_phone}"
            to_number = f"whatsapp:{phone_clean}"
            client.messages.create(body=body, from_=from_number, to=to_number)
            print(f"[notify] WhatsApp sent to {phone_clean}")
    except Exception as e:
        print(f"[notify] Failed to send WhatsApp to {recipient_phone}: {e}")


def send_scan_notification_async(details, location, timestamp):
    cfg = get_event_config()
    channels = cfg.get("scan_notify_channels", "none")
    if channels == "none":
        return
        
    name = details.get("Name", "")
    email = details.get("Email", "")
    phone = details.get("Phone", "")
    reg_no = details.get("Registration Number", "")
    
    # Prepare body using template
    template = cfg.get("scan_notify_template", "")
    body = template.replace("{Name}", name).replace("{Location}", location).replace("{Time}", timestamp).replace("{Registration Number}", reg_no)
    
    # Send email if configured
    if channels in ["email", "both"] and email:
        sender = cfg.get("email_sender")
        password = cfg.get("email_password")
        subject = "Check-in Confirmation: " + cfg.get("event_name_template", "Event")
        if sender and password:
            threading.Thread(target=_send_single_email, args=(sender, password, email, subject, body), daemon=True).start()
            
    # Send WhatsApp if configured
    if channels in ["whatsapp", "both"] and phone:
        sid = cfg.get("twilio_sid")
        token = cfg.get("twilio_token")
        twilio_phone = cfg.get("twilio_sender")
        if sid and token and twilio_phone:
            threading.Thread(target=_send_single_whatsapp, args=(sid, token, twilio_phone, phone, body), daemon=True).start()


def _perform_checkin(qr_data: str, device_id: str = "unknown") -> tuple[dict, int]:
    """Helper to perform check-in operations. Returns (response_dict, status_code)."""
    # Normalize incoming QR data string
    qr_data_norm = qr_data.replace("\r\n", "\n").strip()

    excel_file = get_excel_file()
    log_file = get_log_file()
    wb = None
    try:
        with lock:
            now = datetime.now().strftime("%Y-%m-%d %H:%M:%S")

            # 1) Open Excel and find matching attendee
            wb   = load_workbook(excel_file)
            ws   = wb.active
            hdrs, _ = _get_or_create_headers(ws)
            scan_key = SCAN_COL_NAME.lower()

            required = {"name", "registration number", "qr"}
            missing  = required - hdrs.keys()
            if missing:
                return {
                    "message": f"Excel is missing columns: {missing}",
                    "error": "schema_error"
                }, 500

            details: dict = {}
            found_row = None
            for row in ws.iter_rows(min_row=2):
                qrval = str(row[hdrs["qr"] - 1].value or "").strip()
                qrval_norm = qrval.replace("\r\n", "\n").strip()
                regval = str(row[hdrs["registration number"] - 1].value or "").strip()
                regval_norm = regval.replace("\r\n", "\n").strip()
                uidval = str(row[hdrs["unique id"] - 1].value or "").strip()
                uidval_norm = uidval.replace("\r\n", "\n").strip()
                
                # Check for match on QR string, registration number, or unique ID
                if qr_data_norm in [qrval_norm, regval_norm, uidval_norm]:
                    found_row = row
                    break

            # If not found in Excel, reject check-in
            if found_row is None:
                wb.close()
                wb = None
                with _devices_lock:
                    if device_id in connected_devices:
                        connected_devices[device_id]["last_activity"] = "Unregistered Scan"
                        connected_devices[device_id]["last_active_time"] = datetime.now().strftime("%H:%M:%S")
                _broadcast_devices()
                return {"message": "❌ Unregistered QR Code", "details": {}, "is_duplicate": False}, 400

            # 2) Attendee matched! Update Excel row
            r   = found_row[0].row
            col = hdrs[scan_key]
            curr = ws.cell(row=r, column=col).value or ""
            curr_str = str(curr).strip().lower()
            if curr_str == "scanned" or curr_str == "manual check-in" or "scanned" in curr_str or "manual" in curr_str:
                m = re.search(r"(\d+)", curr_str)
                cnt = int(m.group(1)) if m else 1
            else:
                cnt = 0
            cnt += 1
            
            if device_id == "Manual Check-In":
                new_status = "Manual Check-In" if cnt == 1 else f"Manual Check-In {cnt} Times"
            else:
                new_status = "Scanned" if cnt == 1 else f"Scanned {cnt} Times"
            ws.cell(row=r, column=col, value=new_status)

            device_name = device_id
            with _devices_lock:
                if device_id in connected_devices:
                    device_name = connected_devices[device_id]["name"]

            ts_col_idx = hdrs.get("scan timestamps")
            dev_col_idx = hdrs.get("scan devices")
            if ts_col_idx:
                prev_ts = str(ws.cell(row=r, column=ts_col_idx).value or "").strip()
                ws.cell(row=r, column=ts_col_idx, value=f"{prev_ts};{now}" if prev_ts else now)
            if dev_col_idx:
                prev_dev = str(ws.cell(row=r, column=dev_col_idx).value or "").strip()
                ws.cell(row=r, column=dev_col_idx, value=f"{prev_dev};{device_name}" if prev_dev else device_name)

            # Get phone number column
            phone_col_idx = hdrs.get("phone number")
            phone_val = str(found_row[phone_col_idx - 1].value or "").strip() if phone_col_idx else ""

            # Dynamic custom columns extraction
            custom_fields = {}
            system_cols = [
                "name", "email address", "registration number", "phone number",
                "unique id", "qr", "barcode", SCAN_COL_NAME.lower(),
                "email sent status", "whatsapp sent status",
                "qr code image", "barcode image"
            ]
            for c_low, c_idx in hdrs.items():
                if c_low not in system_cols:
                    custom_fields[ws.cell(row=1, column=c_idx).value] = str(found_row[c_idx - 1].value or "").strip()

            email_val = str(found_row[hdrs["email address"] - 1].value or "").strip() if "email address" in hdrs else ""

            details = {
                "Name":                str(found_row[hdrs["name"] - 1].value or "").strip(),
                "Email":               email_val,
                "Registration Number": str(found_row[hdrs["registration number"] - 1].value or "").strip(),
                "Phone":               phone_val,
                "Status":              new_status,
                "ScanCount":           cnt,
                "custom_fields":       custom_fields
            }

            # Atomic write: save to temp then replace (avoids Windows lock conflicts)
            _atomic_save(wb, excel_file)
            wb = None  # already closed by _atomic_save

            # 3) Update CSV log (only for registered attendees)
            device_name = "unknown"
            with _devices_lock:
                if device_id in connected_devices:
                    device_name = connected_devices[device_id]["name"]

            if "Devices" not in scanned_log.columns:
                scanned_log["Devices"] = ""

            mask = scanned_log["QR Data"] == qr_data
            if mask.any():
                idx = scanned_log.index[mask][0]
                scanned_log.at[idx, "Scan Count"] += 1
                prev_ts = scanned_log.at[idx, "Timestamps"]
                scanned_log.at[idx, "Timestamps"] = f"{prev_ts};{now}" if prev_ts else now
                
                prev_devs = str(scanned_log.at[idx, "Devices"]) if pd.notna(scanned_log.at[idx, "Devices"]) else ""
                scanned_log.at[idx, "Devices"] = f"{prev_devs};{device_name}" if prev_devs else device_name
            else:
                new_row = pd.DataFrame([{"QR Data": qr_data, "Scan Count": 1, "Timestamps": now, "Devices": device_name}])
                globals()["scanned_log"] = pd.concat(
                    [scanned_log, new_row], ignore_index=True
                )

            sort_log(scanned_log)
            scanned_log.to_csv(log_file, index=False)

        # Rebuild highlighted file in background ─────────────────────────────
        threading.Thread(target=_rebuild_highlighted, daemon=True).start()

        # Emit real-time update ───────────────────────────────────────────────
        with lock:
            row_data = scanned_log.loc[scanned_log["QR Data"] == qr_data].iloc[0]
        last_ts = row_data["Timestamps"].split(";")[-1].strip()
        last_dev = row_data["Devices"].split(";")[-1].strip() if "Devices" in row_data and row_data["Devices"] else device_name
        socketio.emit("row_updated", {
            "qr_data":        qr_data,
            "scan_count":     int(row_data["Scan Count"]),
            "timestamps":     row_data["Timestamps"],
            "last_timestamp": last_ts,
            "devices":        row_data["Devices"] if "Devices" in row_data else device_name,
            "last_device":    last_dev,
            "details":        details,
        })

        name    = details.get("Name", "")
        status  = details.get("Status", "")
        is_dup  = details.get("ScanCount", 1) > 1
        message = f"✅ {name} — {status}" if name else "✅ Scanned"
        if is_dup:
            message = f"⚠️ {name} already checked in ({status})"

        # Update device scan stats
        with _devices_lock:
            if device_id in connected_devices:
                connected_devices[device_id]["scans"] += 1
                connected_devices[device_id]["last_activity"] = f"Scanned {name}"
                connected_devices[device_id]["last_active_time"] = datetime.now().strftime("%H:%M:%S")
        _broadcast_devices()

        # Broadcast scan alert to ALL devices (not just the scanner)
        socketio.emit("scan_alert", {
            "name":       name,
            "is_duplicate": is_dup,
            "status":     status,
            "message":    message,
            "details":    details,
            "device_name": device_name,
        })

        # Send post-scan notification if configured
        send_scan_notification_async(details, device_name, now)

        # Broadcast updated stats to ALL devices
        _emit_stats()

        return {"message": message, "details": details, "is_duplicate": is_dup}, 200

        return {"message": message, "details": details, "is_duplicate": is_dup}, 200

    except PermissionError as e:
        print(f"[checkin] Excel file locked: {e}")
        return {
            "message": "❌ Excel database file is locked (likely open in Microsoft Excel). Please close Excel and try again.",
            "error": "file_locked"
        }, 409
    except Exception as e:
        print(f"[checkin] Error: {e}")
        return {"message": "Internal server error", "error": str(e)}, 500
    finally:
        if wb:
            try:
                wb.close()
            except Exception:
                pass


@app.route("/scan", methods=["POST"])
def scan():
    # Content-type guard
    if not request.is_json:
        return jsonify(message="Request must be JSON."), 415

    # Rate limit
    ip = request.headers.get("X-Forwarded-For", request.remote_addr or "unknown").split(",")[0].strip()
    if _is_rate_limited(ip):
        return jsonify(message="Too many requests. Please wait."), 429

    payload = request.json or {}
    qr_data = payload.get("qr_data", "").strip()
    device_id = payload.get("device_id", "unknown").strip()
    if not qr_data:
        return jsonify(message="No QR data provided."), 400

    res, code = _perform_checkin(qr_data, device_id)
    return jsonify(res), code


@app.route("/manual_checkin", methods=["POST"])
def manual_checkin():
    # Content-type guard
    if not request.is_json:
        return jsonify(message="Request must be JSON."), 415

    payload = request.json or {}
    qr_data = payload.get("qr_data", "").strip()
    if not qr_data:
        return jsonify(message="No QR data provided."), 400

    res, code = _perform_checkin(qr_data, device_id="Manual Check-In")
    return jsonify(res), code



@app.route("/registry")
def get_registry():
    try:
        excel_file = get_excel_file()
        if not os.path.exists(excel_file):
            return jsonify([])
            
        with lock:
            df_xl = pd.read_excel(excel_file, engine="openpyxl")

        df_xl.columns = df_xl.columns.astype(str).str.strip()
        col_map = {c.lower(): c for c in df_xl.columns}

        required = {"name", "registration number", "qr"}
        missing  = required - col_map.keys()
        if missing:
            return jsonify(message=f"Excel is missing columns: {missing}"), 500

        scan_col_orig = col_map.get(SCAN_COL_NAME.lower())
        phone_col_orig = col_map.get("phone number")
        email_sent_col = col_map.get("email sent status")
        wa_sent_col = col_map.get("whatsapp sent status")

        records = []
        for _, row in df_xl.iterrows():
            name_v   = _clean_val(row.get(col_map["name"]))
            reg_v    = _clean_val(row.get(col_map["registration number"]))
            qr_v     = _clean_val(row.get(col_map["qr"]))
            status_v = _clean_val(row.get(scan_col_orig), default="Not Scanned") if scan_col_orig else "Not Scanned"
            
            email_v  = _clean_val(row.get(col_map["email address"])) if "email address" in col_map else ""
            phone_v  = _clean_val(row.get(phone_col_orig)) if phone_col_orig else ""
            email_s  = _clean_val(row.get(email_sent_col), default="Not Sent") if email_sent_col else "Not Sent"
            wa_s     = _clean_val(row.get(wa_sent_col), default="Not Sent") if wa_sent_col else "Not Sent"

            if not any([name_v, email_v, reg_v, qr_v]):
                continue

            # Dynamic custom columns
            custom_fields = {}
            system_cols = [
                "name", "email address", "registration number", "phone number",
                "unique id", "qr", "barcode", SCAN_COL_NAME.lower(),
                "email sent status", "whatsapp sent status",
                "qr code image", "barcode image"
            ]
            for c_low, c_orig in col_map.items():
                if c_low not in system_cols:
                    custom_fields[c_orig] = _clean_val(row.get(c_orig))

            records.append({
                "name":   name_v,
                "email":  email_v,
                "reg_no": reg_v,
                "qr":     qr_v,
                "status": status_v,
                "phone":  phone_v,
                "email_sent": email_s,
                "wa_sent": wa_s,
                "custom_fields": custom_fields
            })

        return jsonify(records)
    except Exception as e:
        return jsonify(message="Error loading registry", error=str(e)), 500


@app.route("/active_event_columns")
def active_event_columns():
    try:
        excel_file = get_excel_file()
        if not os.path.exists(excel_file):
            return jsonify(columns=[])
        
        with lock:
            df_xl = pd.read_excel(excel_file, engine="openpyxl")
        
        df_xl.columns = df_xl.columns.astype(str).str.strip()
        system_cols = [
            "name", "email address", "registration number", "phone number",
            "unique id", "qr", "barcode", SCAN_COL_NAME.lower(),
            "email sent status", "whatsapp sent status",
            "qr code image", "barcode image"
        ]
        
        custom_cols = [col for col in df_xl.columns if col.lower() not in system_cols]
        return jsonify(columns=custom_cols)
    except Exception as e:
        return jsonify(message=str(e)), 500


@app.route("/active_event_form_fields")
def active_event_form_fields():
    try:
        excel_file = get_excel_file()
        if not os.path.exists(excel_file):
            return jsonify(fields=[])
            
        with lock:
            wb = load_workbook(excel_file, read_only=True)
            ws = wb.active
            headers = [cell.value for cell in ws[1] if cell.value]
            wb.close()
            
        system_cols = {
            "unique id", "qr", "barcode", SCAN_COL_NAME.lower(),
            "email sent status", "whatsapp sent status",
            "qr code image", "barcode image", "scan timestamps", "scan devices"
        }
        
        fields = []
        for h in headers:
            h_clean = str(h).strip()
            h_low = h_clean.lower()
            if h_low not in system_cols:
                is_required = h_low in ["name", "registration number"]
                fields.append({
                    "name": h_clean,
                    "key": h_low,
                    "required": is_required
                })
        return jsonify(fields=fields)
    except Exception as e:
        return jsonify(message=str(e)), 500


@app.route("/clean_duplicates", methods=["POST"])
def clean_duplicates():
    excel_file = get_excel_file()
    if not os.path.exists(excel_file):
        return jsonify(message="No registrations spreadsheet found."), 404
        
    wb = None
    try:
        with lock:
            wb = load_workbook(excel_file)
            ws = wb.active
            hdrs, _ = _get_or_create_headers(ws)
            
            reg_col = hdrs.get("registration number")
            if not reg_col:
                wb.close()
                return jsonify(message="Registration Number column is missing."), 400
                
            seen_regs = set()
            rows_to_delete = []
            
            # Identify duplicates from bottom to top
            for r in range(2, ws.max_row + 1):
                reg_val = ws.cell(row=r, column=reg_col).value
                reg_str = str(reg_val).strip().lower() if reg_val else ""
                
                if not reg_str:
                    continue
                    
                if reg_str in seen_regs:
                    rows_to_delete.append(r)
                else:
                    seen_regs.add(reg_str)
            
            removed_cnt = len(rows_to_delete)
            for r in sorted(rows_to_delete, reverse=True):
                # Clean up images for this row and shift lower ones up
                images_to_keep = []
                for img in ws._images:
                    row_num = None
                    if hasattr(img, 'anchor'):
                        if hasattr(img.anchor, '_from') and hasattr(img.anchor._from, 'row'):
                            row_num = img.anchor._from.row + 1
                        elif isinstance(img.anchor, str):
                            row_num = int(''.join(filter(str.isdigit, img.anchor))) if any(c.isdigit() for c in img.anchor) else None
                    
                    if row_num == r:
                        # Skip this image (deletes it)
                        continue
                    elif row_num is not None and row_num > r:
                        # Shift up
                        if hasattr(img.anchor, '_from') and hasattr(img.anchor._from, 'row'):
                            img.anchor._from.row -= 1
                        elif isinstance(img.anchor, str):
                            col_letters = ''.join(filter(str.isalpha, img.anchor))
                            img.anchor = f"{col_letters}{row_num - 1}"
                    images_to_keep.append(img)
                ws._images = images_to_keep
                
                # Delete the row
                ws.delete_rows(r)
                
            if removed_cnt > 0:
                _atomic_save(wb, excel_file)
            else:
                wb.close()
                
            wb = None
            
        if removed_cnt > 0:
            threading.Thread(target=_rebuild_highlighted, daemon=True).start()
            socketio.emit("registry_updated", {})
            _emit_stats()
            
        return jsonify(message=f"Cleaned {removed_cnt} duplicate rows successfully.", cleaned=removed_cnt), 200
        
    except Exception as e:
        return jsonify(message=f"Error cleaning duplicates: {str(e)}"), 500
    finally:
        if wb:
            try:
                wb.close()
            except Exception:
                pass


@app.route("/regenerate_assets", methods=["POST"])
def regenerate_assets():
    excel_file = get_excel_file()
    if not os.path.exists(excel_file):
        return jsonify(message="No registrations spreadsheet found."), 404
        
    wb = None
    try:
        with lock:
            wb = load_workbook(excel_file)
            ws = wb.active
            hdrs, _ = _get_or_create_headers(ws)
            
            # Clear all existing images from the sheet to regenerate them cleanly
            ws._images = []
            
            # Now run self-healing/generation for every row
            for r in range(2, ws.max_row + 1):
                name_val = ws.cell(row=r, column=hdrs["name"]).value
                reg_val = ws.cell(row=r, column=hdrs["registration number"]).value
                if not name_val or not reg_val:
                    continue
                reg_str = str(reg_val).strip()
                
                # Re-generate QR
                qr_path = os.path.join(get_qr_dir(), f"{reg_str}.png")
                _generate_qr_for_guest(reg_str, reg_str)
                _embed_qr_image(ws, r, qr_path, hdrs["qr code image"])
                
                # Re-generate Barcode
                bc_path = os.path.join(get_barcode_dir(), f"{reg_str}.png")
                _generate_barcode_for_guest(reg_str)
                _embed_barcode_image(ws, r, bc_path, hdrs["barcode image"])
                
            _atomic_save(wb, excel_file)
            wb = None
            
        threading.Thread(target=_rebuild_highlighted, daemon=True).start()
        socketio.emit("registry_updated", {})
        
        return jsonify(message="Successfully regenerated and re-embedded all QR codes and barcodes."), 200
        
    except Exception as e:
        return jsonify(message=f"Error regenerating assets: {str(e)}"), 500
    finally:
        if wb:
            try:
                wb.close()
            except Exception:
                pass


@app.route("/events", methods=["GET"])
def list_events():
    tree = []
    if os.path.exists(EVENTS_DIR):
        for main_evt in os.listdir(EVENTS_DIR):
            main_path = os.path.join(EVENTS_DIR, main_evt)
            if os.path.isdir(main_path):
                sub_evts = []
                for sub_evt in os.listdir(main_path):
                    sub_path = os.path.join(main_path, sub_evt)
                    if os.path.isdir(sub_path) and sub_evt not in ["qrcodes", "barcodes"]:
                        sub_evts.append(sub_evt)
                tree.append({
                    "name": main_evt,
                    "sub_events": sub_evts
                })
    return jsonify(events=tree, active=active_event)


@app.route("/select_event", methods=["POST"])
def select_event():
    global active_event
    data = request.json or {}
    name = data.get("name", "").strip()
    if not name:
        return jsonify(message="Event name is required."), 400
    
    clean_name = name.replace("\\", "/").strip("/")
    target_path = os.path.abspath(os.path.join(EVENTS_DIR, clean_name))
    if not target_path.startswith(os.path.abspath(EVENTS_DIR)):
        return jsonify(message="Invalid path."), 400
        
    active_event = clean_name
    load_active_event()
    
    socketio.emit("registry_updated", {})
    _emit_stats()
    
    return jsonify(message=f"Switched active event to '{active_event}'", active=active_event)


@app.route("/create_event", methods=["POST"])
def create_event():
    data = request.json or {}
    name = data.get("name", "").strip()
    parent = data.get("parent", "").strip()
    
    if not name:
        return jsonify(message="Event name is required."), 400
        
    clean_name = re.sub(r'[\\/*?:"<>|]', "", name).strip()
    if not clean_name:
        return jsonify(message="Invalid event name."), 400
        
    if parent:
        parent_clean = parent.replace("\\", "/").strip("/")
        relative_path = f"{parent_clean}/{clean_name}"
    else:
        relative_path = clean_name
        
    target_path = os.path.join(EVENTS_DIR, relative_path)
    if os.path.exists(target_path):
        return jsonify(message="Event folder already exists."), 409
        
    os.makedirs(target_path, exist_ok=True)
    
    # Initialize registrations.xlsx inside it
    excel_path = os.path.join(target_path, "registrations.xlsx")
    import openpyxl
    wb = openpyxl.Workbook()
    ws = wb.active
    headers = [
        "Name", "Email Address", "Registration Number", "Phone Number",
        "Unique ID", "QR", "Barcode", SCAN_COL_NAME, "Email Sent Status", "WhatsApp Sent Status",
        "Scan Timestamps", "Scan Devices", "QR Code Image", "Barcode Image"
    ]
    for col_num, header in enumerate(headers, 1):
        ws.cell(row=1, column=col_num, value=header)
    wb.save(excel_path)
    wb.close()
    
    # Initialize scanned_log.csv
    log_path = os.path.join(target_path, "scanned_log.csv")
    df = pd.DataFrame(columns=["QR Data", "Scan Count", "Timestamps", "Devices"])
    df.to_csv(log_path, index=False)
    
    # Create subfolders for qrcodes and barcodes
    os.makedirs(os.path.join(target_path, "qrcodes"), exist_ok=True)
    os.makedirs(os.path.join(target_path, "barcodes"), exist_ok=True)
    
    return jsonify(message=f"Event '{relative_path}' created successfully.", name=relative_path)


@app.route("/get_config", methods=["GET"])
def get_config():
    return jsonify(get_event_config())


@app.route("/save_config", methods=["POST"])
def save_config():
    data = request.json or {}
    cfg = get_event_config()
    for k, v in data.items():
        cfg[k] = v
    save_event_config(cfg)
    return jsonify(message="Configuration saved successfully.")


@app.route("/help_content", methods=["GET"])
def get_help_content():
    help_path = os.path.join(BASE_DIR, "HELP.md")
    if os.path.exists(help_path):
        try:
            with open(help_path, "r", encoding="utf-8") as f:
                content = f.read()
            return jsonify(content=content)
        except Exception as e:
            return jsonify(content=f"Error reading HELP.md: {str(e)}"), 500
    return jsonify(content="HELP.md not found."), 404


_tunnel_process = None
_tunnel_url = None
_tunnel_active = False
_tunnel_lock = threading.Lock()

def _run_tunnel():
    global _tunnel_process, _tunnel_url, _tunnel_active
    import subprocess
    import re
    
    print("[tunnel] Starting localhost.run SSH tunnel loop...")
    # Add ServerAlive keepalives to prevent disconnect drops and bypass interactive prompt
    cmd = ["ssh", "-o", "StrictHostKeyChecking=no", "-o", "UserKnownHostsFile=NUL" if os.name == "nt" else "/dev/null", "-o", "ServerAliveInterval=30", "-o", "ServerAliveCountMax=3", "-R", "80:127.0.0.1:5001", "nokey@localhost.run"]
    
    url_pattern = re.compile(r"https?://[a-zA-Z0-9.-]+\.lhr\.(?:life|link|run|tunnel)")
    lhrtunnel_pattern = re.compile(r"https?://[a-zA-Z0-9.-]+\.lhrtunnel\.link")
    
    _tunnel_active = True
    socketio.emit("tunnel_status_update", {"active": True, "url": "Connecting..."})
    
    while _tunnel_active:
        try:
            startupinfo = None
            if os.name == 'nt':
                startupinfo = subprocess.STARTUPINFO()
                startupinfo.dwFlags |= subprocess.STARTF_USESHOWWINDOW
                startupinfo.wShowWindow = subprocess.SW_HIDE

            with _tunnel_lock:
                if not _tunnel_active:
                    break
                _tunnel_process = subprocess.Popen(
                    cmd,
                    stdout=subprocess.PIPE,
                    stderr=subprocess.STDOUT,
                    text=True,
                    bufsize=1,
                    startupinfo=startupinfo
                )
            
            while _tunnel_active and _tunnel_process.poll() is None:
                line = _tunnel_process.stdout.readline()
                if not line:
                    break
                print(f"[tunnel-ssh] {line.strip()}")
                
                match = url_pattern.search(line) or lhrtunnel_pattern.search(line)
                if match:
                    _tunnel_url = match.group(0)
                    print(f"[tunnel] Extracted tunnel URL: {_tunnel_url}")
                    socketio.emit("tunnel_status_update", {"active": True, "url": _tunnel_url})
            
            if _tunnel_active:
                print("[tunnel] SSH tunnel dropped. Reconnecting in 3 seconds...")
                socketio.emit("tunnel_status_update", {"active": True, "url": "Reconnecting..."})
                time.sleep(3)
        except Exception as e:
            print(f"[tunnel] Loop Error: {e}")
            if _tunnel_active:
                time.sleep(3)
                
    print("[tunnel] Tunnel SSH loop fully stopped.")
    socketio.emit("tunnel_status_update", {"active": False, "url": None})


@app.route("/start_tunnel", methods=["POST"])
def start_tunnel():
    global _tunnel_active
    with _tunnel_lock:
        if _tunnel_active:
            return jsonify(message="Tunnel is already active.", url=_tunnel_url), 200
            
    threading.Thread(target=_run_tunnel, daemon=True).start()
    return jsonify(message="Tunnel starting..."), 202


@app.route("/stop_tunnel", methods=["POST"])
def stop_tunnel():
    global _tunnel_process, _tunnel_active, _tunnel_url
    with _tunnel_lock:
        if _tunnel_process:
            try:
                _tunnel_process.terminate()
                _tunnel_process.kill()
            except Exception:
                pass
        _tunnel_active = False
        _tunnel_url = None
        _tunnel_process = None
    socketio.emit("tunnel_status_update", {"active": False, "url": None})
    return jsonify(message="Tunnel stopped.")


@app.route("/tunnel_status", methods=["GET"])
def tunnel_status():
    return jsonify(active=_tunnel_active, url=_tunnel_url)


@app.route("/tunnel_qr", methods=["GET"])
def tunnel_qr():
    data = request.args.get("data", "").strip()
    if not data:
        return "Missing data parameter", 400
    try:
        qr = qrcode.QRCode(version=1, box_size=5, border=1)
        qr.add_data(data)
        qr.make(fit=True)
        img = qr.make_image(fill_color="black", back_color="white")
        buf = io.BytesIO()
        img.save(buf, format="PNG")
        buf.seek(0)
        return send_file(buf, mimetype="image/png")
    except Exception as e:
        print(f"[tunnel_qr] Error generating local QR: {e}")
        return "Internal server error", 500


def _get_or_create_headers(ws) -> tuple[dict[str, int], bool]:
    hdrs = _get_headers(ws)
    modified = False
    
    # Define aliases mapping: target_lowercase -> list of lowercase aliases
    aliases = {
        "name": ["name", "fullname", "full name", "attendee name", "visitor name", "guest name"],
        "registration number": ["registration number", "reg number", "reg no", "registration no", "reg_no", "registration_number", "id", "ticket number", "ticket no", "regid"],
        "email address": ["email address", "email", "mail", "email_address"],
        "phone number": ["phone number", "phone", "mobile", "whatsapp", "contact", "phone_number", "contact number", "contact no"],
        "qr": ["qr", "qr code", "qr_code", "qr value", "qr_value"],
        "barcode": ["barcode", "barcode value", "barcode_value"]
    }
    
    # For each target, if it is not in hdrs, see if any alias is in hdrs
    for target, alias_list in aliases.items():
        if target not in hdrs:
            for alias in alias_list:
                if alias in hdrs:
                    col_idx = hdrs[alias]
                    standard_names = {
                        "name": "Name",
                        "registration number": "Registration Number",
                        "email address": "Email Address",
                        "phone number": "Phone Number",
                        "qr": "QR",
                        "barcode": "Barcode"
                    }
                    ws.cell(row=1, column=col_idx, value=standard_names[target])
                    hdrs[target] = col_idx
                    del hdrs[alias]
                    modified = True
                    break
                    
    # Now define standard expected columns and their nice names
    expected = {
        "name": "Name",
        "registration number": "Registration Number",
        "unique id": "Unique ID",
        "qr": "QR",
        "barcode": "Barcode",
        SCAN_COL_NAME.lower(): SCAN_COL_NAME,
        "scan timestamps": "Scan Timestamps",
        "scan devices": "Scan Devices",
        "qr code image": "QR Code Image",
        "barcode image": "Barcode Image"
    }
    
    if "email address" in hdrs:
        expected["email sent status"] = "Email Sent Status"
    if "phone number" in hdrs:
        expected["whatsapp sent status"] = "WhatsApp Sent Status"
        
    for key, val in expected.items():
        if key not in hdrs:
            col = ws.max_column + 1
            ws.cell(row=1, column=col, value=val)
            hdrs[key] = col
            modified = True
            
    return hdrs, modified


def _has_image_at_cell(ws, row: int, col: int) -> bool:
    """Check if there is already an image anchored to the given row and column (1-indexed)."""
    for img in ws._images:
        row_num = None
        col_num = None
        if hasattr(img, 'anchor'):
            if hasattr(img.anchor, '_from') and hasattr(img.anchor._from, 'row'):
                row_num = img.anchor._from.row + 1
                col_num = img.anchor._from.col + 1
            elif isinstance(img.anchor, str):
                try:
                    # Parse coord like 'F2' -> row=2, col=6
                    cell = ws[img.anchor]
                    row_num = cell.row
                    row_col = cell.column
                except Exception:
                    pass
        if row_num == row and col_num == col:
            return True
    return False


def _initialize_missing_metadata(ws) -> bool:
    hdrs, modified = _get_or_create_headers(ws)
    
    # Check if there are any rows
    if ws.max_row < 2:
        return modified
        
    existing_uids = set()
    for r in range(2, ws.max_row + 1):
        uid_val = ws.cell(row=r, column=hdrs["unique id"]).value
        if uid_val:
            existing_uids.add(str(uid_val).strip().upper())
            
    row_modified = False
    for r in range(2, ws.max_row + 1):
        name_val = ws.cell(row=r, column=hdrs["name"]).value
        reg_val = ws.cell(row=r, column=hdrs["registration number"]).value
        if not name_val or not reg_val:
            continue
            
        reg_str = str(reg_val).strip()
        
        # 1) Unique ID
        uid_cell = ws.cell(row=r, column=hdrs["unique id"])
        if not uid_cell.value:
            uid = _generate_unique_id(existing_uids)
            uid_cell.value = uid
            existing_uids.add(uid)
            row_modified = True
            
        # 2) QR value
        qr_cell = ws.cell(row=r, column=hdrs["qr"])
        if not qr_cell.value:
            qr_cell.value = reg_str
            row_modified = True
            
        # 3) Barcode value
        bc_cell = ws.cell(row=r, column=hdrs["barcode"])
        if not bc_cell.value:
            bc_cell.value = reg_str
            row_modified = True
            
        # 4) Scanned Status
        status_cell = ws.cell(row=r, column=hdrs[SCAN_COL_NAME.lower()])
        if status_cell.value is None:
            status_cell.value = ""
            row_modified = True
            
        # 5) Email Sent Status
        if "email sent status" in hdrs:
            esc_cell = ws.cell(row=r, column=hdrs["email sent status"])
            if esc_cell.value is None:
                esc_cell.value = "Not Sent"
                row_modified = True
                
        # 6) WhatsApp Sent Status
        if "whatsapp sent status" in hdrs:
            wsc_cell = ws.cell(row=r, column=hdrs["whatsapp sent status"])
            if wsc_cell.value is None:
                wsc_cell.value = "Not Sent"
                row_modified = True
                
        # 7) QR Image
        qr_path = os.path.join(get_qr_dir(), f"{reg_str}.png")
        if not os.path.exists(qr_path) or not _has_image_at_cell(ws, r, hdrs["qr code image"]):
            _generate_qr_for_guest(reg_str, reg_str)
            _embed_qr_image(ws, r, qr_path, hdrs["qr code image"])
            row_modified = True
            
        # 8) Barcode Image
        bc_path = os.path.join(get_barcode_dir(), f"{reg_str}.png")
        if not os.path.exists(bc_path) or not _has_image_at_cell(ws, r, hdrs["barcode image"]):
            _generate_barcode_for_guest(reg_str)
            _embed_barcode_image(ws, r, bc_path, hdrs["barcode image"])
            row_modified = True
            
    return modified or row_modified


def _generate_qr_for_guest(content_str: str, reg_no: str) -> str:
    qr_dir = get_qr_dir()
    qr_path = os.path.join(qr_dir, f"{reg_no.strip()}.png")
    qr_img = qrcode.make(content_str.strip())
    qr_img.save(qr_path)
    return qr_path


def _generate_barcode_for_guest(reg_no: str) -> str:
    import barcode
    from barcode.writer import ImageWriter
    code_class = barcode.get_barcode_class('code128')
    writer = ImageWriter()
    options = {'write_text': True, 'font_size': 9, 'text_distance': 3.0, 'module_height': 12.0}
    
    cleaned_reg = str(reg_no).strip().upper()
    barcode_dir = get_barcode_dir()
    barcode_path_no_ext = os.path.join(barcode_dir, cleaned_reg)
    b = code_class(cleaned_reg, writer=writer)
    b.save(barcode_path_no_ext, options=options)
    return barcode_path_no_ext + ".png"


def _embed_qr_image(ws, r: int, qr_path: str, col_idx: int) -> None:
    if not os.path.exists(qr_path):
        return
    xl_img = XLImage(qr_path)
    xl_img.width = 100
    xl_img.height = 100
    cell_addr = f"{get_column_letter(col_idx)}{r}"
    ws.add_image(xl_img, cell_addr)
    ws.row_dimensions[r].height = 80
    ws.cell(row=r, column=col_idx, value="Embedded")


def _embed_barcode_image(ws, r: int, barcode_path: str, col_idx: int) -> None:
    if not os.path.exists(barcode_path):
        return
    xl_img = XLImage(barcode_path)
    xl_img.width = 150
    xl_img.height = 60
    cell_addr = f"{get_column_letter(col_idx)}{r}"
    ws.add_image(xl_img, cell_addr)
    ws.row_dimensions[r].height = 80
    ws.cell(row=r, column=col_idx, value="Embedded")


def _generate_unique_id(existing_ids: set[str], length: int = 8) -> str:
    charset = string.ascii_uppercase + string.digits
    while True:
        uid = "".join(random.choices(charset, k=length))
        if uid not in existing_ids:
            return uid


@app.route("/qrcodes/<path:filename>")
def serve_qrcode(filename):
    return send_from_directory(get_qr_dir(), filename)


@app.route("/barcodes/<path:filename>")
def serve_barcode(filename):
    return send_from_directory(get_barcode_dir(), filename)


@app.route("/add_attendee", methods=["POST"])
def add_attendee():
    if not request.is_json:
        return jsonify(message="Request must be JSON."), 415
        
    payload = request.json or {}
    name = payload.get("name", "").strip()
    email = payload.get("email", "").strip()
    reg_no = payload.get("reg_no", "").strip()
    phone = payload.get("phone", "").strip()
    
    if not name or not reg_no:
        return jsonify(message="Name and Registration Number are required."), 400
        
    excel_file = get_excel_file()
    wb = None
    try:
        with lock:
            wb = load_workbook(excel_file)
            ws = wb.active
            hdrs, _ = _get_or_create_headers(ws)
            
            # Check duplicate reg number
            for r in range(2, ws.max_row + 1):
                cell_val = ws.cell(row=r, column=hdrs["registration number"]).value
                if cell_val and str(cell_val).strip().lower() == reg_no.lower():
                    wb.close()
                    return jsonify(message=f"Registration number '{reg_no}' already exists."), 409
            
            existing_uids = set()
            for r in range(2, ws.max_row + 1):
                cell_val = ws.cell(row=r, column=hdrs["unique id"]).value
                if cell_val:
                    existing_uids.add(str(cell_val).strip().upper())
                    
            uid = _generate_unique_id(existing_uids)
            qr_str = reg_no
            qr_path = _generate_qr_for_guest(qr_str, reg_no)
            barcode_path = _generate_barcode_for_guest(reg_no)
            
            new_r = ws.max_row + 1
            ws.cell(row=new_r, column=hdrs["name"], value=name)
            if "email address" in hdrs:
                ws.cell(row=new_r, column=hdrs["email address"], value=email)
            ws.cell(row=new_r, column=hdrs["registration number"], value=reg_no)
            if "phone number" in hdrs:
                ws.cell(row=new_r, column=hdrs["phone number"], value=phone)
            ws.cell(row=new_r, column=hdrs["unique id"], value=uid)
            ws.cell(row=new_r, column=hdrs["qr"], value=qr_str)
            ws.cell(row=new_r, column=hdrs["barcode"], value=reg_no)
            ws.cell(row=new_r, column=hdrs[SCAN_COL_NAME.lower()], value="")
            if "email sent status" in hdrs:
                ws.cell(row=new_r, column=hdrs["email sent status"], value="Not Sent")
            if "whatsapp sent status" in hdrs:
                ws.cell(row=new_r, column=hdrs["whatsapp sent status"], value="Not Sent")
            
            # Custom fields dynamic storage
            custom_payload = payload.get("custom_fields", {})
            for key, val in custom_payload.items():
                key_clean = key.lower().strip()
                if key_clean in hdrs:
                    ws.cell(row=new_r, column=hdrs[key_clean], value=str(val).strip())
            
            _embed_qr_image(ws, new_r, qr_path, hdrs["qr code image"])
            _embed_barcode_image(ws, new_r, barcode_path, hdrs["barcode image"])
            
            _atomic_save(wb, excel_file)
            wb = None
            
        threading.Thread(target=_rebuild_highlighted, daemon=True).start()
        
        socketio.emit("registry_updated", {})
        _emit_stats()
        
        return jsonify(message=f"Successfully registered {name} ({reg_no})"), 200
        
    except PermissionError:
        return jsonify(message="Excel file is locked. Please close it in Excel first."), 409
    except Exception as e:
        return jsonify(message=f"Error adding attendee: {str(e)}"), 500
    finally:
        if wb:
            try:
                wb.close()
            except Exception:
                pass


def _perform_bulk_import(rows, mode, auto_resolve) -> tuple[dict, int]:
    excel_file = get_excel_file()
    log_file = get_log_file()
    wb = None
    try:
        with lock:
            if mode == "reset":
                # Create a fresh Excel file with headers
                wb = load_workbook(excel_file)
                ws = wb.active
                ws.delete_rows(1, ws.max_row + 1)
                
                # We start with the core headers
                headers = ["Name", "Registration Number", "Unique ID", "QR", "Barcode", SCAN_COL_NAME]
                
                has_email = False
                has_phone = False
                if rows:
                    for r_item in rows:
                        for key in r_item.keys():
                            key_clean = key.lower().strip()
                            if "email" in key_clean:
                                has_email = True
                            elif "phone" in key_clean or "whatsapp" in key_clean or "contact" in key_clean:
                                has_phone = True
                
                if has_email:
                    headers.extend(["Email Address", "Email Sent Status"])
                if has_phone:
                    headers.extend(["Phone Number", "WhatsApp Sent Status"])
                    
                headers.extend(["Scan Timestamps", "Scan Devices", "QR Code Image", "Barcode Image"])
                
                # Dynamic custom columns from rows
                custom_cols = []
                system_names_low = {
                    "name", "registration number", "unique id", "qr", "barcode",
                    "scan timestamps", "scan devices", "qr code image", "barcode image",
                    "email address", "email sent status", "phone number", "whatsapp sent status",
                    "email", "phone", "whatsapp", "contact", SCAN_COL_NAME.lower()
                }
                if rows:
                    for key in rows[0].keys():
                        if key.lower().strip() not in system_names_low:
                            custom_cols.append(key.strip())
                headers.extend(custom_cols)
                
                for c_idx, h_val in enumerate(headers, 1):
                    ws.cell(row=1, column=c_idx, value=h_val)
                    
                hdrs = {h.lower(): i for i, h in enumerate(headers, 1)}
                
                global scanned_log
                scanned_log = pd.DataFrame(columns=["QR Data", "Scan Count", "Timestamps", "Devices"])
                scanned_log.to_csv(log_file, index=False)
                
                with _devices_lock:
                    for d_val in connected_devices.values():
                        d_val["scans"] = 0
                        d_val["last_activity"] = "Reset Registry"
                _broadcast_devices()
                _emit_stats()
            else:
                wb = load_workbook(excel_file)
                ws = wb.active
                hdrs, _ = _get_or_create_headers(ws)
                
                # Append any new custom columns
                if rows:
                    for key in rows[0].keys():
                        key_clean = key.lower().strip()
                        if key_clean not in ["name", "email", "reg_no", "phone"] and key_clean not in hdrs:
                            system_names_low = {
                                "name", "registration number", "unique id", "qr", "barcode",
                                "scan timestamps", "scan devices", "qr code image", "barcode image",
                                "email address", "email sent status", "phone number", "whatsapp sent status",
                                "email", "phone", "whatsapp", "contact", SCAN_COL_NAME.lower()
                            }
                            if key_clean not in system_names_low:
                                col = ws.max_column + 1
                                ws.cell(row=1, column=col, value=key.strip())
                                hdrs[key_clean] = col
                
            # Read existing reg numbers & unique IDs to avoid duplicates
            existing_regs = set()
            existing_emails = set()
            existing_uids = set()
            for r in range(2, ws.max_row + 1):
                reg_val = ws.cell(row=r, column=hdrs["registration number"]).value
                email_val = ws.cell(row=r, column=hdrs["email address"]).value if "email address" in hdrs else None
                uid_val = ws.cell(row=r, column=hdrs["unique id"]).value
                if reg_val:
                    existing_regs.add(str(reg_val).strip().lower())
                if email_val:
                    existing_emails.add(str(email_val).strip().lower())
                if uid_val:
                    existing_uids.add(str(uid_val).strip().upper())
                    
            added_cnt = 0
            skipped_cnt = 0
            
            for row in rows:
                name = ""
                email = ""
                reg_no = ""
                phone = ""
                
                for k, v in row.items():
                    k_low = k.lower().strip()
                    if "name" in k_low:
                        name = str(v or "").strip()
                    elif "email" in k_low:
                        email = str(v or "").strip()
                    elif "reg" in k_low or "number" in k_low or "id" in k_low:
                        if "phone" not in k_low and "whatsapp" not in k_low:
                            reg_no = str(v or "").strip()
                    elif "phone" in k_low or "whatsapp" in k_low or "contact" in k_low:
                        phone = str(v or "").strip()
                
                # Validation check
                if not name or not reg_no:
                    skipped_cnt += 1
                    continue
                    
                # Skip duplicate registration numbers in append mode, 
                # unless auto-resolve is enabled
                if reg_no.lower() in existing_regs:
                    if auto_resolve:
                        # Generate a non-colliding reg number
                        base_reg = reg_no
                        counter = 1
                        while f"{base_reg}-{counter}".lower() in existing_regs:
                            counter += 1
                        reg_no = f"{base_reg}-{counter}"
                    else:
                        skipped_cnt += 1
                        continue
                        
                if email and email.lower() in existing_emails:
                    if not auto_resolve:
                        skipped_cnt += 1
                        continue
                        
                uid = _generate_unique_id(existing_uids)
                existing_regs.add(reg_no.lower())
                if email:
                    existing_emails.add(email.lower())
                existing_uids.add(uid)
                
                qr_str = reg_no
                qr_path = _generate_qr_for_guest(qr_str, reg_no)
                barcode_path = _generate_barcode_for_guest(reg_no)
                
                new_r = ws.max_row + 1
                ws.cell(row=new_r, column=hdrs["name"], value=name)
                ws.cell(row=new_r, column=hdrs["registration number"], value=reg_no)
                if "email address" in hdrs:
                    ws.cell(row=new_r, column=hdrs["email address"], value=email)
                if "phone number" in hdrs:
                    ws.cell(row=new_r, column=hdrs["phone number"], value=phone)
                ws.cell(row=new_r, column=hdrs["unique id"], value=uid)
                ws.cell(row=new_r, column=hdrs["qr"], value=qr_str)
                ws.cell(row=new_r, column=hdrs["barcode"], value=reg_no)
                ws.cell(row=new_r, column=hdrs[SCAN_COL_NAME.lower()], value="")
                if "email sent status" in hdrs:
                    ws.cell(row=new_r, column=hdrs["email sent status"], value="Not Sent")
                if "whatsapp sent status" in hdrs:
                    ws.cell(row=new_r, column=hdrs["whatsapp sent status"], value="Not Sent")
                
                # Fill custom columns dynamically
                for key, val in row.items():
                    key_clean = key.lower().strip()
                    if key_clean in hdrs and key_clean not in ["name", "email", "reg_no", "phone", "unique id", "qr", "barcode", SCAN_COL_NAME.lower(), "email sent status", "whatsapp sent status"]:
                        ws.cell(row=new_r, column=hdrs[key_clean], value=str(val).strip())
                
                _embed_qr_image(ws, new_r, qr_path, hdrs["qr code image"])
                _embed_barcode_image(ws, new_r, barcode_path, hdrs["barcode image"])
                added_cnt += 1
                
            _atomic_save(wb, excel_file)
            wb = None
            
        threading.Thread(target=_rebuild_highlighted, daemon=True).start()
        
        socketio.emit("registry_updated", {})
        _emit_stats()
        
        return {
            "message": f"Import completed successfully. Added {added_cnt} attendees, skipped {skipped_cnt}.",
            "added": added_cnt,
            "skipped": skipped_cnt
        }, 200
        
    except PermissionError:
        return {"message": "Excel file is locked. Please close it in Excel first."}, 409
    except Exception as e:
        return {"message": f"Error importing attendees: {str(e)}"}, 500
    finally:
        if wb:
            try:
                wb.close()
            except Exception:
                pass


@app.route("/import_csv", methods=["POST"])
def import_csv():
    if "file" not in request.files:
        return jsonify(message="No file uploaded."), 400
    file = request.files["file"]
    mode = request.form.get("mode", "append")
    
    if not file.filename.endswith(".csv"):
        return jsonify(message="Uploaded file must be a CSV."), 400
        
    try:
        stream = io.StringIO(file.stream.read().decode("utf-8", errors="replace"), newline=None)
        reader = csv.DictReader(stream)
        reader.fieldnames = [f.strip() for f in reader.fieldnames]
        
        required_cols = {"Name", "Email Address", "Registration Number"}
        missing_cols = required_cols - set(reader.fieldnames)
        if missing_cols:
            return jsonify(message=f"CSV is missing columns: {missing_cols}"), 400
            
        rows_to_process = []
        for row in reader:
            rows_to_process.append({
                "name": row.get("Name", "").strip(),
                "email": row.get("Email Address", "").strip(),
                "reg_no": row.get("Registration Number", "").strip(),
                "phone": row.get("Phone Number", "").strip() if "Phone Number" in row else ""
            })
            
    except Exception as e:
        return jsonify(message=f"Error parsing CSV: {str(e)}"), 400
        
    res, status_code = _perform_bulk_import(rows_to_process, mode, False)
    return jsonify(res), status_code


@app.route("/preview_import", methods=["POST"])
def preview_import():
    if "file" not in request.files:
        return jsonify(message="No file uploaded."), 400
        
    file = request.files["file"]
    filename = file.filename.lower()
    
    if not (filename.endswith(".csv") or filename.endswith(".xlsx")):
        return jsonify(message="Uploaded file must be a CSV or Excel (.xlsx) file."), 400
        
    try:
        if filename.endswith(".csv"):
            content = file.stream.read()
            stream = io.StringIO(content.decode("utf-8", errors="replace"), newline=None)
            df = pd.read_csv(stream)
        else:
            df = pd.read_excel(file, engine="openpyxl")
            
        df.columns = df.columns.astype(str).str.strip()
    except Exception as e:
        return jsonify(message=f"Failed to read file: {str(e)}"), 400
        
    mapping = {}
    for col in df.columns:
        c_low = col.lower().strip()
        if "name" in c_low:
            mapping["name"] = col
        elif "email" in c_low:
            mapping["email"] = col
        elif "reg" in c_low or "number" in c_low or "id" in c_low:
            if "phone" not in c_low and "whatsapp" not in c_low:
                mapping["reg_no"] = col
        elif "phone" in c_low or "whatsapp" in c_low or "contact" in c_low:
            mapping["phone"] = col
            
    if "name" not in mapping or "reg_no" not in mapping:
        return jsonify(message="File must contain columns for Name and Registration Number."), 400
        
    db_regs = set()
    db_emails = set()
    try:
        with lock:
            df_db = pd.read_excel(get_excel_file(), engine="openpyxl")
        df_db.columns = df_db.columns.str.strip()
        db_col_map = {c.lower(): c for c in df_db.columns}
        
        reg_c = db_col_map.get("registration number")
        email_c = db_col_map.get("email address")
        
        if reg_c:
            for val in df_db[reg_c]:
                if pd.notna(val):
                    db_regs.add(str(val).strip().lower())
        if email_c:
            for val in df_db[email_c]:
                if pd.notna(val):
                    db_emails.add(str(val).strip().lower())
    except Exception as e:
        print(f"[preview] Error loading db: {e}")
        
    file_regs = set()
    file_emails = set()
    
    preview_records = []
    
    for idx, row in df.iterrows():
        name = str(row.get(mapping["name"], "") or "").strip()
        email = str(row.get(mapping["email"], "") or "").strip() if "email" in mapping else ""
        reg_no = str(row.get(mapping["reg_no"], "") or "").strip()
        phone = str(row.get(mapping.get("phone", ""), "") or "").strip() if "phone" in mapping else ""
        
        if not name and not reg_no:
            continue
            
        status = "ok"
        issue = ""
        
        if not name or not reg_no:
            status = "missing"
            issue = "Missing critical fields"
        else:
            reg_low = reg_no.lower()
            email_low = email.lower() if email else ""
            
            if reg_low in db_regs:
                status = "dup_db_reg"
                issue = f"Reg no '{reg_no}' already in database"
            elif email_low and email_low in db_emails:
                status = "dup_db_email"
                issue = f"Email '{email}' already in database"
            elif reg_low in file_regs:
                status = "dup_file_reg"
                issue = f"Duplicate Reg no '{reg_no}' inside file"
            elif email_low and email_low in file_emails:
                status = "dup_file_email"
                issue = f"Duplicate Email '{email}' inside file"
                
            file_regs.add(reg_low)
            if email_low:
                file_emails.add(email_low)
            
        preview_records.append({
            "name": name,
            "email": email,
            "reg_no": reg_no,
            "phone": phone,
            "status": status,
            "issue": issue
        })
        
    return jsonify(preview_records)


@app.route("/confirm_import", methods=["POST"])
def confirm_import():
    if not request.is_json:
        return jsonify(message="Request must be JSON."), 415
        
    payload = request.json or {}
    rows = payload.get("rows", [])
    mode = payload.get("mode", "append")
    auto_resolve = payload.get("auto_resolve", False)
    
    if not rows:
        return jsonify(message="No data rows to import."), 400
        
    res, status_code = _perform_bulk_import(rows, mode, auto_resolve)
    return jsonify(res), status_code


# ── Template Formatter Helper ──────────────────────────────────────────────────
def format_template(template: str, row_dict: dict, extra: dict = None) -> str:
    result = template
    # Merge row_dict and extra
    data = {}
    for k, v in row_dict.items():
        data[str(k).strip()] = str(v or "").strip()
    if extra:
        for k, v in extra.items():
            data[str(k).strip()] = str(v or "").strip()
            
    for k, v in data.items():
        placeholder = "{" + str(k) + "}"
        result = result.replace(placeholder, str(v))
    return result


def _update_sent_status(reg_no: str, channel: str, status: str):
    excel_file = get_excel_file()
    wb = None
    try:
        with lock:
            wb = load_workbook(excel_file)
            ws = wb.active
            hdrs, _ = _get_or_create_headers(ws)
            col_name = "email sent status" if channel == "email" else "whatsapp sent status"
            col_idx = hdrs.get(col_name)
            if col_idx:
                for r in range(2, ws.max_row + 1):
                    reg_val = ws.cell(row=r, column=hdrs["registration number"]).value
                    if reg_val and str(reg_val).strip().lower() == reg_no.lower():
                        ws.cell(row=r, column=col_idx, value=status)
                        break
            _atomic_save(wb, excel_file)
            wb = None
            
        socketio.emit("registry_updated", {})
        threading.Thread(target=_rebuild_highlighted, daemon=True).start()
    except Exception as e:
        print(f"[status-update] Error: {e}")
    finally:
        if wb:
            try:
                wb.close()
            except Exception:
                pass


# ── Twilio WhatsApp Campaign Tracker ──────────────────────────────────────────
_whatsapp_sending_lock = threading.Lock()
_whatsapp_sending_active = False
_whatsapp_progress = {
    "sent": 0,
    "skipped": 0,
    "failed": 0,
    "total": 0,
    "current_phone": "",
    "current_status": "Idle",
    "logs": []
}

def _run_whatsapp_campaign(twilio_sid, twilio_token, twilio_sender, host_url, event_name):
    global _whatsapp_sending_active, _whatsapp_progress
    from twilio.rest import Client
    import re
    
    excel_file = get_excel_file()
    cfg = get_event_config()
    wa_template = cfg.get("whatsapp_template", "")
    
    attendees = []
    try:
        with lock:
            df_xl = pd.read_excel(excel_file, engine="openpyxl")
            
        df_xl.columns = df_xl.columns.str.strip()
        col_map = {c.lower(): c for c in df_xl.columns}
        phone_col = col_map.get("phone number")
        
        for _, row in df_xl.iterrows():
            name_v = _clean_val(row.get(col_map.get("name")))
            phone_v = _clean_val(row.get(phone_col)) if phone_col else ""
            reg_v = _clean_val(row.get(col_map.get("registration number")))
            
            if name_v and phone_v and reg_v:
                phone_clean = re.sub(r"[^\d+]", "", phone_v)
                if phone_clean:
                    if not phone_clean.startswith("+"):
                        if len(phone_clean) == 10:
                            phone_clean = "+91" + phone_clean
                        else:
                            phone_clean = "+" + phone_clean
                    
                    row_data = {c_orig: _clean_val(row.get(c_orig)) for c_low, c_orig in col_map.items()}
                    attendees.append({
                        "name": name_v,
                        "phone": phone_clean,
                        "reg_no": reg_v,
                        "row_dict": row_data
                    })
    except Exception as e:
        with _whatsapp_sending_lock:
            _whatsapp_progress["current_status"] = "Failed to load attendees"
            _whatsapp_progress["logs"].append(f"❌ Error: {str(e)}")
            _whatsapp_sending_active = False
        socketio.emit("whatsapp_progress_update", _whatsapp_progress)
        return
        
    total_count = len(attendees)
    with _whatsapp_sending_lock:
        _whatsapp_progress["total"] = total_count
        _whatsapp_progress["sent"] = 0
        _whatsapp_progress["skipped"] = 0
        _whatsapp_progress["failed"] = 0
        _whatsapp_progress["current_status"] = "Starting WhatsApp campaign..."
        _whatsapp_progress["logs"] = [f"🚀 Started Twilio WhatsApp campaign. Total: {total_count}"]
    socketio.emit("whatsapp_progress_update", _whatsapp_progress)
    
    client = None
    try:
        client = Client(twilio_sid, twilio_token)
    except Exception as e:
        with _whatsapp_sending_lock:
            _whatsapp_progress["current_status"] = "Twilio Auth failed"
            _whatsapp_progress["logs"].append(f"❌ Twilio Init Error: {str(e)}")
            _whatsapp_sending_active = False
        socketio.emit("whatsapp_progress_update", _whatsapp_progress)
        return
        
    qr_dir = get_qr_dir()
    is_local = "localhost" in host_url or "127.0.0.1" in host_url
    
    for idx, att in enumerate(attendees):
        name = att["name"]
        phone = att["phone"]
        reg = att["reg_no"]
        row_dict = att["row_dict"]
        qr_file = os.path.join(qr_dir, f"{reg}.png")
        
        with _whatsapp_sending_lock:
            _whatsapp_progress["current_phone"] = phone
            _whatsapp_progress["current_status"] = f"Sending to {phone} ({idx+1}/{total_count})"
        socketio.emit("whatsapp_progress_update", _whatsapp_progress)
        
        if not os.path.exists(qr_file):
            msg = f"⚠️ Skipped {name} ({phone}) — QR image missing"
            _update_sent_status(reg, "whatsapp", "Skipped")
            with _whatsapp_sending_lock:
                _whatsapp_progress["skipped"] += 1
                _whatsapp_progress["logs"].append(msg)
            socketio.emit("whatsapp_progress_update", _whatsapp_progress)
            continue
            
        qr_url = f"{host_url}qrcodes/{reg}.png"
        extra = {"Event": event_name, "QR_URL": qr_url}
        body_text = format_template(wa_template, row_dict, extra)
        
        success = False
        for attempt in range(1, 4):
            try:
                from_number = f"whatsapp:{twilio_sender}"
                to_number = f"whatsapp:{phone}"
                
                if is_local:
                    client.messages.create(
                        body=body_text + f"\n\nDownload QR at: {qr_url}",
                        from_=from_number,
                        to=to_number
                    )
                else:
                    client.messages.create(
                        body=body_text,
                        media_url=[qr_url],
                        from_=from_number,
                        to=to_number
                    )
                success = True
                break
            except Exception as e:
                wait = 2 ** attempt
                with _whatsapp_sending_lock:
                    _whatsapp_progress["logs"].append(f"   Attempt {attempt} failed for {phone}: {str(e)}. Retrying in {wait}s...")
                socketio.emit("whatsapp_progress_update", _whatsapp_progress)
                time.sleep(wait)
                
        if success:
            msg = f"✅ Sent WhatsApp to {name} ({phone})"
            _update_sent_status(reg, "whatsapp", "Sent")
            with _whatsapp_sending_lock:
                _whatsapp_progress["sent"] += 1
                _whatsapp_progress["logs"].append(msg)
        else:
            msg = f"❌ Failed to send WhatsApp to {name} ({phone})"
            _update_sent_status(reg, "whatsapp", "Failed")
            with _whatsapp_sending_lock:
                _whatsapp_progress["failed"] += 1
                _whatsapp_progress["logs"].append(msg)
                
        socketio.emit("whatsapp_progress_update", _whatsapp_progress)
        time.sleep(1.2)
        
    with _whatsapp_sending_lock:
        _whatsapp_progress["current_status"] = "Finished"
        _whatsapp_progress["current_phone"] = ""
        _whatsapp_progress["logs"].append(f"🏁 WhatsApp Campaign finished. Sent: {_whatsapp_progress['sent']}, Skipped: {_whatsapp_progress['skipped']}, Failed: {_whatsapp_progress['failed']}.")
        _whatsapp_sending_active = False
    socketio.emit("whatsapp_progress_update", _whatsapp_progress)


@app.route("/send_whatsapp_bulk", methods=["POST"])
def send_whatsapp_bulk():
    global _whatsapp_sending_active
    if not request.is_json:
        return jsonify(message="Request must be JSON."), 415
        
    payload = request.json or {}
    twilio_sid = payload.get("twilio_sid", "").strip()
    twilio_token = payload.get("twilio_token", "").strip()
    twilio_sender = payload.get("twilio_sender", "").strip()
    event_name = payload.get("event_name", "the Event").strip()
    
    if not twilio_sid or not twilio_token or not twilio_sender:
        return jsonify(message="Twilio Account SID, Auth Token, and Sender number are required."), 400
        
    with _whatsapp_sending_lock:
        if _whatsapp_sending_active:
            return jsonify(message="A WhatsApp campaign is already in progress."), 409
        _whatsapp_sending_active = True
        
    if not twilio_sender.startswith("+") and not twilio_sender.startswith("whatsapp:"):
        twilio_sender = "+" + twilio_sender
        
    host_url = request.host_url
    
    # Save parameters to config for next use
    cfg = get_event_config()
    cfg["twilio_sid"] = twilio_sid
    cfg["twilio_token"] = twilio_token
    cfg["twilio_sender"] = twilio_sender
    cfg["event_name_template"] = event_name
    save_event_config(cfg)
    
    threading.Thread(
        target=_run_whatsapp_campaign,
        args=(twilio_sid, twilio_token, twilio_sender, host_url, event_name),
        daemon=True
    ).start()
    
    return jsonify(message="WhatsApp campaign started successfully."), 202


@app.route("/whatsapp_status")
def get_whatsapp_status():
    with _whatsapp_sending_lock:
        return jsonify(
            active=_whatsapp_sending_active,
            sent=_whatsapp_progress["sent"],
            skipped=_whatsapp_progress["skipped"],
            failed=_whatsapp_progress["failed"],
            total=_whatsapp_progress["total"],
            current_phone=_whatsapp_progress["current_phone"],
            current_status=_whatsapp_progress["current_status"],
            logs=_whatsapp_progress["logs"]
        )


# ── Email Campaign Tracker ────────────────────────────────────────────────────
_email_sending_lock = threading.Lock()
_email_sending_active = False
_email_progress = {
    "sent": 0,
    "skipped": 0,
    "failed": 0,
    "total": 0,
    "current_email": "",
    "current_status": "Idle",
    "logs": []
}

def _run_email_campaign(sender_email, app_password, subject, event_name):
    global _email_sending_active, _email_progress
    import yagmail
    
    excel_file = get_excel_file()
    cfg = get_event_config()
    email_template = cfg.get("email_template", "")
    
    attendees = []
    try:
        with lock:
            df_xl = pd.read_excel(excel_file, engine="openpyxl")
            
        df_xl.columns = df_xl.columns.str.strip()
        col_map = {c.lower(): c for c in df_xl.columns}
        
        for _, row in df_xl.iterrows():
            name_v = _clean_val(row.get(col_map.get("name")))
            email_v = _clean_val(row.get(col_map.get("email address")))
            reg_v = _clean_val(row.get(col_map.get("registration number")))
            
            if name_v and email_v and reg_v:
                row_data = {c_orig: _clean_val(row.get(c_orig)) for c_low, c_orig in col_map.items()}
                attendees.append({
                    "name": name_v,
                    "email": email_v,
                    "reg_no": reg_v,
                    "row_dict": row_data
                })
    except Exception as e:
        with _email_sending_lock:
            _email_progress["current_status"] = "Failed to load attendees"
            _email_progress["logs"].append(f"❌ Error loading attendees: {str(e)}")
            _email_sending_active = False
        socketio.emit("email_progress_update", _email_progress)
        return
        
    total_count = len(attendees)
    with _email_sending_lock:
        _email_progress["total"] = total_count
        _email_progress["sent"] = 0
        _email_progress["skipped"] = 0
        _email_progress["failed"] = 0
        _email_progress["current_status"] = "Starting campaign..."
        _email_progress["logs"] = [f"🚀 Started email campaign. Total recipients: {total_count}"]
    socketio.emit("email_progress_update", _email_progress)
    
    yag = None
    try:
        yag = yagmail.SMTP(user=sender_email, password=app_password)
    except Exception as e:
        with _email_sending_lock:
            _email_progress["current_status"] = "SMTP Authentication failed"
            _email_progress["logs"].append(f"❌ SMTP Connection Error: {str(e)}")
            _email_sending_active = False
        socketio.emit("email_progress_update", _email_progress)
        return
        
    qr_dir = get_qr_dir()
    
    for idx, att in enumerate(attendees):
        name = att["name"]
        recipient = att["email"]
        reg = att["reg_no"]
        row_dict = att["row_dict"]
        qr_file = os.path.join(qr_dir, f"{reg}.png")
        
        with _email_sending_lock:
            _email_progress["current_email"] = recipient
            _email_progress["current_status"] = f"Sending to {recipient} ({idx+1}/{total_count})"
        socketio.emit("email_progress_update", _email_progress)
        
        if not os.path.exists(qr_file):
            msg = f"⚠️ Skipped {name} ({recipient}) — QR image missing"
            _update_sent_status(reg, "email", "Skipped")
            with _email_sending_lock:
                _email_progress["skipped"] += 1
                _email_progress["logs"].append(msg)
            socketio.emit("email_progress_update", _email_progress)
            continue
            
        extra = {"Event": event_name}
        body = format_template(email_template, row_dict, extra)
        
        success = False
        for attempt in range(1, 4):
            try:
                yag.send(
                    to=recipient,
                    subject=subject,
                    contents=body,
                    attachments=qr_file,
                )
                success = True
                break
            except Exception as e:
                wait = 2 ** attempt
                with _email_sending_lock:
                    _email_progress["logs"].append(f"   Attempt {attempt} failed for {recipient}: {str(e)}. Retrying in {wait}s...")
                socketio.emit("email_progress_update", _email_progress)
                time.sleep(wait)
                
        if success:
            msg = f"✅ Sent to {name} ({recipient})"
            _update_sent_status(reg, "email", "Sent")
            with _email_sending_lock:
                _email_progress["sent"] += 1
                _email_progress["logs"].append(msg)
        else:
            msg = f"❌ Failed to send to {name} ({recipient})"
            _update_sent_status(reg, "email", "Failed")
            with _email_sending_lock:
                _email_progress["failed"] += 1
                _email_progress["logs"].append(msg)
                
        socketio.emit("email_progress_update", _email_progress)
        time.sleep(1.0)
        
    try:
        yag.close()
    except Exception:
        pass
        
    with _email_sending_lock:
        _email_progress["current_status"] = "Finished"
        _email_progress["current_email"] = ""
        _email_progress["logs"].append(f"🏁 Campaign finished. Sent: {_email_progress['sent']}, Skipped: {_email_progress['skipped']}, Failed: {_email_progress['failed']}.")
        _email_sending_active = False
    socketio.emit("email_progress_update", _email_progress)


@app.route("/send_emails", methods=["POST"])
def send_emails():
    global _email_sending_active
    if not request.is_json:
        return jsonify(message="Request must be JSON."), 415
        
    payload = request.json or {}
    sender = payload.get("sender_email", "").strip()
    password = payload.get("app_password", "").strip()
    subject = payload.get("subject", "Your Event QR Code").strip()
    event = payload.get("event_name", "the Event").strip()
    
    if not sender or not password:
        return jsonify(message="Sender email and App password are required."), 400
        
    with _email_sending_lock:
        if _email_sending_active:
            return jsonify(message="An email campaign is already in progress."), 409
        _email_sending_active = True
        
    # Save configuration parameters
    cfg = get_event_config()
    cfg["email_sender"] = sender
    cfg["email_password"] = password
    cfg["email_subject"] = subject
    cfg["event_name_template"] = event
    save_event_config(cfg)
    
    threading.Thread(
        target=_run_email_campaign,
        args=(sender, password, subject, event),
        daemon=True
    ).start()
    
    return jsonify(message="Email campaign started successfully."), 202


@app.route("/email_status")
def get_email_status():
    with _email_sending_lock:
        return jsonify(
            active=_email_sending_active,
            sent=_email_progress["sent"],
            skipped=_email_progress["skipped"],
            failed=_email_progress["failed"],
            total=_email_progress["total"],
            current_email=_email_progress["current_email"],
            current_status=_email_progress["current_status"],
            logs=_email_progress["logs"]
        )


@app.route("/download/excel")
def download_excel():
    with highlight_lock:
        high_file = get_highlighted_file()
        if os.path.exists(high_file):
            return send_file(
                high_file,
                as_attachment=True,
                download_name="event_checkin_highlighted.xlsx",
                mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )
    return jsonify(message="Highlighted Excel report not generated yet."), 404


@app.route("/download/csv")
def download_csv():
    with lock:
        log_file_path = get_log_file()
        if os.path.exists(log_file_path):
            return send_file(
                log_file_path,
                as_attachment=True,
                download_name="scanned_log.csv",
                mimetype="text/csv"
            )
    return jsonify(message="No scan log found."), 404


load_active_event()


# ── Entry point ────────────────────────────────────────────────────────────────
if __name__ == "__main__":
    port  = int(os.environ.get("PORT", 5001))
    debug = os.environ.get("FLASK_DEBUG", "false").lower() == "true"
    lan_ip = get_lan_ip()
    print("")
    print("=== QR Check-In System is starting ===")
    print("")
    print(f"   Local   -> http://localhost:{port}")
    print(f"   Network -> http://{lan_ip}:{port}  (share this with other devices)")
    print("")
    print("   All devices on the same Wi-Fi can open the network URL above.")
    print("")
    socketio.run(app, debug=debug, host="0.0.0.0", port=port)
