import eventlet
eventlet.monkey_patch()

import unittest
import json
import os
import shutil
import pandas as pd
from openpyxl import Workbook
from datetime import datetime, timedelta

# Mock debug mode for app import
os.environ["FLASK_DEBUG"] = "false"
import app

class TestUpgradedFeatures(unittest.TestCase):
    def setUp(self):
        app.app.config["TESTING"] = True
        self.app_client = app.app.test_client()
        self.excel_file = "test_event_data.xlsx"
        
        # Create a mock spreadsheet
        wb = Workbook()
        ws = wb.active
        ws.title = "Roster"
        # Headers
        ws.append(["Name", "Email Address", "Registration Number", "Phone Number", "Unique ID", "QR", "Level", "Scan Status"])
        # Rows
        ws.append(["Alice Doe", "alice@example.com", "REG001", "+1234567890", "UID001", "Alice Doe | REG001", "VIP", "Not Scanned"])
        ws.append(["Bob Smith", "bob@example.com", "REG002", "+1234567891", "UID002", "Bob Smith | REG002", "Guest", "Not Scanned"])
        wb.save(self.excel_file)
        
        # Override paths in app module
        self.original_get_excel = app.get_excel_file
        self.original_get_config = app.get_config_file
        app.get_excel_file = lambda: self.excel_file
        app.get_config_file = lambda: "test_config.json"
        
        # Initialise configuration
        self.config_file = "test_config.json"
        self.test_config = app.get_event_config()

    def tearDown(self):
        # Restore original settings
        app.get_excel_file = self.original_get_excel
        app.get_config_file = self.original_get_config
        
        if os.path.exists(self.excel_file):
            try:
                os.remove(self.excel_file)
            except Exception:
                pass
        if os.path.exists(self.config_file):
            try:
                os.remove(self.config_file)
            except Exception:
                pass

    def test_subgroup_rename(self):
        # Rename "VIP" to "Super VIP" in column "Level"
        payload = {
            "column": "Level",
            "old_value": "VIP",
            "new_value": "Super VIP"
        }
        res = self.app_client.post("/rename_subgroup", json=payload)
        self.assertEqual(res.status_code, 200)
        data = res.get_json()
        self.assertIn("Successfully renamed", data["message"])
        
        # Verify Excel sheet has changed
        df = pd.read_excel(self.excel_file, engine="openpyxl")
        self.assertEqual(df.loc[df["Registration Number"] == "REG001", "Level"].values[0], "Super VIP")
        self.assertEqual(df.loc[df["Registration Number"] == "REG002", "Level"].values[0], "Guest")

    def test_id_card_preview(self):
        res = self.app_client.get("/preview_id_card?theme=cyber_neon")
        self.assertEqual(res.status_code, 200)
        self.assertEqual(res.mimetype, "image/png")

    def test_checkin_date_time_restrictions(self):
        # Configure checkin date restriction in config to be in the future
        cfg = app.get_event_config()
        future_start = datetime.now() + timedelta(days=1)
        cfg["checkin_start_date"] = future_start.strftime("%Y-%m-%d")
        cfg["checkin_start_time"] = future_start.strftime("%H:%M")
        app.save_event_config(cfg)
        
        # Attempt checkin
        res = self.app_client.post("/scan", json={
            "qr_data": "Alice Doe | REG001",
            "device_id": "test_device"
        })
        self.assertEqual(res.status_code, 400)
        data = res.get_json()
        self.assertIn("Check-In Not Started", data["message"])

        # Configure checkin date restriction in config to be in the past, and checkin end in the past
        past_end = datetime.now() - timedelta(days=1)
        cfg["checkin_start_date"] = ""
        cfg["checkin_start_time"] = ""
        cfg["checkin_end_date"] = past_end.strftime("%Y-%m-%d")
        cfg["checkin_end_time"] = past_end.strftime("%H:%M")
        app.save_event_config(cfg)
        
        # Attempt checkin
        res = self.app_client.post("/scan", json={
            "qr_data": "Alice Doe | REG001",
            "device_id": "test_device"
        })
        self.assertEqual(res.status_code, 400)
        data = res.get_json()
        self.assertIn("Check-In Closed", data["message"])

    def test_phone_normalization_edge_cases(self):
        # 1. Impossible numbers
        self.assertEqual(app._normalize_phone("0000000000"), "")
        self.assertEqual(app._normalize_phone("123"), "")
        self.assertEqual(app._normalize_phone("1234567890123456"), "") # >15 digits
        
        # 2. Priorities based on config
        cfg = app.get_event_config()
        cfg["allowed_country_codes"] = "91,1"
        app.save_event_config(cfg)
        
        # Test default country code matching
        self.assertEqual(app._normalize_phone("9876543210"), "+919876543210")
        
        # Test alternate allowed country code matching
        self.assertEqual(app._normalize_phone("12345678901"), "+12345678901")

    def test_path_traversal_guards(self):
        # Path traversal checks
        res = self.app_client.get("/qrcodes/../app.py")
        self.assertEqual(res.status_code, 400)
        res = self.app_client.get("/barcodes/..%2f..%2fapp.py")
        self.assertEqual(res.status_code, 400)

    def test_quarantine_endpoints(self):
        # Add item to quarantine manually
        app._quarantine_scan("InvalidQRData", "TestScanner", "Invalid Cryptographic Signature")
        
        # 1. GET /quarantine
        res = self.app_client.get("/quarantine")
        self.assertEqual(res.status_code, 200)
        data = res.get_json()
        self.assertIn("quarantine", data)
        self.assertGreater(len(data["quarantine"]), 0)
        
        q_item = data["quarantine"][0]
        q_id = q_item["id"]
        
        # 2. POST /quarantine/approve
        res = self.app_client.post("/quarantine/approve", json={
            "id": q_id,
            "device_name": "TestDashboard"
        })
        self.assertEqual(res.status_code, 200)
        
        # Verify approved scan is no longer in quarantine
        res_after = self.app_client.get("/quarantine")
        q_after = res_after.get_json()["quarantine"]
        self.assertNotIn(q_id, [item["id"] for item in q_after])

        # Test Reject
        app._quarantine_scan("AnotherInvalid", "TestScanner", "Time Restricted")
        res_list = self.app_client.get("/quarantine")
        q_list = res_list.get_json()["quarantine"]
        another_q_id = [item["id"] for item in q_list if item["qr_data"] == "AnotherInvalid"][0]
        
        res_reject = self.app_client.post("/quarantine/reject", json={"id": another_q_id})
        self.assertEqual(res_reject.status_code, 200)
        
        res_final = self.app_client.get("/quarantine")
        self.assertNotIn(another_q_id, [item["id"] for item in res_final.get_json()["quarantine"]])

    def test_audit_log_and_integrity(self):
        # Trigger an administrative action to write to the audit log
        app._log_audit("Test Admin Action", "Details of test action", "TestDevice")
        
        audit_file = os.path.join(app.get_active_event_path(), "audit_log.csv")
        self.assertTrue(os.path.exists(audit_file))
        
        # Verify integrity chain
        is_valid, failing_row = app._verify_audit_log_integrity(audit_file)
        self.assertTrue(is_valid)
        self.assertEqual(failing_row, 0)
        
        # Check export audit log route
        res = self.app_client.get("/export_audit_log")
        self.assertEqual(res.status_code, 200)
        self.assertEqual(res.mimetype, "text/csv")

    def test_timeline_stats(self):
        res = self.app_client.get("/stats/timeline")
        self.assertEqual(res.status_code, 200)
        data = res.get_json()
        self.assertIn("timeline", data)

    def test_checkin_revocation(self):
        # Check-in Alice first
        self.app_client.post("/manual_checkin", json={
            "qr_data": "Alice Doe | REG001",
            "device_id": "test_device"
        })
        
        # Revoke check-in
        res = self.app_client.post("/revoke_checkin", json={
            "reg_no": "REG001",
            "reason": "Accidental Scan"
        })
        self.assertEqual(res.status_code, 200)
        
        # Verify Excel sheet status is empty
        df = pd.read_excel(self.excel_file, engine="openpyxl")
        alice_status = df.loc[df["Registration Number"] == "REG001", "Scan Status"].values[0]
        self.assertTrue(pd.isna(alice_status) or alice_status == "" or "not scanned" in str(alice_status).lower())

    def test_stats_timeline_intervals(self):
        for interval in ["5m", "10m", "30m", "1h", "2h", "4h", "6h", "1d"]:
            res = self.app_client.get(f"/stats/timeline?interval={interval}")
            self.assertEqual(res.status_code, 200)
            data = res.get_json()
            self.assertIn("timeline", data)

    def test_send_id_card_single(self):
        res = self.app_client.post("/send_id_card_single", json={
            "reg_no": "REG001",
            "channel": "email"
        })
        self.assertIn(res.status_code, [200, 400])

if __name__ == "__main__":
    unittest.main()
