"""Isolated regression tests. Never import the live app against its event folder.

Run: python tests/test_badge_workflows.py
Optional disposable UI sandbox: python tests/test_badge_workflows.py --serve
"""
import sys
from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(ROOT))

# Patch before Flask/threading imports, matching application startup.
import eventlet
eventlet.monkey_patch()
import io
import os
import tempfile
import types
import unittest
from PIL import Image
from openpyxl import Workbook, load_workbook
from card_design import color, field_text, render_custom, resolve_design, validate_design
from roster_schema import match_columns, is_identity_column, validate_headers


def sandbox_app():
    (ROOT / 'runtime').mkdir(exist_ok=True)
    temp = tempfile.TemporaryDirectory(prefix='badge_test_', dir=ROOT / 'runtime')
    module = types.ModuleType('badge_sandbox')
    module.__file__ = str(Path(temp.name) / 'app.py')
    sys.modules[module.__name__] = module
    exec(compile((ROOT / 'app.py').read_text(encoding='utf-8'), module.__file__, 'exec'), module.__dict__)
    assert Path(module.EVENTS_DIR).is_relative_to(temp.name)
    module.app.template_folder = str(ROOT / 'templates')
    module.app.static_folder = str(ROOT / 'static')
    return temp, module


def design():
    return {'mode': 'custom', 'bg_start': [11,22,39], 'bg_end': [27,44,64],
            'accent_color': '#68e3c0', 'border_color': 'rgba(255,255,255,0.12)',
            'photo': {'enabled': True}, 'fields': [
                {'text': '{Event}', 'x': 300, 'y': 82, 'font_size': 22, 'color': [104,227,192]},
                {'text': '{Name}', 'bind_column': 'Name', 'x': 300, 'y': 352, 'font_size': 42, 'bold': True},
                {'text': '{Registration Number}', 'x': 300, 'y': 412, 'font_size': 21},
                {'text': '{Company Name}', 'x': 300, 'y': 472, 'font_size': 22},
                {'text': 'SCAN AT CHECK-IN', 'x': 300, 'y': 839, 'font_size': 15}]}


def seed(module):
    cfg = module.get_event_config()
    cfg.update(enable_id_card_generation=False, reg_notify_email=False, reg_notify_whatsapp=False,
               reg_notify_sms=False, id_card_designer_settings={}, require_scanner_auth=False,
               event_name_template='DESIGN SUMMIT 2026', subgroup_designer_settings={},
               custom_group_designer_settings={}, group_column='')
    module.save_event_config(cfg)
    wb = Workbook()
    ws = wb.active
    ws.append(['Name', 'Registration Number', 'Email Address', 'Phone Number', 'Unique ID', 'QR', 'Company Name', 'Emergency Contact', 'Seat Number'])
    ws.append(['Alex Morgan', '0007', 'alex@example.com', '+919876543210', 'UID7', '0007', 'North Studio', 'Support Desk', 'NA'])
    wb.save(module.get_excel_file())
    wb.close()


class DesignUnitTests(unittest.TestCase):
    def test_css_colors(self):
        self.assertEqual(color('#2457a7'), (36,87,167,255))
        self.assertEqual(color('rgba(255,255,255,0.12)'), (255,255,255,31))
        self.assertEqual(color(None), (255,255,255,255))

    def test_binding_and_missing_value(self):
        self.assertEqual(field_text({'bind_column': 'Company_Name', 'text': 'Wrong'}, {'company name': 'North Studio'}), 'North Studio')
        self.assertEqual(field_text({'bind_column': 'Pass Type', 'text': 'VIP PASS'}, {}), '')
        self.assertEqual(field_text({'text': '{Name} / {Seat Number}'}, {'name': 'Alex', 'seat number': 'NA'}), 'Alex / NA')

    def test_schema_does_not_guess(self):
        mapping = match_columns(['Company Name', 'Emergency Contact', 'Seat Number', 'Name', 'Registration Number', 'Phone Number'])
        self.assertEqual(mapping, {'name': 'Name', 'reg_no': 'Registration Number', 'phone': 'Phone Number'})
        self.assertFalse(is_identity_column('Company Name'))
        self.assertFalse(is_identity_column('Emergency Contact'))
        self.assertEqual(match_columns(['Name', 'guest name'])['name'], 'Name')

    def test_duplicate_headers_rejected(self):
        for headers in ([], ['Name', 'name'], ['Name', '']):
            with self.assertRaises(ValueError):
                validate_headers(headers)

    def test_design_validation(self):
        for invalid in (None, [], {'fields': 'bad'}, {'fields': [None]}, {'photo': 'bad'}, {'mode': 'unknown'}, {'fields': [{}] * 61}):
            with self.assertRaises(ValueError):
                validate_design(invalid)

    def test_group_precedence(self):
        cfg = {'id_card_designer_settings': {'global': True}, 'group_column': 'Team',
               'subgroup_designer_settings': {'Team:Design': {'subgroup': True}},
               'custom_group_designer_settings': {'Speakers': {'named': True}}}
        row = {'Registration Number': '0007', 'Team': 'Design'}
        self.assertEqual(resolve_design(cfg, row, [{'name': 'Speakers', 'reg_nos': ['0007']}]), {'named': True})
        self.assertEqual(resolve_design(cfg, row), {'subgroup': True})
        self.assertEqual(resolve_design(cfg, {}), {'global': True})


class WorkflowTests(unittest.TestCase):
    @classmethod
    def setUpClass(cls):
        cls.temp, cls.module = sandbox_app()
        cls.module.app.config['TESTING'] = True

    @classmethod
    def tearDownClass(cls):
        eventlet.sleep(0.1)
        cls.temp.cleanup()

    def setUp(self):
        seed(self.module)
        self.client = self.module.app.test_client()

    def test_preview_valid_and_does_not_overwrite_cards(self):
        cards = Path(self.module.get_qr_dir()) / 'id_cards'
        cards.mkdir(exist_ok=True)
        saved = cards / 'id_0007.png'
        saved.write_bytes(b'previously-saved-card')
        response = self.client.post('/preview_id_card?reg_no=0007', json=design())
        self.assertEqual(response.status_code, 200, response.data[:300])
        self.assertIn('no-store', response.headers['Cache-Control'])
        image = Image.open(io.BytesIO(response.data))
        self.assertEqual(image.size, (600,900))
        self.assertEqual(image.mode, 'RGB')
        self.assertEqual(saved.read_bytes(), b'previously-saved-card')
        image.save(ROOT / 'runtime' / 'id-card-proof.png')
        # A visually plausible QR is not enough: decode the rendered pixels.
        import cv2
        import numpy as np
        decoded, _, _ = cv2.QRCodeDetector().detectAndDecode(np.array(image))
        self.assertEqual(decoded, self.module._get_qr_payload('0007', self.module.get_event_config()))

    def test_long_names_and_extreme_sizes_still_render(self):
        settings = design()
        settings['fields'][1].update(font_size=999999, text='A' * 500, bind_column='')
        settings['qr'] = {'size': 999999, 'x': -900, 'y': 9000}
        settings['border_color'] = '#2457a7'
        response = self.client.post('/preview_id_card', json=settings)
        self.assertEqual(response.status_code, 200, response.data[:300])

    def test_saved_design_download(self):
        response = self.client.post('/save_config', json={'id_card_designer_settings': design()})
        self.assertEqual(response.status_code, 200)
        response = self.client.get('/download/id_card/0007')
        self.assertEqual(response.status_code, 200, response.data[:300])
        self.assertTrue(response.data.startswith(b'%PDF'))
        response.close()

    def test_preview_errors(self):
        self.assertEqual(self.client.post('/preview_id_card', json={'fields': 'bad'}).status_code, 400)
        self.assertEqual(self.client.post('/preview_id_card?reg_no=missing', json=design()).status_code, 404)

    def test_standard_preview_is_isolated(self):
        response = self.client.get('/preview_id_card?theme=classic_corporate&reg_no=0007')
        self.assertEqual(response.status_code, 200, response.data[:300])
        self.assertEqual(Image.open(io.BytesIO(response.data)).size, (600,900))
        self.assertFalse((Path(self.module.get_qr_dir()) / 'id_cards' / 'id_DEMO12345.png').exists())

    def test_text_roster_and_custom_columns(self):
        rows = self.client.get('/registry').get_json()
        self.assertEqual(rows[0]['reg_no'], '0007')
        self.assertEqual(rows[0]['phone'], '+919876543210')
        self.assertEqual(rows[0]['custom_fields']['Seat Number'], 'NA')
        columns = self.client.get('/active_event_columns').get_json()['columns']
        self.assertIn('Company Name', columns)
        self.assertNotIn('SMS Sent Status', columns)

    def test_import_preview_confirm_keeps_project_fields(self):
        content = '\ufeffName,Registration Number,Email Address,Phone Number,Company Name,Emergency Contact,Seat Number\nTaylor Reed,0008,taylor@example.com,+919876543211,West Studio,Main Desk,NA\n'
        response = self.client.post('/preview_import', data={'file': (io.BytesIO(content.encode()), 'roster.csv')})
        self.assertEqual(response.status_code, 200, response.data)
        preview = response.get_json()
        self.assertEqual(preview[0]['name'], 'Taylor Reed')
        self.assertEqual(preview[0]['reg_no'], '0008')
        self.assertEqual(preview[0]['phone'], '+919876543211')
        result = self.client.post('/confirm_import', json={'rows': preview, 'mode': 'append'})
        self.assertEqual(result.status_code, 200, result.data)
        self.assertEqual(result.get_json()['added'], 1)
        rows = self.client.get('/registry').get_json()
        added = next(row for row in rows if row['reg_no'] == '0008')
        self.assertEqual(added['custom_fields']['Company Name'], 'West Studio')
        self.assertEqual(added['custom_fields']['Emergency Contact'], 'Main Desk')
        self.assertEqual(added['custom_fields']['Seat Number'], 'NA')
        self.assertNotIn('raw', added['custom_fields'])
        repeated = self.client.post('/confirm_import', json={'rows': preview, 'mode': 'append'}).get_json()
        self.assertEqual(repeated['added'], 0)
        self.assertEqual(repeated['skipped'], 1)

    def test_duplicate_import_headers_rejected(self):
        response = self.client.post('/preview_import', data={'file': (io.BytesIO(b'Name,Name,Registration Number\nA,B,8'), 'roster.csv')})
        self.assertEqual(response.status_code, 400)

    def test_invalid_save_rejected(self):
        response = self.client.post('/save_config', json={'id_card_designer_settings': {'fields': 'bad'}})
        self.assertEqual(response.status_code, 400)

    def test_remote_preview_requires_login(self):
        response = self.client.post('/preview_id_card', json=design(), headers={'X-Forwarded-For': '203.0.113.7'})
        self.assertIn(response.status_code, (401,403))

    def test_socket_respects_scanner_passcode(self):
        cfg = self.module.get_event_config()
        cfg['require_scanner_auth'] = True
        self.module.save_event_config(cfg)
        socket = self.module.socketio.test_client(self.module.app, headers={'X-Forwarded-For': '203.0.113.7'})
        self.assertFalse(socket.is_connected())


if __name__ == '__main__':
    if '--serve' in sys.argv:
        sandbox, app_module = sandbox_app()
        seed(app_module)
        print('Disposable preview: http://127.0.0.1:5057/dashboard', flush=True)
        app_module.socketio.run(app_module.app, host='127.0.0.1', port=5057, use_reloader=False)
    else:
        unittest.main(verbosity=2)
