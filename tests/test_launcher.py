"""Launcher checks with no live event writes or public tunnel connections."""
import os
from pathlib import Path
import shutil
import subprocess
import sys
import tempfile
import unittest
from unittest.mock import MagicMock, patch

ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(ROOT))
import launcher


class LauncherTests(unittest.TestCase):
    def test_check_never_starts_server(self):
        with patch.object(launcher, 'preflight', return_value=5123), patch.object(launcher.subprocess, 'Popen') as start:
            self.assertEqual(launcher.main(['--check']), 0)
            start.assert_not_called()

    def test_invalid_port(self):
        for port in ('bad', '0', '65536'):
            with patch.dict(os.environ, {'PORT': port}), self.assertRaises(ValueError):
                launcher.preflight()

    def test_missing_dependency_explains_install(self):
        with patch.dict(os.environ, {'PORT': '5001'}), patch.object(launcher.importlib.util, 'find_spec', return_value=None):
            with self.assertRaisesRegex(ValueError, 'pip install'):
                launcher.preflight()

    def test_busy_port_does_not_spawn_or_kill(self):
        with patch.object(launcher, 'preflight', return_value=5001), patch.object(launcher, 'port_in_use', return_value=True), patch.object(launcher.subprocess, 'Popen') as start:
            self.assertEqual(launcher.main(['--no-browser']), 1)
            start.assert_not_called()

    def test_start_uses_selected_interpreter_and_directory(self):
        process = MagicMock()
        process.wait.return_value = 7
        with patch.object(launcher, 'preflight', return_value=5001), patch.object(launcher, 'port_in_use', return_value=False), patch.object(launcher.subprocess, 'Popen', return_value=process) as start:
            self.assertEqual(launcher.main(['--no-browser']), 7)
            start.assert_called_once_with([sys.executable, str(ROOT / 'app.py')], cwd=ROOT)

    def test_browser_opens_only_when_ready(self):
        process = MagicMock()
        process.poll.return_value = None
        opener = MagicMock()
        response = opener.open.return_value.__enter__.return_value
        response.status = 200
        response.read.return_value = b'<title>Manager Control Center</title>'
        with patch.object(launcher.urllib.request, 'build_opener', return_value=opener), patch.object(launcher.webbrowser, 'open') as browser:
            launcher.open_when_ready(process, 5123)
            browser.assert_called_once_with('http://127.0.0.1:5123/dashboard')
        process.poll.return_value = 1
        with patch.object(launcher.webbrowser, 'open') as browser:
            launcher.open_when_ready(process, 5123)
            browser.assert_not_called()

    def test_tunnel_requires_confirmation(self):
        with patch.object(launcher, 'preflight', return_value=5001), patch.object(launcher, 'port_in_use', return_value=True), patch.object(launcher.shutil, 'which', return_value='ssh'), patch('builtins.input', return_value='n'), patch.object(launcher.subprocess, 'call') as ssh:
            self.assertEqual(launcher.main(['--tunnel']), 0)
            ssh.assert_not_called()

    @unittest.skipUnless(os.name == 'nt', 'Windows BAT integration')
    def test_batch_paths_spaces_custom_port_and_exit_code(self):
        # Package a copy in a path with spaces and invoke it from another folder.
        (ROOT / 'runtime').mkdir(exist_ok=True)
        with tempfile.TemporaryDirectory(prefix='launcher check ', dir=ROOT / 'runtime') as directory:
            package = Path(directory)
            app_dir = package / 'qr_checkin_system'
            files = ['run_system.bat', 'start_public_tunnel.bat', 'launcher.py', 'app.py',
                     'card_design.py', 'roster_schema.py', 'templates/dashboard.html', 'static/js/card-preview.js']
            for filename in files:
                target = app_dir / filename
                target.parent.mkdir(parents=True, exist_ok=True)
                shutil.copy2(ROOT / filename, target)
            wrapper = ROOT.parent / 'START_SYSTEM.bat'
            if wrapper.exists():
                shutil.copy2(wrapper, package / wrapper.name)
            entrypoints = [app_dir / 'run_system.bat', app_dir / 'start_public_tunnel.bat']
            if wrapper.exists():
                entrypoints.append(package / wrapper.name)
            for entrypoint in entrypoints:
                result = subprocess.run(['cmd', '/d', '/c', 'call', str(entrypoint), '--check'],
                    cwd=ROOT.parent, env={**os.environ, 'PORT': '5123'}, capture_output=True, text=True, timeout=20)
                self.assertEqual(result.returncode, 0, result.stdout + result.stderr)
                self.assertIn('127.0.0.1:5123/dashboard', result.stdout)
            result = subprocess.run(['cmd', '/d', '/c', 'call', str(app_dir / 'run_system.bat'), '--check'],
                cwd=ROOT.parent, env={**os.environ, 'PORT': 'bad'}, capture_output=True, text=True, timeout=20)
            self.assertEqual(result.returncode, 1, result.stdout + result.stderr)


if __name__ == '__main__':
    unittest.main(verbosity=2)
