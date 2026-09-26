"""Standard-library startup checks for the Windows BAT launchers."""
import argparse
import importlib.util
import os
from pathlib import Path
import shutil
import socket
import subprocess
import sys
import threading
import time
import urllib.request
import webbrowser

ROOT = Path(__file__).resolve().parent
REQUIRED = {'flask': 'flask', 'flask_socketio': 'flask-socketio', 'eventlet': 'eventlet',
            'pandas': 'pandas', 'numpy': 'numpy', 'openpyxl': 'openpyxl',
            'PIL': 'Pillow', 'qrcode': 'qrcode[pil]', 'barcode': 'python-barcode',
            'cryptography': 'cryptography', 'requests': 'requests'}


def preflight():
    if sys.version_info < (3, 10):
        raise ValueError('Python 3.10 or newer is required.')
    try:
        port = int(os.environ.get('PORT', '5001'))
    except ValueError:
        raise ValueError('PORT must be a number from 1 to 65535.') from None
    if not 1 <= port <= 65535:
        raise ValueError('PORT must be a number from 1 to 65535.')
    for filename in ('app.py', 'card_design.py', 'roster_schema.py', 'templates/dashboard.html', 'static/js/card-preview.js'):
        if not (ROOT / filename).is_file():
            raise ValueError(f'Missing application file: {filename}. Restore the complete app folder.')
    missing = [package for module, package in REQUIRED.items() if importlib.util.find_spec(module) is None]
    if missing:
        raise ValueError('Missing packages: ' + ', '.join(missing) + '\nInstall them with:\n'
                         f'"{sys.executable}" -m pip install -r "{ROOT / "requirements.txt"}"')
    return port


def port_in_use(port):
    with socket.socket() as probe:
        probe.settimeout(0.5)
        return probe.connect_ex(('127.0.0.1', port)) == 0


def open_when_ready(process, port, timeout=60):
    url = f'http://127.0.0.1:{port}/dashboard'
    opener = urllib.request.build_opener(urllib.request.ProxyHandler({}))
    deadline = time.monotonic() + timeout
    while process.poll() is None and time.monotonic() < deadline:
        try:
            with opener.open(url, timeout=1) as response:
                ready = response.status == 200 and b'Manager Control Center' in response.read(200000)
            if ready and process.poll() is None:
                webbrowser.open(url)
                print(f'\nDashboard ready: {url}', flush=True)
                return
        except (OSError, ValueError):
            pass
        time.sleep(0.5)
    if process.poll() is None:
        print(f'\nStartup is slow. Check the server messages, then open {url}', flush=True)


def main(argv=None):
    parser = argparse.ArgumentParser(description='Start the QR Check-In app safely.')
    parser.add_argument('--check', action='store_true', help='Check installation without starting or changing event data.')
    parser.add_argument('--no-browser', action='store_true', help='Do not open the dashboard automatically.')
    parser.add_argument('--tunnel', action='store_true', help='Start the optional public tunnel after confirmation.')
    args = parser.parse_args(argv)
    try:
        port = preflight()
        if args.tunnel and not shutil.which('ssh'):
            raise ValueError('OpenSSH Client is missing. Enable it in Windows Optional Features first.')
        print(f'Python: {sys.executable}\nDashboard: http://127.0.0.1:{port}/dashboard', flush=True)
        if args.check:
            print('Installation checks passed. No server or public tunnel was started.')
            return 0
        if args.tunnel:
            if not port_in_use(port):
                raise ValueError('Start the app with START_SYSTEM.bat or run_system.bat before opening a tunnel.')
            print('WARNING: This exposes the local app to the public internet.\n'
                  'Enable manager and scanner authentication first. Keep this window open while sharing.')
            if input('Start public access now? [y/N] ').strip().lower() != 'y':
                return 0
            # Keep SSH host-key verification enabled, including on first connection.
            return subprocess.call(['ssh', '-T', '-o', 'ServerAliveInterval=30', '-o',
                                    'ServerAliveCountMax=3', '-o', 'ExitOnForwardFailure=yes',
                                    '-R', f'80:127.0.0.1:{port}', 'nokey@localhost.run'])
        if port_in_use(port):
            raise ValueError(f'Port {port} is already in use. If this is your app, open the dashboard URL above.\n'
                             'Otherwise stop the old server yourself or set another PORT. No process was stopped.')
        print('Starting the server. Keep this window open; press Ctrl+C to stop.', flush=True)
        process = subprocess.Popen([sys.executable, str(ROOT / 'app.py')], cwd=ROOT)
        if not args.no_browser:
            threading.Thread(target=open_when_ready, args=(process, port), daemon=True).start()
        try:
            return process.wait()
        except KeyboardInterrupt:
            if process.poll() is None:
                process.terminate()
                try:
                    process.wait(timeout=10)
                except subprocess.TimeoutExpired:
                    process.kill()
                    process.wait()
            return 0
    except (ValueError, OSError) as error:
        print(f'\n[ERROR] {error}', file=sys.stderr)
        return 1
    except (KeyboardInterrupt, EOFError):
        return 0


if __name__ == '__main__':
    sys.exit(main())
