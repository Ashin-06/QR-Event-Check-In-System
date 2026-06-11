import os
import sys
import urllib.request

# Ensure console supports Unicode printing
if hasattr(sys.stdout, "reconfigure"):
    sys.stdout.reconfigure(encoding="utf-8")

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
static_js = os.path.join(BASE_DIR, "static", "js")
static_css = os.path.join(BASE_DIR, "static", "css")

os.makedirs(static_js, exist_ok=True)
os.makedirs(static_css, exist_ok=True)

libraries = {
    "https://code.jquery.com/jquery-3.7.1.min.js": os.path.join(static_js, "jquery.min.js"),
    "https://cdn.socket.io/4.5.4/socket.io.min.js": os.path.join(static_js, "socket.io.min.js"),
    "https://cdn.datatables.net/1.13.6/js/jquery.dataTables.min.js": os.path.join(static_js, "jquery.dataTables.min.js"),
    "https://cdn.datatables.net/1.13.6/css/jquery.dataTables.min.css": os.path.join(static_css, "jquery.dataTables.min.css"),
}

headers = {
    "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/120.0.0.0 Safari/537.36"
}

for url, path in libraries.items():
    try:
        print(f"Downloading {url}...")
        req = urllib.request.Request(url, headers=headers)
        with urllib.request.urlopen(req) as response:
            with open(path, "wb") as f:
                f.write(response.read())
        print(f" Saved to {path}")
    except Exception as e:
        print(f"[-] Failed to download {url}: {e}")

print("Done!")
