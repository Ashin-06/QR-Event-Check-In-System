import os
import sys

def search_files():
    keywords = ["register_dashboard", "register_device", "devices_updated", "update_device_fields"]
    for root, dirs, files in os.walk("."):
        if ".git" in root or "__pycache__" in root:
            continue
        for file in files:
            if file.endswith((".py", ".html", ".js")) and file != "search_sockets.py":
                path = os.path.join(root, file)
                try:
                    with open(path, "r", encoding="utf-8", errors="ignore") as f:
                        lines = f.readlines()
                    for idx, line in enumerate(lines, 1):
                        for kw in keywords:
                            if kw in line:
                                out_line = f"{path}:{idx}: {line.strip()}"
                                # Print as ascii compatible representation to avoid terminal encoding errors
                                print(out_line.encode("ascii", errors="replace").decode("ascii"))
                except Exception as e:
                    print(f"Error reading {path}: {e}")

if __name__ == "__main__":
    search_files()
