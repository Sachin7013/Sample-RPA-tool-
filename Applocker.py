"""
collect_installed_apps.py
Simple evidence collector for Windows 11:
 - enumerates installed apps (registry + Store apps)
 - tries to infer a 'LastModified' timestamp (from installed files)
 - writes CSV and HTML report to /evidence
 - writes an evidence log (evidence_log.csv)
 - optional: create a full-screen screenshot (use --screenshot)
"""

import os
import csv
import json
import re
import subprocess
from pathlib import Path
from datetime import datetime
import winreg  # built-in on Windows
import glob

# -------------------------
EVIDENCE_DIR = Path("evidence")
EVIDENCE_DIR.mkdir(exist_ok=True)
CSV_OUT = EVIDENCE_DIR / "installed_apps.csv"
HTML_OUT = EVIDENCE_DIR / "installed_apps.html"
EVIDENCE_LOG = EVIDENCE_DIR / "evidence_log.csv"

# -------------------------
def timestamp_now():
    return datetime.now().strftime("%Y-%m-%d %H:%M:%S")

def parse_install_date(val):
    # registry InstallDate often YYYYMMDD
    if not val:
        return ""
    val = str(val).strip()
    m = re.match(r"^(\d{4})(\d{2})(\d{2})", val)
    if m:
        return f"{m.group(1)}-{m.group(2)}-{m.group(3)}"
    # fallback: return raw
    return val

def guess_last_modified(install_location, uninstall_string):
    """
    Try to infer last-modified time from install folder or an exe path inside uninstall string.
    Returns ISO timestamp string or empty string.
    """
    paths_to_check = []
    if install_location and os.path.isdir(install_location):
        # find executables in that folder (recursively)
        for ext in ("*.exe", "*.dll", "*.sys"):
            paths_to_check += glob.glob(os.path.join(install_location, "**", ext), recursive=True)
    if uninstall_string:
        # try to extract a quoted path or .exe from uninstall string
        q = re.search(r'"([^"]+\.exe)"', uninstall_string)
        if q:
            paths_to_check.append(q.group(1))
        else:
            q2 = re.search(r"([A-Za-z]:\\[^ ]+\.exe)", uninstall_string)
            if q2:
                paths_to_check.append(q2.group(1))
    # check file dates
    latest = 0
    for p in paths_to_check:
        try:
            m = os.path.getmtime(p)
            if m > latest:
                latest = m
        except Exception:
            continue
    if latest:
        return datetime.fromtimestamp(latest).isoformat(sep=' ', timespec='seconds')
    return ""

def read_uninstall_key(hive, keypath):
    apps = []
    try:
        with winreg.OpenKey(hive, keypath) as regkey:
            for i in range(0, winreg.QueryInfoKey(regkey)[0]):
                try:
                    subname = winreg.EnumKey(regkey, i)
                    with winreg.OpenKey(regkey, subname) as subk:
                        def rv(name):
                            try:
                                v, _ = winreg.QueryValueEx(subk, name)
                                return v
                            except FileNotFoundError:
                                return None
                        name = rv("DisplayName")
                        if not name:
                            continue
                        app = {
                            "Name": name,
                            "Version": rv("DisplayVersion") or "",
                            "InstallDate": parse_install_date(rv("InstallDate")),
                            "Publisher": rv("Publisher") or "",
                            "InstallLocation": rv("InstallLocation") or "",
                            "UninstallString": rv("UninstallString") or "",
                            "Source": "Registry",
                        }
                        app["LastModified"] = guess_last_modified(app["InstallLocation"], app["UninstallString"])
                        apps.append(app)
                except OSError:
                    continue
    except FileNotFoundError:
        pass
    return apps

def get_store_apps_via_powershell():
    """Get user Store (Appx) packages using PowerShell -> JSON (best-effort)."""
    try:
        cmd = [
            "powershell", "-NoProfile", "-Command",
            "Get-AppxPackage | Select Name, PackageFullName, InstallLocation | ConvertTo-Json -Depth 2"
        ]
        res = subprocess.run(cmd, capture_output=True, text=True, check=True)
        if not res.stdout:
            return []
        js = json.loads(res.stdout)
        if isinstance(js, dict):
            js = [js]
        apps = []
        for item in js:
            name = item.get("Name") or item.get("PackageFullName")
            installloc = item.get("InstallLocation") or ""
            app = {
                "Name": name,
                "Version": "",
                "InstallDate": "",
                "Publisher": "",
                "InstallLocation": installloc,
                "UninstallString": "",
                "Source": "StoreApp",
                "LastModified": guess_last_modified(installloc, "")
            }
            apps.append(app)
        return apps
    except Exception:
        return []

def dedupe_apps(list_of_apps):
    """Simple dedupe by name, keep the entry with most fields filled."""
    repo = {}
    for a in list_of_apps:
        key = a["Name"].strip()
        if not key:
            continue
        score = sum(1 for v in (a.get("Version"), a.get("InstallDate"), a.get("InstallLocation"), a.get("LastModified")) if v)
        if key not in repo or score > repo[key][0]:
            repo[key] = (score, a)
    return [v[1] for v in repo.values()]

def write_csv(items, path):
    headers = ["Name","Version","InstallDate","Publisher","InstallLocation","UninstallString","Source","LastModified"]
    with open(path, "w", newline="", encoding="utf-8") as f:
        w = csv.DictWriter(f, fieldnames=headers)
        w.writeheader()
        for it in items:
            w.writerow({h: it.get(h,"") for h in headers})

def write_html(items, path):
    rows = []
    for it in items:
        rows.append("<tr>" + "".join(f"<td>{(it.get(h) or '')}</td>" for h in ["Name","Version","InstallDate","Publisher","InstallLocation","Source","LastModified"]) + "</tr>")
    html = f"""<!doctype html>
<html><head><meta charset="utf-8"><title>Installed apps evidence</title>
<style>
body{{font-family:Segoe UI,Arial;margin:20px}}table{{border-collapse:collapse;width:100%}}
th,td{{border:1px solid #ddd;padding:6px;text-align:left;font-size:13px}}
th{{background:#f2f2f2}}
</style>
</head><body>
<h2>Installed Applications (collected: {timestamp_now()})</h2>
<table><thead><tr><th>Name</th><th>Version</th><th>InstallDate</th><th>Publisher</th><th>InstallLocation</th><th>Source</th><th>LastModified (inferred)</th></tr></thead><tbody>
{''.join(rows)}
</tbody></table>
</body></html>"""
    with open(path, "w", encoding="utf-8") as f:
        f.write(html)

def append_evidence_log(action, filepath, note=""):
    new = False
    if not EVIDENCE_LOG.exists():
        new = True
    with open(EVIDENCE_LOG, "a", newline="", encoding="utf-8") as f:
        w = csv.writer(f)
        if new:
            w.writerow(["Timestamp", "Action", "Filepath", "Note"])
        w.writerow([timestamp_now(), action, str(filepath), note])

def take_screenshot(save_path):
    # optional; requires pyautogui + pillow
    try:
        import pyautogui
    except Exception as e:
        return False, f"pyautogui not available: {e}"
    try:
        img = pyautogui.screenshot()
        img.save(save_path)
        return True, ""
    except Exception as e:
        return False, str(e)

# -------------------------
def main():
    print("Collecting from registry...")
    apps = []
    apps += read_uninstall_key(winreg.HKEY_LOCAL_MACHINE, r"SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall")
    apps += read_uninstall_key(winreg.HKEY_LOCAL_MACHINE, r"SOFTWARE\WOW6432Node\Microsoft\Windows\CurrentVersion\Uninstall")
    apps += read_uninstall_key(winreg.HKEY_CURRENT_USER, r"SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall")
    print(f"Found (registry): {len(apps)} entries")

    print("Collecting Store apps via PowerShell (user scope, best-effort)...")
    store = get_store_apps_via_powershell()
    print(f"Found (store): {len(store)} entries")
    apps += store

    print("De-duplicating...")
    apps = dedupe_apps(apps)
    print(f"Unique apps: {len(apps)}")

    print(f"Writing CSV -> {CSV_OUT}")
    write_csv(apps, CSV_OUT)
    append_evidence_log("write_csv", CSV_OUT, f"{len(apps)} rows")

    print(f"Writing HTML -> {HTML_OUT}")
    write_html(apps, HTML_OUT)
    append_evidence_log("write_html", HTML_OUT)

    # Save one screenshot if user asked via environment var or simple prompt:
    import argparse
    parser = argparse.ArgumentParser()
    parser.add_argument("--screenshot", action="store_true", help="Also capture a full-screen screenshot (requires pyautogui)")
    args = parser.parse_args()
    if args.screenshot:
        outshot = EVIDENCE_DIR / f"screenshot_{datetime.now().strftime('%Y%m%d_%H%M%S')}.png"
        ok, err = take_screenshot(outshot)
        if ok:
            print(f"Screenshot saved: {outshot}")
            append_evidence_log("screenshot", outshot)
        else:
            print("Screenshot failed:", err)
            append_evidence_log("screenshot_failed", outshot, err)

    print("Done. Evidence saved in 'evidence' folder.")
    print("You can open the HTML report with your browser to visually inspect the results.")

if __name__ == "__main__":
    main()
