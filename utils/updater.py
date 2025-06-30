import os
import sys
import requests
from PySide6.QtWidgets import QMessageBox
from utils.i18n import tr
from __init__ import __version__

GITHUB_API_LATEST = "https://api.github.com/repos/Slipfaith/xlMerger/releases/latest"

def get_latest_release():
    response = requests.get(GITHUB_API_LATEST, timeout=10)
    response.raise_for_status()
    return response.json()

def compare_versions(current: str, latest: str) -> bool:
    def normalize(v):
        return [int(x) for x in v.strip('v').split('.') if x.isdigit()]
    return normalize(latest) > normalize(current)

def download_asset(url: str, name: str) -> str:
    exe_dir = os.path.dirname(sys.argv[0])
    path = os.path.join(exe_dir, name)
    with requests.get(url, stream=True, timeout=10) as r:
        r.raise_for_status()
        with open(path, 'wb') as f:
            for chunk in r.iter_content(chunk_size=8192):
                if chunk:
                    f.write(chunk)
    return path

def check_for_update(parent, auto=False):
    try:
        data = get_latest_release()
        latest_version = data.get("tag_name", "").lstrip('v')
        if latest_version and compare_versions(__version__, latest_version):
            asset = next(
                (a for a in data.get("assets", []) if a["name"].endswith(".exe")),
                None,
            )
            if not asset:
                QMessageBox.warning(parent, "Update", "No update file found in the release.")
                return

            new_exe_name = "xlMerger_new.exe"
            file_path = download_asset(asset["browser_download_url"], new_exe_name)

            answer = QMessageBox.question(
                parent,
                "Update Available",
                f"A new version {latest_version} is available.\n"
                f"The update has been downloaded as {new_exe_name} in the application folder.\n\n"
                "Update now?\n\n"
                "If you click Yes, the app will close. Please launch the new version manually.",
                QMessageBox.Yes | QMessageBox.No,
            )
            if answer == QMessageBox.Yes:
                QMessageBox.information(
                    parent,
                    "Updating",
                    f"The application will now close.\n\n"
                    f"Please start {new_exe_name} manually.\n"
                    "You can delete the old file after updating."
                )
                sys.exit(0)
            else:
                QMessageBox.information(
                    parent,
                    "Update Deferred",
                    f"You can update later by launching {new_exe_name} from the app folder.",
                )
        else:
            if not auto:
                QMessageBox.information(
                    parent,
                    "Check for Updates",
                    "You have the latest version.",
                )
    except Exception as e:
        if not auto:
            QMessageBox.critical(parent, "Update Error", str(e))
