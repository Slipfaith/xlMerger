import os
import sys
import subprocess
import tempfile
import requests
from PySide6.QtWidgets import QMessageBox, QProgressDialog, QApplication
from PySide6.QtCore import Qt
from utils.i18n import tr
from __init__ import __version__

pending_update = None

GITHUB_API_LATEST = "https://api.github.com/repos/Slipfaith/xlMerger/releases/latest"

PUBLIC_GPG_KEY = """-----BEGIN PGP PUBLIC KEY BLOCK-----

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
0c8=
=ohJm
-----END PGP PUBLIC KEY BLOCK-----"""

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


def _verify_signature(exe_path: str, sig_path: str):
    key_file = tempfile.NamedTemporaryFile(delete=False, suffix=".asc")
    try:
        key_file.write(PUBLIC_GPG_KEY.encode("utf-8"))
        key_file.close()
        subprocess.run(["gpg", "--import", key_file.name], check=True, stdout=subprocess.PIPE, stderr=subprocess.PIPE)
        subprocess.run(["gpg", "--verify", sig_path, exe_path], check=True, stdout=subprocess.PIPE, stderr=subprocess.PIPE)
    finally:
        try:
            os.remove(key_file.name)
        except OSError:
            pass

def _run_update_script(old_exe: str, new_exe: str):
    if os.name != "nt":
        try:
            if os.path.exists(old_exe):
                os.remove(old_exe)
        except PermissionError:
            pass
        os.replace(new_exe, old_exe)
        return

    script_path = os.path.join(os.path.dirname(new_exe), "update.bat")
    script = (
        f"@echo off\n"
        f"set new=\"{new_exe}\"\n"
        f"set old=\"{old_exe}\"\n"
        f":loop\n"
        f"del %old% >nul 2>&1\n"
        f"if exist %old% (\n"
        f"  ping 127.0.0.1 -n 2 >nul\n"
        f"  goto loop\n"
        f")\n"
        f"move /Y %new% %old%\n"
        f"del %~f0\n"
    )
    with open(script_path, "w", encoding="utf-8") as f:
        f.write(script)
    subprocess.Popen(["cmd", "/c", "start", "", script_path], shell=True)

def _schedule_update_on_exit(new_exe: str):
    global pending_update
    pending_update = new_exe
    app = QApplication.instance()
    if app is not None:
        def _on_quit():
            global pending_update
            if pending_update:
                _run_update_script(sys.argv[0], pending_update)
                pending_update = None
        app.aboutToQuit.connect(_on_quit)

def check_for_update(parent, auto=False):
    progress = QProgressDialog(
        tr("Checking for updates..."), "", 0, 0, parent
    )
    progress.setWindowTitle(tr("Update"))
    progress.setCancelButton(None)
    progress.setWindowModality(Qt.ApplicationModal)
    progress.show()
    QApplication.processEvents()
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

            sig_asset = next(
                (
                    a
                    for a in data.get("assets", [])
                    if a["name"] == asset["name"] + ".asc"
                ),
                None,
            )
            if not sig_asset:
                QMessageBox.critical(parent, tr("Update Error"), tr("Signature file not found."))
                return

            new_exe_name = "xlMerger_new.exe"
            file_path = download_asset(asset["browser_download_url"], new_exe_name)
            sig_path = download_asset(sig_asset["browser_download_url"], new_exe_name + ".asc")
            try:
                _verify_signature(file_path, sig_path)
            except Exception:
                for p in (file_path, sig_path):
                    try:
                        os.remove(p)
                    except OSError:
                        pass
                QMessageBox.critical(parent, tr("Update Error"), tr("Failed to verify update signature."))
                return
            try:
                os.remove(sig_path)
            except OSError:
                pass
            answer = QMessageBox.question(
                parent,
                tr("Update Available"),
                f"{tr('A new version')} {latest_version} {tr('is available.')}\n"
                f"{tr('The update has been downloaded as')} {new_exe_name}.\n\n"
                f"{tr('Update now?')}",
                QMessageBox.Yes | QMessageBox.No,
            )
            if answer == QMessageBox.Yes:
                _run_update_script(sys.argv[0], file_path)
                QApplication.instance().quit()
            else:
                QMessageBox.information(
                    parent,
                    tr("Update Deferred"),
                    tr("The update will be installed after you close the program."),
                )
                _schedule_update_on_exit(file_path)
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
    finally:
        progress.close()
