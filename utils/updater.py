import hashlib
import os
import sys
import subprocess
import requests

PREVIOUS_VERSION = "2.2.1"
CURRENT_VERSION = "2.2.2"
__version__ = CURRENT_VERSION

try:  # pragma: no cover - optional GUI dependencies
    from PySide6.QtWidgets import QMessageBox, QProgressDialog, QApplication
    from PySide6.QtCore import Qt
except Exception:  # when PySide6 or system libraries are missing
    QMessageBox = QProgressDialog = QApplication = None  # type: ignore
    Qt = None  # type: ignore

try:
    from .i18n import tr  # type: ignore
except ImportError:  # pragma: no cover - fallback for script execution
    from utils.i18n import tr

pending_update = None

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


def _verify_checksum(exe_path: str, hash_path: str):
    """Verify a SHA256 checksum for the downloaded update."""
    with open(hash_path, "r", encoding="utf-8") as f:
        expected = f.read().strip().split()[0]
    with open(exe_path, "rb") as f:
        actual = hashlib.sha256(f.read()).hexdigest()
    if actual != expected:
        raise ValueError("Invalid update checksum")

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
    if any(obj is None for obj in (QProgressDialog, QMessageBox, QApplication, Qt)):
        raise RuntimeError("GUI libraries not available")
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

            hash_asset = next(
                (
                    a
                    for a in data.get("assets", [])
                    if a["name"] == asset["name"] + ".sha256"
                ),
                None,
            )
            if not hash_asset:
                QMessageBox.critical(parent, tr("Update Error"), tr("Checksum file not found."))
                return

            new_exe_name = "xlMerger_new.exe"
            file_path = download_asset(asset["browser_download_url"], new_exe_name)
            hash_path = download_asset(hash_asset["browser_download_url"], new_exe_name + ".sha256")
            try:
                _verify_checksum(file_path, hash_path)
            except Exception:
                for p in (file_path, hash_path):
                    try:
                        os.remove(p)
                    except OSError:
                        pass
                QMessageBox.critical(parent, tr("Update Error"), tr("Failed to verify update checksum."))
                return
            try:
                os.remove(hash_path)
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
