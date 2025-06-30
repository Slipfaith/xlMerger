import os
import requests
import tempfile
import webbrowser
from PySide6.QtWidgets import QMessageBox
from utils.i18n import tr
from __init__ import __version__

GITHUB_API_LATEST = "https://api.github.com/repos/{owner}/{repo}/releases/latest"
REPO_OWNER = "yourusername"
REPO_NAME = "xlMerger"


def get_latest_release():
    url = GITHUB_API_LATEST.format(owner=REPO_OWNER, repo=REPO_NAME)
    response = requests.get(url, timeout=10)
    response.raise_for_status()
    return response.json()


def compare_versions(current: str, latest: str) -> bool:
    def normalize(v):
        return [int(x) for x in v.strip('v').split('.') if x.isdigit()]
    return normalize(latest) > normalize(current)


def download_asset(url: str, name: str) -> str:
    path = os.path.join(tempfile.gettempdir(), name)
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
            if QMessageBox.question(
                parent,
                tr("Update Available"),
                tr("A newer version is available. Download?"),
                QMessageBox.Yes | QMessageBox.No,
            ) == QMessageBox.Yes:
                asset = next(
                    (a for a in data.get("assets", []) if a["name"].endswith(".exe")),
                    None,
                )
                if asset:
                    file_path = download_asset(asset["browser_download_url"], asset["name"])
                    QMessageBox.information(
                        parent,
                        tr("Downloaded"),
                        tr("Installer downloaded to {path}. Run it to update.").format(path=file_path),
                    )
                else:
                    webbrowser.open(data.get("html_url"))
        else:
            if not auto:
                QMessageBox.information(
                    parent,
                    tr("Check for Updates"),
                    tr("You have the latest version."),
                )
    except Exception as e:
        if not auto:
            QMessageBox.critical(parent, tr("Error"), str(e))
