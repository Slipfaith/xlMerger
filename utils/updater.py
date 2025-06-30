import appdirs
from pyupdater.client import Client
from PySide6.QtWidgets import QMessageBox

from utils.i18n import tr
from __init__ import __version__


class ClientConfig(object):
    APP_NAME = "xlMerger"
    COMPANY_NAME = "xlMerger"
    UPDATE_URLS = [
        "https://github.com/yourusername/xlMerger/releases/download/"
    ]
    PUBLIC_KEY = ""  # Add your public key here when using signing
    DATA_DIR = appdirs.user_data_dir(APP_NAME, COMPANY_NAME)


def check_for_update(parent, auto=False):
    try:
        client = Client(ClientConfig(), refresh=True)
        app_update = client.update_check(ClientConfig.APP_NAME, __version__)
        if app_update:
            if QMessageBox.question(
                parent,
                tr("Update Available"),
                tr("A newer version is available. Download?"),
                QMessageBox.Yes | QMessageBox.No,
            ) == QMessageBox.Yes:
                app_update.download()
                if app_update.is_downloaded():
                    QMessageBox.information(
                        parent,
                        tr("Update Downloaded"),
                        tr("Update downloaded. Restart application to install."),
                    )
                    app_update.extract_restart()
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
