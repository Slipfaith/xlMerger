import sys
import pytest

pytest.importorskip("PySide6.QtWidgets")
from PySide6.QtWidgets import QApplication

from gui.merge_tab import MergeTab


@pytest.fixture(scope="session")
def qapp():
    app = QApplication.instance()
    if app is None:
        app = QApplication(sys.argv)
    return app


def test_handle_files_selected(qapp, tmp_path):
    tab = MergeTab()
    files = []
    for i in range(2):
        p = tmp_path / f"f{i}.xlsx"
        p.write_text("")
        files.append(str(p))
    tab.handle_files_selected(files)
    assert tab.source_files == files
