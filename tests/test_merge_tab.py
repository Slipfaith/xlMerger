# -*- coding: utf-8 -*-
import sys
import pytest

pytest.importorskip("PySide6.QtWidgets")
from PySide6.QtCore import QUrl
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


def test_open_file_location_uses_clicked_link(qapp, tmp_path, monkeypatch):
    tab = MergeTab()
    first = tmp_path / "first.xlsx"
    second = tmp_path / "second file [draft]&.xlsx"
    first.write_text("")
    second.write_text("")
    tab.output_files = [("src1", str(first)), ("src2", str(second))]

    calls = []

    def fake_run(cmd):
        calls.append(cmd)

    monkeypatch.setattr("gui.merge_tab.platform.system", lambda: "Windows")
    monkeypatch.setattr("gui.merge_tab.subprocess.run", fake_run)

    tab.open_file_location(QUrl.fromLocalFile(str(second)).toString())

    assert len(calls) == 1
    assert calls[0] == ["explorer", "/select,", str(second)]
    assert calls[0][2] != str(first)
