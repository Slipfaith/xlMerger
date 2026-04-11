# -*- coding: utf-8 -*-
import sys

import pytest

pytest.importorskip("PySide6.QtWidgets")
from PySide6.QtWidgets import QApplication

from gui.main_page import MainPageWidget
from core.main_page_logic import MainPageLogic
from utils.i18n import i18n, tr


@pytest.fixture(scope="session")
def qapp():
    app = QApplication.instance()
    if app is None:
        app = QApplication(sys.argv)
    return app


def test_copy_column_field_has_helpful_placeholder(qapp):
    page = MainPageWidget()
    text = page.copy_column_entry.placeholderText()
    assert text
    assert "A" in text
    page.close()


def test_start_disabled_and_warning_visible_when_copy_column_missing(qapp, monkeypatch):
    source = r"C:\fake\source.xlsx"
    target = r"C:\fake\target.xlsx"
    monkeypatch.setattr(
        "core.main_page_logic.os.path.isfile",
        lambda path: path in {source, target},
    )

    page = MainPageWidget()
    logic = MainPageLogic(page)
    logic.selected_files = [source]
    page.excel_file_entry.setText(target)
    page.copy_column_entry.setText("")

    logic.update_process_button_state()

    assert not page.process_button.isEnabled()
    assert page.process_button.property("variant") == "secondary"
    assert not page.copy_column_status_label.isHidden()
    assert page.copy_column_status_label.text() == tr("не выбрана колонка с переводом.")
    page.close()


def test_start_enabled_and_warning_hidden_when_copy_column_set(qapp, monkeypatch):
    source = r"C:\fake\source.xlsx"
    target = r"C:\fake\target.xlsx"
    monkeypatch.setattr(
        "core.main_page_logic.os.path.isfile",
        lambda path: path in {source, target},
    )

    page = MainPageWidget()
    logic = MainPageLogic(page)
    logic.selected_files = [source]
    page.excel_file_entry.setText(target)
    page.copy_column_entry.setText("A")

    logic.update_process_button_state()

    assert page.process_button.isEnabled()
    assert page.process_button.property("variant") == "orange"
    assert page.copy_column_status_label.isHidden()
    page.close()


def test_english_warning_label_is_not_clipped(qapp):
    previous_language = i18n.language
    page = None
    try:
        i18n.set_language("en")
        page = MainPageWidget()
        page.resize(640, 480)
        page.show()
        page.copy_column_status_label.setHidden(False)
        qapp.processEvents()

        assert page.copy_column_status_label.text() == "translation column is not selected."
        assert page.copy_column_status_label.width() >= 180
    finally:
        if page is not None:
            page.close()
        i18n.set_language(previous_language)
