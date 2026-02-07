# -*- coding: utf-8 -*-
from PySide6.QtGui import QFont


BASE_FONT_FAMILY = "Segoe UI Variable"
BASE_FONT_SIZE = 10


APP_QSS = """
QWidget {
    background-color: #F7F9FC;
    color: #0F172A;
    font-family: "Segoe UI Variable", "Segoe UI", sans-serif;
}

QMainWindow, QDialog {
    background-color: #F7F9FC;
}

QGroupBox,
QFrame#mappingCard,
QFrame[card="true"] {
    background-color: #FFFFFF;
    border: 1px solid #E8EDF4;
    border-radius: 12px;
    margin-top: 10px;
    padding: 10px 10px 8px 10px;
}

QGroupBox::title {
    subcontrol-origin: margin;
    left: 12px;
    padding: 0 4px;
    color: #64748B;
    font-size: 11px;
    letter-spacing: 0.5px;
    font-weight: 600;
}

QLabel[role="heading"] {
    color: #0F172A;
    font-size: 16px;
    font-weight: 700;
}

QLabel[role="eyebrow"] {
    color: #64748B;
    font-size: 11px;
    letter-spacing: 0.7px;
    font-weight: 600;
}

QLabel[role="muted"] {
    color: #64748B;
    font-size: 12px;
}

QLabel[role="link"] {
    color: #5B5BD6;
    text-decoration: underline;
}

QLabel[state="success"] {
    color: #15803D;
    font-weight: 600;
}

QLabel[state="error"] {
    color: #B91C1C;
    font-weight: 600;
}

QPushButton {
    min-height: 32px;
    padding: 0 14px;
    border-radius: 16px;
    border: 1px solid #D8E0EC;
    background-color: #EEF2F8;
    color: #0F172A;
    font-size: 12px;
    font-weight: 600;
}

QPushButton:hover {
    background-color: #E6EBF4;
    border-color: #CCD7E7;
}

QPushButton:pressed {
    background-color: #DEE5F1;
}

QPushButton:disabled {
    background-color: #EDF1F6;
    color: #97A4B6;
    border-color: #E1E7F0;
}

QPushButton[variant="primary"] {
    background-color: #5B5BD6;
    color: #FFFFFF;
    border-color: #5B5BD6;
}

QPushButton[variant="primary"]:hover {
    background-color: #4F4FC8;
    border-color: #4F4FC8;
}

QPushButton[variant="primary"]:pressed {
    background-color: #4545B7;
    border-color: #4545B7;
}

QPushButton[variant="secondary"] {
    background-color: #EFF3F8;
    color: #0F172A;
    border-color: #D7DFEA;
}

QPushButton[variant="secondary"]:hover {
    background-color: #E8EDF5;
}

QPushButton[variant="orange"] {
    background-color: #F59E0B;
    color: #FFFFFF;
    border-color: #F59E0B;
}

QPushButton[variant="orange"]:hover {
    background-color: #EA8A00;
    border-color: #EA8A00;
}

QPushButton[variant="orange"]:pressed {
    background-color: #D97706;
    border-color: #D97706;
}

QPushButton[variant="ghost"] {
    background-color: transparent;
    border-color: transparent;
    color: #64748B;
}

QPushButton[variant="ghost"]:hover {
    background-color: #EEF2F7;
    color: #334155;
}

QPushButton[variant="danger"] {
    background-color: #DC2626;
    border-color: #DC2626;
    color: #FFFFFF;
}

QPushButton[variant="danger"]:hover {
    background-color: #B91C1C;
    border-color: #B91C1C;
}

QPushButton[shape="circle"] {
    min-width: 24px;
    max-width: 24px;
    min-height: 24px;
    max-height: 24px;
    padding: 0;
    border-radius: 12px;
}

QPushButton[shape="circleLarge"] {
    min-width: 28px;
    max-width: 28px;
    min-height: 28px;
    max-height: 28px;
    padding: 0;
    border-radius: 14px;
}

QLineEdit,
QComboBox,
QTextEdit,
QPlainTextEdit,
QListWidget,
QTableWidget,
QTableView {
    background-color: #F1F5F9;
    border: 1px solid transparent;
    border-radius: 10px;
    padding: 6px 8px;
    selection-background-color: #DCE3FF;
    selection-color: #0F172A;
}

QLineEdit:focus,
QComboBox:focus,
QTextEdit:focus,
QPlainTextEdit:focus,
QListWidget:focus,
QTableWidget:focus,
QTableView:focus {
    background-color: #FFFFFF;
    border: 1px solid #A5B4FC;
}

QComboBox::drop-down {
    border: none;
    width: 18px;
}

QCheckBox {
    spacing: 6px;
    color: #334155;
    font-size: 12px;
}

QHeaderView::section {
    background-color: transparent;
    color: #64748B;
    border: none;
    border-bottom: 1px solid #E6EDF5;
    padding: 6px 8px;
    font-size: 11px;
    letter-spacing: 0.4px;
    font-weight: 600;
}

QProgressBar {
    border: 1px solid #D8E0EC;
    border-radius: 8px;
    text-align: center;
    background-color: #EEF2F8;
    min-height: 8px;
}

QProgressBar::chunk {
    background-color: #5B5BD6;
    border-radius: 7px;
}

QTabWidget::pane {
    border: 1px solid #E6EDF5;
    border-radius: 12px;
    background-color: #FFFFFF;
    margin-top: 6px;
}

QTabBar::tab {
    background-color: #EEF2F8;
    color: #64748B;
    border-radius: 12px;
    border: 1px solid transparent;
    padding: 8px 14px;
    margin-right: 6px;
    min-height: 16px;
}

QTabBar::tab:selected {
    background-color: #FFFFFF;
    color: #0F172A;
    border: 1px solid #DCE4F0;
    font-weight: 600;
}

QMenuBar {
    background-color: #FFFFFF;
    border: none;
    border-bottom: 1px solid #E8EDF4;
    padding: 2px 6px;
}

QMenuBar::item {
    border-radius: 8px;
    padding: 4px 8px;
}

QMenuBar::item:selected {
    background-color: #EEF2F8;
}

QMenu {
    background-color: #FFFFFF;
    border: 1px solid #E1E8F3;
    border-radius: 10px;
    padding: 4px;
}

QMenu::item {
    border-radius: 6px;
    padding: 6px 10px;
}

QMenu::item:selected {
    background-color: #EEF2F8;
}

QScrollArea {
    border: none;
    background: transparent;
}

QTableWidget#sheetMappingTable {
    border: none;
    background-color: transparent;
    gridline-color: transparent;
}

QTableWidget#sheetMappingTable::item {
    border: none;
    padding: 0 8px;
    background-color: transparent;
}

QComboBox[flatSelect="true"] {
    border: 1px solid transparent;
    border-radius: 8px;
    background-color: transparent;
    padding: 4px 6px;
}

QComboBox[flatSelect="true"]:focus {
    background-color: #FFFFFF;
    border: 1px solid #A5B4FC;
}
"""


def _refresh_widget(widget):
    if widget is None:
        return
    style = widget.style()
    style.unpolish(widget)
    style.polish(widget)
    widget.update()


def set_button_variant(button, variant: str):
    button.setProperty("variant", variant)
    _refresh_widget(button)


def set_button_shape(button, shape: str):
    button.setProperty("shape", shape)
    _refresh_widget(button)


def set_label_role(label, role: str):
    label.setProperty("role", role)
    _refresh_widget(label)


def set_label_state(label, state: str):
    label.setProperty("state", state)
    _refresh_widget(label)


def set_card(widget, enabled: bool = True):
    widget.setProperty("card", bool(enabled))
    _refresh_widget(widget)


def apply_app_style(app):
    app.setFont(QFont(BASE_FONT_FAMILY, BASE_FONT_SIZE))
    app.setStyleSheet(APP_QSS)
