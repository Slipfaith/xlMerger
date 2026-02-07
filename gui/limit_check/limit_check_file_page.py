# -*- coding: utf-8 -*-
from PySide6.QtWidgets import (
    QWidget, QVBoxLayout, QHBoxLayout, QGroupBox, QComboBox,
    QPushButton
)
from PySide6.QtCore import Signal
from core.drag_drop import DragDropLineEdit
from utils.i18n import tr
from gui.style_system import set_button_variant

class FileSelectionPage(QWidget):
    file_selected = Signal(str)
    sheet_selected = Signal(str)
    mapping_clicked = Signal()
    next_clicked = Signal()

    def __init__(self, parent=None):
        super().__init__(parent)
        self.init_ui()

    def init_ui(self):
        layout = QVBoxLayout(self)

        file_group = QGroupBox(tr("Выбор файла (перетащи Excel или дважды кликни)"))
        file_layout = QHBoxLayout()
        self.file_line = DragDropLineEdit(mode='file')
        self.file_line.fileSelected.connect(self.on_file_dropped)
        file_layout.addWidget(self.file_line)
        file_group.setLayout(file_layout)
        layout.addWidget(file_group)

        sheet_group = QGroupBox(tr("Выбор листа"))
        sheet_layout = QHBoxLayout()
        self.sheet_combo = QComboBox()
        self.sheet_combo.currentTextChanged.connect(self.on_sheet_changed)
        sheet_layout.addWidget(self.sheet_combo)
        sheet_group.setLayout(sheet_layout)
        layout.addWidget(sheet_group)

        btn_layout = QHBoxLayout()
        self.map_button = QPushButton(tr("Лимиты"))
        set_button_variant(self.map_button, "secondary")
        self.map_button.clicked.connect(self.mapping_clicked.emit)
        self.next_button = QPushButton(tr("Далее"))
        set_button_variant(self.next_button, "orange")
        self.next_button.clicked.connect(self.next_clicked.emit)
        btn_layout.addWidget(self.map_button)
        btn_layout.addWidget(self.next_button)
        layout.addLayout(btn_layout)

    def on_file_dropped(self, file_path):
        # Тут можно просто передавать полный путь дальше
        self.file_selected.emit(file_path)

    def set_sheets(self, sheets: list):
        self.sheet_combo.clear()
        self.sheet_combo.addItems(sheets)

    def set_selected_sheet(self, name: str):
        idx = self.sheet_combo.findText(name)
        if idx >= 0:
            self.sheet_combo.setCurrentIndex(idx)

    def current_sheet(self):
        return self.sheet_combo.currentText()

    def on_sheet_changed(self, name):
        self.sheet_selected.emit(name)

    def get_selected_file(self):
        return self.file_line.text()
