import os
from PySide6.QtWidgets import (
    QWidget, QVBoxLayout, QLabel, QScrollArea, QFrame, QGridLayout,
    QHBoxLayout, QComboBox, QPushButton
)
from PySide6.QtCore import Qt, Signal

class HeaderRowPage(QWidget):
    backClicked = Signal()
    nextClicked = Signal(dict)  # dict: {sheet_name: header_row_number}

    def __init__(self, selected_sheets, parent=None):
        super().__init__(parent)
        self.selected_sheets = selected_sheets
        self._setup_ui()

    def _setup_ui(self):
        layout = QVBoxLayout(self)
        layout.addWidget(QLabel("Выбери номер строки заголовка для каждого листа:"))
        scroll_area = QScrollArea()
        scroll_area.setWidgetResizable(True)
        scroll_content = QFrame()
        scroll_layout = QGridLayout(scroll_content)
        scroll_content.setLayout(scroll_layout)
        scroll_area.setWidget(scroll_content)

        self.sheet_to_combo = {}
        for row, sheet_name in enumerate(self.selected_sheets):
            label = QLabel(sheet_name)
            combo = QComboBox()
            combo.setMaximumWidth(100)
            combo.addItems([str(i) for i in range(1, 11)])
            combo.setCurrentIndex(0)
            scroll_layout.addWidget(label, row, 0)
            scroll_layout.addWidget(combo, row, 1)
            self.sheet_to_combo[sheet_name] = combo

        layout.addWidget(scroll_area)

        button_layout = QHBoxLayout()
        self.back_btn = QPushButton("Назад")
        self.next_btn = QPushButton("Далее")
        self.back_btn.clicked.connect(self.backClicked.emit)
        self.next_btn.clicked.connect(self._on_next_clicked)
        button_layout.addWidget(self.back_btn)
        button_layout.addWidget(self.next_btn)
        layout.addLayout(button_layout)

    def _on_next_clicked(self):
        result = {sheet: int(combo.currentText()) - 1 for sheet, combo in self.sheet_to_combo.items()}
        self.nextClicked.emit(result)