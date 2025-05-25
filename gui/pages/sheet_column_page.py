# gui/pages/sheet_column_page.py

from PySide6.QtWidgets import (
    QWidget, QVBoxLayout, QLabel, QScrollArea, QFrame, QGridLayout, QLineEdit,
    QPushButton, QHBoxLayout
)
from PySide6.QtCore import Signal, Qt

class SheetColumnPage(QWidget):
    backClicked = Signal()
    nextClicked = Signal(dict)

    def __init__(self, selected_sheets, default_column=""):
        super().__init__()
        self.setWindowTitle("Соответствие лист-столбец")
        self.selected_sheets = selected_sheets
        self.default_column = default_column
        self.sheet_to_column_widgets = {}

        layout = QVBoxLayout()
        layout.addWidget(QLabel("Из какого столбца на каждом листе копировать?"))

        scroll_area = QScrollArea()
        scroll_area.setWidgetResizable(True)
        scroll_content = QFrame()
        scroll_layout = QGridLayout(scroll_content)
        scroll_content.setLayout(scroll_layout)
        scroll_area.setWidget(scroll_content)

        for row, sheet_name in enumerate(self.selected_sheets):
            sheet_label = QLabel(sheet_name)
            column_entry = QLineEdit()
            column_entry.setMaximumWidth(100)
            column_entry.setText(self.default_column)
            scroll_layout.addWidget(sheet_label, row, 0)
            scroll_layout.addWidget(column_entry, row, 1)
            self.sheet_to_column_widgets[sheet_name] = column_entry

        layout.addWidget(scroll_area)
        button_layout = QHBoxLayout()
        back_button = QPushButton("Назад")
        next_button = QPushButton("Далее")
        back_button.clicked.connect(self.backClicked.emit)
        next_button.clicked.connect(self._on_next)
        button_layout.addWidget(back_button)
        button_layout.addWidget(next_button)
        layout.addLayout(button_layout)
        self.setLayout(layout)

    def _on_next(self):
        sheet_to_column = {k: v.text().strip() for k, v in self.sheet_to_column_widgets.items()}
        self.nextClicked.emit(sheet_to_column)