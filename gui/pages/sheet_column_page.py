from PySide6.QtWidgets import (
    QWidget, QVBoxLayout, QLabel, QScrollArea, QFrame, QLineEdit,
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

        layout = QVBoxLayout(self)
        layout.addWidget(QLabel("Из какого столбца на каждом листе копировать?"))

        scroll_area = QScrollArea(self)
        scroll_area.setWidgetResizable(True)
        scroll_content = QWidget()
        scroll_content_layout = QVBoxLayout(scroll_content)
        scroll_content_layout.setSpacing(12)  # чуть воздуха между парами

        for sheet_name in self.selected_sheets:
            # ОДИН БЛОК: имя + поле под ним, компактно!
            block = QWidget()
            block_layout = QVBoxLayout(block)
            block_layout.setSpacing(2)
            block_layout.setContentsMargins(0, 0, 0, 0)

            label = QLabel(sheet_name)
            label.setAlignment(Qt.AlignLeft)
            label.setStyleSheet("font-size: 11pt;")  # <--- обязательно! Чтобы совпадал с entry

            entry = QLineEdit()
            entry.setMaximumWidth(120)
            entry.setText(self.default_column)
            entry.setStyleSheet("font-size: 11pt; padding: 0px; margin: 0px;")

            block_layout.addWidget(label)
            block_layout.addWidget(entry)
            block_layout.addStretch(0)  # нет искусственного растяжения!

            self.sheet_to_column_widgets[sheet_name] = entry
            scroll_content_layout.addWidget(block, alignment=Qt.AlignLeft)

        scroll_content_layout.addStretch(1)  # чтобы все блоки были сверху!
        scroll_area.setWidget(scroll_content)
        layout.addWidget(scroll_area)

        button_layout = QHBoxLayout()
        back_button = QPushButton("Назад")
        next_button = QPushButton("Далее")
        back_button.clicked.connect(self.backClicked.emit)
        next_button.clicked.connect(self._on_next)
        button_layout.addWidget(back_button)
        button_layout.addWidget(next_button)
        layout.addLayout(button_layout)

    def _on_next(self):
        sheet_to_column = {k: v.text().strip() for k, v in self.sheet_to_column_widgets.items()}
        self.nextClicked.emit(sheet_to_column)