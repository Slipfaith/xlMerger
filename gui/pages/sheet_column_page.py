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
        scroll_content = QFrame()
        vbox = QVBoxLayout(scroll_content)
        vbox.setSpacing(0)
        vbox.setContentsMargins(20, 12, 20, 12)
        scroll_content.setLayout(vbox)
        scroll_area.setWidget(scroll_content)

        for sheet_name in self.selected_sheets:
            label = QLabel(sheet_name)
            label.setAlignment(Qt.AlignLeft)
            label.setStyleSheet("font-size: 14px; margin-bottom: 0px; padding-bottom: 0px;")

            entry = QLineEdit()
            entry.setMaximumWidth(120)
            entry.setText(self.default_column)
            entry.setStyleSheet("font-size: 11pt; padding: 0px; margin: 0px;")

            vbox.addWidget(label, alignment=Qt.AlignLeft)
            vbox.addWidget(entry, alignment=Qt.AlignLeft)
            vbox.addSpacing(2)
            self.sheet_to_column_widgets[sheet_name] = entry

        layout.addWidget(scroll_area)

        # Кнопки - строго справа
        button_layout = QHBoxLayout()
        self.back_btn = QPushButton("Назад")
        self.next_btn = QPushButton("Далее")
        self.back_btn.clicked.connect(self.backClicked.emit)
        self.next_btn.clicked.connect(self._on_next_clicked)
        button_layout.addStretch()
        button_layout.addWidget(self.back_btn)
        button_layout.addWidget(self.next_btn)
        layout.addLayout(button_layout)
        vbox.addStretch(1)

    def _on_next_clicked(self):
        sheet_to_column = {k: v.text().strip() for k, v in self.sheet_to_column_widgets.items()}
        self.nextClicked.emit(sheet_to_column)
