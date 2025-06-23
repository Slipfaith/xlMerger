from PySide6.QtWidgets import (
    QWidget, QVBoxLayout, QLabel, QScrollArea, QFrame, QComboBox,
    QPushButton, QHBoxLayout, QSizePolicy
)
from PySide6.QtCore import Qt, Signal

class HeaderRowPage(QWidget):
    backClicked = Signal()
    nextClicked = Signal(dict)

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
        vbox = QVBoxLayout(scroll_content)
        vbox.setSpacing(0)
        vbox.setContentsMargins(20, 12, 20, 12)
        scroll_content.setLayout(vbox)
        scroll_area.setWidget(scroll_content)

        self.sheet_to_combo = {}

        for sheet_name in self.selected_sheets:
            label = QLabel(sheet_name)
            label.setAlignment(Qt.AlignLeft)
            label.setStyleSheet("font-size: 14px; margin-bottom: 0px; padding-bottom: 0px;")
            combo = QComboBox()
            combo.setStyleSheet("font-size: 14pt;")
            combo.setMaximumWidth(72)
            combo.setMinimumWidth(48)
            combo.setStyleSheet(
                "margin-top: 0px; margin-bottom: 0px; padding: 0px;"
            )
            combo.addItems([str(i) for i in range(1, 21)])
            combo.setCurrentIndex(0)
            combo.setSizePolicy(QSizePolicy.Fixed, QSizePolicy.Fixed)

            vbox.addWidget(label, alignment=Qt.AlignLeft)
            vbox.addWidget(combo, alignment=Qt.AlignLeft)
            # Минимизируем зазор: добавляем spacer всего 2px или вообще убираем
            vbox.addSpacing(2)
            self.sheet_to_combo[sheet_name] = combo

        layout.addWidget(scroll_area)

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
        result = {sheet: int(combo.currentText()) - 1 for sheet, combo in self.sheet_to_combo.items()}
        self.nextClicked.emit(result)