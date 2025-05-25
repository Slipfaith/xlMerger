from PySide6.QtWidgets import (
    QWidget, QVBoxLayout, QLabel, QHBoxLayout, QPushButton,
    QTableWidget, QTableWidgetItem, QHeaderView
)
from PySide6.QtCore import Qt, Signal

class ConfirmPage(QWidget):
    backClicked = Signal()
    startClicked = Signal()

    def __init__(self, items, parent=None):
        super().__init__(parent)
        self.items = items
        self._build_ui()

    def _build_ui(self):
        layout = QVBoxLayout(self)

        total = len(self.items)
        missing = sum(1 for _, col in self.items if not col)
        summary = f"Всего: {total}. Сопоставлено: {total - missing}. Не выбрано: {missing}."
        lbl = QLabel(summary)
        if missing:
            lbl.setStyleSheet("color: #b91c1c; font-weight: 500;")
        else:
            lbl.setStyleSheet("color: #15803d; font-weight: 500;")
        layout.addWidget(lbl)

        # Таблица сопоставлений
        table = QTableWidget(len(self.items), 2, self)
        table.setHorizontalHeaderLabels(["Файл/Папка", "Столбец"])
        table.verticalHeader().setVisible(False)
        table.horizontalHeader().setSectionResizeMode(QHeaderView.Stretch)
        table.setEditTriggers(QTableWidget.NoEditTriggers)
        table.setSelectionMode(QTableWidget.NoSelection)
        table.setFocusPolicy(Qt.NoFocus)
        table.setStyleSheet("""
            QTableWidget {
                border: none;
                font-size: 14px;
            }
            QTableWidget::item {
                padding: 6px;
            }
        """)

        for row, (name, col_name) in enumerate(self.items):
            # Короткое имя + тултип
            short = name
            if "\\" in name or "/" in name:
                short = name.split("\\")[-1].split("/")[-1]
            item_file = QTableWidgetItem(short)
            item_file.setToolTip(name)
            table.setItem(row, 0, item_file)

            item_col = QTableWidgetItem(col_name if col_name else "(не выбрано)")
            if not col_name:
                item_col.setForeground(Qt.red)
                font = item_col.font()
                font.setBold(True)
                item_col.setFont(font)
            else:
                item_col.setForeground(Qt.darkGreen)
                font = item_col.font()
                font.setBold(True)
                item_col.setFont(font)
            table.setItem(row, 1, item_col)

        layout.addWidget(table)

        # Кнопки
        btn_layout = QHBoxLayout()
        btn_back = QPushButton("Назад")
        btn_start = QPushButton("Начать")

        # Стили только для кнопки "Начать"
        btn_start.setStyleSheet("""
            QPushButton {
                background-color: #f47929;
                color: white;
                border-radius: 6px;
                padding: 4px 14px;
                font-size: 13px;
                font-weight: bold;
            }
            QPushButton:hover {
                background-color: #65d88f;
                color: #222;
            }
            QPushButton:pressed {
                background-color: #41bb6f;
                color: white;
            }
            QPushButton:disabled {
                background-color: #cccccc;
                color: #888888;
            }
        """)

        # Чтобы обе кнопки были одинаковые по размеру (ширине):
        max_width = max(btn_back.sizeHint().width(), btn_start.sizeHint().width())
        btn_back.setMinimumWidth(max_width)
        btn_start.setMinimumWidth(max_width)

        btn_back.clicked.connect(self.backClicked.emit)
        btn_start.clicked.connect(self.startClicked.emit)
        btn_start.setEnabled(True)
        btn_layout.addWidget(btn_back)
        btn_layout.addWidget(btn_start)
        layout.addLayout(btn_layout)

        self.setLayout(layout)