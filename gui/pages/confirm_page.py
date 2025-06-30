from PySide6.QtWidgets import (
    QWidget, QVBoxLayout, QLabel, QHBoxLayout, QPushButton,
    QTableWidget, QHeaderView, QComboBox, QCheckBox, QTableWidgetItem
)
from PySide6.QtCore import Qt, Signal
from utils.i18n import tr, i18n

class ConfirmPage(QWidget):
    backClicked = Signal()
    startClicked = Signal()

    def __init__(self, items, available_columns, preserve_formatting=False, parent=None):
        super().__init__(parent)
        self.items = items
        self.available_columns = [''] + sorted(set(available_columns))
        self._preserve_formatting = preserve_formatting
        self._build_ui()
        i18n.language_changed.connect(self.retranslate_ui)
        self.retranslate_ui()

    def _build_ui(self):
        layout = QVBoxLayout(self)

        total = len(self.items)
        missing = sum(1 for _, col in self.items if not col)
        summary = f"Всего: {total}. Сопоставлено: {total - missing}. Не выбрано: {missing}."
        self.summary_label = QLabel(summary)
        if missing:
            self.summary_label.setStyleSheet("color: #b91c1c; font-weight: 500;")
        else:
            self.summary_label.setStyleSheet("color: #15803d; font-weight: 500;")
        layout.addWidget(self.summary_label)

        # Таблица сопоставлений
        self.table = QTableWidget(len(self.items), 2, self)
        self.table.setHorizontalHeaderLabels(["Файл/Папка", "Столбец"])
        self.table.verticalHeader().setVisible(False)
        self.table.horizontalHeader().setSectionResizeMode(QHeaderView.Stretch)
        self.table.setEditTriggers(QTableWidget.NoEditTriggers)
        self.table.setSelectionMode(QTableWidget.NoSelection)
        self.table.setFocusPolicy(Qt.NoFocus)
        self.table.setStyleSheet("""
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
            self.table.setItem(row, 0, item_file)

            combo = QComboBox()
            combo.addItems(self.available_columns)
            if col_name:
                idx = combo.findText(col_name)
                if idx != -1:
                    combo.setCurrentIndex(idx)
            self.table.setCellWidget(row, 1, combo)

        layout.addWidget(self.table)

        self.format_checkbox = QCheckBox(tr("Копировать с сохранением форматирования"))
        self.format_checkbox.setChecked(self._preserve_formatting)
        self.format_checkbox.setStyleSheet("QCheckBox { padding: 4px; }")
        layout.addWidget(self.format_checkbox)

        # Кнопки
        btn_layout = QHBoxLayout()
        self.btn_back = QPushButton(tr("Назад"))
        self.btn_start = QPushButton(tr("Начать"))

        # Стили убраны, используем системные границы кнопок

        # Чтобы обе кнопки были одинаковые по размеру (ширине):
        max_width = max(self.btn_back.sizeHint().width(), self.btn_start.sizeHint().width())
        self.btn_back.setMinimumWidth(max_width)
        self.btn_start.setMinimumWidth(max_width)

        self.btn_back.clicked.connect(self.backClicked.emit)
        self.btn_start.clicked.connect(self.startClicked.emit)
        self.btn_start.setEnabled(True)
        btn_layout.addWidget(self.btn_back)
        btn_layout.addWidget(self.btn_start)
        layout.addLayout(btn_layout)

        self.setLayout(layout)

    def retranslate_ui(self):
        total = len(self.items)
        missing = sum(1 for _, col in self.items if not col)
        self.summary_label.setText(tr("Всего: {total}. Сопоставлено: {mapped}. Не выбрано: {missing}." ).format(total=total, mapped=total - missing, missing=missing))
        self.btn_back.setText(tr("Назад"))
        self.btn_start.setText(tr("Начать"))

    def get_current_mapping(self):
        mapping = {}
        for row, (name, _) in enumerate(self.items):
            combo = self.table.cellWidget(row, 1)
            if combo:
                mapping[name] = combo.currentText()
        return mapping

    def is_format_preserved(self):
        return self.format_checkbox.isChecked()

