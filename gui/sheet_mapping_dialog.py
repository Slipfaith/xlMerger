from PySide6.QtWidgets import (
    QDialog, QVBoxLayout, QGridLayout, QLabel, QComboBox,
    QPushButton, QHBoxLayout
)
from utils.i18n import tr
import os


class SheetMappingDialog(QDialog):
    """Dialog to map sheet names from translation files."""

    def __init__(self, main_sheets, file_to_sheets, auto_map=None, parent=None):
        super().__init__(parent)
        self.main_sheets = list(main_sheets)
        self.file_to_sheets = file_to_sheets
        self.auto_map = auto_map or {}
        self.comboboxes = {}
        self._build_ui()

    def _build_ui(self):
        self.setWindowTitle(tr("Сопоставление листов"))
        layout = QVBoxLayout(self)

        grid = QGridLayout()
        grid.addWidget(QLabel(tr("Файл")), 0, 0)
        for col, ms in enumerate(self.main_sheets, start=1):
            grid.addWidget(QLabel(ms), 0, col)

        for row, (file, sheets) in enumerate(self.file_to_sheets.items(), start=1):
            grid.addWidget(QLabel(os.path.basename(file)), row, 0)
            for col, ms in enumerate(self.main_sheets, start=1):
                combo = QComboBox()
                combo.addItems(sheets)
                pre = self.auto_map.get(file, {}).get(ms)
                if pre and pre in sheets:
                    combo.setCurrentText(pre)
                elif ms in sheets:
                    combo.setCurrentText(ms)
                self.comboboxes[(file, ms)] = combo
                grid.addWidget(combo, row, col)
        layout.addLayout(grid)

        btn_layout = QHBoxLayout()
        ok_btn = QPushButton(tr("Готово"))
        cancel_btn = QPushButton(tr("Отмена"))
        ok_btn.clicked.connect(self.accept)
        cancel_btn.clicked.connect(self.reject)
        btn_layout.addStretch()
        btn_layout.addWidget(ok_btn)
        btn_layout.addWidget(cancel_btn)
        layout.addLayout(btn_layout)

    def get_mapping(self):
        mapping = {}
        for (file, ms), combo in self.comboboxes.items():
            mapping.setdefault(file, {})[ms] = combo.currentText()
        return mapping
