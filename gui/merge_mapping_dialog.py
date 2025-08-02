from __future__ import annotations

from typing import List

from PySide6.QtWidgets import (
    QDialog, QVBoxLayout, QHBoxLayout, QPushButton, QWidget,
    QLineEdit
)

from utils.i18n import tr, i18n
from core.drag_drop import DragDropLineEdit


class MappingRow(QWidget):
    def __init__(self, parent: QWidget | None = None):
        super().__init__(parent)
        layout = QHBoxLayout(self)
        layout.setSpacing(10)

        self.file_input = DragDropLineEdit(mode='file')
        self.file_input.setPlaceholderText(tr("Файл источник"))
        layout.addWidget(self.file_input)

        self.src_cols = QLineEdit()
        self.src_cols.setPlaceholderText(tr("Колонки источника"))
        layout.addWidget(self.src_cols)

        self.target_sheet = QLineEdit()
        self.target_sheet.setPlaceholderText(tr("Лист назначения"))
        layout.addWidget(self.target_sheet)

        self.target_cols = QLineEdit()
        self.target_cols.setPlaceholderText(tr("Колонки назначения"))
        layout.addWidget(self.target_cols)

        self.remove_btn = QPushButton("✕")
        layout.addWidget(self.remove_btn)

    def get_mapping(self):
        src = self.file_input.text()
        src_cols = [c.strip() for c in self.src_cols.text().split(',') if c.strip()]
        tgt_sheet = self.target_sheet.text().strip()
        tgt_cols = [c.strip() for c in self.target_cols.text().split(',') if c.strip()]
        return {
            'source': src,
            'source_columns': src_cols,
            'target_sheet': tgt_sheet,
            'target_columns': tgt_cols,
        }

    def retranslate(self):
        self.file_input.setPlaceholderText(tr("Файл источник"))
        self.src_cols.setPlaceholderText(tr("Колонки источника"))
        self.target_sheet.setPlaceholderText(tr("Лист назначения"))
        self.target_cols.setPlaceholderText(tr("Колонки назначения"))


class MergeMappingDialog(QDialog):
    """Dialog for configuring column mappings for multiple source files."""

    def __init__(self, parent: QWidget | None = None):
        super().__init__(parent)
        self.rows: List[MappingRow] = []
        self.setWindowTitle(tr("Настройки сопоставления"))
        layout = QVBoxLayout(self)

        self.mappings_layout = QVBoxLayout()
        layout.addLayout(self.mappings_layout)

        add_btn = QPushButton(tr("Добавить источник"))
        add_btn.clicked.connect(self.add_row)
        layout.addWidget(add_btn)
        self.add_btn = add_btn

        btn_layout = QHBoxLayout()
        ok_btn = QPushButton(tr("Готово"))
        cancel_btn = QPushButton(tr("Отмена"))
        ok_btn.clicked.connect(self.accept)
        cancel_btn.clicked.connect(self.reject)
        btn_layout.addStretch()
        btn_layout.addWidget(ok_btn)
        btn_layout.addWidget(cancel_btn)
        layout.addLayout(btn_layout)
        self.ok_btn = ok_btn
        self.cancel_btn = cancel_btn

        i18n.language_changed.connect(self.retranslate_ui)
        self.retranslate_ui()

    def add_row(self, file_path: str | None = None):
        row = MappingRow(self)
        row.remove_btn.clicked.connect(lambda: self.remove_row(row))
        self.mappings_layout.addWidget(row)
        self.rows.append(row)
        if file_path:
            row.file_input.setText(file_path)
        return row

    def remove_row(self, row: MappingRow):
        row.setParent(None)
        self.rows.remove(row)

    def get_mappings(self):
        mappings = []
        for row in self.rows:
            mp = row.get_mapping()
            if mp['source'] and mp['source_columns'] and mp['target_sheet'] and mp['target_columns']:
                mappings.append(mp)
        return mappings

    def retranslate_ui(self):
        self.setWindowTitle(tr("Настройки сопоставления"))
        self.add_btn.setText(tr("Добавить источник"))
        self.ok_btn.setText(tr("Готово"))
        self.cancel_btn.setText(tr("Отмена"))
        for r in self.rows:
            r.retranslate()
