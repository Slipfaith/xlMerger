from __future__ import annotations

from typing import List

from PySide6.QtWidgets import (
    QWidget, QVBoxLayout, QHBoxLayout, QPushButton, QMessageBox, QLineEdit
)

from utils.i18n import tr, i18n
from core.drag_drop import DragDropLineEdit
from core.merge_columns import merge_excel_columns


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


class MergeTab(QWidget):
    def __init__(self):
        super().__init__()
        self.rows: List[MappingRow] = []
        self.init_ui()
        i18n.language_changed.connect(self.retranslate_ui)
        self.retranslate_ui()

    def init_ui(self):
        layout = QVBoxLayout(self)
        layout.setSpacing(15)

        self.main_file_input = DragDropLineEdit(mode='file')
        self.main_file_input.setPlaceholderText(tr("Общий Excel"))
        layout.addWidget(self.main_file_input)

        self.sources_input = DragDropLineEdit(mode='files_or_folder')
        self.sources_input.setPlaceholderText(tr("Файлы источников"))
        self.sources_input.filesSelected.connect(self.add_rows_from_files)
        layout.addWidget(self.sources_input)

        self.mappings_layout = QVBoxLayout()
        layout.addLayout(self.mappings_layout)

        add_btn = QPushButton(tr("Добавить источник"))
        add_btn.clicked.connect(self.add_row)
        layout.addWidget(add_btn)
        self.add_btn = add_btn

        layout.addStretch()

        merge_btn = QPushButton(tr("Объединить"))
        merge_btn.clicked.connect(self.run_merge)
        layout.addWidget(merge_btn)
        self.merge_btn = merge_btn

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

    def add_rows_from_files(self, files: List[str]):
        for f in files:
            self.add_row(f)

    def run_merge(self):
        main_file = self.main_file_input.text()
        if not main_file:
            QMessageBox.critical(self, tr("Ошибка"), tr("Укажи общий Excel."))
            return

        mappings = []
        for row in self.rows:
            mp = row.get_mapping()
            if not (mp['source'] and mp['source_columns'] and mp['target_sheet'] and mp['target_columns']):
                QMessageBox.critical(self, tr("Ошибка"), tr("Заполни все поля."))
                return
            mappings.append(mp)

        try:
            output = merge_excel_columns(main_file, mappings)
            QMessageBox.information(self, tr("Успех"), tr("Файл сохранён: {output}").format(output=output))
        except Exception as e:
            QMessageBox.critical(self, tr("Ошибка"), str(e))

    def retranslate_ui(self):
        self.main_file_input.setPlaceholderText(tr("Общий Excel"))
        self.sources_input.setPlaceholderText(tr("Файлы источников"))
        self.add_btn.setText(tr("Добавить источник"))
        self.merge_btn.setText(tr("Объединить"))
        for r in self.rows:
            r.retranslate()
