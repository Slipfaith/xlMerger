from __future__ import annotations

from typing import List
import os

from PySide6.QtWidgets import (
    QWidget, QVBoxLayout, QPushButton, QMessageBox
)

from utils.i18n import tr, i18n
from core.drag_drop import DragDropLineEdit
from core.merge_columns import merge_excel_columns
from .merge_mapping_dialog import MergeMappingDialog


class MergeTab(QWidget):
    def __init__(self):
        super().__init__()
        self.source_files: List[str] = []
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
        self.sources_input.filesSelected.connect(self.handle_files_selected)
        self.sources_input.folderSelected.connect(self.handle_folder_selected)
        layout.addWidget(self.sources_input)

        layout.addStretch()

        merge_btn = QPushButton(tr("Объединить"))
        merge_btn.clicked.connect(self.run_merge)
        layout.addWidget(merge_btn)
        self.merge_btn = merge_btn

    def handle_files_selected(self, files: List[str]):
        self.source_files = files

    def handle_folder_selected(self, folder: str):
        excel_exts = ('.xlsx', '.xls')
        if os.path.isdir(folder):
            self.source_files = [
                os.path.join(folder, f)
                for f in os.listdir(folder)
                if f.lower().endswith(excel_exts)
            ]
        else:
            self.source_files = []

    def run_merge(self):
        main_file = self.main_file_input.text()
        if not main_file:
            QMessageBox.critical(self, tr("Ошибка"), tr("Укажи общий Excel."))
            return
        if not self.source_files:
            QMessageBox.critical(self, tr("Ошибка"), tr("Укажи файлы источников."))
            return

        dialog = MergeMappingDialog(main_file, self)
        for f in self.source_files:
            dialog.add_row(f)
        if dialog.exec():
            mappings = dialog.get_mappings()
            try:
                output = merge_excel_columns(main_file, mappings)
                QMessageBox.information(
                    self, tr("Успех"), tr("Файл сохранён: {output}").format(output=output)
                )
            except Exception as e:
                QMessageBox.critical(self, tr("Ошибка"), str(e))

    def retranslate_ui(self):
        self.main_file_input.setPlaceholderText(tr("Общий Excel"))
        self.sources_input.setPlaceholderText(tr("Файлы источников"))
        self.merge_btn.setText(tr("Объединить"))
