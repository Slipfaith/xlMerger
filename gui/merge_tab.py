from __future__ import annotations

from typing import List
import os

from PySide6.QtWidgets import (
    QWidget, QVBoxLayout, QHBoxLayout, QPushButton, QMessageBox, QLabel, QGroupBox, QProgressBar
)

from PySide6.QtCore import QTimer, Qt

from utils.i18n import tr, i18n
from core.drag_drop import DragDropLineEdit
from core.merge_columns import merge_excel_columns
from .merge_mapping_dialog import MergeMappingDialog


class MergeTab(QWidget):
    def __init__(self):
        super().__init__()
        self.source_files: List[str] = []
        self.main_file = ""
        self.mappings = []
        self.is_processing = False
        self.worker = None
        self.init_ui()
        i18n.language_changed.connect(self.retranslate_ui)
        self.retranslate_ui()

    def init_ui(self):
        layout = QVBoxLayout(self)
        layout.setSpacing(20)
        layout.setContentsMargins(20, 20, 20, 20)

        # Группа для основного файла
        main_group = QGroupBox()
        main_layout = QVBoxLayout()
        self.main_label = QLabel()
        self.main_file_input = DragDropLineEdit(mode='file')
        self.main_file_input.fileSelected.connect(self.handle_main_file_selected)
        main_layout.addWidget(self.main_label)
        main_layout.addWidget(self.main_file_input)
        main_group.setLayout(main_layout)
        layout.addWidget(main_group)

        # Группа для файлов источников
        sources_group = QGroupBox()
        sources_layout = QVBoxLayout()
        self.sources_label = QLabel()
        self.sources_input = DragDropLineEdit(mode='files_or_folder')
        self.sources_input.filesSelected.connect(self.handle_files_selected)
        self.sources_input.folderSelected.connect(self.handle_folder_selected)
        sources_layout.addWidget(self.sources_label)
        sources_layout.addWidget(self.sources_input)
        sources_group.setLayout(sources_layout)
        layout.addWidget(sources_group)

        # Кнопки и прогресс
        button_layout = QVBoxLayout()
        button_layout.addStretch()

        self.configure_btn = QPushButton()
        self.configure_btn.clicked.connect(self.open_preview)
        self.configure_btn.setMinimumWidth(120)
        self.configure_btn.setMinimumHeight(35)

        button_layout.addWidget(self.configure_btn)

        # Прогресс бар и статус
        self.progress_widget = QWidget()
        progress_layout = QVBoxLayout(self.progress_widget)
        progress_layout.setContentsMargins(0, 10, 0, 10)

        self.status_label = QLabel()
        self.status_label.setAlignment(Qt.AlignCenter)
        self.status_label.setStyleSheet("color: #666; font-style: italic;")
        progress_layout.addWidget(self.status_label)

        self.progress_bar = QProgressBar()
        self.progress_bar.setVisible(False)
        self.progress_bar.setMinimumHeight(8)
        self.progress_bar.setStyleSheet("""
            QProgressBar {
                border: 1px solid #ccc;
                border-radius: 4px;
                text-align: center;
                background-color: #f0f0f0;
            }
            QProgressBar::chunk {
                background-color: #007bff;
                border-radius: 3px;
            }
        """)
        progress_layout.addWidget(self.progress_bar)

        button_layout.addWidget(self.progress_widget)

        self.merge_btn = QPushButton()
        self.merge_btn.clicked.connect(self.run_merge)
        self.merge_btn.setMinimumWidth(120)
        self.merge_btn.setMinimumHeight(35)
        self.merge_btn.setEnabled(False)

        button_layout.addWidget(self.merge_btn)

        layout.addLayout(button_layout)
        layout.addStretch()

    def handle_main_file_selected(self, file_path: str):
        self.main_file = file_path

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

    def open_preview(self):
        if not self.main_file:
            QMessageBox.critical(self, tr("Ошибка"), tr("Укажи основной Excel."))
            return
        if not self.source_files:
            QMessageBox.critical(self, tr("Ошибка"), tr("Укажи файлы источников."))
            return

        dialog = MergeMappingDialog(self.main_file, self)
        for f in self.source_files:
            dialog.add_row_with_file(f)
        if dialog.exec():
            self.mappings = dialog.get_mappings()
            self.merge_btn.setEnabled(True)
            self.status_label.setText(tr("Настройки сохранены. Готово к объединению."))

    def run_merge(self):
        if not self.main_file:
            QMessageBox.critical(self, tr("Ошибка"), tr("Укажи основной Excel."))
            return
        if not self.source_files:
            QMessageBox.critical(self, tr("Ошибка"), tr("Укажи файлы источников."))
            return
        if not self.mappings:
            QMessageBox.warning(self, tr("Предупреждение"),
                                tr("Сначала настройте сопоставления через кнопку Настроить."))
            return

        if self.is_processing:
            return

        self.is_processing = True
        self.merge_btn.setEnabled(False)
        self.configure_btn.setEnabled(False)

        # Показываем прогресс бар
        self.progress_bar.setVisible(True)
        self.progress_bar.setRange(0, 100)
        self.progress_bar.setValue(0)
        self.status_label.setText(tr("Начинаем объединение..."))

        # Создаем и запускаем worker в отдельном потоке
        self.worker = MergeWorker(self.main_file, self.mappings)
        self.worker.finished.connect(self.on_merge_finished)
        self.worker.error.connect(self.on_merge_error)
        self.worker.progress.connect(self.on_progress_update)
        self.worker.start()

    def on_progress_update(self, value, message):
        self.progress_bar.setValue(value)
        self.status_label.setText(message)

    def on_merge_finished(self, output_file):
        self.worker = None
        self.progress_bar.setVisible(False)
        self.status_label.setText(tr("Объединение завершено успешно!"))

        self.is_processing = False
        self.merge_btn.setEnabled(True)
        self.configure_btn.setEnabled(True)

        QMessageBox.information(
            self, tr("Успех"), tr("Файл сохранён: {output}").format(output=output_file)
        )

    def on_merge_error(self, error_message):
        self.worker = None
        self.progress_bar.setVisible(False)
        self.status_label.setText(tr("Ошибка при объединении"))

        self.is_processing = False
        self.merge_btn.setEnabled(True)
        self.configure_btn.setEnabled(True)

        QMessageBox.critical(self, tr("Ошибка"), error_message)

    def retranslate_ui(self):
        self.main_label.setText(tr("Основной Excel файл:"))
        self.main_file_input.setPlaceholderText(tr("Перетащи или кликни дважды"))
        self.sources_label.setText(tr("Файлы источников:"))
        self.sources_input.setPlaceholderText(tr("Перетащи или кликни дважды"))
        self.configure_btn.setText(tr("Настроить"))
        self.merge_btn.setText(tr("Объединить"))

        if hasattr(self, 'status_label') and not self.mappings:
            self.status_label.setText(tr("Сначала настройте сопоставления"))