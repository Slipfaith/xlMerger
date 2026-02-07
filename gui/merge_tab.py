# -*- coding: utf-8 -*-
from __future__ import annotations

from typing import List
import os
import subprocess
import platform

from PySide6.QtWidgets import (
    QWidget, QVBoxLayout, QHBoxLayout, QPushButton, QMessageBox, QLabel, QGroupBox, QProgressBar
)

from PySide6.QtCore import Qt, QThread, Signal, QUrl

from utils.i18n import tr, i18n
from core.drag_drop import DragDropLineEdit
from core.merge_columns import merge_excel_columns
from .multi_merge_mapping_dialog import MultiMergeMappingDialog
from .style_system import set_button_variant, set_label_role, set_label_state


class MergeWorker(QThread):
    finished = Signal(list)
    error = Signal(str)
    progress = Signal(int, str)

    def __init__(self, tasks):
        super().__init__()
        self.tasks = tasks
        self.outputs = []

    def run(self):
        try:
            total_steps = sum(len(task.get("mappings", [])) for task in self.tasks) or 1
            completed = 0
            for idx_task, task in enumerate(self.tasks):
                main_file = task.get("target")
                mappings = task.get("mappings", [])
                self.progress.emit(1, tr("Подготовка {file}").format(file=os.path.basename(main_file)))

                def progress_callback(idx, total, mapping):
                    source_file = os.path.basename(mapping.get("source", ""))
                    target_sheet = mapping.get("target_sheet", "")
                    step = int(((completed + idx) / total_steps) * 99) + 1
                    message = tr("{file}: лист {sheet}, источник {src}").format(
                        file=os.path.basename(main_file),
                        sheet=target_sheet,
                        src=source_file
                    )
                    self.progress.emit(step, message)

                output = merge_excel_columns(main_file, mappings, progress_callback=progress_callback)
                completed += len(mappings)
                self.outputs.append((main_file, output))

            self.progress.emit(100, tr("Объединение завершено!"))
            self.finished.emit(self.outputs)

        except Exception as e:
            self.error.emit(str(e))


class MergeTab(QWidget):
    def __init__(self):
        super().__init__()
        self.source_files: List[str] = []
        self.target_files: List[str] = []
        self.merge_tasks = []
        self.is_processing = False
        self.worker = None
        self.init_ui()
        i18n.language_changed.connect(self.retranslate_ui)
        self.retranslate_ui()

    def init_ui(self):
        layout = QVBoxLayout(self)
        layout.setSpacing(20)
        layout.setContentsMargins(20, 20, 20, 20)

        main_group = QGroupBox()
        main_layout = QVBoxLayout()
        self.main_label = QLabel()
        self.target_files_input = DragDropLineEdit(mode='files')
        self.target_files_input.filesSelected.connect(self.handle_target_files_selected)
        main_layout.addWidget(self.main_label)
        main_layout.addWidget(self.target_files_input)
        main_group.setLayout(main_layout)
        layout.addWidget(main_group)

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

        button_layout = QVBoxLayout()

        self.configure_btn = QPushButton()
        self.configure_btn.clicked.connect(self.open_preview)
        self.configure_btn.setMinimumWidth(120)
        self.configure_btn.setMinimumHeight(35)
        set_button_variant(self.configure_btn, "secondary")

        self.progress_widget = QWidget()
        progress_layout = QVBoxLayout(self.progress_widget)
        progress_layout.setContentsMargins(0, 10, 0, 10)

        self.status_label = QLabel()
        self.status_label.setAlignment(Qt.AlignCenter)
        set_label_role(self.status_label, "muted")
        progress_layout.addWidget(self.status_label)

        self.progress_bar = QProgressBar()
        self.progress_bar.setVisible(False)
        self.progress_bar.setMinimumHeight(8)
        progress_layout.addWidget(self.progress_bar)

        self.file_link = QLabel()
        self.file_link.setAlignment(Qt.AlignCenter)
        self.file_link.setVisible(False)
        self.file_link.setTextFormat(Qt.RichText)
        self.file_link.setOpenExternalLinks(False)
        self.file_link.setTextInteractionFlags(Qt.TextBrowserInteraction)
        self.file_link.linkActivated.connect(self.open_file_location)
        set_label_role(self.file_link, "link")
        progress_layout.addWidget(self.file_link)

        button_layout.addWidget(self.progress_widget)

        self.merge_btn = QPushButton()
        self.merge_btn.clicked.connect(self.run_merge)
        self.merge_btn.setMinimumWidth(120)
        self.merge_btn.setMinimumHeight(35)
        self.merge_btn.setEnabled(False)
        set_button_variant(self.merge_btn, "orange")

        action_row = QHBoxLayout()
        action_row.setSpacing(10)
        action_row.addWidget(self.configure_btn)
        action_row.addWidget(self.merge_btn)
        button_layout.addLayout(action_row)

        layout.addStretch()
        layout.addLayout(button_layout)

    def handle_target_files_selected(self, files: List[str]):
        combined = []
        for path in [*self.target_files, *files]:
            if path not in combined:
                combined.append(path)
        self.target_files = combined
        self.target_files_input.setText('; '.join([self.target_files_input._short_name(f) for f in combined]))

    def handle_files_selected(self, files: List[str]):
        self.source_files = files

    def handle_folder_selected(self, folder: str):
        excel_exts = ('.xlsx', '.xls')
        if os.path.isdir(folder):
            collected = []
            for root, _, files in os.walk(folder):
                for fname in files:
                    if fname.lower().endswith(excel_exts):
                        collected.append(os.path.join(root, fname))
            self.source_files = collected
        else:
            self.source_files = []

    def open_preview(self):
        if not self.target_files:
            QMessageBox.critical(self, tr("Ошибка"), tr("Укажи целевые Excel файлы."))
            return
        if not self.source_files:
            QMessageBox.critical(self, tr("Ошибка"), tr("Укажи файлы или папки с переводами."))
            return

        dialog = MultiMergeMappingDialog(self.target_files, self.source_files, self)
        if dialog.exec():
            self.merge_tasks = dialog.get_tasks()
            if self.merge_tasks:
                self.merge_btn.setEnabled(True)
                self.status_label.setText(tr("Настройки сохранены. Готово к объединению."))

    def run_merge(self):
        if not self.target_files:
            QMessageBox.critical(self, tr("Ошибка"), tr("Укажи целевые Excel файлы."))
            return
        if not self.source_files:
            QMessageBox.critical(self, tr("Ошибка"), tr("Укажи файлы или папки с переводами."))
            return
        if not self.merge_tasks:
            QMessageBox.warning(self, tr("Предупреждение"),
                                tr("Сначала настройте сопоставления через кнопку Настроить."))
            return

        if self.is_processing:
            return

        self.is_processing = True
        self.merge_btn.setEnabled(False)
        self.configure_btn.setEnabled(False)
        self.file_link.setVisible(False)

        self.progress_bar.setVisible(True)
        self.progress_bar.setRange(0, 100)
        self.progress_bar.setValue(0)
        self.status_label.setText(tr("Начинаем объединение..."))

        self.worker = MergeWorker(self.merge_tasks)
        self.worker.finished.connect(self.on_merge_finished)
        self.worker.error.connect(self.on_merge_error)
        self.worker.progress.connect(self.on_progress_update)
        self.worker.start()

    def on_progress_update(self, value, message):
        self.progress_bar.setValue(value)
        self.status_label.setText(message)

    def on_merge_finished(self, outputs):
        self.worker = None
        self.output_files = outputs
        self.progress_bar.setValue(100)
        self.status_label.setText(tr("Объединение завершено успешно!"))
        set_label_state(self.status_label, "success")

        links = []
        for _, path in outputs:
            filename = os.path.basename(path)
            href = QUrl.fromLocalFile(path).toString()
            links.append(f'<a href="{href}">{filename}</a>')
        if links:
            self.file_link.setText("<br>".join(links))
            self.file_link.setVisible(True)

        self.is_processing = False
        self.merge_btn.setEnabled(True)
        self.configure_btn.setEnabled(True)

    def on_merge_error(self, error_message):
        self.worker = None
        self.progress_bar.setVisible(False)
        self.file_link.setVisible(False)
        self.status_label.setText(tr("Ошибка при объединении"))
        set_label_state(self.status_label, "error")

        self.is_processing = False
        self.merge_btn.setEnabled(True)
        self.configure_btn.setEnabled(True)

        QMessageBox.critical(self, tr("Ошибка"), error_message)

    def open_file_location(self, link):
        if not hasattr(self, 'output_files') or not self.output_files:
            return

        clicked_path = QUrl(link).toLocalFile() if link else ""
        if not clicked_path:
            clicked_path = self.output_files[0][1]

        folder = os.path.dirname(clicked_path)
        if platform.system() == 'Windows':
            subprocess.run(['explorer', '/select,', os.path.normpath(clicked_path)])
        elif platform.system() == 'Darwin':
            subprocess.run(['open', '-R', clicked_path])
        else:
            subprocess.run(['xdg-open', folder])

    def retranslate_ui(self):
        self.main_label.setText(tr("Целевые Excel файлы:"))
        self.target_files_input.setPlaceholderText(tr("Перетащи или кликни дважды"))
        self.sources_label.setText(tr("Файлы источников:"))
        self.sources_input.setPlaceholderText(tr("Перетащи или кликни дважды"))
        self.configure_btn.setText(tr("Настроить"))
        self.merge_btn.setText(tr("Объединить"))

        if hasattr(self, 'status_label') and not self.merge_tasks:
            set_label_state(self.status_label, "")
            self.status_label.setText(tr("Сначала настройте сопоставления"))
