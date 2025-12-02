from __future__ import annotations

from typing import Dict, List
import os
import subprocess
import platform

from PySide6.QtWidgets import (
    QWidget, QVBoxLayout, QHBoxLayout, QPushButton, QMessageBox, QLabel, QGroupBox, QProgressBar
)

from PySide6.QtCore import Qt, QThread, Signal

from utils.i18n import tr, i18n
from core.drag_drop import DragDropLineEdit
from core.merge_columns import merge_excel_columns
from core.multi_target_merge import (
    TargetMergePlan,
    merge_translations_to_many_targets,
    suggest_target_mapping,
)
from .merge_mapping_dialog import MergeMappingDialog


class MergeWorker(QThread):
    finished = Signal(str)
    error = Signal(str)
    progress = Signal(int, str)

    def __init__(self, main_file, mappings):
        super().__init__()
        self.main_file = main_file
        self.mappings = mappings
        self.output_file = None

    def run(self):
        try:
            self.progress.emit(1, tr("Подготовка файлов..."))

            def progress_callback(idx, total, mapping):
                source_file = os.path.basename(mapping.get("source", ""))
                target_sheet = mapping.get("target_sheet", "")
                progress_percent = int((idx / total) * 99) + 1
                message = tr("Обрабатывается файл: {file}, лист: {sheet}").format(
                    file=source_file,
                    sheet=target_sheet
                )
                self.progress.emit(progress_percent, message)

            self.output_file = merge_excel_columns(self.main_file, self.mappings, progress_callback=progress_callback)

            self.progress.emit(100, tr("Объединение завершено!"))
            self.finished.emit(self.output_file)

        except Exception as e:
            self.error.emit(str(e))


class MultiTargetMergeWorker(QThread):
    finished = Signal(list)
    error = Signal(str)
    progress = Signal(int, str)

    def __init__(self, plans: List[TargetMergePlan], sources: List[str], manual_mapping: Dict[str, str] | None = None):
        super().__init__()
        self.plans = plans
        self.sources = sources
        self.manual_mapping = manual_mapping or {}
        self.outputs: List[str] = []

    def run(self):
        try:
            self.progress.emit(1, tr("Подготовка файлов..."))

            def progress_callback(percent: int, message: str):
                self.progress.emit(percent, message)

            self.outputs = merge_translations_to_many_targets(
                self.plans,
                self.sources,
                manual_mapping=self.manual_mapping,
                progress_callback=progress_callback,
            )

            self.progress.emit(100, tr("Объединение завершено!"))
            self.finished.emit(self.outputs)
        except Exception as e:
            self.error.emit(str(e))


class MergeTab(QWidget):
    def __init__(self):
        super().__init__()
        self.source_files: List[str] = []
        self.main_files: List[str] = []
        self.mappings: List[dict] = []
        self.plans: List[TargetMergePlan] = []
        self.manual_mapping: Dict[str, str] = {}
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
        self.main_file_input = DragDropLineEdit(mode='files_or_folder')
        self.main_file_input.filesSelected.connect(self.handle_main_files_selected)
        self.main_file_input.folderSelected.connect(self.handle_main_folder_selected)
        main_layout.addWidget(self.main_label)
        main_layout.addWidget(self.main_file_input)
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
        button_layout.addStretch()

        self.configure_btn = QPushButton()
        self.configure_btn.clicked.connect(self.open_preview)
        self.configure_btn.setMinimumWidth(120)
        self.configure_btn.setMinimumHeight(35)

        button_layout.addWidget(self.configure_btn)

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

        self.file_link = QLabel()
        self.file_link.setAlignment(Qt.AlignCenter)
        self.file_link.setVisible(False)
        self.file_link.setTextFormat(Qt.RichText)
        self.file_link.setOpenExternalLinks(False)
        self.file_link.setTextInteractionFlags(Qt.TextBrowserInteraction)
        self.file_link.linkActivated.connect(self.open_file_location)
        self.file_link.setStyleSheet("color: #007bff; text-decoration: underline; cursor: pointer;")
        progress_layout.addWidget(self.file_link)

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
        # Поддержка обратной совместимости, если signal fileSelected сработает
        self.main_files = [file_path]

    def handle_main_files_selected(self, files: List[str]):
        self.main_files = list(files)
        self._reset_config()
        if len(self.main_files) == 1:
            self.main_file_input.setText(os.path.basename(self.main_files[0]))
        else:
            short_names = [os.path.basename(f) for f in self.main_files]
            self.main_file_input.setText(
                tr("{count} файлов выбрано").format(count=len(short_names))
            )

    def handle_main_folder_selected(self, folder: str):
        excel_exts = ('.xlsx', '.xls')
        if os.path.isdir(folder):
            self.main_files = [
                os.path.join(folder, f)
                for f in os.listdir(folder)
                if f.lower().endswith(excel_exts)
            ]
            self._reset_config()
            self.main_file_input.setText(
                tr("Найдено файлов: {count}").format(count=len(self.main_files))
            )
        else:
            self.main_files = []

    def handle_files_selected(self, files: List[str]):
        self.source_files = files
        self._reset_config()

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
        self._reset_config()

    def open_preview(self):
        if not self.main_files:
            QMessageBox.critical(self, tr("Ошибка"), tr("Укажи основной(ые) Excel."))
            return
        if not self.source_files:
            QMessageBox.critical(self, tr("Ошибка"), tr("Укажи файлы источников."))
            return

        if len(self.main_files) == 1:
            dialog = MergeMappingDialog(self.main_files[0], self)
            for f in self.source_files:
                dialog.add_row_with_file(f)
            if dialog.exec():
                self.mappings = dialog.get_mappings()
                self.merge_btn.setEnabled(True)
                self.status_label.setText(tr("Настройки сохранены. Готово к объединению."))
                self.plans = [TargetMergePlan(main_file=self.main_files[0], mappings=self.mappings)]
                self.manual_mapping = self._build_manual_mapping(self.plans)
        else:
            assignments = suggest_target_mapping(self.source_files, self.main_files)
            plans: List[TargetMergePlan] = []
            for main_file in self.main_files:
                dialog = MergeMappingDialog(main_file, self)
                for f in assignments.get(main_file, self.source_files):
                    dialog.add_row_with_file(f)
                if not dialog.exec():
                    return
                mappings = dialog.get_mappings()
                if mappings:
                    plans.append(TargetMergePlan(main_file=main_file, mappings=mappings))
            if plans:
                self.plans = plans
                self.manual_mapping = self._build_manual_mapping(plans)
                self.merge_btn.setEnabled(True)
                self.status_label.setText(tr("Настройки сохранены для нескольких целей."))

    def run_merge(self):
        if not self.main_files:
            QMessageBox.critical(self, tr("Ошибка"), tr("Укажи основной(ые) Excel."))
            return
        if not self.source_files:
            QMessageBox.critical(self, tr("Ошибка"), tr("Укажи файлы источников."))
            return
        if len(self.main_files) == 1:
            if not self.mappings:
                QMessageBox.warning(self, tr("Предупреждение"),
                                    tr("Сначала настройте сопоставления через кнопку Настроить."))
                return
        else:
            if not self.plans:
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

        if len(self.main_files) == 1:
            self.worker = MergeWorker(self.main_files[0], self.mappings)
            self.worker.finished.connect(self.on_merge_finished)
            self.worker.error.connect(self.on_merge_error)
            self.worker.progress.connect(self.on_progress_update)
        else:
            self.worker = MultiTargetMergeWorker(self.plans, self.source_files, manual_mapping=self.manual_mapping)
            self.worker.finished.connect(self.on_multi_merge_finished)
            self.worker.error.connect(self.on_merge_error)
            self.worker.progress.connect(self.on_progress_update)
        self.worker.start()

    def on_progress_update(self, value, message):
        self.progress_bar.setValue(value)
        self.status_label.setText(message)

    def on_merge_finished(self, output_file):
        self.worker = None
        self.output_file = output_file
        self.progress_bar.setValue(100)
        self.status_label.setText(tr("Объединение завершено успешно!"))
        self.status_label.setStyleSheet("color: #28a745; font-weight: bold;")

        filename = os.path.basename(output_file)
        self.file_link.setText(f'<a href="file:///{output_file}">{filename}</a>')
        self.file_link.setVisible(True)

        self.is_processing = False
        self.merge_btn.setEnabled(True)
        self.configure_btn.setEnabled(True)

    def on_multi_merge_finished(self, output_files: List[str]):
        self.worker = None
        self.output_files = output_files
        self.progress_bar.setValue(100)
        self.status_label.setText(tr("Объединение завершено успешно!"))
        self.status_label.setStyleSheet("color: #28a745; font-weight: bold;")

        links = "<br>".join(
            f'<a href="file:///{path}">{os.path.basename(path)}</a>' for path in output_files
        )
        self.file_link.setText(links)
        self.file_link.setVisible(True)

        self.is_processing = False
        self.merge_btn.setEnabled(True)
        self.configure_btn.setEnabled(True)

    def on_merge_error(self, error_message):
        self.worker = None
        self.progress_bar.setVisible(False)
        self.file_link.setVisible(False)
        self.status_label.setText(tr("Ошибка при объединении"))
        self.status_label.setStyleSheet("color: #dc3545; font-weight: bold;")

        self.is_processing = False
        self.merge_btn.setEnabled(True)
        self.configure_btn.setEnabled(True)

        QMessageBox.critical(self, tr("Ошибка"), error_message)

    def open_file_location(self, link):
        target_path = link
        if not target_path and hasattr(self, 'output_file'):
            target_path = self.output_file

        if target_path:
            folder = os.path.dirname(target_path)
            if platform.system() == 'Windows':
                subprocess.run(['explorer', '/select,', os.path.normpath(target_path)])
            elif platform.system() == 'Darwin':
                subprocess.run(['open', '-R', target_path])
            else:
                subprocess.run(['xdg-open', folder])

    def _build_manual_mapping(self, plans: List[TargetMergePlan]) -> Dict[str, str]:
        mapping: Dict[str, str] = {}
        for plan in plans:
            for m in plan.mappings:
                src = m.get("source")
                if src:
                    mapping[src] = plan.main_file
        return mapping

    def _reset_config(self):
        self.mappings = []
        self.plans = []
        self.manual_mapping = {}
        self.merge_btn.setEnabled(False)
        if hasattr(self, "status_label"):
            self.status_label.setText(tr("Сначала настройте сопоставления"))
        self.file_link.setVisible(False)

    def retranslate_ui(self):
        self.main_label.setText(tr("Основной Excel файл(ы):"))
        self.main_file_input.setPlaceholderText(tr("Перетащи или кликни дважды (можно несколько)"))
        self.sources_label.setText(tr("Файлы источников:"))
        self.sources_input.setPlaceholderText(tr("Перетащи или кликни дважды"))
        self.configure_btn.setText(tr("Настроить"))
        self.merge_btn.setText(tr("Объединить"))

        if hasattr(self, 'status_label') and not self.mappings:
            self.status_label.setText(tr("Сначала настройте сопоставления"))

