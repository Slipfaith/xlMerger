import os
from typing import Dict, List

import pandas as pd
from PySide6.QtCore import Qt, QUrl
from PySide6.QtWidgets import (
    QWidget,
    QApplication,
    QVBoxLayout,
    QHBoxLayout,
    QLabel,
    QPushButton,
    QListWidget,
    QListWidgetItem,
    QGroupBox,
    QFileDialog,
    QMessageBox,
    QComboBox,
    QLineEdit,
    QSpinBox,
    QTableWidget,
    QTableWidgetItem,
    QCheckBox,
    QTextEdit,
    QProgressDialog,
    QScrollArea,
    QSizePolicy,
    QSlider,
    QToolButton,
)

from utils.i18n import tr
from utils.logger import logger
from core.drag_drop import DragDropLineEdit
from excel_builder import ExcelBuilderExecutor, ExcelFilesManager


class ExcelBuilderTab(QWidget):
    def __init__(self):
        super().__init__()
        self.manager = ExcelFilesManager()
        self.executor = ExcelBuilderExecutor(log_callback=self._log_line)
        self.operations: List[Dict] = []
        self._last_preview_df = pd.DataFrame()
        self.init_ui()

    def init_ui(self):
        layout = QVBoxLayout(self)
        layout.setContentsMargins(0, 0, 0, 0)

        scroll = QScrollArea()
        scroll.setWidgetResizable(True)
        scroll_widget = QWidget()
        scroll.setWidget(scroll_widget)
        layout.addWidget(scroll)

        content_layout = QVBoxLayout(scroll_widget)
        content_layout.setContentsMargins(20, 20, 20, 20)
        content_layout.setSpacing(15)

        content_layout.addWidget(self._create_loader_box())
        content_layout.addWidget(self._create_preview_box())
        content_layout.addWidget(self._create_operations_box())
        content_layout.addWidget(self._create_actions_box())
        content_layout.addStretch()

    # region UI builders
    def _create_loader_box(self):
        box = QGroupBox(tr("Загрузка входных данных"))
        vbox = QVBoxLayout(box)
        vbox.setSpacing(10)

        self.drop_input = DragDropLineEdit(mode="files_or_folder")
        self.drop_input.setPlaceholderText(tr("Перетащи файлы или папку сюда"))
        self.drop_input.filesSelected.connect(self.add_files)
        self.drop_input.folderSelected.connect(self.add_folder)
        vbox.addWidget(self.drop_input)

        btn_row = QHBoxLayout()
        add_files_btn = QPushButton(tr("Добавить файлы"))
        add_files_btn.clicked.connect(self.pick_files)
        add_folder_btn = QPushButton(tr("Добавить папку"))
        add_folder_btn.clicked.connect(self.pick_folder)
        btn_row.addWidget(add_files_btn)
        btn_row.addWidget(add_folder_btn)
        btn_row.addStretch()
        vbox.addLayout(btn_row)

        self.file_list = QListWidget()
        vbox.addWidget(QLabel(tr("Список загруженных файлов")))
        vbox.addWidget(self.file_list)
        self._update_list_rows(self.file_list, 0, max_rows=5)

        remove_row = QHBoxLayout()
        remove_btn = QPushButton(tr("Удалить выбранные"))
        remove_btn.clicked.connect(self.remove_selected_files)
        remove_row.addWidget(remove_btn)
        remove_row.addStretch()
        vbox.addLayout(remove_row)
        return box

    def _create_preview_box(self):
        box = QGroupBox(tr("Превью"))
        vbox = QVBoxLayout(box)
        vbox.setSpacing(8)

        selection_row = QHBoxLayout()
        selection_row.addWidget(QLabel(tr("Файл:")))
        self.preview_file_combo = QComboBox()
        self.preview_file_combo.currentIndexChanged.connect(self._on_preview_file_changed)
        selection_row.addWidget(self.preview_file_combo, 1)
        selection_row.addWidget(QLabel(tr("Лист:")))
        self.preview_sheet_combo = QComboBox()
        self.preview_sheet_combo.currentIndexChanged.connect(self.refresh_preview)
        selection_row.addWidget(self.preview_sheet_combo, 1)
        vbox.addLayout(selection_row)

        header_row = QHBoxLayout()
        header_row.addWidget(QLabel(tr("Строка заголовков")))
        self.header_spin = QSpinBox()
        self.header_spin.setMinimum(1)
        self.header_spin.setValue(1)
        self.header_spin.valueChanged.connect(self.refresh_preview)
        header_row.addWidget(self.header_spin)
        header_row.addStretch()
        vbox.addLayout(header_row)

        preview_controls = QHBoxLayout()
        preview_controls.addWidget(QLabel(tr("Высота превью")))
        self.preview_height_slider = QSlider(Qt.Horizontal)
        self.preview_height_slider.setRange(120, 800)
        self.preview_height_slider.setValue(320)
        self.preview_height_slider.setSingleStep(10)
        self.preview_height_slider.valueChanged.connect(self._update_preview_height)
        preview_controls.addWidget(self.preview_height_slider, 1)
        self.preview_height_label = QLabel("320 px")
        preview_controls.addWidget(self.preview_height_label)

        self.preview_toggle = QToolButton()
        self.preview_toggle.setCheckable(True)
        self.preview_toggle.setChecked(True)
        self.preview_toggle.setText(tr("Свернуть"))
        self.preview_toggle.toggled.connect(self._toggle_preview_visibility)
        preview_controls.addWidget(self.preview_toggle)
        vbox.addLayout(preview_controls)

        self.preview_table = QTableWidget()
        self.preview_table.setRowCount(0)
        self.preview_table.setColumnCount(0)
        self.preview_table.setMinimumHeight(320)
        self.preview_table.setSizePolicy(QSizePolicy.Expanding, QSizePolicy.Expanding)
        vbox.addWidget(self.preview_table)
        self._update_preview_height(self.preview_height_slider.value())
        return box

    def _create_operations_box(self):
        box = QGroupBox(tr("Операции"))
        vbox = QVBoxLayout(box)
        vbox.setSpacing(10)

        scope_row = QHBoxLayout()
        scope_row.addWidget(QLabel(tr("Применить к:")))
        self.scope_combo = QComboBox()
        self.scope_combo.addItem(tr("Все файлы"), userData="all")
        scope_row.addWidget(self.scope_combo, 1)
        vbox.addLayout(scope_row)

        # Header rename
        header_box = QGroupBox(tr("Редактирование заголовков"))
        header_layout = QHBoxLayout(header_box)
        self.header_identifier = QComboBox()
        self.header_identifier.setEditable(True)
        self.header_identifier.lineEdit().setPlaceholderText(tr("A или Текущий заголовок"))
        self.header_new_value = QLineEdit()
        self.header_new_value.setPlaceholderText(tr("Новый заголовок"))
        self.header_mode = QComboBox()
        self.header_mode.addItems([tr("По букве"), tr("По тексту")])
        self.header_mode.currentIndexChanged.connect(self._on_header_mode_changed)
        header_btn = QPushButton(tr("Добавить"))
        header_btn.clicked.connect(self.add_header_operation)
        header_layout.addWidget(self.header_identifier)
        header_layout.addWidget(self.header_new_value)
        header_layout.addWidget(self.header_mode)
        header_layout.addWidget(header_btn)
        vbox.addWidget(header_box)

        # Fill cell
        fill_box = QGroupBox(tr("Заполнение значения"))
        fill_layout = QVBoxLayout(fill_box)
        fill_description = QLabel(
            tr("Позволяет задать конкретную ячейку и новое значение. Можно заполнить только пустые ячейки."),
        )
        fill_description.setWordWrap(True)
        self.fill_cell = QLineEdit()
        self.fill_cell.setPlaceholderText("C1")
        self.fill_value = QLineEdit()
        self.fill_value.setPlaceholderText(tr("Значение"))
        self.fill_only_empty = QCheckBox(tr("Только пустые"))
        fill_btn = QPushButton(tr("Добавить"))
        fill_btn.clicked.connect(self.add_fill_operation)
        fill_row = QHBoxLayout()
        fill_row.addWidget(self.fill_cell)
        fill_row.addWidget(self.fill_value)
        fill_row.addWidget(self.fill_only_empty)
        fill_row.addWidget(fill_btn)
        fill_layout.addWidget(fill_description)
        fill_layout.addLayout(fill_row)
        vbox.addWidget(fill_box)

        # Rename sheets
        sheet_box = QGroupBox(tr("Переименование листов"))
        sheet_layout = QHBoxLayout(sheet_box)
        self.old_sheet = QComboBox()
        self.old_sheet.setEditable(True)
        self.old_sheet.lineEdit().setPlaceholderText(tr("Старое имя"))
        self.new_sheet = QLineEdit()
        self.new_sheet.setPlaceholderText(tr("Новое имя"))
        sheet_btn = QPushButton(tr("Добавить"))
        sheet_btn.clicked.connect(self.add_sheet_rename_operation)
        sheet_layout.addWidget(self.old_sheet)
        sheet_layout.addWidget(self.new_sheet)
        sheet_layout.addWidget(sheet_btn)
        vbox.addWidget(sheet_box)

        # Clear column
        clear_box = QGroupBox(tr("Очистка колонки"))
        clear_layout = QHBoxLayout(clear_box)
        self.clear_identifier = QComboBox()
        self.clear_identifier.setEditable(True)
        self.clear_identifier.lineEdit().setPlaceholderText(tr("Буква или заголовок"))
        self.clear_mode = QComboBox()
        self.clear_mode.addItems([tr("По букве"), tr("По тексту")])
        self.clear_mode.currentIndexChanged.connect(self._on_header_mode_changed)
        self.clear_format = QCheckBox(tr("Очистить формат"))
        clear_btn = QPushButton(tr("Добавить"))
        clear_btn.clicked.connect(self.add_clear_operation)
        clear_layout.addWidget(self.clear_identifier)
        clear_layout.addWidget(self.clear_mode)
        clear_layout.addWidget(self.clear_format)
        clear_layout.addWidget(clear_btn)
        vbox.addWidget(clear_box)

        vbox.addWidget(QLabel(tr("Запланированные операции")))
        self.operations_list = QListWidget()
        vbox.addWidget(self.operations_list)
        self._update_list_rows(self.operations_list, 0, max_rows=5)
        return box

    def _create_actions_box(self):
        box = QGroupBox(tr("Применение и сохранение"))
        vbox = QVBoxLayout(box)
        vbox.setSpacing(8)

        btn_row = QHBoxLayout()
        self.execute_btn = QPushButton(tr("Выполнить"))
        self.execute_btn.clicked.connect(self.execute)
        btn_row.addWidget(self.execute_btn)
        btn_row.addStretch()
        vbox.addLayout(btn_row)

        self.output_path_label = QLabel(tr("Папка сохранения: не выбрана"))
        self.output_path_label.setTextInteractionFlags(Qt.TextBrowserInteraction)
        self.output_path_label.setOpenExternalLinks(True)
        self.output_path_label.setWordWrap(True)
        vbox.addWidget(self.output_path_label)

        self.log_output = QTextEdit()
        self.log_output.setReadOnly(True)
        vbox.addWidget(self.log_output)
        self._limit_text_rows(self.log_output, 5)
        return box

    # endregion

    # region loaders
    def pick_files(self):
        files, _ = QFileDialog.getOpenFileNames(self, tr("Выбери файлы"), "", "Excel (*.xlsx *.xls)")
        if files:
            self.add_files(files)

    def pick_folder(self):
        folder = QFileDialog.getExistingDirectory(self, tr("Выбери папку"))
        if folder:
            self.add_folder(folder)

    def add_files(self, files: List[str]):
        self.manager.add_files(files)
        self._refresh_file_views()

    def add_folder(self, folder: str):
        self.manager.add_folder(folder)
        self._refresh_file_views()

    def _refresh_file_views(self):
        self.file_list.clear()
        self.preview_file_combo.blockSignals(True)
        self.preview_file_combo.clear()
        self.scope_combo.blockSignals(True)
        # keep "all" entry
        self.scope_combo.clear()
        self.scope_combo.addItem(tr("Все файлы"), userData="all")
        for f in self.manager.files:
            name = self._display_name(f["path"])
            item = QListWidgetItem(name)
            item.setToolTip(f["path"])
            self.file_list.addItem(item)

            self.preview_file_combo.addItem(name, userData=f)
            index = self.preview_file_combo.count() - 1
            self.preview_file_combo.setItemData(index, name, role=Qt.ToolTipRole)

            self.scope_combo.addItem(name, userData=f["path"])
            scope_index = self.scope_combo.count() - 1
            self.scope_combo.setItemData(scope_index, name, role=Qt.ToolTipRole)
        self.preview_file_combo.blockSignals(False)
        self.scope_combo.blockSignals(False)
        self._on_preview_file_changed()
        self._update_list_rows(self.file_list, self.file_list.count(), max_rows=5)

    def remove_selected_files(self):
        rows = sorted({index.row() for index in self.file_list.selectedIndexes()}, reverse=True)
        self.manager.remove_indices(rows)
        self._refresh_file_views()

    # endregion

    def _on_preview_file_changed(self):
        self.preview_sheet_combo.blockSignals(True)
        self.preview_sheet_combo.clear()
        file_info = self.preview_file_combo.currentData()
        if not file_info:
            self.preview_sheet_combo.blockSignals(False)
            self.preview_sheet_combo.clear()
            self._clear_preview_table()
            self._update_sheet_suggestions([])
            return
        sheets = self._read_sheets(file_info["path"], preview=True)
        self._update_sheet_suggestions(list(sheets.keys()))
        for name in sheets.keys():
            self.preview_sheet_combo.addItem(name)
        self.preview_sheet_combo.blockSignals(False)
        self.refresh_preview()

    def refresh_preview(self):
        file_info = self.preview_file_combo.currentData()
        sheet_name = self.preview_sheet_combo.currentText()
        if not file_info or not sheet_name:
            self._clear_preview_table()
            return
        sheets = self._read_sheets(file_info["path"], preview=True)
        if sheet_name not in sheets:
            self._clear_preview_table()
            return
        if sheet_name:
            self.old_sheet.setCurrentText(sheet_name)
        df = sheets[sheet_name].head(30)
        self._last_preview_df = df
        self._populate_table(df)
        self._update_header_suggestions(df)

    def _read_sheets(self, path: str, preview: bool = False) -> Dict[str, pd.DataFrame]:
        return self.executor.read_sheets(path, preview=preview)

    def _populate_table(self, df: pd.DataFrame):
        self.preview_table.clear()
        self.preview_table.setRowCount(len(df))
        self.preview_table.setColumnCount(len(df.columns))
        for i in range(len(df)):
            for j in range(len(df.columns)):
                value = df.iat[i, j]
                item = QTableWidgetItem("" if pd.isna(value) else str(value))
                self.preview_table.setItem(i, j, item)
        self.preview_table.resizeColumnsToContents()

    def _clear_preview_table(self):
        self.preview_table.clear()
        self.preview_table.setRowCount(0)
        self.preview_table.setColumnCount(0)
        self._last_preview_df = pd.DataFrame()
        self._update_header_suggestions(pd.DataFrame())

    # region operation adders
    def _current_scope(self):
        data = self.scope_combo.currentData()
        if data == "all":
            return "all"
        return data

    def _current_sheet(self):
        return self.preview_sheet_combo.currentText()

    def add_header_operation(self):
        identifier = self.header_identifier.currentText().strip()
        new_val = self.header_new_value.text().strip()
        if not identifier or not new_val:
            QMessageBox.warning(self, tr("Ошибка"), tr("Укажите колонку и новый заголовок"))
            return
        operation = {
            "type": "rename_header",
            "identifier": identifier,
            "mode": "letter" if self.header_mode.currentIndex() == 0 else "text",
            "new": new_val,
            "scope": self._current_scope(),
            "sheet": self._current_sheet(),
            "header_row": self.header_spin.value(),
        }
        self.operations.append(operation)
        self._append_operation_item(tr("Заголовок"), operation)
        self.header_identifier.setCurrentText("")
        self.header_new_value.clear()

    def add_fill_operation(self):
        cell = self.fill_cell.text().strip()
        value = self.fill_value.text()
        if not cell:
            QMessageBox.warning(self, tr("Ошибка"), tr("Укажите ячейку"))
            return
        operation = {
            "type": "fill_cell",
            "cell": cell,
            "value": value,
            "only_empty": self.fill_only_empty.isChecked(),
            "scope": self._current_scope(),
            "sheet": self._current_sheet(),
        }
        self.operations.append(operation)
        self._append_operation_item(tr("Заполнить"), operation)
        self.fill_cell.clear()
        self.fill_value.clear()
        self.fill_only_empty.setChecked(False)

    def add_sheet_rename_operation(self):
        old = self.old_sheet.currentText().strip()
        new = self.new_sheet.text().strip()
        if not old or not new:
            QMessageBox.warning(self, tr("Ошибка"), tr("Укажите старое и новое имя"))
            return
        operation = {
            "type": "rename_sheet",
            "old": old,
            "new": new,
            "scope": self._current_scope(),
        }
        self.operations.append(operation)
        self._append_operation_item(tr("Лист"), operation)
        self.old_sheet.setCurrentText("")
        self.new_sheet.clear()

    def add_clear_operation(self):
        identifier = self.clear_identifier.currentText().strip()
        if not identifier:
            QMessageBox.warning(self, tr("Ошибка"), tr("Укажите колонку"))
            return
        operation = {
            "type": "clear_column",
            "identifier": identifier,
            "mode": "letter" if self.clear_mode.currentIndex() == 0 else "text",
            "clear_format": self.clear_format.isChecked(),
            "scope": self._current_scope(),
            "sheet": self._current_sheet(),
            "header_row": self.header_spin.value(),
        }
        self.operations.append(operation)
        self._append_operation_item(tr("Очистка"), operation)
        self.clear_identifier.setCurrentText("")
        self.clear_format.setChecked(False)

    def _append_operation_item(self, prefix: str, operation: Dict):
        desc = f"{prefix}: {self._describe_operation(operation)}"
        self.operations_list.addItem(desc)
        self._update_list_rows(self.operations_list, self.operations_list.count(), max_rows=5)

    def _describe_operation(self, op: Dict) -> str:
        scope = tr("все") if op.get("scope") == "all" else self._display_name(op.get("scope", ""))
        if op["type"] == "rename_header":
            return f"[{scope}] {op['sheet']}: {op['identifier']} -> {op['new']}"
        if op["type"] == "fill_cell":
            suffix = tr(" только пустые") if op.get("only_empty") else ""
            return f"[{scope}] {op['sheet']}: {op['cell']} = {op['value']}{suffix}"
        if op["type"] == "rename_sheet":
            return f"[{scope}] {op['old']} -> {op['new']}"
        if op["type"] == "clear_column":
            return f"[{scope}] {op['sheet']}: {op['identifier']}"
        return str(op)

    # endregion

    # region execution
    def execute(self):
        if not self.manager.files:
            QMessageBox.warning(self, tr("Ошибка"), tr("Добавьте файлы"))
            return
        target_ops = self.operations
        if not target_ops:
            reply = QMessageBox.question(
                self,
                tr("Операции"),
                tr("Операции не заданы. Просто сохранить копии?"),
            )
            if reply != QMessageBox.Yes:
                return
        progress = QProgressDialog(tr("Обработка..."), tr("Отмена"), 0, len(self.manager.files), self)
        progress.setWindowModality(Qt.ApplicationModal)
        progress.show()
        self.log_output.clear()

        output_root = self.manager.build_output_root()
        os.makedirs(output_root, exist_ok=True)
        self._update_output_path_link(output_root)
        logger.info(f"Output root: {output_root}")
        for idx, f in enumerate(self.manager.files, start=1):
            progress.setValue(idx - 1)
            progress.setLabelText(self._display_name(f["path"]))
            QApplication.processEvents()
            try:
                self.executor.process_file(f, output_root, target_ops)
                self._log_line(f"✓ {f['path']}")
            except Exception as exc:  # noqa: BLE001
                self._log_line(f"✗ {f['path']}: {exc}")
            if progress.wasCanceled():
                break
        progress.setValue(len(self.manager.files))
        QMessageBox.information(self, tr("Готово"), tr("Обработка завершена"))

    def _display_name(self, path: str) -> str:
        return os.path.basename(path)

    def _log_line(self, text: str):
        logger.info(text)
        self.log_output.append(text)

    def _update_sheet_suggestions(self, sheets: List[str]):
        self.old_sheet.blockSignals(True)
        self.old_sheet.clear()
        for name in sheets:
            self.old_sheet.addItem(name)
        self.old_sheet.blockSignals(False)
        if sheets:
            preferred = self.preview_sheet_combo.currentText()
            self.old_sheet.setCurrentText(preferred if preferred else sheets[0])

    def _update_header_suggestions(self, df: pd.DataFrame):
        header_row = self.header_spin.value() - 1
        headers: List[str] = []
        if header_row < len(df):
            headers = [
                str(val)
                for val in df.iloc[header_row].tolist()
                if pd.notna(val) and str(val)
            ]
        letters = [self._column_letter(idx) for idx in range(len(df.columns))]

        self._set_identifier_options(self.header_identifier, self.header_mode.currentIndex(), letters, headers)
        self._set_identifier_options(self.clear_identifier, self.clear_mode.currentIndex(), letters, headers)

    def _set_identifier_options(self, combo: QComboBox, mode: int, letters: List[str], headers: List[str]):
        options = letters if mode == 0 else headers
        combo.blockSignals(True)
        text = combo.currentText()
        combo.clear()
        for item in options:
            combo.addItem(item)
        if text and text in options:
            combo.setCurrentText(text)
        elif options:
            combo.setCurrentIndex(0)
        else:
            combo.setCurrentText("")
        combo.blockSignals(False)

    def _on_header_mode_changed(self):
        self._update_header_suggestions(self._last_preview_df)

    def _update_preview_height(self, value: int):
        self.preview_height_label.setText(f"{value} px")
        if self.preview_toggle.isChecked():
            self.preview_table.setFixedHeight(value)
        self.preview_table.setMinimumHeight(value)

    def _toggle_preview_visibility(self, checked: bool):
        self.preview_table.setVisible(checked)
        self.preview_height_slider.setEnabled(checked)
        self.preview_height_label.setEnabled(checked)
        self.preview_toggle.setText(tr("Свернуть") if checked else tr("Показать"))
        self.preview_height_label.setText(f"{self.preview_height_slider.value()} px" if checked else tr("Скрыто"))
        if checked:
            self._update_preview_height(self.preview_height_slider.value())

    def _column_letter(self, idx: int) -> str:
        idx += 1
        letters = ""
        while idx > 0:
            idx, remainder = divmod(idx - 1, 26)
            letters = chr(65 + remainder) + letters
        return letters

    def _update_list_rows(self, list_widget: QListWidget, items_count: int, max_rows: int, min_rows: int = 1):
        row_height = list_widget.sizeHintForRow(0)
        if row_height <= 0:
            row_height = list_widget.fontMetrics().height() + 8
        visible_rows = min(max(items_count, min_rows), max_rows)
        height = visible_rows * row_height + 2 * list_widget.frameWidth()
        list_widget.setFixedHeight(height)
        list_widget.setVerticalScrollBarPolicy(Qt.ScrollBarAsNeeded)

    def _limit_text_rows(self, text_widget: QTextEdit, rows: int):
        line_height = text_widget.fontMetrics().lineSpacing()
        margins = int(text_widget.document().documentMargin() * 2)
        height = rows * line_height + 2 * text_widget.frameWidth() + margins
        text_widget.setFixedHeight(height)
        text_widget.setVerticalScrollBarPolicy(Qt.ScrollBarAsNeeded)

    def _update_output_path_link(self, path: str):
        url = QUrl.fromLocalFile(path)
        link = f'<a href="{url.toString()}">{path}</a>'
        self.output_path_label.setText(f"{tr('Папка сохранения')}: {link}")

    # endregion
