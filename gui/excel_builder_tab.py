import os
import string
from typing import Dict, List

import pandas as pd
from PySide6.QtCore import Qt
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
)

from utils.i18n import tr
from utils.logger import logger
from core.drag_drop import DragDropLineEdit


class ExcelBuilderTab(QWidget):
    def __init__(self):
        super().__init__()
        self.files: List[Dict[str, str]] = []
        self.operations: List[Dict] = []
        self.base_folder: str | None = None
        self.init_ui()

    def init_ui(self):
        layout = QVBoxLayout(self)
        layout.setContentsMargins(20, 20, 20, 20)
        layout.setSpacing(15)

        layout.addWidget(self._create_loader_box())
        layout.addWidget(self._create_preview_box())
        layout.addWidget(self._create_operations_box())
        layout.addWidget(self._create_actions_box())

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

        self.preview_table = QTableWidget()
        self.preview_table.setRowCount(0)
        self.preview_table.setColumnCount(0)
        vbox.addWidget(self.preview_table)
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
        self.header_identifier = QLineEdit()
        self.header_identifier.setPlaceholderText(tr("A или Текущий заголовок"))
        self.header_new_value = QLineEdit()
        self.header_new_value.setPlaceholderText(tr("Новый заголовок"))
        self.header_mode = QComboBox()
        self.header_mode.addItems([tr("По букве"), tr("По тексту")])
        header_btn = QPushButton(tr("Добавить"))
        header_btn.clicked.connect(self.add_header_operation)
        header_layout.addWidget(self.header_identifier)
        header_layout.addWidget(self.header_new_value)
        header_layout.addWidget(self.header_mode)
        header_layout.addWidget(header_btn)
        vbox.addWidget(header_box)

        # Fill cell
        fill_box = QGroupBox(tr("Заполнение значения"))
        fill_layout = QHBoxLayout(fill_box)
        self.fill_cell = QLineEdit()
        self.fill_cell.setPlaceholderText("C1")
        self.fill_value = QLineEdit()
        self.fill_value.setPlaceholderText(tr("Значение"))
        self.fill_only_empty = QCheckBox(tr("Только пустые"))
        fill_btn = QPushButton(tr("Добавить"))
        fill_btn.clicked.connect(self.add_fill_operation)
        fill_layout.addWidget(self.fill_cell)
        fill_layout.addWidget(self.fill_value)
        fill_layout.addWidget(self.fill_only_empty)
        fill_layout.addWidget(fill_btn)
        vbox.addWidget(fill_box)

        # Rename sheets
        sheet_box = QGroupBox(tr("Переименование листов"))
        sheet_layout = QHBoxLayout(sheet_box)
        self.old_sheet = QLineEdit()
        self.old_sheet.setPlaceholderText(tr("Старое имя"))
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
        self.clear_identifier = QLineEdit()
        self.clear_identifier.setPlaceholderText(tr("Буква или заголовок"))
        self.clear_mode = QComboBox()
        self.clear_mode.addItems([tr("По букве"), tr("По тексту")])
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

        self.log_output = QTextEdit()
        self.log_output.setReadOnly(True)
        vbox.addWidget(self.log_output)
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
        if not files:
            return
        if self.base_folder is None:
            self.base_folder = os.path.dirname(files[0])
        for path in files:
            if not os.path.isfile(path) or not path.lower().endswith((".xlsx", ".xls")):
                continue
            if any(f["path"] == path for f in self.files):
                continue
            rel = self._relative_path(path)
            self.files.append({"path": path, "rel": rel})
        self._refresh_file_views()

    def add_folder(self, folder: str):
        if not os.path.isdir(folder):
            return
        if self.base_folder is None:
            self.base_folder = folder
        for root, _, filenames in os.walk(folder):
            for name in filenames:
                if not name.lower().endswith((".xlsx", ".xls")):
                    continue
                path = os.path.join(root, name)
                if any(f["path"] == path for f in self.files):
                    continue
                rel = os.path.relpath(path, folder)
                self.files.append({"path": path, "rel": rel})
        self._refresh_file_views()

    def _relative_path(self, path: str) -> str:
        if self.base_folder and os.path.commonpath([self.base_folder, path]) == self.base_folder:
            return os.path.relpath(path, self.base_folder)
        return os.path.basename(path)

    def _refresh_file_views(self):
        self.file_list.clear()
        self.preview_file_combo.blockSignals(True)
        self.preview_file_combo.clear()
        self.scope_combo.blockSignals(True)
        # keep "all" entry
        self.scope_combo.clear()
        self.scope_combo.addItem(tr("Все файлы"), userData="all")
        for f in self.files:
            item = QListWidgetItem(self._short_name(os.path.basename(f["path"])))
            item.setToolTip(f["path"])
            self.file_list.addItem(item)

            self.preview_file_combo.addItem(self._short_name(os.path.basename(f["path"])), userData=f)
            self.scope_combo.addItem(self._short_name(os.path.basename(f["path"])), userData=f["path"])
        self.preview_file_combo.blockSignals(False)
        self.scope_combo.blockSignals(False)
        self._on_preview_file_changed()

    def remove_selected_files(self):
        rows = sorted({index.row() for index in self.file_list.selectedIndexes()}, reverse=True)
        for row in rows:
            if 0 <= row < len(self.files):
                del self.files[row]
        if not self.files:
            self.base_folder = None
        self._refresh_file_views()

    # endregion

    def _on_preview_file_changed(self):
        self.preview_sheet_combo.blockSignals(True)
        self.preview_sheet_combo.clear()
        file_info = self.preview_file_combo.currentData()
        if not file_info:
            self.preview_sheet_combo.blockSignals(False)
            self.preview_sheet_combo.clear()
            self.preview_table.setRowCount(0)
            self.preview_table.setColumnCount(0)
            return
        sheets = self._read_sheets(file_info["path"], preview=True)
        for name in sheets.keys():
            self.preview_sheet_combo.addItem(name)
        self.preview_sheet_combo.blockSignals(False)
        self.refresh_preview()

    def refresh_preview(self):
        file_info = self.preview_file_combo.currentData()
        sheet_name = self.preview_sheet_combo.currentText()
        if not file_info or not sheet_name:
            return
        sheets = self._read_sheets(file_info["path"], preview=True)
        if sheet_name not in sheets:
            return
        df = sheets[sheet_name].head(30)
        self._populate_table(df)

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

    # region operation adders
    def _current_scope(self):
        data = self.scope_combo.currentData()
        if data == "all":
            return "all"
        return data

    def _current_sheet(self):
        return self.preview_sheet_combo.currentText()

    def add_header_operation(self):
        identifier = self.header_identifier.text().strip()
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
        self.header_identifier.clear()
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
        old = self.old_sheet.text().strip()
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
        self.old_sheet.clear()
        self.new_sheet.clear()

    def add_clear_operation(self):
        identifier = self.clear_identifier.text().strip()
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
        self.clear_identifier.clear()
        self.clear_format.setChecked(False)

    def _append_operation_item(self, prefix: str, operation: Dict):
        desc = f"{prefix}: {self._describe_operation(operation)}"
        self.operations_list.addItem(desc)

    def _describe_operation(self, op: Dict) -> str:
        scope = tr("все") if op.get("scope") == "all" else self._short_name(os.path.basename(op.get("scope", "")))
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
        if not self.files:
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
        progress = QProgressDialog(tr("Обработка..."), tr("Отмена"), 0, len(self.files), self)
        progress.setWindowModality(Qt.ApplicationModal)
        progress.show()
        self.log_output.clear()

        output_root = self._output_root()
        os.makedirs(output_root, exist_ok=True)
        logger.info(f"Output root: {output_root}")
        for idx, f in enumerate(self.files, start=1):
            progress.setValue(idx - 1)
            progress.setLabelText(os.path.basename(f["path"]))
            QApplication.processEvents()
            try:
                self._process_file(f, output_root, target_ops)
                self._log_line(f"✓ {f['path']}")
            except Exception as exc:  # noqa: BLE001
                self._log_line(f"✗ {f['path']}: {exc}")
            if progress.wasCanceled():
                break
        progress.setValue(len(self.files))
        QMessageBox.information(self, tr("Готово"), tr("Обработка завершена"))

    def _output_root(self) -> str:
        if self.base_folder:
            parent = os.path.dirname(self.base_folder)
            name = os.path.basename(self.base_folder)
            return os.path.join(parent, f"{name}_upd")
        first_file_dir = os.path.dirname(self.files[0]["path"])
        first_name = os.path.splitext(os.path.basename(self.files[0]["path"]))[0]
        return os.path.join(first_file_dir, f"{first_name}_upd")

    def _process_file(self, file_info: Dict[str, str], output_root: str, operations: List[Dict]):
        sheets = self._read_sheets(file_info["path"])
        sheets = self._apply_operations(sheets, operations, file_info["path"])
        rel_path = file_info["rel"]
        dest_path = os.path.join(output_root, rel_path)
        os.makedirs(os.path.dirname(dest_path), exist_ok=True)
        if dest_path.lower().endswith(".xls"):
            dest_path = dest_path + "x"  # normalize to xlsx for writer
        with pd.ExcelWriter(dest_path, engine="xlsxwriter") as writer:
            for name, df in sheets.items():
                df.to_excel(writer, sheet_name=name, index=False, header=False)

    def _apply_operations(self, sheets: Dict[str, pd.DataFrame], operations: List[Dict], file_path: str):
        # Apply sheet renames first so other operations use updated names
        rename_ops = [op for op in operations if op["type"] == "rename_sheet" and self._op_matches(op, file_path)]
        if rename_ops:
            sheets = self._rename_sheets(sheets, rename_ops)

        for op in operations:
            if op["type"] == "rename_sheet":
                continue
            if not self._op_matches(op, file_path):
                continue
            sheet_name = op.get("sheet")
            if not sheet_name or sheet_name not in sheets:
                self._log_line(f"Пропуск операции {op['type']} для файла {file_path}: лист '{sheet_name}' не найден")
                continue
            df = sheets[sheet_name]
            if op["type"] == "rename_header":
                df = self._rename_header(df, op)
            elif op["type"] == "fill_cell":
                df = self._fill_cell(df, op)
            elif op["type"] == "clear_column":
                df = self._clear_column(df, op)
            sheets[sheet_name] = df
        return sheets

    def _rename_sheets(self, sheets: Dict[str, pd.DataFrame], operations: List[Dict]):
        result: Dict[str, pd.DataFrame] = {}
        for name, df in sheets.items():
            new_name = name
            for op in operations:
                if op["old"] == name:
                    candidate = op["new"]
                    counter = 2
                    while candidate in result or candidate in sheets:
                        candidate = f"{op['new']}_{counter}"
                        counter += 1
                    new_name = candidate
                    break
            result[new_name] = df
        return result

    def _rename_header(self, df: pd.DataFrame, op: Dict):
        header_row = op.get("header_row", 1) - 1
        col_idx = self._find_column(df, op["identifier"], op["mode"], header_row)
        if col_idx is None:
            self._log_line(f"Колонка {op['identifier']} не найдена")
            return df
        self._ensure_size(df, header_row, col_idx)
        df.iat[header_row, col_idx] = op["new"]
        return df

    def _fill_cell(self, df: pd.DataFrame, op: Dict):
        col_letter = ''.join([c for c in op["cell"] if c.isalpha()])
        row_part = ''.join([c for c in op["cell"] if c.isdigit()])
        if not col_letter or not row_part:
            self._log_line(f"Неверная ячейка: {op['cell']}")
            return df
        row_idx = int(row_part) - 1
        col_idx = self._column_from_letter(col_letter)
        self._ensure_size(df, row_idx, col_idx)
        if op.get("only_empty") and pd.notna(df.iat[row_idx, col_idx]):
            return df
        df.iat[row_idx, col_idx] = op.get("value")
        return df

    def _clear_column(self, df: pd.DataFrame, op: Dict):
        header_row = op.get("header_row", 1) - 1
        col_idx = self._find_column(df, op["identifier"], op["mode"], header_row)
        if col_idx is None:
            self._log_line(f"Колонка {op['identifier']} не найдена")
            return df
        self._ensure_size(df, header_row + 1, col_idx)
        df.iloc[header_row + 1 :, col_idx] = ""
        return df

    def _ensure_size(self, df: pd.DataFrame, row_idx: int, col_idx: int):
        # expand rows
        while row_idx >= len(df):
            df.loc[len(df)] = [None] * len(df.columns)
        while col_idx >= len(df.columns):
            df[len(df.columns)] = None

    def _find_column(self, df: pd.DataFrame, identifier: str, mode: str, header_row: int):
        if mode == "letter":
            return self._column_from_letter(identifier)
        if header_row < len(df):
            headers = df.iloc[header_row].tolist()
            for idx, val in enumerate(headers):
                if str(val) == identifier:
                    return idx
        return None

    def _column_from_letter(self, letter: str) -> int:
        letter = letter.upper()
        idx = 0
        for char in letter:
            if char in string.ascii_uppercase:
                idx = idx * 26 + (ord(char) - ord('A') + 1)
        return idx - 1

    def _read_sheets(self, path: str, preview: bool = False) -> Dict[str, pd.DataFrame]:
        try:
            if preview:
                sheets = pd.read_excel(path, sheet_name=None, header=None, nrows=30)
            else:
                sheets = pd.read_excel(path, sheet_name=None, header=None)
            return {name: df.fillna("") for name, df in sheets.items()}
        except Exception as exc:  # noqa: BLE001
            self._log_line(f"Не удалось прочитать {path}: {exc}")
            return {}

    def _op_matches(self, op: Dict, file_path: str) -> bool:
        if op.get("scope") == "all":
            return True
        return op.get("scope") == file_path

    def _short_name(self, path: str, n: int = 5) -> str:
        return path if len(path) <= 2 * n else f"{path[:n]}...{path[-n:]}"

    def _log_line(self, text: str):
        logger.info(text)
        self.log_output.append(text)

    # endregion
