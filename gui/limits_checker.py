# limits_checker.py

import os, sys
from PyQt5 import QtWidgets, QtCore
from PyQt5.QtWidgets import (
    QWidget, QVBoxLayout, QHBoxLayout, QGroupBox, QLabel, QLineEdit,
    QPushButton, QFileDialog, QComboBox, QMessageBox, QStackedWidget, QTextEdit,
    QSplitter, QListWidget, QMenu
)
from PyQt5.QtCore import Qt, pyqtSignal
from PyQt5.QtGui import QStandardItemModel, QStandardItem, QColor, QBrush
from openpyxl import load_workbook
from openpyxl.styles import PatternFill


# Универсальная функция для приведения значения к целому числу
def _get_int_value(value):
    try:
        s = str(value).strip()
        return int(s) if s != "" else None
    except Exception:
        return None


# ==================== DragDropLineEdit ====================
class DragDropLineEdit(QLineEdit):
    def __init__(self, update_callback, parent=None):
        super().__init__(parent)
        self.setAcceptDrops(True)
        self.update_callback = update_callback

    def dragEnterEvent(self, event):
        if event.mimeData().hasUrls():
            event.acceptProposedAction()

    def dropEvent(self, event):
        urls = event.mimeData().urls()
        if urls and urls[0].isLocalFile():
            file_path = urls[0].toLocalFile()
            self.setText(file_path)
            self.update_callback(file_path)


# ==================== DraggableHeaderView ====================
class DraggableHeaderView(QtWidgets.QHeaderView):
    dragSelectionChanged = pyqtSignal(set)

    def __init__(self, orientation, parent=None):
        super().__init__(orientation, parent)
        self.setSectionsClickable(True)
        self._dragging = False
        self._drag_start = None
        self._drag_current = None

    def mousePressEvent(self, event):
        index = self.logicalIndexAt(event.pos())
        if index >= 0:
            self._drag_start = index
            self._drag_current = index
            self._dragging = True
            self.dragSelectionChanged.emit({index})
        super().mousePressEvent(event)

    def mouseMoveEvent(self, event):
        if self._dragging:
            index = self.logicalIndexAt(event.pos())
            if index >= 0 and index != self._drag_current:
                self._drag_current = index
                start = min(self._drag_start, self._drag_current)
                end = max(self._drag_start, self._drag_current)
                selection = set(range(start, end + 1))
                self.dragSelectionChanged.emit(selection)
        super().mouseMoveEvent(event)

    def mouseReleaseEvent(self, event):
        self._dragging = False
        super().mouseReleaseEvent(event)


# ==================== LimitsMappingPreviewDialog ====================
class LimitsMappingPreviewDialog(QtWidgets.QDialog):
    def __init__(self, model: QStandardItemModel, headers: list, parent=None):
        """
        :param model: QStandardItemModel для превью (например, первые N строк Excel)
        :param headers: Список заголовков столбцов
        """
        super().__init__(parent)
        self.setWindowTitle("Интерактивное сопоставление лимитов")
        self.resize(800, 600)
        self.headers = headers
        self.model = model
        self.manual = False  # Ручной режим выключен по умолчанию
        self.current_limit = None
        self.current_texts = set()
        # Для режима столбцов: (limit_header, [text_headers], False, None, None, "column")
        # Для ручного режима: (selected_cells, True, upper, lower, "cell")
        self.mappings = []
        self.init_ui()

    def init_ui(self):
        main_layout = QVBoxLayout(self)
        self.manual_checkbox = QtWidgets.QCheckBox("Вручную задавать лимиты")
        self.manual_checkbox.stateChanged.connect(self.toggle_manual_mode)
        main_layout.addWidget(self.manual_checkbox)

        splitter = QtWidgets.QSplitter(Qt.Vertical)
        top_widget = QWidget()
        top_layout = QVBoxLayout(top_widget)
        instruction = QLabel(
            "Если ручной режим не активен, кликните по заголовку для выбора столбца с лимитом (синим),\n"
            "а затем, перетаскивая, выберите текстовые столбцы (зелёным).\n\n"
            "Если ручной режим активен, выделяйте нужный диапазон ячеек – стандартное выделение будет видно."
        )
        top_layout.addWidget(instruction)
        self.table_view = QtWidgets.QTableView()
        self.table_view.setModel(self.model)
        self.table_view.setStyleSheet("QTableView::item { color: black; }")
        self.header_view = DraggableHeaderView(Qt.Horizontal, self.table_view)
        self.table_view.setHorizontalHeader(self.header_view)
        self.header_view.dragSelectionChanged.connect(self.handle_drag_selection)
        self.table_view.setEditTriggers(QtWidgets.QAbstractItemView.NoEditTriggers)
        self.table_view.selectionModel().selectionChanged.connect(self.update_header_colors)
        top_layout.addWidget(self.table_view)
        self.current_mapping_label = QLabel("Текущая настройка: Не выбрано")
        top_layout.addWidget(self.current_mapping_label)

        manual_group = QtWidgets.QGroupBox("Ручная настройка лимитов")
        manual_layout = QHBoxLayout()
        self.upper_limit_edit = QLineEdit()
        self.upper_limit_edit.setPlaceholderText("Верхний лимит")
        self.upper_limit_edit.setFixedWidth(100)
        self.upper_limit_edit.setEnabled(False)
        manual_layout.addWidget(self.upper_limit_edit)
        self.lower_limit_edit = QLineEdit()
        self.lower_limit_edit.setPlaceholderText("Нижний лимит")
        self.lower_limit_edit.setFixedWidth(100)
        self.lower_limit_edit.setEnabled(False)
        manual_layout.addWidget(self.lower_limit_edit)
        manual_group.setLayout(manual_layout)
        top_layout.addWidget(manual_group)

        btn_layout = QtWidgets.QHBoxLayout()
        self.save_mapping_btn = QtWidgets.QPushButton("Подтвердить")
        self.save_mapping_btn.clicked.connect(self.save_current_mapping)
        self.clear_selection_btn = QtWidgets.QPushButton("Очистить выбор")
        self.clear_selection_btn.clicked.connect(self.clear_selection)
        btn_layout.addWidget(self.save_mapping_btn)
        btn_layout.addWidget(self.clear_selection_btn)
        top_layout.addLayout(btn_layout)
        splitter.addWidget(top_widget)

        bottom_widget = QWidget()
        bottom_layout = QtWidgets.QVBoxLayout(bottom_widget)
        bottom_layout.addWidget(QtWidgets.QLabel("Сохранённые сопоставления:"))
        self.mapping_list = QListWidget()
        self.mapping_list.setContextMenuPolicy(Qt.CustomContextMenu)
        self.mapping_list.customContextMenuRequested.connect(self.show_mapping_context_menu)
        bottom_layout.addWidget(self.mapping_list)
        splitter.addWidget(bottom_widget)
        splitter.setStretchFactor(0, 2)
        splitter.setStretchFactor(1, 1)
        bottom_widget.setMaximumHeight(int(self.height() * 0.35))

        main_layout.addWidget(splitter)

        bottom_btn_layout = QtWidgets.QHBoxLayout()
        self.done_btn = QtWidgets.QPushButton("Готово")
        self.done_btn.clicked.connect(self.accept)
        self.cancel_btn = QtWidgets.QPushButton("Отмена")
        self.cancel_btn.clicked.connect(self.reject)
        bottom_btn_layout.addWidget(self.done_btn)
        bottom_btn_layout.addWidget(self.cancel_btn)
        main_layout.addLayout(bottom_btn_layout)
        self.setLayout(main_layout)

    def accept(self):
        duplicates = self.check_duplicates()
        if duplicates:
            QMessageBox.critical(self, "Ошибка дублирования", "\n".join(duplicates))
            return
        super().accept()

    def check_duplicates(self):
        errors = []
        for i, mapping1 in enumerate(self.mappings):
            if mapping1[-1] == "column":
                limit1 = mapping1[0]
                texts1 = set(mapping1[1])
                for j, mapping2 in enumerate(self.mappings):
                    if i >= j:
                        continue
                    if mapping2[-1] == "column":
                        limit2 = mapping2[0]
                        texts2 = set(mapping2[1])
                        if limit1 == limit2:
                            errors.append(
                                f"Дублирование: лимитный столбец '{limit1}' используется в сопоставлениях {i + 1} и {j + 1}.")
                        common = texts1.intersection(texts2)
                        if common:
                            errors.append(
                                f"Дублирование: текстовые столбцы {', '.join(common)} используются в сопоставлениях {i + 1} и {j + 1}.")
                    else:
                        for (r, col) in mapping2[0]:
                            header = self.headers[col]
                            if header == limit1 or header in texts1:
                                errors.append(
                                    f"Дублирование: ячейка ({r},{header}) присутствует в сопоставлении {j + 1} (ячейки) и {i + 1} (столбцы).")
            else:
                cells1 = set(mapping1[0])
                for j, mapping2 in enumerate(self.mappings):
                    if i >= j:
                        continue
                    if mapping2[-1] == "cell":
                        cells2 = set(mapping2[0])
                        common = cells1.intersection(cells2)
                        if common:
                            errors.append(
                                f"Дублирование: ячейки {common} присутствуют в сопоставлениях {i + 1} и {j + 1}.")
                    else:
                        for (r, col) in cells1:
                            header = self.headers[col]
                            if header == mapping2[0] or header in mapping2[1]:
                                errors.append(
                                    f"Дублирование: ячейка ({r},{header}) присутствует в сопоставлении {i + 1} (ячейки) и {j + 1} (столбцы).")
        return errors

    def toggle_manual_mode(self, state):
        self.manual = (state == Qt.Checked)
        if self.manual:
            self.table_view.setSelectionMode(QtWidgets.QAbstractItemView.ExtendedSelection)
            self.upper_limit_edit.setEnabled(True)
            self.lower_limit_edit.setEnabled(True)
        else:
            self.table_view.setSelectionMode(QtWidgets.QAbstractItemView.NoEditSelection)
            self.upper_limit_edit.setEnabled(False)
            self.lower_limit_edit.setEnabled(False)
            self.clear_selection()
        self.update_current_mapping_label()

    def handle_drag_selection(self, selection: set):
        if self.manual:
            return
        if self.current_limit is None:
            if len(selection) == 1:
                self.current_limit = next(iter(selection))
        else:
            sel = set(selection)
            if self.current_limit in sel:
                sel.remove(self.current_limit)
            self.current_texts = sel
        self.update_header_colors()
        self.update_current_mapping_label()

    def update_header_colors(self):
        for row in range(self.model.rowCount()):
            for col in range(self.model.columnCount()):
                index = self.model.index(row, col)
                self.model.setData(index, QBrush(QColor("white")), role=Qt.BackgroundRole)
        for mapping in self.mappings:
            if mapping[-1] == "column":
                try:
                    idx_limit = self.headers.index(mapping[0])
                except ValueError:
                    continue
                for row in range(self.model.rowCount()):
                    index = self.model.index(row, idx_limit)
                    self.model.setData(index, QBrush(QColor("lightblue")), role=Qt.BackgroundRole)
                for text in mapping[1]:
                    try:
                        idx_text = self.headers.index(text)
                    except ValueError:
                        continue
                    for row in range(self.model.rowCount()):
                        index = self.model.index(row, idx_text)
                        self.model.setData(index, QBrush(QColor("lightgreen")), role=Qt.BackgroundRole)
            else:
                for (sheet_row, col) in mapping[0]:
                    model_row = sheet_row - 2
                    if 0 <= model_row < self.model.rowCount() and 0 <= col < self.model.columnCount():
                        index = self.model.index(model_row, col)
                        self.model.setData(index, QBrush(QColor("lightgreen")), role=Qt.BackgroundRole)
        if not self.manual:
            if self.current_limit is not None:
                for row in range(self.model.rowCount()):
                    index = self.model.index(row, self.current_limit)
                    current_brush = self.model.data(index, role=Qt.BackgroundRole)
                    if not current_brush or current_brush.color().name() == "#ffffff":
                        self.model.setData(index, QBrush(QColor("lightblue")), role=Qt.BackgroundRole)
            for col in self.current_texts:
                for row in range(self.model.rowCount()):
                    index = self.model.index(row, col)
                    current_brush = self.model.data(index, role=Qt.BackgroundRole)
                    if not current_brush or current_brush.color().name() == "#ffffff":
                        self.model.setData(index, QBrush(QColor("lightgreen")), role=Qt.BackgroundRole)
        else:
            pass

    def update_current_mapping_label(self):
        if self.manual:
            indexes = self.table_view.selectionModel().selectedIndexes()
            if indexes:
                cells = [f"({index.row() + 2},{self.headers[index.column()]})" for index in
                         sorted(indexes, key=lambda x: (x.row(), x.column()))]
                mapping_str = f"Выбрано ячеек: {', '.join(cells)}"
            else:
                mapping_str = "Не выбрано"
        else:
            if self.current_limit is not None:
                limit_text = self.headers[self.current_limit]
            else:
                limit_text = "Не выбрано"
            text_list = [self.headers[i] for i in sorted(self.current_texts)]
            mapping_str = f"Лимит: {limit_text}; Тексты: {', '.join(text_list) if text_list else 'Не выбрано'}"
        if self.manual:
            upper = self.upper_limit_edit.text()
            lower = self.lower_limit_edit.text()
            mapping_str += f"; Верхний: {upper if upper else '—'}; Нижний: {lower if lower else '—'}"
        self.current_mapping_label.setText("Текущая настройка: " + mapping_str)

    def clear_selection(self):
        self.table_view.clearSelection()
        if not self.manual:
            self.current_limit = None
            self.current_texts.clear()
        if not self.manual:
            self.manual_checkbox.setChecked(False)
        self.upper_limit_edit.clear()
        self.lower_limit_edit.clear()
        self.update_header_colors()
        self.update_current_mapping_label()

    def save_current_mapping(self):
        manual = self.manual
        if not manual:
            if self.current_limit is None or not self.current_texts:
                QMessageBox.critical(self, "Ошибка", "Выберите столбец с лимитом и хотя бы один столбец с текстом.")
                return
            mapping = (
            self.headers[self.current_limit], [self.headers[i] for i in sorted(self.current_texts)], False, None, None,
            "column")
        else:
            indexes = self.table_view.selectionModel().selectedIndexes()
            if not indexes:
                QMessageBox.critical(self, "Ошибка", "Выберите диапазон ячеек для проверки.")
                return
            # При сохранении в ручном режиме уже добавляем смещение строк
            cells = set((index.row() + 2, index.column()) for index in indexes)
            upper = _get_int_value(self.upper_limit_edit.text())
            lower = _get_int_value(self.lower_limit_edit.text())
            mapping = (list(cells), True, upper, lower, "cell")
        self.mappings.append(mapping)
        if mapping[-1] == "column":
            mapping_str = f"Лимит: {mapping[0]} -> Тексты: {', '.join(mapping[1])}"
        else:
            cells_str = ", ".join([f"({r},{self.headers[c]})" for (r, c) in mapping[0]])
            mapping_str = f"Ячейки: {cells_str}"
        if mapping[1] is True or mapping[-1] == "cell":
            mapping_str += f" (Вручную: верхний={mapping[2] if mapping[2] is not None else '—'}, нижний={mapping[3] if mapping[3] is not None else '—'})"
        self.mapping_list.addItem(mapping_str)
        self.clear_selection()

    def show_mapping_context_menu(self, pos):
        item = self.mapping_list.itemAt(pos)
        if item:
            menu = QMenu()
            delete_action = menu.addAction("Удалить")
            edit_action = menu.addAction("Редактировать")
            action = menu.exec_(self.mapping_list.mapToGlobal(pos))
            row = self.mapping_list.row(item)
            if action == delete_action:
                self.mapping_list.takeItem(row)
                if 0 <= row < len(self.mappings):
                    del self.mappings[row]
            elif action == edit_action:
                if 0 <= row < len(self.mappings):
                    mapping = self.mappings.pop(row)
                    self.mapping_list.takeItem(row)
                    if mapping[-1] == "column":
                        idx_limit = self.headers.index(mapping[0])
                        idx_texts = [self.headers.index(t) for t in mapping[1]]
                        self.current_limit = idx_limit
                        self.current_texts = set(idx_texts)
                        self.manual_checkbox.setChecked(False)
                        self.table_view.setSelectionMode(QtWidgets.QAbstractItemView.NoSelection)
                        self.table_view.horizontalHeader().show()
                    else:
                        self.manual_checkbox.setChecked(True)
                        self.table_view.setSelectionMode(QtWidgets.QAbstractItemView.ExtendedSelection)
                    if mapping[-1] == "cell":
                        if mapping[2] is not None:
                            self.upper_limit_edit.setText(str(mapping[2]))
                        if mapping[3] is not None:
                            self.lower_limit_edit.setText(str(mapping[3]))
                    self.update_header_colors()
                    self.update_current_mapping_label()

    def get_mappings(self):
        return self.mappings

    def accept(self):
        duplicates = self.check_duplicates()
        if duplicates:
            QMessageBox.critical(self, "Ошибка дублирования", "\n".join(duplicates))
            return
        super().accept()


# ==================== Основной класс LimitsChecker ====================
class LimitsChecker(QWidget):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("Проверка лимитов")
        self.resize(800, 600)
        self.stack = QStackedWidget()
        main_layout = QVBoxLayout()
        main_layout.addWidget(self.stack)
        self.setLayout(main_layout)
        self.selected_file = ""
        self.workbook = None
        self.sheet = None
        self.sheet_name = ""
        self.headers = []
        self.mappings = []
        self.report_text = ""
        self.file_page = self.create_file_selection_page()
        self.stack.addWidget(self.file_page)

    def create_file_selection_page(self):
        page = QWidget()
        layout = QVBoxLayout()
        file_group = QGroupBox("Выбор файла (нажмите 'Обзор...' или перетащите файл сюда)")
        file_layout = QHBoxLayout()
        self.file_line = DragDropLineEdit(self.update_sheet_list)
        self.browse_button = QPushButton("Обзор...")
        self.browse_button.clicked.connect(self.select_file)
        file_layout.addWidget(self.file_line)
        file_layout.addWidget(self.browse_button)
        file_group.setLayout(file_layout)
        layout.addWidget(file_group)
        sheet_group = QGroupBox("Выбор листа")
        sheet_layout = QHBoxLayout()
        self.sheet_combo = QComboBox()
        sheet_layout.addWidget(self.sheet_combo)
        sheet_group.setLayout(sheet_layout)
        layout.addWidget(sheet_group)
        btn_layout = QHBoxLayout()
        self.map_button = QPushButton("Лимиты")
        self.map_button.clicked.connect(self.open_mapping_dialog)
        self.next_button = QPushButton("Далее")
        self.next_button.clicked.connect(self.goto_results_page)
        btn_layout.addWidget(self.map_button)
        btn_layout.addWidget(self.next_button)
        layout.addLayout(btn_layout)
        page.setLayout(layout)
        return page

    def select_file(self):
        file_path, _ = QFileDialog.getOpenFileName(self, "Выберите Excel файл", "", "Excel Files (*.xlsx *.xls)")
        if file_path:
            self.file_line.setText(file_path)
            self.update_sheet_list(file_path)

    def update_sheet_list(self, file_path):
        try:
            wb = load_workbook(file_path, read_only=True)
            self.sheet_combo.clear()
            self.sheet_combo.addItems(wb.sheetnames)
            self.selected_file = file_path
        except Exception as e:
            QMessageBox.critical(self, "Ошибка", f"Не удалось загрузить листы: {e}")

    def open_mapping_dialog(self):
        if not self.selected_file:
            QMessageBox.critical(self, "Ошибка", "Выберите файл Excel.")
            return
        self.sheet_name = self.sheet_combo.currentText()
        try:
            self.workbook = load_workbook(self.selected_file)
            self.sheet = self.workbook[self.sheet_name]
            self.headers = [str(cell.value) if cell.value is not None else ""
                            for cell in next(self.sheet.iter_rows(min_row=1, max_row=1))]
            if not any(self.headers):
                QMessageBox.critical(self, "Ошибка", "Не найдены заголовки в первой строке.")
                return
        except Exception as e:
            QMessageBox.critical(self, "Ошибка", f"Ошибка при открытии файла: {e}")
            return
        from PyQt5.QtGui import QStandardItemModel, QStandardItem
        model = QStandardItemModel()
        model.setHorizontalHeaderLabels(self.headers)
        row_count = 10
        rows = list(self.sheet.iter_rows(min_row=2, max_row=row_count + 1, values_only=True))
        for row in rows:
            items = [QStandardItem(str(cell)) if cell is not None else QStandardItem("")
                     for cell in row]
            model.appendRow(items)
        dialog = LimitsMappingPreviewDialog(model, self.headers, self)
        if dialog.exec_() == QtWidgets.QDialog.Accepted:
            self.mappings = dialog.get_mappings()

    def goto_results_page(self):
        if not self.mappings:
            QMessageBox.critical(self, "Ошибка", "Сначала создайте сопоставления через кнопку 'Лимиты'.")
            return
        self.run_limit_check()

    def run_limit_check(self):
        if not self.mappings:
            QMessageBox.critical(self, "Ошибка", "Не заданы сопоставления лимитов.")
            return
        report_lines = []
        total_violations = 0
        mapping_indices = []
        for mapping in self.mappings:
            if mapping[-1] == "column":
                try:
                    limit_index = self.headers.index(mapping[0]) + 1
                except ValueError:
                    QMessageBox.critical(self, "Ошибка", f"Столбец '{mapping[0]}' не найден.")
                    return
                text_indices = []
                for txt in mapping[1]:
                    try:
                        text_indices.append(self.headers.index(txt) + 1)
                    except ValueError:
                        QMessageBox.critical(self, "Ошибка", f"Столбец '{txt}' не найден.")
                        return
                mapping_indices.append(("column", limit_index, text_indices, mapping[2], mapping[3], mapping[4]))
            else:
                mapping_indices.append(("cell", mapping[0], mapping[1], mapping[2], mapping[3], mapping[4]))
        # Автоматический режим: проверяем каждую строку, сравниваем с лимитом из ячейки
        for row in self.sheet.iter_rows(min_row=2, values_only=False):
            row_num = row[0].row
            for m in mapping_indices:
                if m[0] == "column":
                    _, limit_idx, text_idxs, manual, upper, lower = m
                    limit_cell = row[limit_idx - 1]
                    if manual:
                        current_limit = _get_int_value(upper)
                        current_lower = _get_int_value(lower)
                    else:
                        current_limit = _get_int_value(limit_cell.value)
                        current_lower = None
                    for txt_idx in text_idxs:
                        text_cell = row[txt_idx - 1]
                        cell_text = text_cell.value
                        if cell_text is None:
                            continue
                        text_str = str(cell_text)
                        text_length = len(text_str)
                        violation = False
                        detail = ""
                        if current_limit is not None and text_length > current_limit:
                            violation = True
                            detail += f"длина = {text_length} (лимит {current_limit})"
                        if current_lower is not None and text_length < current_lower:
                            violation = True
                            detail += f", длина = {text_length} (нижний лимит {current_lower})"
                        if violation:
                            fill = PatternFill(start_color="FF9999", end_color="FF9999", fill_type="solid")
                            text_cell.fill = fill
                            total_violations += 1
                            header = self.headers[txt_idx - 1]
                            report_lines.append(f"Строка {row_num}, столбец '{header}': {detail}")
        # Ручной режим: проверяем каждый сохранённый диапазон ячеек, без дополнительного смещения (при сохранении уже прибавлено +2)
        for m in mapping_indices:
            if m[0] == "cell":
                _, selected_cells, manual, upper, lower, _ = m
                current_limit = _get_int_value(upper)
                current_lower = _get_int_value(lower)
                for (model_row, col) in selected_cells:
                    sheet_row = model_row  # Используем номер строки как сохранённый
                    cell_obj = self.sheet.cell(row=sheet_row, column=col + 1)
                    cell_text = cell_obj.value
                    if cell_text is None:
                        continue
                    text_str = str(cell_text)
                    text_length = len(text_str)
                    violation = False
                    detail = ""
                    # Если превышает верхний лимит, окрашиваем в бледно-красный
                    if current_limit is not None and text_length > current_limit:
                        violation = True
                        detail += f"длина = {text_length} (лимит {current_limit})"
                        fill = PatternFill(start_color="FF9999", end_color="FF9999", fill_type="solid")
                    # Если меньше нижнего лимита, окрашиваем в бледно-оранжевый
                    elif current_lower is not None and text_length < current_lower:
                        violation = True
                        detail += f"длина = {text_length} (нижний лимит {current_lower})"
                        fill = PatternFill(start_color="FFD699", end_color="FFD699", fill_type="solid")
                    if violation:
                        cell_obj.fill = fill
                        total_violations += 1
                        header = self.headers[col]
                        report_lines.append(f"Строка {sheet_row}, столбец '{header}': {detail}")
        base, ext = os.path.splitext(self.selected_file)
        output_file = f"{base}_checked{ext}"
        try:
            self.workbook.save(output_file)
        except Exception as e:
            QMessageBox.critical(self, "Ошибка", f"Не удалось сохранить файл: {e}")
            return
        report = "Проверка лимитов завершена.\n"
        report += f"Всего нарушений: {total_violations}\n\n"
        if report_lines:
            report += "Детали:\n" + "\n".join(report_lines)
        else:
            report += "Нарушений не обнаружено."
        self.report_text = report
        self.results_page = self.create_results_page(output_file)
        self.stack.addWidget(self.results_page)
        self.stack.setCurrentWidget(self.results_page)

    def create_results_page(self, output_file):
        page = QWidget()
        layout = QVBoxLayout()
        report_label = QLabel("Результаты проверки:")
        layout.addWidget(report_label)
        self.report_text_edit = QTextEdit()
        self.report_text_edit.setReadOnly(True)
        self.report_text_edit.setText(self.report_text)
        layout.addWidget(self.report_text_edit)
        file_info = QLabel(f"Изменённый файл сохранён как:\n{output_file}")
        layout.addWidget(file_info)
        btn_layout = QHBoxLayout()
        back_btn = QPushButton("Вернуться к сопоставлению")
        back_btn.clicked.connect(self.go_back_to_file_page)
        close_btn = QPushButton("Закрыть")
        close_btn.clicked.connect(self.close)
        btn_layout.addWidget(back_btn)
        btn_layout.addWidget(close_btn)
        layout.addLayout(btn_layout)
        page.setLayout(layout)
        return page

    def go_back_to_file_page(self):
        self.stack.setCurrentWidget(self.file_page)


if __name__ == '__main__':
    app = QtWidgets.QApplication(sys.argv)
    window = LimitsChecker()
    window.show()
    sys.exit(app.exec_())
