from PySide6.QtWidgets import (
    QDialog,
    QVBoxLayout,
    QHBoxLayout,
    QLabel,
    QPushButton,
    QTableView,
    QListWidget,
    QListWidgetItem,
    QComboBox,
    QListView,
)
from PySide6.QtCore import Qt
from PySide6.QtGui import QStandardItemModel, QStandardItem, QColor, QBrush
from openpyxl import load_workbook
from itertools import islice

from gui.limits_checker import DraggableHeaderView
from utils.i18n import tr


class SplitMappingDialog(QDialog):
    """Dialog to select source/target columns for one or many sheets."""

    def __init__(self, excel_path: str, sheet_names, parent=None):
        super().__init__(parent)
        if isinstance(sheet_names, str):
            sheet_names = [sheet_names]
        self.excel_path = excel_path
        self.sheet_names = sheet_names
        self.current_sheet = sheet_names[0]

        self.headers_map: dict[str, list[str]] = {}
        self.models: dict[str, QStandardItemModel] = {}
        # columns that contain at least one non-empty value for each sheet
        self.non_empty_cols: dict[str, set[int]] = {}
        self.configs = {
            name: {"src": None, "targets": set(), "extra": set()}
            for name in sheet_names
        }

        self.source_col: int | None = None
        self.target_cols: set[int] = set()
        self.extra_cols: set[int] = set()

        self._load_all_previews()
        self._init_ui()

    def _load_all_previews(self):
        wb = load_workbook(self.excel_path, read_only=True)
        for name in self.sheet_names:
            sheet = wb[name]
            headers = [
                str(cell.value) if cell.value is not None else ""
                for cell in next(sheet.iter_rows(min_row=1, max_row=1))
            ]
            rows_for_preview = list(sheet.iter_rows(min_row=2, max_row=11, values_only=True))
            rows_for_check = list(islice(sheet.iter_rows(min_row=2, values_only=True), 30))

            non_empty = {idx for idx, h in enumerate(headers) if str(h).strip() != ""}
            if rows_for_check:
                for idx, col in enumerate(zip(*rows_for_check)):
                    if any(v not in (None, "") for v in col):
                        non_empty.add(idx)

            keep_idx = sorted(non_empty)
            filtered_headers = [headers[i] for i in keep_idx]
            model = QStandardItemModel()
            model.setHorizontalHeaderLabels([str(h) if h is not None else "" for h in filtered_headers])
            for row in rows_for_preview:
                filtered_row = [row[i] if i < len(row) else None for i in keep_idx]
                items = [
                    QStandardItem(str(cell)) if cell is not None else QStandardItem("")
                    for cell in filtered_row
                ]
                model.appendRow(items)

            self.headers_map[name] = [str(h) if h is not None else "" for h in filtered_headers]
            self.models[name] = model
            self.non_empty_cols[name] = set(range(len(keep_idx)))
        wb.close()
        self.model = self.models[self.current_sheet]
        self.headers = self.headers_map[self.current_sheet]

    def _init_ui(self):
        self.setWindowTitle(tr("Настройка разделения"))
        layout = QVBoxLayout(self)

        if len(self.sheet_names) > 1:
            self.sheet_combo = QComboBox()
            self.sheet_combo.addItems(self.sheet_names)
            self.sheet_combo.currentTextChanged.connect(self.switch_sheet)
            layout.addWidget(self.sheet_combo)

        info = QLabel(tr("Выбери исходный столбец (синий) и столбцы перевода (зелёный)."))
        info.setWordWrap(True)
        layout.addWidget(info)

        self.table = QTableView()
        self.table.setModel(self.model)
        self.table.setEditTriggers(QTableView.NoEditTriggers)
        self.table.setSelectionMode(QTableView.NoSelection)
        self.table.setSelectionBehavior(QTableView.SelectItems)
        self.header_view = DraggableHeaderView(Qt.Horizontal, self.table)
        self.table.setHorizontalHeader(self.header_view)
        self.header_view.dragSelectionChanged.connect(self.handle_drag)
        self.header_view.rightClicked.connect(self.handle_right_click)
        self.table.horizontalHeader().setStretchLastSection(False)
        layout.addWidget(self.table)

        layout.addWidget(QLabel(tr("Дополнительные столбцы:")))
        self.extra_list = QListWidget()
        self.extra_list.setFlow(QListView.LeftToRight)
        self.extra_list.setWrapping(True)
        self.extra_list.setMaximumHeight(80)
        non_empty = self.non_empty_cols.get(self.current_sheet, set())
        for idx, h in enumerate(self.headers):
            item = QListWidgetItem(h)
            if idx not in non_empty:
                item.setFlags(item.flags() & ~Qt.ItemIsEnabled)
                item.setForeground(QBrush(QColor("gray")))
            item.setCheckState(Qt.Unchecked)
            self.extra_list.addItem(item)
        self.extra_list.itemChanged.connect(self.update_label)
        layout.addWidget(self.extra_list)

        self.count_label = QLabel()
        layout.addWidget(self.count_label)

        self.current_label = QLabel()
        self.current_label.setWordWrap(True)
        layout.addWidget(self.current_label)

        btn_layout = QHBoxLayout()
        ok_btn = QPushButton(tr("Готово"))
        ok_btn.clicked.connect(self.accept)
        cancel_btn = QPushButton(tr("Отмена"))
        cancel_btn.clicked.connect(self.reject)
        btn_layout.addWidget(ok_btn)
        btn_layout.addWidget(cancel_btn)
        layout.addLayout(btn_layout)

        self.update_colors()
        self.update_label()
        self.resize(700, 500)
        self.setFixedSize(self.size())

    def _save_current(self):
        cfg = self.configs[self.current_sheet]
        cfg["src"] = self.source_col
        cfg["targets"] = set(self.target_cols)
        cfg["extra"] = set(self.extra_cols)

    def switch_sheet(self, name: str):
        self._save_current()
        self.current_sheet = name
        self.model = self.models[name]
        self.headers = self.headers_map[name]
        self.table.setModel(self.model)
        cfg = self.configs[name]
        self.source_col = cfg["src"]
        self.target_cols = set(cfg["targets"])
        self.extra_cols = set(cfg["extra"])
        self._rebuild_extra_list()
        self.update_colors()
        self.update_label()

    def _rebuild_extra_list(self):
        self.extra_list.blockSignals(True)
        self.extra_list.clear()
        non_empty = self.non_empty_cols.get(self.current_sheet, set())
        for idx, h in enumerate(self.headers):
            item = QListWidgetItem(h)
            if idx not in non_empty:
                item.setFlags(item.flags() & ~Qt.ItemIsEnabled)
                item.setForeground(QBrush(QColor("gray")))
            if idx in self.extra_cols:
                item.setCheckState(Qt.Checked)
            else:
                item.setCheckState(Qt.Unchecked)
            self.extra_list.addItem(item)
        self.extra_list.blockSignals(False)

    def handle_drag(self, selection: set[int]):
        start = getattr(self.header_view, "_drag_start", None)
        allowed = self.non_empty_cols.get(self.current_sheet, set())
        if self.source_col is None:
            if start is None:
                return
            if start not in allowed:
                return
            self.source_col = start
            sel = {c for c in selection if c in allowed and c != self.source_col}
            self.target_cols.update(sel)
        else:
            sel = set(selection)
            if self.source_col in sel:
                sel.remove(self.source_col)
            filtered = {c for c in sel if c in allowed}
            self.target_cols.update(filtered)
        self.update_colors()
        self.update_label()

    def handle_right_click(self, index: int):
        if index == self.source_col:
            self.source_col = None
            self.target_cols.clear()
        elif index in self.target_cols:
            self.target_cols.remove(index)
        else:
            return
        self.update_colors()
        self.update_label()

    def update_colors(self):
        for row in range(self.model.rowCount()):
            for col in range(self.model.columnCount()):
                idx = self.model.index(row, col)
                self.model.setData(idx, QBrush(QColor("white")), Qt.BackgroundRole)
        if self.source_col is not None:
            for row in range(self.model.rowCount()):
                idx = self.model.index(row, self.source_col)
                self.model.setData(idx, QBrush(QColor("#a2cffe")), Qt.BackgroundRole)
        for col in self.target_cols:
            for row in range(self.model.rowCount()):
                idx = self.model.index(row, col)
                self.model.setData(idx, QBrush(QColor("#b6fcb6")), Qt.BackgroundRole)
        for col in self.extra_cols:
            for row in range(self.model.rowCount()):
                idx = self.model.index(row, col)
                self.model.setData(idx, QBrush(QColor("#fff2a8")), Qt.BackgroundRole)

    def _collect_extras(self):
        self.extra_cols.clear()
        for i in range(self.extra_list.count()):
            item = self.extra_list.item(i)
            if not item.flags() & Qt.ItemIsEnabled:
                continue
            if item.checkState() == Qt.Checked:
                if i != self.source_col and i not in self.target_cols:
                    self.extra_cols.add(i)

    def update_label(self):
        self._collect_extras()
        if self.source_col is None:
            txt = f"{tr('Источник')}: —\n{tr('Цели')}: —\n{tr('Доп')}: —"
        else:
            src = self.headers[self.source_col]
            tgts = [self.headers[c] for c in sorted(self.target_cols)] if self.target_cols else ['—']
            extras = [self.headers[c] for c in sorted(self.extra_cols)] if self.extra_cols else ['—']
            txt = (
                f"{tr('Источник')}: {src}\n"
                f"{tr('Цели')}: {', '.join(tgts)}\n"
                f"{tr('Доп')}: {', '.join(extras)}"
            )
        self.count_label.setText(tr("Выбрано целей: {n}").format(n=len(self.target_cols)))
        self.current_label.setText(tr("Текущая настройка:\n{txt}").format(txt=txt))

    def get_selection(self):
        self._save_current()
        result = {}
        for sheet, cfg in self.configs.items():
            if cfg["src"] is None:
                continue
            headers = self.headers_map[sheet]
            src = headers[cfg["src"]]
            tgts = [headers[c] for c in sorted(cfg["targets"])]
            extras = [headers[c] for c in sorted(cfg["extra"])]
            result[sheet] = (src, tgts, extras)
        return result
