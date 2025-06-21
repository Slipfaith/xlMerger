from PySide6.QtWidgets import (
    QDialog, QVBoxLayout, QHBoxLayout, QLabel, QPushButton, QTableView,
    QListWidget, QListWidgetItem
)
from PySide6.QtCore import Qt
from PySide6.QtGui import QStandardItemModel, QStandardItem, QColor, QBrush
from openpyxl import load_workbook

from gui.limits_checker import DraggableHeaderView
from utils.i18n import tr


class SplitMappingDialog(QDialog):
    """Dialog to select source and target columns using a preview."""

    def __init__(self, excel_path: str, sheet_name: str, parent=None):
        super().__init__(parent)
        self.excel_path = excel_path
        self.sheet_name = sheet_name
        self.headers: list[str] = []
        self.model = QStandardItemModel()
        self.source_col: int | None = None
        self.target_cols: set[int] = set()
        self.extra_cols: set[int] = set()
        self._load_preview()
        self._init_ui()

    def _load_preview(self):
        wb = load_workbook(self.excel_path, read_only=True)
        sheet = wb[self.sheet_name]
        self.headers = [
            str(cell.value) if cell.value is not None else ""
            for cell in next(sheet.iter_rows(min_row=1, max_row=1))
        ]
        self.model.setHorizontalHeaderLabels(self.headers)
        rows = list(sheet.iter_rows(min_row=2, max_row=11, values_only=True))
        for row in rows:
            items = [
                QStandardItem(str(cell)) if cell is not None else QStandardItem("")
                for cell in row
            ]
            self.model.appendRow(items)
        wb.close()

    def _init_ui(self):
        self.setWindowTitle(tr("Настройка разделения"))
        layout = QVBoxLayout(self)
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
        self.table.horizontalHeader().setStretchLastSection(False)
        layout.addWidget(self.table)

        layout.addWidget(QLabel(tr("Дополнительные столбцы:")))
        self.extra_list = QListWidget()
        for h in self.headers:
            item = QListWidgetItem(h)
            item.setCheckState(Qt.Unchecked)
            self.extra_list.addItem(item)
        self.extra_list.itemChanged.connect(self.update_label)
        layout.addWidget(self.extra_list)

        self.current_label = QLabel()
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

    def handle_drag(self, selection: set[int]):
        start = getattr(self.header_view, "_drag_start", None)
        if self.source_col is None:
            if start is None:
                return
            self.source_col = start
            self.target_cols = set(selection)
            self.target_cols.discard(self.source_col)
        else:
            sel = set(selection)
            if self.source_col in sel:
                sel.remove(self.source_col)
            self.target_cols = sel
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
            if item.checkState() == Qt.Checked:
                if i != self.source_col and i not in self.target_cols:
                    self.extra_cols.add(i)

    def update_label(self):
        self._collect_extras()
        if self.source_col is None:
            txt = f"{tr('Источник')}: —; {tr('Цели')}: —; {tr('Доп')}: —"
        else:
            src = self.headers[self.source_col]
            tgts = [self.headers[c] for c in sorted(self.target_cols)] if self.target_cols else ['—']
            extras = [self.headers[c] for c in sorted(self.extra_cols)] if self.extra_cols else ['—']
            txt = f"{tr('Источник')}: {src}; {tr('Цели')}: {', '.join(tgts)}; {tr('Доп')}: {', '.join(extras)}"
        self.current_label.setText(tr("Текущая настройка: {txt}").format(txt=txt))

    def get_selection(self):
        if self.source_col is None:
            return None, [], []
        source = self.headers[self.source_col]
        targets = [self.headers[c] for c in sorted(self.target_cols)]
        extras = [self.headers[c] for c in sorted(self.extra_cols)]
        return source, targets, extras
