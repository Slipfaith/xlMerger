from PySide6.QtWidgets import (
    QDialog, QVBoxLayout, QHBoxLayout, QLabel, QPushButton, QTableView
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
        layout.addWidget(self.table)

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

    def handle_drag(self, selection: set[int]):
        if self.source_col is None and len(selection) == 1:
            self.source_col = next(iter(selection))
            self.target_cols.clear()
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

    def update_label(self):
        if self.source_col is None:
            txt = f"{tr('Источник')}: —; {tr('Цели')}: —"
        else:
            src = self.headers[self.source_col]
            tgts = [self.headers[c] for c in sorted(self.target_cols)] if self.target_cols else ['—']
            txt = f"{tr('Источник')}: {src}; {tr('Цели')}: {', '.join(tgts)}"
        self.current_label.setText(tr("Текущая настройка: {txt}").format(txt=txt))

    def get_selection(self):
        if self.source_col is None:
            return None, []
        source = self.headers[self.source_col]
        targets = [self.headers[c] for c in sorted(self.target_cols)]
        return source, targets
