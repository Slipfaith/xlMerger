# -*- coding: utf-8 -*-
from PySide6.QtCore import Qt
from PySide6.QtWidgets import (
    QAbstractItemView,
    QDialog,
    QFrame,
    QHBoxLayout,
    QHeaderView,
    QLabel,
    QListWidget,
    QPushButton,
    QStackedWidget,
    QTableWidget,
    QTableWidgetItem,
    QVBoxLayout,
    QComboBox,
)
from .style_system import set_button_variant, set_label_role
import os


class SheetMappingDialog(QDialog):
    """Dialog to map source sheet names to target sheet names."""

    def __init__(self, main_sheets, file_to_sheets, auto_map=None, parent=None):
        super().__init__(parent)
        self.main_sheets = list(main_sheets)
        self.file_to_sheets = dict(file_to_sheets)
        self.auto_map = auto_map or {}
        self.comboboxes = {}
        self._file_order = list(self.file_to_sheets.keys())
        self._build_ui()

    def _build_ui(self):
        self.setWindowTitle("Sheet mapping")
        self.resize(880, 560)
        self.setMinimumSize(760, 420)
        self.setMaximumHeight(700)

        layout = QVBoxLayout(self)
        layout.setContentsMargins(16, 16, 16, 16)
        layout.setSpacing(12)

        title = QLabel("Sheet mapping")
        set_label_role(title, "heading")
        subtitle = QLabel("Map source sheets to target sheets")
        set_label_role(subtitle, "muted")
        layout.addWidget(title)
        layout.addWidget(subtitle)

        if len(self._file_order) == 1:
            only_file = self._file_order[0]
            source_info = QLabel(f"Source file: {os.path.basename(only_file)}")
            set_label_role(source_info, "muted")
            layout.addWidget(source_info)
            layout.addWidget(self._build_table_for_file(only_file), 1)
        else:
            split = QHBoxLayout()
            split.setSpacing(12)

            self.file_list = QListWidget()
            self.file_list.setMinimumWidth(240)
            self.file_list.setMaximumWidth(320)
            self.file_list.setSelectionMode(QAbstractItemView.SingleSelection)

            self.stack = QStackedWidget()

            for file_path in self._file_order:
                self.file_list.addItem(os.path.basename(file_path))
                page = self._build_file_page(file_path)
                self.stack.addWidget(page)

            self.file_list.currentRowChanged.connect(self._on_file_row_changed)
            self.file_list.setCurrentRow(0)

            split.addWidget(self.file_list)
            split.addWidget(self.stack, 1)
            layout.addLayout(split, 1)

        divider = QFrame()
        divider.setFrameShape(QFrame.HLine)
        divider.setFrameShadow(QFrame.Plain)
        layout.addWidget(divider)

        buttons = QHBoxLayout()
        buttons.addStretch()
        cancel_btn = QPushButton("Cancel")
        apply_btn = QPushButton("Apply")
        set_button_variant(cancel_btn, "ghost")
        set_button_variant(apply_btn, "primary")
        apply_btn.setDefault(True)
        cancel_btn.clicked.connect(self.reject)
        apply_btn.clicked.connect(self.accept)
        buttons.addWidget(cancel_btn)
        buttons.addWidget(apply_btn)
        layout.addLayout(buttons)

    def _build_file_page(self, file_path):
        page = QFrame()
        page_layout = QVBoxLayout(page)
        page_layout.setContentsMargins(0, 0, 0, 0)
        page_layout.setSpacing(8)

        header = QLabel(f"Source file: {os.path.basename(file_path)}")
        set_label_role(header, "muted")
        page_layout.addWidget(header)
        page_layout.addWidget(self._build_table_for_file(file_path), 1)
        return page

    def _build_table_for_file(self, file_path):
        source_sheets = list(self.file_to_sheets.get(file_path, []))
        table = QTableWidget(len(source_sheets), 2)
        table.setObjectName("sheetMappingTable")
        table.setHorizontalHeaderLabels(["Source sheet", "Target sheet"])
        table.verticalHeader().setVisible(False)
        table.horizontalHeader().setSectionResizeMode(0, QHeaderView.Stretch)
        table.horizontalHeader().setSectionResizeMode(1, QHeaderView.Stretch)
        table.horizontalHeader().setMinimumSectionSize(120)
        table.setShowGrid(False)
        table.setSelectionMode(QAbstractItemView.NoSelection)
        table.setEditTriggers(QAbstractItemView.NoEditTriggers)
        table.setFocusPolicy(Qt.NoFocus)
        table.setFrameShape(QFrame.NoFrame)

        for row, source_sheet in enumerate(source_sheets):
            source_item = QTableWidgetItem(source_sheet)
            source_item.setFlags(Qt.ItemIsEnabled)
            table.setItem(row, 0, source_item)

            combo = QComboBox(table)
            combo.setProperty("flatSelect", True)
            combo.addItem("")
            combo.addItems(self.main_sheets)
            initial_target = self._get_initial_target(file_path, source_sheet)
            if initial_target:
                combo.setCurrentText(initial_target)
            table.setCellWidget(row, 1, combo)
            self.comboboxes[(file_path, source_sheet)] = combo
            table.setRowHeight(row, 34)

        table.verticalHeader().setDefaultSectionSize(34)
        return table

    def _get_initial_target(self, file_path, source_sheet):
        prefilled = self.auto_map.get(file_path, {})
        for target_sheet, mapped_source in prefilled.items():
            if mapped_source == source_sheet and target_sheet in self.main_sheets:
                return target_sheet
        if source_sheet in self.main_sheets:
            return source_sheet
        return ""

    def _on_file_row_changed(self, row):
        if row >= 0:
            self.stack.setCurrentIndex(row)

    def get_mapping(self):
        mapping = {}
        for file_path in self._file_order:
            source_sheets = list(self.file_to_sheets.get(file_path, []))
            file_mapping = {}
            prefilled = self.auto_map.get(file_path, {})

            # Fallbacks keep previous behavior: each target sheet gets a source.
            for target_sheet in self.main_sheets:
                if target_sheet in prefilled and prefilled[target_sheet] in source_sheets:
                    file_mapping[target_sheet] = prefilled[target_sheet]
                elif target_sheet in source_sheets:
                    file_mapping[target_sheet] = target_sheet
                elif source_sheets:
                    file_mapping[target_sheet] = source_sheets[0]
                else:
                    file_mapping[target_sheet] = ""

            # Explicit UI selections override fallbacks.
            for source_sheet in source_sheets:
                combo = self.comboboxes.get((file_path, source_sheet))
                if combo is None:
                    continue
                target_sheet = combo.currentText().strip()
                if target_sheet in self.main_sheets:
                    file_mapping[target_sheet] = source_sheet

            mapping[file_path] = file_mapping

        return mapping
