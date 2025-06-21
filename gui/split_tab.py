from PySide6.QtWidgets import (
    QWidget, QVBoxLayout, QLabel, QPushButton, QComboBox, QMessageBox,
    QApplication, QProgressDialog
)
from PySide6.QtCore import Qt
from utils.i18n import tr, i18n
from core.drag_drop import DragDropLineEdit
from core.split_excel import split_excel_by_languages
from gui.split_mapping_dialog import SplitMappingDialog
from openpyxl import load_workbook


class SplitTab(QWidget):
    def __init__(self):
        super().__init__()
        self.excel_path = ''
        self.sheet_name = ''
        self.headers = []
        # mapping per sheet: {sheet: (src, [targets], [extras])}
        self.sheet_mappings: dict[str, tuple[str, list[str], list[str]]] = {}
        self.init_ui()
        i18n.language_changed.connect(self.retranslate_ui)
        self.retranslate_ui()

    def init_ui(self):
        layout = QVBoxLayout(self)

        self.file_input = DragDropLineEdit(mode='file')
        self.file_input.fileSelected.connect(self.on_file_selected)
        layout.addWidget(QLabel(tr("Файл Excel:")))
        layout.addWidget(self.file_input)

        self.sheet_combo = QComboBox()
        self.sheet_combo.currentTextChanged.connect(self.on_sheet_changed)
        layout.addWidget(self.sheet_combo)

        layout.addWidget(QLabel(tr("Исходный язык:")))
        self.source_combo = QComboBox()
        layout.addWidget(self.source_combo)

        self.config_btn = QPushButton(tr("Превью"))
        self.config_btn.clicked.connect(self.open_mapping_dialog)
        layout.addWidget(self.config_btn)

        self.current_label = QLabel(tr("Текущая настройка: —"))
        self.current_label.setWordWrap(True)
        layout.addWidget(self.current_label)

        self.split_btn = QPushButton(tr("Разделить"))
        self.split_btn.clicked.connect(self.run_split)
        layout.addWidget(self.split_btn)

    def on_file_selected(self, path):
        try:
            wb = load_workbook(path, read_only=True)
            self.excel_path = path
            self.sheet_combo.clear()
            self.sheet_combo.addItems(wb.sheetnames)
            self.sheet_name = wb.sheetnames[0] if wb.sheetnames else ''
            wb.close()
            self.load_headers()
        except Exception as e:
            QMessageBox.critical(self, tr("Ошибка"), str(e))

    def on_sheet_changed(self, name):
        self.sheet_name = name
        self.load_headers()

    def load_headers(self):
        if not self.excel_path or not self.sheet_name:
            return
        wb = load_workbook(self.excel_path, read_only=True)
        sheet = wb[self.sheet_name]
        self.headers = [
            str(cell.value) if cell.value is not None else ''
            for cell in next(sheet.iter_rows(min_row=1, max_row=1))
        ]
        wb.close()
        self.source_combo.clear()
        self.source_combo.addItems([h for h in self.headers if h])

        # reset current selection
        self.sheet_mappings = {}
        self.current_label.setText(tr("Текущая настройка: —"))

    def open_mapping_dialog(self):
        if not self.excel_path:
            QMessageBox.critical(self, tr("Ошибка"), tr("Выберите файл Excel."))
            return
        wb = load_workbook(self.excel_path, read_only=True)
        sheets = wb.sheetnames
        wb.close()
        dialog = SplitMappingDialog(self.excel_path, sheets, self)
        if dialog.exec():
            self.sheet_mappings = dialog.get_selection()
            if self.sheet_mappings:
                total_targets = sum(len(cfg[1]) if cfg[1] else 0 for cfg in self.sheet_mappings.values())
                self.current_label.setText(
                    tr("Настроено листов: {n}; выбрано целей: {m}").format(
                        n=len(self.sheet_mappings), m=total_targets
                    )
                )

    def run_split(self):
        if not self.excel_path:
            QMessageBox.critical(self, tr("Ошибка"), tr("Выберите файл Excel."))
            return
        if not self.sheet_mappings:
            QMessageBox.critical(self, tr("Ошибка"), tr("Сначала настройте листы."))
            return
        try:
            progress = QProgressDialog(tr("Сохранение..."), tr("Отмена"), 0, 0, self)
            progress.setWindowTitle(tr("Прогресс"))
            progress.setWindowModality(Qt.ApplicationModal)

            for sheet, (src, targets, extras) in self.sheet_mappings.items():
                def cb(i, total, name, sh=sheet):
                    progress.setMaximum(total)
                    progress.setValue(i)
                    progress.setLabelText(tr("{sheet}: {name}").format(sheet=sh, name=name))
                    QApplication.processEvents()

                split_excel_by_languages(
                    self.excel_path,
                    sheet,
                    src,
                    target_langs=targets if targets else None,
                    extra_columns=extras,
                    progress_callback=cb,
                )

            progress.close()
            QMessageBox.information(self, tr("Успех"), tr("Файлы успешно сохранены."))
        except Exception as e:
            QMessageBox.critical(self, tr("Ошибка"), str(e))

    def retranslate_ui(self):
        self.setWindowTitle(tr("Разделение"))
        self.split_btn.setText(tr("Разделить"))
        self.config_btn.setText(tr("Превью"))
        if self.sheet_mappings:
            total_targets = sum(len(cfg[1]) if cfg[1] else 0 for cfg in self.sheet_mappings.values())
            self.current_label.setText(
                tr("Настроено листов: {n}; выбрано целей: {m}").format(
                    n=len(self.sheet_mappings), m=total_targets
                )
            )
        else:
            self.current_label.setText(tr("Текущая настройка: —"))
        # update labels - they are static but to refresh we need to re-add them? Not necessary
