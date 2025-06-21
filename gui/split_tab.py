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
        self.source_lang = ''
        self.target_langs: list[str] = []
        self.extra_columns: list[str] = []
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
        self.source_lang = ''
        self.target_langs = []
        self.extra_columns = []
        self.current_label.setText(tr("Текущая настройка: —"))

    def open_mapping_dialog(self):
        if not self.excel_path:
            QMessageBox.critical(self, tr("Ошибка"), tr("Выберите файл Excel."))
            return
        dialog = SplitMappingDialog(self.excel_path, self.sheet_name, self)
        if dialog.exec():
            src, targets, extras = dialog.get_selection()
            if src:
                self.source_lang = src
                self.target_langs = targets
                self.extra_columns = extras
                self.current_label.setText(
                    tr("Текущая настройка: {txt}").format(
                        txt=f"{src} -> {', '.join(targets) if targets else '—'}; {tr('Доп')}: {', '.join(extras) if extras else '—'}"
                    )
                )

    def run_split(self):
        if not self.excel_path:
            QMessageBox.critical(self, tr("Ошибка"), tr("Выберите файл Excel."))
            return
        src = self.source_lang or self.source_combo.currentText()
        targets = self.target_langs if self.target_langs else None
        extras = self.extra_columns
        try:
            progress = QProgressDialog(tr("Сохранение..."), tr("Отмена"), 0, 0, self)
            progress.setWindowTitle(tr("Прогресс"))
            progress.setWindowModality(Qt.ApplicationModal)

            def cb(i, total, name):
                progress.setMaximum(total)
                progress.setValue(i)
                progress.setLabelText(tr("Сохраняется: {name}").format(name=name))
                QApplication.processEvents()

            split_excel_by_languages(
                self.excel_path,
                self.sheet_name,
                src,
                target_langs=targets,
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
        if self.source_lang:
            txt = f"{self.source_lang} -> {', '.join(self.target_langs) if self.target_langs else '—'}; {tr('Доп')}: {', '.join(self.extra_columns) if self.extra_columns else '—'}"
            self.current_label.setText(tr("Текущая настройка: {txt}").format(txt=txt))
        else:
            self.current_label.setText(tr("Текущая настройка: —"))
        # update labels - they are static but to refresh we need to re-add them? Not necessary
