# gui/pages/progress_page.py

from PySide6.QtWidgets import QWidget, QVBoxLayout, QLabel, QProgressBar, QHBoxLayout
from PySide6.QtCore import Qt, QTimer, QEasingCurve, Property
from PySide6.QtGui import QFont, QMovie

class ProgressPage(QWidget):
    def __init__(self, parent=None):
        super().__init__(parent)
        self.setWindowTitle("Копирование переводов")

        # --- Текущий файл и статус ---
        self.label_file_info = QLabel("")
        self.label_file_info.setAlignment(Qt.AlignCenter)
        font = QFont()
        font.setPointSize(10)
        font.setBold(True)
        self.label_file_info.setFont(font)

        # --- Прогресс-бар ---
        self.progress_bar = QProgressBar(self)
        self.progress_bar.setAlignment(Qt.AlignCenter)
        self.progress_bar.setStyleSheet("""
            QProgressBar {
                text-align: center;
                height: 30px;
                border-radius: 7px;
                background-color: #F3F3F3;
                font-size: 14px;
            }
            QProgressBar::chunk {
                background-color: #f47929;
                border-radius: 7px;
            }
        """)

        # --- Галочка завершения ---
        self.label_done = QLabel("")
        self.label_done.setAlignment(Qt.AlignCenter)
        done_font = QFont()
        done_font.setPointSize(16)
        done_font.setBold(True)
        self.label_done.setFont(done_font)
        self.label_done.hide()

        # --- Анимация барa ---
        self._current_value = 0
        self._target_value = 0
        self._anim_timer = QTimer(self)
        self._anim_timer.timeout.connect(self._animate_progress)

        # --- Layout ---
        layout = QVBoxLayout(self)
        layout.addSpacing(12)
        layout.addWidget(self.label_file_info)
        layout.addSpacing(8)
        layout.addWidget(self.progress_bar)
        layout.addSpacing(10)
        layout.addWidget(self.label_done)
        self.setLayout(layout)

    def set_progress(self, value, maximum=None, file_index=None, file_total=None, filename=None):
        # Ставим максимум если надо
        if maximum is not None:
            self.progress_bar.setMaximum(maximum)
        # Плавная анимация
        self._target_value = int(value)
        if not self._anim_timer.isActive():
            self._anim_timer.start(10)
        # Подпись текущего файла
        if file_index is not None and file_total is not None and filename:
            self.label_file_info.setText(
                f"Файл {file_index} из {file_total}: {self._short_name(filename)}"
            )
        elif filename:
            self.label_file_info.setText(f"Сейчас: {self._short_name(filename)}")
        else:
            self.label_file_info.clear()
        self.label_done.hide()

    def set_complete(self):
        self.label_done.setText("✔ Готово!")
        self.label_done.setStyleSheet("color: #47A447;")
        self.label_done.show()
        self.progress_bar.setValue(self.progress_bar.maximum())
        self.label_file_info.setText("Файлы успешно скопированы!")

    def _animate_progress(self):
        step = max(1, abs(self._target_value - self._current_value) // 10)
        if self._current_value < self._target_value:
            self._current_value += step
            if self._current_value > self._target_value:
                self._current_value = self._target_value
        elif self._current_value > self._target_value:
            self._current_value -= step
            if self._current_value < self._target_value:
                self._current_value = self._target_value
        self.progress_bar.setValue(self._current_value)
        if self._current_value == self._target_value:
            self._anim_timer.stop()

    def get_progressbar(self):
        return self.progress_bar

    @staticmethod
    def _short_name(name, n=12):
        base = name
        if len(base) <= 2 * n + 3:
            return base
        return f"{base[:n]}...{base[-n:]}"
