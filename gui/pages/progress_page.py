# gui/pages/progress_page.py

from PySide6.QtWidgets import QWidget, QVBoxLayout, QProgressBar
from PySide6.QtCore import Qt

class ProgressPage(QWidget):
    def __init__(self, parent=None):
        super().__init__(parent)
        self.setWindowTitle("Копирование переводов")
        layout = QVBoxLayout(self)
        self.progress_bar = QProgressBar(self)
        self.progress_bar.setAlignment(Qt.AlignCenter)
        self.progress_bar.setStyleSheet("""
            QProgressBar { text-align: center; }
            QProgressBar::chunk { background-color: #f47929; }
        """)
        layout.addWidget(self.progress_bar)
        self.setLayout(layout)

    def set_progress(self, value, maximum=None):
        if maximum is not None:
            self.progress_bar.setMaximum(maximum)
        self.progress_bar.setValue(value)

    def get_progressbar(self):
        return self.progress_bar
