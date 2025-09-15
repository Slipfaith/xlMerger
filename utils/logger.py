import logging
import datetime
import os

logger = logging.getLogger("app_logger")
logger.setLevel(logging.DEBUG)
if not logger.hasHandlers():
    handler = logging.StreamHandler()
    formatter = logging.Formatter('[%(asctime)s] %(message)s', "%Y-%m-%d %H:%M:%S")
    handler.setFormatter(formatter)
    logger.addHandler(handler)

class Logger:
    def __init__(self, log_file="copy_log.txt"):
        self.log_file = log_file
        self.entries = []

    def log(self, message, level=logging.INFO):
        timestamp = datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        entry = f"[{timestamp}] {message}"
        self.entries.append(entry)
        logger.log(level, message)  # Дублируем запись в стандартный логгер

    def log_copy(self, sheet, row, col, value):
        msg = f"COPY: {sheet} R{row}C{col} -> {repr(value)}"
        self.log(msg, logging.INFO)

    def log_error(self, sheet, row, col, value):
        msg = f"ERROR: {sheet} R{row}C{col} -> {repr(value)}"
        self.log(msg, logging.ERROR)

    def log_info(self, text):
        self.log(f"INFO: {text}", logging.INFO)

    def log_warning(self, text):
        self.log(f"WARNING: {text}", logging.WARNING)

    def save(self):
        if not self.entries:
            return
        os.makedirs(os.path.dirname(self.log_file) or ".", exist_ok=True)
        with open(self.log_file, "a", encoding="utf-8") as f:
            for entry in self.entries:
                f.write(entry + "\n")
        self.entries.clear()
