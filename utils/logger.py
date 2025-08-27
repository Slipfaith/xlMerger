"""Simple logging helper used across the project.

The original implementation only wrote to a file and duplicated messages to the
root logger as errors.  For debugging from the command line we want the output
to appear immediately and we also need a convenient way to log arbitrary error
messages.  This module now prints log messages to the console using the
``app_logger`` logger and allows ``log_error`` to receive a message first which
matches how it is used throughout the codebase.
"""

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

    def log(self, message):
        timestamp = datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        entry = f"[{timestamp}] {message}"
        self.entries.append(entry)
        # also output to the console immediately
        logger.info(message)

    def log_copy(self, sheet, row, col, value):
        msg = f"COPY: {sheet} R{row}C{col} -> {repr(value)}"
        self.log(msg)

    def log_error(self, message, sheet="", row="", col="", value=""):
        """Log an error message with optional cell context.

        Most call sites simply pass a text description.  Some also provide
        additional information such as sheet name, row/column or the value that
        caused the issue.  To support both use‑cases the arguments are optional
        and the human readable message is placed first.
        """
        location = f" {sheet} R{row}C{col} -> {repr(value)}" if any(
            [sheet, row, col, value]
        ) else ""
        self.log(f"ERROR: {message}{location}")

    def log_info(self, text):
        self.log(f"INFO: {text}")

    def log_warning(self, text):
        self.log(f"WARNING: {text}")

    def save(self):
        if not self.entries:
            return
        os.makedirs(os.path.dirname(self.log_file) or ".", exist_ok=True)
        with open(self.log_file, "a", encoding="utf-8") as f:
            for entry in self.entries:
                f.write(entry + "\n")
        self.entries.clear()
