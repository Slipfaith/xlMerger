import os
from typing import Dict, List

from utils.logger import logger


class ExcelFilesManager:
    """Tracks selected Excel files and their relationships."""

    def __init__(self):
        self.files: List[Dict[str, str]] = []
        self.base_folder: str | None = None

    def add_files(self, files: List[str]):
        if not files:
            return
        if self.base_folder is None:
            self.base_folder = os.path.dirname(files[0])
        for path in files:
            if not os.path.isfile(path) or not path.lower().endswith((".xlsx", ".xls")):
                continue
            if any(f["path"] == path for f in self.files):
                continue
            rel = self._relative_path(path)
            self.files.append({"path": path, "rel": rel})
            logger.debug("Added file %s (rel %s)", path, rel)

    def add_folder(self, folder: str):
        if not os.path.isdir(folder):
            logger.warning("Folder does not exist: %s", folder)
            return
        if self.base_folder is None:
            self.base_folder = folder
        for root, _, filenames in os.walk(folder):
            for name in filenames:
                if not name.lower().endswith((".xlsx", ".xls")):
                    continue
                path = os.path.join(root, name)
                if any(f["path"] == path for f in self.files):
                    continue
                rel = os.path.relpath(path, folder)
                self.files.append({"path": path, "rel": rel})
                logger.debug("Added file %s (rel %s)", path, rel)

    def remove_indices(self, rows: List[int]):
        for row in sorted(rows, reverse=True):
            if 0 <= row < len(self.files):
                removed = self.files.pop(row)
                logger.debug("Removed file %s", removed["path"])
        if not self.files:
            self.base_folder = None

    def reset(self):
        self.files.clear()
        self.base_folder = None

    def _relative_path(self, path: str) -> str:
        if self.base_folder and os.path.commonpath([self.base_folder, path]) == self.base_folder:
            return os.path.relpath(path, self.base_folder)
        return os.path.basename(path)

    def build_output_root(self) -> str:
        if self.base_folder:
            parent = os.path.dirname(self.base_folder)
            name = os.path.basename(self.base_folder)
            return os.path.join(parent, f"{name}_upd")
        first_file_dir = os.path.dirname(self.files[0]["path"])
        first_name = os.path.splitext(os.path.basename(self.files[0]["path"]))[0]
        return os.path.join(first_file_dir, f"{first_name}_upd")
