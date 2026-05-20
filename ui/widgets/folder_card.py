"""Output folder picker — compact row."""
from __future__ import annotations

from pathlib import Path

from PySide6.QtCore import Qt, Signal
from PySide6.QtWidgets import (
    QFileDialog, QFrame, QHBoxLayout, QLabel, QPushButton, QWidget,
)


class FolderCard(QFrame):
    folder_changed = Signal(str)

    def __init__(self, label: str, initial: str = "", parent: QWidget | None = None):
        super().__init__(parent)
        self.setObjectName("card")
        self.setProperty("filled", bool(initial))

        row = QHBoxLayout(self)
        row.setContentsMargins(16, 10, 16, 10)
        row.setSpacing(10)

        self._label = QLabel(label)
        self._label.setObjectName("cardTitle")
        self._path_lbl = QLabel(initial or "—")
        self._path_lbl.setObjectName("fileName")
        self._path_lbl.setTextInteractionFlags(Qt.TextSelectableByMouse)

        self._btn = QPushButton("Choose folder")
        self._btn.setObjectName("secondary")
        self._btn.setCursor(Qt.PointingHandCursor)
        self._btn.clicked.connect(self._on_pick)

        row.addWidget(self._label)
        row.addWidget(self._path_lbl, 1)
        row.addWidget(self._btn)

        self._path = initial

    def _on_pick(self) -> None:
        path = QFileDialog.getExistingDirectory(self, "Select folder", self._path or "")
        if path:
            self.set_path(path)

    def set_path(self, path: str) -> None:
        self._path = path
        self._path_lbl.setText(self._shorten(path))
        self.setProperty("filled", True)
        self.style().unpolish(self)
        self.style().polish(self)
        self.folder_changed.emit(path)

    def path(self) -> str:
        return self._path

    @staticmethod
    def _shorten(path: str, max_len: int = 60) -> str:
        if len(path) <= max_len:
            return path
        p = Path(path)
        parts = p.parts
        if len(parts) <= 2:
            return path
        return f"{parts[0]}/…/{parts[-2]}/{parts[-1]}"
