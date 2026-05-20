"""Drag-and-drop file picker card (single file)."""
from __future__ import annotations

from pathlib import Path

from PySide6.QtCore import Qt, Signal
from PySide6.QtWidgets import (
    QFileDialog, QFrame, QHBoxLayout, QLabel, QPushButton,
    QVBoxLayout, QWidget,
)


class FileCard(QFrame):
    file_selected = Signal(str)
    cleared = Signal()

    def __init__(self, title: str, *, hint: str = "Drop file or click Browse",
                 extensions: list[str] | None = None,
                 parent: QWidget | None = None):
        super().__init__(parent)
        self.setObjectName("card")
        self.setProperty("filled", False)
        self.setProperty("drop", False)
        self.setAcceptDrops(True)
        self.setMinimumHeight(86)

        self._extensions = extensions or []
        self._path: str | None = None

        wrap = QVBoxLayout(self)
        wrap.setContentsMargins(16, 12, 16, 12)
        wrap.setSpacing(4)

        self._title = QLabel(title)
        self._title.setObjectName("cardTitle")

        self._subtitle = QLabel(hint)
        self._subtitle.setObjectName("cardSubtitle")

        self._row = QHBoxLayout()
        self._row.setContentsMargins(0, 0, 0, 0)
        self._row.setSpacing(10)
        self._name_lbl = QLabel("")
        self._name_lbl.setObjectName("fileName")
        self._meta_lbl = QLabel("")
        self._meta_lbl.setObjectName("fileMeta")
        self._row.addWidget(self._name_lbl, 1)
        self._row.addWidget(self._meta_lbl)

        btns = QHBoxLayout()
        btns.setContentsMargins(0, 4, 0, 0)
        btns.setSpacing(8)
        self._browse_btn = QPushButton("Browse")
        self._browse_btn.setObjectName("secondary")
        self._browse_btn.setCursor(Qt.PointingHandCursor)
        self._browse_btn.clicked.connect(self._on_browse)
        self._clear_btn = QPushButton("✕ Remove")
        self._clear_btn.setObjectName("danger")
        self._clear_btn.setCursor(Qt.PointingHandCursor)
        self._clear_btn.clicked.connect(self._on_clear)
        self._clear_btn.hide()
        btns.addWidget(self._browse_btn)
        btns.addWidget(self._clear_btn)
        btns.addStretch()

        wrap.addWidget(self._title)
        wrap.addWidget(self._subtitle)
        wrap.addLayout(self._row)
        wrap.addLayout(btns)

    # ─── Drag & drop ─────────────────────────────────────────────────────────
    def dragEnterEvent(self, event):
        if event.mimeData().hasUrls():
            event.acceptProposedAction()
            self._set_drop(True)

    def dragLeaveEvent(self, event):
        self._set_drop(False)

    def dropEvent(self, event):
        self._set_drop(False)
        urls = event.mimeData().urls()
        if not urls:
            return
        path = urls[0].toLocalFile()
        if self._is_allowed(path):
            self._set_path(path)

    def _set_drop(self, on: bool) -> None:
        self.setProperty("drop", on)
        self.style().unpolish(self)
        self.style().polish(self)

    # ─── State ───────────────────────────────────────────────────────────────
    def _on_browse(self) -> None:
        filt = self._build_filter()
        path, _ = QFileDialog.getOpenFileName(self, self._title.text(), "", filt)
        if path:
            self._set_path(path)

    def _on_clear(self) -> None:
        self._path = None
        self._name_lbl.setText("")
        self._meta_lbl.setText("")
        self._subtitle.show()
        self._clear_btn.hide()
        self.setProperty("filled", False)
        self.style().unpolish(self)
        self.style().polish(self)
        self.cleared.emit()

    def _set_path(self, path: str) -> None:
        self._path = path
        p = Path(path)
        self._name_lbl.setText(p.name)
        try:
            size_kb = p.stat().st_size / 1024
            self._meta_lbl.setText(
                f"{size_kb:.0f} KB" if size_kb < 1024 else f"{size_kb/1024:.1f} MB")
        except OSError:
            self._meta_lbl.setText("")
        self._subtitle.hide()
        self._clear_btn.show()
        self.setProperty("filled", True)
        self.style().unpolish(self)
        self.style().polish(self)
        self.file_selected.emit(path)

    def _is_allowed(self, path: str) -> bool:
        if not self._extensions:
            return True
        return any(path.lower().endswith(e.lower()) for e in self._extensions)

    def _build_filter(self) -> str:
        if not self._extensions:
            return "All files (*)"
        patterns = " ".join(f"*{e}" for e in self._extensions)
        return f"Files ({patterns});;All files (*)"

    def path(self) -> str | None:
        return self._path
