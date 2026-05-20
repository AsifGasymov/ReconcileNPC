"""Multi-file picker card with list + add/remove."""
from __future__ import annotations

from pathlib import Path

from PySide6.QtCore import Qt, Signal
from PySide6.QtWidgets import (
    QFileDialog, QFrame, QHBoxLayout, QLabel, QListWidget, QListWidgetItem,
    QPushButton, QVBoxLayout, QWidget,
)


class MultiFileCard(QFrame):
    paths_changed = Signal(list)

    def __init__(self, title: str, *, hint: str = "Drop files or click Add",
                 extensions: list[str] | None = None, max_items: int = 9,
                 parent: QWidget | None = None):
        super().__init__(parent)
        self.setObjectName("card")
        self.setProperty("filled", False)
        self.setProperty("drop", False)
        self.setAcceptDrops(True)

        self._extensions = extensions or []
        self._max = max_items
        self._paths: list[str] = []

        wrap = QVBoxLayout(self)
        wrap.setContentsMargins(16, 12, 16, 12)
        wrap.setSpacing(6)

        head = QHBoxLayout()
        head.setSpacing(6)
        title_lbl = QLabel(title)
        title_lbl.setObjectName("cardTitle")
        self._counter = QLabel(f"0 / {max_items}")
        self._counter.setObjectName("muted")
        head.addWidget(title_lbl)
        head.addStretch()
        head.addWidget(self._counter)
        wrap.addLayout(head)

        self._subtitle = QLabel(hint)
        self._subtitle.setObjectName("cardSubtitle")
        wrap.addWidget(self._subtitle)

        self._list = QListWidget()
        self._list.setMaximumHeight(140)
        self._list.hide()
        wrap.addWidget(self._list)

        btns = QHBoxLayout()
        btns.setSpacing(8)
        btns.setContentsMargins(0, 6, 0, 0)
        self._add_btn = QPushButton("+ Add files")
        self._add_btn.setObjectName("secondary")
        self._add_btn.setCursor(Qt.PointingHandCursor)
        self._add_btn.clicked.connect(self._on_add)
        self._rm_btn = QPushButton("✕ Remove selected")
        self._rm_btn.setObjectName("danger")
        self._rm_btn.setCursor(Qt.PointingHandCursor)
        self._rm_btn.clicked.connect(self._on_remove)
        self._clear_btn = QPushButton("Clear all")
        self._clear_btn.setObjectName("ghost")
        self._clear_btn.setCursor(Qt.PointingHandCursor)
        self._clear_btn.clicked.connect(self._on_clear_all)
        btns.addWidget(self._add_btn)
        btns.addWidget(self._rm_btn)
        btns.addWidget(self._clear_btn)
        btns.addStretch()
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
        for url in event.mimeData().urls():
            path = url.toLocalFile()
            if path and self._is_allowed(path):
                self._add_path(path)

    def _set_drop(self, on: bool) -> None:
        self.setProperty("drop", on)
        self.style().unpolish(self)
        self.style().polish(self)

    # ─── Actions ─────────────────────────────────────────────────────────────
    def _on_add(self) -> None:
        filt = self._build_filter()
        paths, _ = QFileDialog.getOpenFileNames(self, "Select files", "", filt)
        for p in paths:
            self._add_path(p)

    def _on_remove(self) -> None:
        for item in self._list.selectedItems():
            row = self._list.row(item)
            self._list.takeItem(row)
            self._paths.pop(row)
        self._refresh()

    def _on_clear_all(self) -> None:
        self._paths.clear()
        self._list.clear()
        self._refresh()

    def _add_path(self, path: str) -> None:
        if path in self._paths or len(self._paths) >= self._max:
            return
        self._paths.append(path)
        item = QListWidgetItem(Path(path).name)
        item.setToolTip(path)
        self._list.addItem(item)
        self._refresh()

    def _refresh(self) -> None:
        n = len(self._paths)
        self._counter.setText(f"{n} / {self._max}")
        has = n > 0
        self._list.setVisible(has)
        self._subtitle.setVisible(not has)
        self.setProperty("filled", has)
        self.style().unpolish(self)
        self.style().polish(self)
        self.paths_changed.emit(list(self._paths))

    def _is_allowed(self, path: str) -> bool:
        if not self._extensions:
            return True
        return any(path.lower().endswith(e.lower()) for e in self._extensions)

    def _build_filter(self) -> str:
        if not self._extensions:
            return "All files (*)"
        patterns = " ".join(f"*{e}" for e in self._extensions)
        return f"Files ({patterns});;All files (*)"

    def paths(self) -> list[str]:
        return list(self._paths)
