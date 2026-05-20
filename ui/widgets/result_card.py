"""Result card displayed after successful run."""
from __future__ import annotations

import subprocess
import sys
from pathlib import Path

from PySide6.QtCore import Qt, Signal
from PySide6.QtWidgets import (
    QFrame, QGridLayout, QHBoxLayout, QLabel, QPushButton, QVBoxLayout, QWidget,
)


class ResultCard(QFrame):
    rerun_requested = Signal()

    def __init__(self, parent: QWidget | None = None):
        super().__init__(parent)
        self.setObjectName("resultCard")

        wrap = QVBoxLayout(self)
        wrap.setContentsMargins(20, 18, 20, 18)
        wrap.setSpacing(12)

        head = QHBoxLayout()
        title = QLabel("✓  Done")
        title.setObjectName("resultTitle")
        head.addWidget(title)
        head.addStretch()
        wrap.addLayout(head)

        self._stats_grid = QGridLayout()
        self._stats_grid.setHorizontalSpacing(28)
        self._stats_grid.setVerticalSpacing(2)
        wrap.addLayout(self._stats_grid)

        self._path_lbl = QLabel("")
        self._path_lbl.setObjectName("hint")
        self._path_lbl.setTextInteractionFlags(Qt.TextSelectableByMouse)
        self._path_lbl.setWordWrap(True)
        wrap.addWidget(self._path_lbl)

        btns = QHBoxLayout()
        btns.setSpacing(8)
        self._open_btn = QPushButton("Open file")
        self._open_btn.setObjectName("primary")
        self._open_btn.setCursor(Qt.PointingHandCursor)
        self._open_btn.clicked.connect(self._on_open)
        self._reveal_btn = QPushButton("Show in folder")
        self._reveal_btn.setObjectName("secondary")
        self._reveal_btn.setCursor(Qt.PointingHandCursor)
        self._reveal_btn.clicked.connect(self._on_reveal)
        self._rerun_btn = QPushButton("New run")
        self._rerun_btn.setObjectName("secondary")
        self._rerun_btn.setCursor(Qt.PointingHandCursor)
        self._rerun_btn.clicked.connect(self.rerun_requested.emit)
        btns.addWidget(self._open_btn)
        btns.addWidget(self._reveal_btn)
        btns.addWidget(self._rerun_btn)
        btns.addStretch()
        wrap.addLayout(btns)

        self._out_path: str | None = None

    def show_result(self, *, stats: list[tuple[str, str]], out_path: str) -> None:
        # clear grid
        while self._stats_grid.count():
            child = self._stats_grid.takeAt(0)
            if child.widget():
                child.widget().deleteLater()

        for i, (label, value) in enumerate(stats):
            num = QLabel(value)
            num.setObjectName("resultStatNum")
            cap = QLabel(label)
            cap.setObjectName("muted")
            self._stats_grid.addWidget(num, 0, i, alignment=Qt.AlignLeft)
            self._stats_grid.addWidget(cap, 1, i, alignment=Qt.AlignLeft)

        self._out_path = out_path
        self._path_lbl.setText(out_path)

    def _on_open(self) -> None:
        if not self._out_path:
            return
        path = self._out_path
        if sys.platform == "darwin":
            subprocess.Popen(["open", path])
        elif sys.platform == "win32":
            subprocess.Popen(["start", "", path], shell=True)
        else:
            subprocess.Popen(["xdg-open", path])

    def _on_reveal(self) -> None:
        if not self._out_path:
            return
        path = self._out_path
        if sys.platform == "darwin":
            subprocess.Popen(["open", "-R", path])
        elif sys.platform == "win32":
            subprocess.Popen(["explorer", "/select,", path])
        else:
            subprocess.Popen(["xdg-open", str(Path(path).parent)])
