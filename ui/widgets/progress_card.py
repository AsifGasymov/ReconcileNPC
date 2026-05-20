"""Progress card with bar, status, and scrolling log stream."""
from __future__ import annotations

from PySide6.QtCore import Qt
from PySide6.QtWidgets import (
    QFrame, QHBoxLayout, QLabel, QPlainTextEdit, QProgressBar,
    QVBoxLayout, QWidget,
)


class ProgressCard(QFrame):
    def __init__(self, parent: QWidget | None = None):
        super().__init__(parent)
        self.setObjectName("card")

        wrap = QVBoxLayout(self)
        wrap.setContentsMargins(16, 14, 16, 14)
        wrap.setSpacing(10)

        head = QHBoxLayout()
        head.setSpacing(10)
        self._title = QLabel("Processing")
        self._title.setObjectName("cardTitle")
        self._pct = QLabel("0%")
        self._pct.setObjectName("hint")
        head.addWidget(self._title)
        head.addStretch()
        head.addWidget(self._pct)

        self._bar = QProgressBar()
        self._bar.setRange(0, 100)
        self._bar.setValue(0)
        self._bar.setTextVisible(False)
        self._bar.setFixedHeight(8)

        self._status = QLabel("—")
        self._status.setObjectName("hint")

        self._log = QPlainTextEdit()
        self._log.setObjectName("logStream")
        self._log.setReadOnly(True)
        self._log.setFixedHeight(140)
        self._log.setVerticalScrollBarPolicy(Qt.ScrollBarAsNeeded)

        wrap.addLayout(head)
        wrap.addWidget(self._bar)
        wrap.addWidget(self._status)
        wrap.addWidget(self._log)

    def reset(self) -> None:
        self._bar.setValue(0)
        self._pct.setText("0%")
        self._status.setText("Starting…")
        self._status.setObjectName("hint")
        self._log.clear()
        self._restyle(self._status)

    def set_progress(self, pct: int) -> None:
        pct = max(0, min(100, int(pct)))
        self._bar.setValue(pct)
        self._pct.setText(f"{pct}%")

    def set_status(self, text: str, kind: str = "info") -> None:
        self._status.setText(text)
        self._status.setObjectName({
            "info": "hint", "ok": "statusOk",
            "warn": "statusWarn", "err": "statusErr",
        }.get(kind, "hint"))
        self._restyle(self._status)

    def append_log(self, line: str) -> None:
        self._log.appendPlainText(line)
        self._log.verticalScrollBar().setValue(
            self._log.verticalScrollBar().maximum())

    @staticmethod
    def _restyle(w: QWidget) -> None:
        w.style().unpolish(w)
        w.style().polish(w)
