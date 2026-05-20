"""Sidebar — two sections (CDQ / DCT) with nav buttons and footer."""
from __future__ import annotations

from PySide6.QtCore import Qt, Signal
from PySide6.QtWidgets import (
    QFrame, QLabel, QPushButton, QVBoxLayout, QWidget,
)

from ui.theme import ACCENT, BORDER, PANEL, TEXT, TEXT_DIM, TEXT_MUTED


class _SectionLabel(QLabel):
    def __init__(self, text: str, parent: QWidget | None = None):
        super().__init__(text, parent)
        self.setObjectName("muted")
        self.setContentsMargins(4, 0, 0, 0)


class NavButton(QPushButton):
    def __init__(self, label: str, parent: QWidget | None = None):
        super().__init__(label, parent)
        self.setProperty("nav", True)
        self.setProperty("active", False)
        self.setCheckable(True)
        self.setCursor(Qt.PointingHandCursor)
        self.setMinimumHeight(36)

    def set_active(self, active: bool) -> None:
        self.setProperty("active", active)
        self.style().unpolish(self)
        self.style().polish(self)


class Sidebar(QWidget):
    # emits the stack index (0-3) corresponding to the selected stage
    stage_selected = Signal(int)
    settings_clicked = Signal()
    backups_clicked = Signal()

    # Stack layout:
    # 0 = CDQ · Reconciliation
    # 1 = CDQ · ERP Merger
    # 2 = DCT · Reconciliation
    # 3 = DCT · Rate Tool

    def __init__(self, parent: QWidget | None = None):
        super().__init__(parent)
        self.setObjectName("sidebar")
        self.setFixedWidth(240)

        layout = QVBoxLayout(self)
        layout.setContentsMargins(16, 20, 16, 16)
        layout.setSpacing(4)

        # Brand
        brand = QLabel("NPCMode")
        brand.setObjectName("brand")
        sub = QLabel("RECONCILIATION SUITE")
        sub.setObjectName("brandSub")
        layout.addWidget(brand)
        layout.addWidget(sub)
        layout.addSpacing(20)

        # ── CDQ section ───────────────────────────────────────────────────────
        layout.addWidget(_SectionLabel("CDQ"))
        layout.addSpacing(2)

        self._cdq1 = NavButton("  1  ·  Reconciliation")
        self._cdq2 = NavButton("  2  ·  ERP Merger")
        self._cdq1.clicked.connect(lambda: self._select(0))
        self._cdq2.clicked.connect(lambda: self._select(1))
        layout.addWidget(self._cdq1)
        layout.addWidget(self._cdq2)

        layout.addSpacing(14)

        # ── DCT section ───────────────────────────────────────────────────────
        layout.addWidget(_SectionLabel("DCT"))
        layout.addSpacing(2)

        self._dct1 = NavButton("  1  ·  Reconciliation")
        self._dct2 = NavButton("  2  ·  Rate Tool")
        self._dct1.clicked.connect(lambda: self._select(2))
        self._dct2.clicked.connect(lambda: self._select(3))
        layout.addWidget(self._dct1)
        layout.addWidget(self._dct2)

        layout.addStretch()

        # Footer
        sep = QFrame()
        sep.setFrameShape(QFrame.HLine)
        sep.setStyleSheet(f"color: {BORDER}; background: {BORDER}; max-height: 1px;")
        layout.addWidget(sep)
        layout.addSpacing(4)

        backups_btn = NavButton("  ◷   Backups")
        backups_btn.setCheckable(False)
        backups_btn.clicked.connect(self.backups_clicked.emit)
        layout.addWidget(backups_btn)

        settings_btn = NavButton("  ⚙   Settings")
        settings_btn.setCheckable(False)
        settings_btn.clicked.connect(self.settings_clicked.emit)
        layout.addWidget(settings_btn)

        version = QLabel("v0.1.0")
        version.setObjectName("muted")
        version.setAlignment(Qt.AlignCenter)
        layout.addSpacing(4)
        layout.addWidget(version)

        self._all_btns = [self._cdq1, self._cdq2, self._dct1, self._dct2]
        self._select(0)

    def _select(self, idx: int) -> None:
        for i, btn in enumerate(self._all_btns):
            btn.set_active(i == idx)
            btn.setChecked(i == idx)
        self.stage_selected.emit(idx)
