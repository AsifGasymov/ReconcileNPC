"""Settings view — backup folder, retention, info."""
from __future__ import annotations

from PySide6.QtCore import Qt
from PySide6.QtWidgets import (
    QFileDialog, QFrame, QHBoxLayout, QLabel, QPushButton, QSpinBox,
    QVBoxLayout, QWidget,
)

from services.settings import (
    get_backup_dir, get_backup_retention_days,
    set_backup_dir, set_backup_retention_days,
)


class SettingsView(QWidget):
    def __init__(self, parent: QWidget | None = None):
        super().__init__(parent)

        outer = QVBoxLayout(self)
        outer.setContentsMargins(40, 32, 40, 32)
        outer.setSpacing(20)

        title = QLabel("Settings")
        title.setObjectName("h1")
        outer.addWidget(title)

        sub = QLabel("Backup configuration and global preferences.")
        sub.setObjectName("hint")
        outer.addWidget(sub)
        outer.addSpacing(12)

        # ─── Backup folder card ──────────────────────────────────────────────
        backup_card = QFrame()
        backup_card.setObjectName("card")
        bl = QVBoxLayout(backup_card)
        bl.setContentsMargins(20, 16, 20, 16)
        bl.setSpacing(8)

        h2 = QLabel("Backup folder")
        h2.setObjectName("cardTitle")
        sub2 = QLabel("Output files are archived here under YYYY-MM-DD/stage_N/.")
        sub2.setObjectName("cardSubtitle")
        bl.addWidget(h2)
        bl.addWidget(sub2)

        path_row = QHBoxLayout()
        path_row.setSpacing(10)
        self._path_lbl = QLabel(str(get_backup_dir()))
        self._path_lbl.setObjectName("fileName")
        self._path_lbl.setTextInteractionFlags(Qt.TextSelectableByMouse)
        change_btn = QPushButton("Change")
        change_btn.setObjectName("secondary")
        change_btn.setCursor(Qt.PointingHandCursor)
        change_btn.clicked.connect(self._on_change_backup)
        path_row.addWidget(self._path_lbl, 1)
        path_row.addWidget(change_btn)
        bl.addLayout(path_row)

        outer.addWidget(backup_card)

        # ─── Retention card ──────────────────────────────────────────────────
        retention_card = QFrame()
        retention_card.setObjectName("card")
        rl = QVBoxLayout(retention_card)
        rl.setContentsMargins(20, 16, 20, 16)
        rl.setSpacing(8)

        h3 = QLabel("Retention")
        h3.setObjectName("cardTitle")
        sub3 = QLabel("Backups older than this are removed automatically.")
        sub3.setObjectName("cardSubtitle")
        rl.addWidget(h3)
        rl.addWidget(sub3)

        ret_row = QHBoxLayout()
        ret_row.setSpacing(10)
        self._spin = QSpinBox()
        self._spin.setMinimum(1)
        self._spin.setMaximum(365)
        self._spin.setValue(get_backup_retention_days())
        self._spin.setSuffix(" days")
        self._spin.setFixedWidth(120)
        self._spin.valueChanged.connect(set_backup_retention_days)
        ret_row.addWidget(self._spin)
        ret_row.addStretch()
        rl.addLayout(ret_row)

        outer.addWidget(retention_card)

        outer.addStretch()

    def _on_change_backup(self) -> None:
        path = QFileDialog.getExistingDirectory(
            self, "Select backup folder", str(get_backup_dir()))
        if path:
            set_backup_dir(path)
            self._path_lbl.setText(path)
