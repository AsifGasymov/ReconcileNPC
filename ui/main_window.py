"""Main application window — sidebar + stacked stage views."""
from __future__ import annotations

from PySide6.QtCore import Qt
from PySide6.QtWidgets import (
    QHBoxLayout, QMainWindow, QStackedWidget, QWidget,
)

from .stages.recon_view import ReconView
from .stages.erp_view import ERPView
from .stages.decta_recon_view import DectaReconView
from .stages.dct_view import DCTView
from .settings_view import SettingsView
from .widgets.sidebar import Sidebar

# Stack indices
CDQ_RECON  = 0
CDQ_ERP    = 1
DCT_RECON  = 2
DCT_RATE   = 3
SETTINGS   = 4


class MainWindow(QMainWindow):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("NPCMode — Reconciliation Suite")
        self.setMinimumSize(1100, 720)
        self.resize(1240, 800)

        root = QWidget()
        root.setObjectName("root")
        layout = QHBoxLayout(root)
        layout.setContentsMargins(0, 0, 0, 0)
        layout.setSpacing(0)

        self._sidebar = Sidebar()
        layout.addWidget(self._sidebar)

        self._stack = QStackedWidget()
        self._stack.addWidget(ReconView())       # 0 CDQ Reconciliation
        self._stack.addWidget(ERPView())         # 1 CDQ ERP Merger
        self._stack.addWidget(DectaReconView())  # 2 DCT Reconciliation
        self._stack.addWidget(DCTView())         # 3 DCT Rate Tool
        self._stack.addWidget(SettingsView())    # 4 Settings
        layout.addWidget(self._stack, 1)

        self.setCentralWidget(root)

        self._sidebar.stage_selected.connect(self._stack.setCurrentIndex)
        self._sidebar.settings_clicked.connect(
            lambda: self._stack.setCurrentIndex(SETTINGS))
        self._sidebar.backups_clicked.connect(self._open_backups)

        self._center_on_screen()

    def _center_on_screen(self) -> None:
        screen = self.screen().availableGeometry()
        frame = self.frameGeometry()
        frame.moveCenter(screen.center())
        self.move(frame.topLeft())

    def _open_backups(self) -> None:
        from services.settings import get_backup_dir
        from PySide6.QtGui import QDesktopServices
        from PySide6.QtCore import QUrl
        path = get_backup_dir()
        path.mkdir(parents=True, exist_ok=True)
        QDesktopServices.openUrl(QUrl.fromLocalFile(str(path)))
