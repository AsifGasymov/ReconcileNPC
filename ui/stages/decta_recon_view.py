"""DCT Stage 1 — Reconciliation view."""
from __future__ import annotations

from PySide6.QtCore import Qt
from PySide6.QtWidgets import (
    QLabel, QPushButton, QScrollArea, QVBoxLayout, QWidget,
)

from core.decta_recon import DectaReconResult, run_decta_recon
from services.backup import archive_output, prune_old_backups
from services.settings import get_output_dir, set_output_dir
from services.worker import start_worker

from ..widgets.file_card import FileCard
from ..widgets.folder_card import FolderCard
from ..widgets.multi_file_card import MultiFileCard
from ..widgets.progress_card import ProgressCard
from ..widgets.result_card import ResultCard


STAGE_KEY = "dct_stage_1"


class DectaReconView(QWidget):
    def __init__(self, parent: QWidget | None = None):
        super().__init__(parent)

        scroll = QScrollArea(self)
        scroll.setWidgetResizable(True)
        scroll.setFrameShape(QScrollArea.NoFrame)

        outer = QVBoxLayout(self)
        outer.setContentsMargins(0, 0, 0, 0)
        outer.addWidget(scroll)

        content = QWidget()
        scroll.setWidget(content)

        layout = QVBoxLayout(content)
        layout.setContentsMargins(40, 32, 40, 32)
        layout.setSpacing(16)

        title = QLabel("DCT · Stage 1 · Reconciliation")
        title.setObjectName("h1")
        sub = QLabel("Main × Registry × Providers  →  New Transactions + DCT workbook")
        sub.setObjectName("hint")
        layout.addWidget(title)
        layout.addWidget(sub)
        layout.addSpacing(8)

        self._main = FileCard(
            "Main file",
            hint="XLSX — existing transactions (reference)",
            extensions=[".xlsx"],
        )
        self._registry = FileCard(
            "Payment Registry",
            hint="XLSX, CSV or .numbers — new transactions to check",
            extensions=[".xlsx", ".csv", ".numbers"],
        )
        self._providers = MultiFileCard(
            "Provider files",
            hint="XLSX files — CDQ/provider exports (up to 9)",
            extensions=[".xlsx", ".xls", ".csv"],
            max_items=9,
        )
        self._output = FolderCard(
            "Output folder",
            initial=str(get_output_dir(STAGE_KEY)),
        )
        self._output.folder_changed.connect(
            lambda p: set_output_dir(STAGE_KEY, p))

        layout.addWidget(self._main)
        layout.addWidget(self._registry)
        layout.addWidget(self._providers)
        layout.addWidget(self._output)

        self._run_btn = QPushButton("▶  Run reconciliation")
        self._run_btn.setObjectName("primary")
        self._run_btn.setCursor(Qt.PointingHandCursor)
        self._run_btn.setMinimumHeight(48)
        self._run_btn.setEnabled(False)
        self._run_btn.clicked.connect(self._on_run)
        layout.addSpacing(8)
        layout.addWidget(self._run_btn)

        self._progress = ProgressCard()
        self._progress.hide()
        layout.addWidget(self._progress)

        self._result = ResultCard()
        self._result.hide()
        self._result.rerun_requested.connect(self._on_rerun)
        layout.addWidget(self._result)

        layout.addStretch()

        self._main.file_selected.connect(lambda *_: self._refresh_enabled())
        self._main.cleared.connect(self._refresh_enabled)
        self._registry.file_selected.connect(lambda *_: self._refresh_enabled())
        self._registry.cleared.connect(self._refresh_enabled)

        self._thread = None
        self._worker = None

    def _refresh_enabled(self) -> None:
        ready = bool(
            self._main.path() and self._registry.path() and self._output.path()
        )
        self._run_btn.setEnabled(ready)

    def _on_run(self) -> None:
        self._run_btn.setEnabled(False)
        self._result.hide()
        self._progress.show()
        self._progress.reset()
        self._progress.set_status("Starting…", "info")

        self._thread, self._worker = start_worker(
            self,
            run_decta_recon,
            main_path=self._main.path(),
            registry_path=self._registry.path(),
            provider_paths=self._providers.paths(),
            out_dir=self._output.path(),
        )
        self._worker.log.connect(self._on_log)
        self._worker.finished.connect(self._on_finished)
        self._worker.failed.connect(self._on_failed)
        self._thread.start()

    def _on_log(self, msg: str) -> None:
        self._progress.append_log(msg)
        self._progress.set_status(msg, "info")
        text = msg.lower()
        if "reading main" in text:
            self._progress.set_progress(10)
        elif "registry" in text:
            self._progress.set_progress(25)
        elif "filtering" in text:
            self._progress.set_progress(45)
        elif "provider" in text:
            self._progress.set_progress(60)
        elif "transforming" in text:
            self._progress.set_progress(80)
        elif "writing" in text:
            self._progress.set_progress(92)

    def _on_finished(self, result: DectaReconResult) -> None:
        self._progress.set_progress(100)
        self._progress.set_status("Done", "ok")

        archived = archive_output(result.out_path, STAGE_KEY)
        prune_old_backups()
        if archived:
            self._progress.append_log(f"Archived → {archived}")

        self._result.show_result(
            stats=[
                ("New rows", str(result.new_rows)),
                ("Skipped", str(result.skipped)),
                ("DCT rows", str(result.dct_rows)),
            ],
            out_path=result.out_path,
        )
        self._result.show()
        self._run_btn.setEnabled(True)

    def _on_failed(self, error: str) -> None:
        self._progress.set_status("Error", "err")
        self._progress.append_log(error)
        self._run_btn.setEnabled(True)

    def _on_rerun(self) -> None:
        self._result.hide()
        self._progress.hide()
