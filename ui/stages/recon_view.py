"""Stage 1 — Reconciliation view."""
from __future__ import annotations

from PySide6.QtCore import Qt
from PySide6.QtWidgets import (
    QHBoxLayout, QLabel, QPushButton, QScrollArea, QVBoxLayout, QWidget,
)

from core.reconciliation import ReconResult, run_reconciliation
from services.backup import archive_output, prune_old_backups
from services.settings import get_output_dir, set_output_dir
from services.worker import start_worker

from ..widgets.file_card import FileCard
from ..widgets.folder_card import FolderCard
from ..widgets.multi_file_card import MultiFileCard
from ..widgets.progress_card import ProgressCard
from ..widgets.result_card import ResultCard


STAGE_KEY = "stage_1"


class ReconView(QWidget):
    def __init__(self, parent: QWidget | None = None):
        super().__init__(parent)

        scroll = QScrollArea(self)
        scroll.setWidgetResizable(True)
        scroll.setFrameShape(QScrollArea.NoFrame)

        outer_wrap = QVBoxLayout(self)
        outer_wrap.setContentsMargins(0, 0, 0, 0)
        outer_wrap.addWidget(scroll)

        content = QWidget()
        scroll.setWidget(content)

        layout = QVBoxLayout(content)
        layout.setContentsMargins(40, 32, 40, 32)
        layout.setSpacing(16)

        # Header
        title = QLabel("Stage 1 · Reconciliation")
        title.setObjectName("h1")
        sub = QLabel("Grafana × CDQ × General  →  unified SPNT/CDQ workbook")
        sub.setObjectName("hint")
        layout.addWidget(title)
        layout.addWidget(sub)
        layout.addSpacing(8)

        # Inputs
        self._grafana = FileCard(
            "Grafana file",
            hint="Drop a CSV or XLSX export here",
            extensions=[".csv", ".xlsx"],
        )
        self._cdq = MultiFileCard(
            "CDQ files",
            hint="Drop up to 9 CSV/XLSX/XLS files",
            extensions=[".csv", ".xlsx", ".xls"],
            max_items=9,
        )
        self._general = FileCard(
            "General · exclusions",
            hint="XLSX with Payment IDs to exclude",
            extensions=[".xlsx"],
        )
        self._output = FolderCard(
            "Output folder",
            initial=str(get_output_dir(STAGE_KEY)),
        )
        self._output.folder_changed.connect(
            lambda p: set_output_dir(STAGE_KEY, p))

        layout.addWidget(self._grafana)
        layout.addWidget(self._cdq)
        layout.addWidget(self._general)
        layout.addWidget(self._output)

        # Run button
        self._run_btn = QPushButton("▶  Run reconciliation")
        self._run_btn.setObjectName("primary")
        self._run_btn.setCursor(Qt.PointingHandCursor)
        self._run_btn.setMinimumHeight(48)
        self._run_btn.setEnabled(False)
        self._run_btn.clicked.connect(self._on_run)
        layout.addSpacing(8)
        layout.addWidget(self._run_btn)

        # Progress + result
        self._progress = ProgressCard()
        self._progress.hide()
        layout.addWidget(self._progress)

        self._result = ResultCard()
        self._result.hide()
        self._result.rerun_requested.connect(self._on_rerun)
        layout.addWidget(self._result)

        layout.addStretch()

        # Wire enable state
        self._grafana.file_selected.connect(lambda *_: self._refresh_enabled())
        self._grafana.cleared.connect(self._refresh_enabled)
        self._cdq.paths_changed.connect(lambda *_: self._refresh_enabled())
        self._general.file_selected.connect(lambda *_: self._refresh_enabled())
        self._general.cleared.connect(self._refresh_enabled)

        self._thread = None
        self._worker = None

    def _refresh_enabled(self) -> None:
        ready = bool(
            self._grafana.path()
            and self._cdq.paths()
            and self._general.path()
            and self._output.path()
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
            run_reconciliation,
            grafana_path=self._grafana.path(),
            cdq_paths=self._cdq.paths(),
            general_path=self._general.path(),
            out_dir=self._output.path(),
        )
        self._worker.log.connect(self._on_log)
        self._worker.finished.connect(self._on_finished)
        self._worker.failed.connect(self._on_failed)
        self._thread.start()

    def _on_log(self, msg: str) -> None:
        self._progress.append_log(msg)
        self._progress.set_status(msg, "info")
        # crude progress estimate by step keywords
        text = msg.lower()
        if "reading grafana" in text:
            self._progress.set_progress(10)
        elif "merging cdq" in text:
            self._progress.set_progress(30)
        elif "general" in text:
            self._progress.set_progress(45)
        elif "filtering" in text:
            self._progress.set_progress(60)
        elif "joining" in text:
            self._progress.set_progress(75)
        elif "writing" in text:
            self._progress.set_progress(90)

    def _on_finished(self, result: ReconResult) -> None:
        self._progress.set_progress(100)
        self._progress.set_status("Done", "ok")

        archived = archive_output(result.out_path, STAGE_KEY)
        prune_old_backups()
        if archived:
            self._progress.append_log(f"Archived → {archived}")

        self._result.show_result(
            stats=[
                ("SPNT rows", str(result.sheet1_rows)),
                ("CDQ rows", str(result.sheet2_rows)),
                ("Excluded", str(result.excluded)),
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
