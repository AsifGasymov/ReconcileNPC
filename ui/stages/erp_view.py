"""Stage 2 — ERP Merger view."""
from __future__ import annotations

from PySide6.QtCore import Qt
from PySide6.QtWidgets import (
    QLabel, QPushButton, QScrollArea, QVBoxLayout, QWidget,
)

from core.erp_merger import ERPResult, run_erp_merger
from services.backup import archive_output, prune_old_backups
from services.settings import get_output_dir, set_output_dir
from services.worker import start_worker

from ..widgets.file_card import FileCard
from ..widgets.folder_card import FolderCard
from ..widgets.progress_card import ProgressCard
from ..widgets.result_card import ResultCard


STAGE_KEY = "stage_2"


class ERPView(QWidget):
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

        title = QLabel("Stage 2 · CDQ")
        title.setObjectName("h1")
        sub = QLabel("Transactions × Statement (FX) × Orders  →  ERP upload workbook")
        sub.setObjectName("hint")
        layout.addWidget(title)
        layout.addWidget(sub)
        layout.addSpacing(8)

        self._trn = FileCard(
            "Transaction report",
            hint="CSV or XLSX",
            extensions=[".csv", ".xlsx"],
        )
        self._stm = FileCard(
            "Statement",
            hint="XLSX with FX rates",
            extensions=[".xlsx"],
        )
        self._ord = FileCard(
            "Order Statement (optional)",
            hint="XLSX — used to map RRN → Payment ID",
            extensions=[".xlsx"],
        )
        self._output = FolderCard(
            "Output folder",
            initial=str(get_output_dir(STAGE_KEY)),
        )
        self._output.folder_changed.connect(
            lambda p: set_output_dir(STAGE_KEY, p))

        layout.addWidget(self._trn)
        layout.addWidget(self._stm)
        layout.addWidget(self._ord)
        layout.addWidget(self._output)

        self._run_btn = QPushButton("▶  Build ERP file")
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

        self._trn.file_selected.connect(lambda *_: self._refresh_enabled())
        self._trn.cleared.connect(self._refresh_enabled)
        self._stm.file_selected.connect(lambda *_: self._refresh_enabled())
        self._stm.cleared.connect(self._refresh_enabled)

        self._thread = None
        self._worker = None

    def _refresh_enabled(self) -> None:
        ready = bool(
            self._trn.path() and self._stm.path() and self._output.path()
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
            run_erp_merger,
            transactions_path=self._trn.path(),
            statement_path=self._stm.path(),
            orders_path=self._ord.path() or None,
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
        if "reading statement" in text:
            self._progress.set_progress(15)
        elif "reading transactions" in text:
            self._progress.set_progress(35)
        elif "order statement" in text:
            self._progress.set_progress(55)
        elif "writing" in text:
            self._progress.set_progress(85)

    def _on_finished(self, result: ERPResult) -> None:
        self._progress.set_progress(100)
        self._progress.set_status("Done", "ok")

        archived = archive_output(result.out_path, STAGE_KEY)
        prune_old_backups()
        if archived:
            self._progress.append_log(f"Archived → {archived}")

        self._result.show_result(
            stats=[
                ("Rows", str(result.rows)),
                ("FX rates", str(result.rates_count)),
                ("Net GBP", f"{result.total_net_gbp:,.2f}"),
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
