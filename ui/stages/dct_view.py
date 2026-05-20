"""Stage 3 — DCT view: Statement × TRX → TRX with rates."""
from __future__ import annotations

from PySide6.QtCore import Qt
from PySide6.QtWidgets import (
    QLabel, QPushButton, QScrollArea, QVBoxLayout, QWidget,
)

from core.dct import DCTResult, run_dct
from services.backup import archive_output, prune_old_backups
from services.settings import get_output_dir, set_output_dir
from services.worker import start_worker

from ..widgets.file_card import FileCard
from ..widgets.folder_card import FolderCard
from ..widgets.progress_card import ProgressCard
from ..widgets.result_card import ResultCard


STAGE_KEY = "stage_3"


class DCTView(QWidget):
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

        title = QLabel("Stage 3 · DCT")
        title.setObjectName("h1")
        sub = QLabel("Statement (SA sheets) × TRX  →  TRX with rates + Summary by merchant")
        sub.setObjectName("hint")
        layout.addWidget(title)
        layout.addWidget(sub)
        layout.addSpacing(8)

        self._stmt = FileCard(
            "Statement file",
            hint="XLSX with SA-EUR_* / SA-USD_* sheets",
            extensions=[".xlsx"],
        )
        self._trx = FileCard(
            "TRX file",
            hint="XLSX with Merchant path, Currency, Shipment date…",
            extensions=[".xlsx"],
        )
        self._recon = FileCard(
            "DCT Stage 1 output (optional)",
            hint="Итоговый файл сверки — добавит Payment ID по ARN",
            extensions=[".xlsx"],
        )
        self._output = FolderCard(
            "Output folder",
            initial=str(get_output_dir(STAGE_KEY)),
        )
        self._output.folder_changed.connect(
            lambda p: set_output_dir(STAGE_KEY, p))

        layout.addWidget(self._stmt)
        layout.addWidget(self._trx)
        layout.addWidget(self._recon)
        layout.addWidget(self._output)

        self._run_btn = QPushButton("▶  Build DCT file")
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

        self._stmt.file_selected.connect(lambda *_: self._refresh_enabled())
        self._stmt.cleared.connect(self._refresh_enabled)
        self._trx.file_selected.connect(lambda *_: self._refresh_enabled())
        self._trx.cleared.connect(self._refresh_enabled)

        self._thread = None
        self._worker = None

    def _refresh_enabled(self) -> None:
        ready = bool(self._stmt.path() and self._trx.path() and self._output.path())
        self._run_btn.setEnabled(ready)

    def _on_run(self) -> None:
        self._run_btn.setEnabled(False)
        self._result.hide()
        self._progress.show()
        self._progress.reset()
        self._progress.set_status("Starting…", "info")

        self._thread, self._worker = start_worker(
            self,
            run_dct,
            statement_path=self._stmt.path(),
            trx_path=self._trx.path(),
            recon_path=self._recon.path() or None,
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
        if "extracting rates" in text:
            self._progress.set_progress(15)
        elif "rate keys" in text:
            self._progress.set_progress(35)
        elif "reading trx" in text:
            self._progress.set_progress(55)
        elif "building" in text:
            self._progress.set_progress(75)
        elif "matched" in text:
            self._progress.set_progress(92)

    def _on_finished(self, result: DCTResult) -> None:
        self._progress.set_progress(100)
        self._progress.set_status("Done", "ok")

        archived = archive_output(result.out_path, STAGE_KEY)
        prune_old_backups()
        if archived:
            self._progress.append_log(f"Archived → {archived}")

        self._result.show_result(
            stats=[
                ("TRX rows", str(result.total_rows)),
                ("Matched", str(result.matched_rows)),
                ("Rate keys", str(result.rates_count)),
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
