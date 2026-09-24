"""XMGate section view — merge exports and check pay ids against a duplicates list."""
from __future__ import annotations

from PySide6.QtCore import Qt
from PySide6.QtWidgets import (
    QLabel, QPushButton, QScrollArea, QVBoxLayout, QWidget,
)

from core.xmgate_dupes import XMGateResult, run_xmgate_dupes
from services.settings import get_output_dir, set_output_dir
from services.worker import start_worker

from ..widgets.file_card import FileCard
from ..widgets.folder_card import FolderCard
from ..widgets.multi_file_card import MultiFileCard
from ..widgets.progress_card import ProgressCard
from ..widgets.result_card import ResultCard

STAGE_KEY_XMGATE = "xmgate_duplicates"


class XMGateView(QWidget):
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

        title = QLabel("XMGate · Duplicates")
        title.setObjectName("h1")
        sub = QLabel("Gateway CSV exports × duplicates list  →  merged workbook with pay id check")
        sub.setObjectName("hint")
        layout.addWidget(title)
        layout.addWidget(sub)
        layout.addSpacing(8)

        self._exports = MultiFileCard(
            "Gateway exports",
            hint="Drop CSV exports (order_id, merchant_order_id, created, …)",
            extensions=[".csv"],
            max_items=20,
        )
        self._dupes = FileCard(
            "Duplicates list",
            hint="XLSX/CSV with a 'pay id' column",
            extensions=[".xlsx", ".csv"],
        )
        self._output = FolderCard(
            "Output folder",
            initial=str(get_output_dir(STAGE_KEY_XMGATE)),
        )
        self._output.folder_changed.connect(
            lambda p: set_output_dir(STAGE_KEY_XMGATE, p))

        layout.addWidget(self._exports)
        layout.addWidget(self._dupes)
        layout.addWidget(self._output)

        self._run_btn = QPushButton("▶  Merge & check duplicates")
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

        self._exports.paths_changed.connect(lambda *_: self._refresh_enabled())
        self._dupes.file_selected.connect(lambda *_: self._refresh_enabled())
        self._dupes.cleared.connect(lambda: self._refresh_enabled())

        self._thread = None
        self._worker = None

    def _refresh_enabled(self) -> None:
        self._run_btn.setEnabled(
            bool(self._exports.paths() and self._dupes.path() and self._output.path())
        )

    def _on_run(self) -> None:
        self._run_btn.setEnabled(False)
        self._result.hide()
        self._progress.show()
        self._progress.reset()
        self._progress.set_status("Starting…", "info")

        self._thread, self._worker = start_worker(
            self,
            run_xmgate_dupes,
            export_paths=self._exports.paths(),
            dupes_path=self._dupes.path(),
            out_dir=self._output.path(),
        )
        self._worker.log.connect(self._on_log)
        self._worker.finished.connect(self._on_finished)
        self._worker.failed.connect(self._on_failed)
        self._thread.start()

    def _on_log(self, msg: str) -> None:
        self._progress.append_log(msg)
        self._progress.set_status(msg, "info")
        low = msg.lower()
        if "gateway" in low:
            self._progress.set_progress(20)
        elif "duplicates list" in low:
            self._progress.set_progress(40)
        elif "checking" in low:
            self._progress.set_progress(60)
        elif "writing" in low:
            self._progress.set_progress(80)

    def _on_finished(self, result: XMGateResult) -> None:
        self._progress.set_progress(100)
        self._progress.set_status("Done", "ok")
        self._result.show_result(
            stats=[
                ("Rows merged",        str(result.total_rows)),
                ("Dup list ids",       str(result.dup_list_size)),
                ("Dup pay ids",        str(result.dup_pay_ids)),
                ("Dup rows",           str(result.dup_rows)),
                ("Dup charges",        str(result.dup_charge_count)),
                ("Dup charge amount",  f"{result.dup_charge_amount:,.2f}"),
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
