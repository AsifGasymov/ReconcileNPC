"""SaltEdge section views — ManoBank, Nexpay, Token."""
from __future__ import annotations

from PySide6.QtCore import Qt
from PySide6.QtWidgets import (
    QFrame, QLabel, QPushButton, QScrollArea, QVBoxLayout, QWidget,
)

from core.saltedge_nexpay import NexpayResult, run_saltedge_nexpay
from services.settings import get_output_dir, set_output_dir
from services.worker import start_worker

from ..widgets.file_card import FileCard
from ..widgets.folder_card import FolderCard
from ..widgets.multi_file_card import MultiFileCard
from ..widgets.progress_card import ProgressCard
from ..widgets.result_card import ResultCard


def _placeholder_view(title: str, subtitle: str) -> QWidget:
    w = QWidget()
    layout = QVBoxLayout(w)
    layout.setContentsMargins(40, 32, 40, 32)
    layout.setSpacing(12)
    lbl = QLabel(title)
    lbl.setObjectName("h1")
    sub = QLabel(subtitle)
    sub.setObjectName("hint")
    layout.addWidget(lbl)
    layout.addWidget(sub)
    layout.addStretch()
    return w


class ManoBankView(QWidget):
    def __init__(self, parent: QWidget | None = None):
        super().__init__(parent)
        layout = QVBoxLayout(self)
        layout.setContentsMargins(40, 32, 40, 32)
        layout.setSpacing(12)
        lbl = QLabel("Saltedge · ManoBank")
        lbl.setObjectName("h1")
        sub = QLabel("ManoBank reconciliation — coming soon")
        sub.setObjectName("hint")
        layout.addWidget(lbl)
        layout.addWidget(sub)
        layout.addStretch()


STAGE_KEY_NEXPAY = "saltedge_nexpay"


class NexpayView(QWidget):
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

        title = QLabel("Saltedge · Nexpay")
        title.setObjectName("h1")
        sub = QLabel(
            "SaltEdge export × Nexpay statement  →  reconciliation workbook"
        )
        sub.setObjectName("hint")
        layout.addWidget(title)
        layout.addWidget(sub)
        layout.addSpacing(8)

        self._se_files = MultiFileCard(
            "System exports (SaltEdge / Token)",
            hint="Drop up to 9 CSV exports from your system",
            extensions=[".csv"],
            max_items=9,
        )
        self._nx_file = FileCard(
            "Nexpay statement",
            hint="XLSX payment history from Nexpay provider",
            extensions=[".xlsx"],
        )
        self._output = FolderCard(
            "Output folder",
            initial=str(get_output_dir(STAGE_KEY_NEXPAY)),
        )
        self._output.folder_changed.connect(
            lambda p: set_output_dir(STAGE_KEY_NEXPAY, p))

        layout.addWidget(self._se_files)
        layout.addWidget(self._nx_file)
        layout.addWidget(self._output)

        self._run_btn = QPushButton("▶  Build Nexpay report")
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

        self._se_files.paths_changed.connect(lambda *_: self._refresh_enabled())
        self._nx_file.file_selected.connect(lambda *_: self._refresh_enabled())
        self._nx_file.cleared.connect(lambda: self._refresh_enabled())

        self._thread = None
        self._worker = None

    def _refresh_enabled(self) -> None:
        self._run_btn.setEnabled(
            bool(self._se_files.paths() and self._nx_file.path() and self._output.path())
        )

    def _on_run(self) -> None:
        self._run_btn.setEnabled(False)
        self._result.hide()
        self._progress.show()
        self._progress.reset()
        self._progress.set_status("Starting…", "info")

        self._thread, self._worker = start_worker(
            self,
            run_saltedge_nexpay,
            system_paths=self._se_files.paths(),
            nexpay_path=self._nx_file.path(),
            out_dir=self._output.path(),
        )
        self._worker.log.connect(self._on_log)
        self._worker.finished.connect(self._on_finished)
        self._worker.failed.connect(self._on_failed)
        self._thread.start()

    def _on_log(self, msg: str) -> None:
        self._progress.append_log(msg)
        self._progress.set_status(msg, "info")
        if "saltedge" in msg.lower():
            self._progress.set_progress(20)
        elif "nexpay" in msg.lower():
            self._progress.set_progress(40)
        elif "merging" in msg.lower():
            self._progress.set_progress(65)
        elif "writing" in msg.lower():
            self._progress.set_progress(85)

    def _on_finished(self, result: NexpayResult) -> None:
        self._progress.set_progress(100)
        self._progress.set_status("Done", "ok")
        self._result.show_result(
            stats=[
                ("Matched",         str(result.matched)),
                ("SE proc / NX miss", str(result.se_processed_nx_missing)),
                ("Exceptions",      str(result.matched_exc)),
                ("Unmatched",       str(result.unmatched)),
                ("SE Total EUR",    f"{result.total_se_amount:,.2f}"),
                ("NX Total EUR",    f"{result.total_nx_amount:,.2f}"),
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
