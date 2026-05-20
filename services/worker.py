"""QThread wrapper for running long-running pipelines off the UI thread."""
from __future__ import annotations

import traceback
from typing import Any, Callable

from PySide6.QtCore import QObject, QThread, Signal


class PipelineWorker(QObject):
    progress = Signal(int)
    log = Signal(str)
    finished = Signal(object)
    failed = Signal(str)

    def __init__(self, fn: Callable[..., Any], **kwargs):
        super().__init__()
        self._fn = fn
        self._kwargs = kwargs

    def run(self) -> None:
        try:
            def log(msg: str) -> None:
                self.log.emit(msg)
            result = self._fn(log=log, **self._kwargs)
            self.finished.emit(result)
        except Exception as exc:  # noqa: BLE001
            tb = traceback.format_exc()
            self.failed.emit(f"{exc}\n\n{tb}")


def start_worker(parent: QObject, fn: Callable[..., Any], **kwargs) -> tuple[QThread, PipelineWorker]:
    """Spin up a QThread + worker, return both so caller can wire signals."""
    thread = QThread(parent)
    worker = PipelineWorker(fn, **kwargs)
    worker.moveToThread(thread)
    thread.started.connect(worker.run)
    worker.finished.connect(thread.quit)
    worker.failed.connect(thread.quit)
    thread.finished.connect(worker.deleteLater)
    thread.finished.connect(thread.deleteLater)
    return thread, worker
