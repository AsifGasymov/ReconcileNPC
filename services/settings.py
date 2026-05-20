"""Persistent user settings via QSettings."""
from __future__ import annotations

from pathlib import Path

from PySide6.QtCore import QSettings


_DEFAULT_BACKUP = Path.home() / "Documents" / "Cardaq" / "Backups"
_DEFAULT_OUTPUT = Path.home() / "Desktop"


def _settings() -> QSettings:
    return QSettings("Cardaq", "NPCMode")


def get_backup_dir() -> Path:
    s = _settings()
    raw = s.value("backup_dir", "", type=str)
    return Path(raw) if raw else _DEFAULT_BACKUP


def set_backup_dir(path: str | Path) -> None:
    _settings().setValue("backup_dir", str(path))


def get_output_dir(stage_key: str) -> Path:
    s = _settings()
    raw = s.value(f"output_dir/{stage_key}", "", type=str)
    return Path(raw) if raw else _DEFAULT_OUTPUT


def set_output_dir(stage_key: str, path: str | Path) -> None:
    _settings().setValue(f"output_dir/{stage_key}", str(path))


def get_backup_retention_days() -> int:
    return int(_settings().value("backup_retention_days", 5, type=int))


def set_backup_retention_days(days: int) -> None:
    _settings().setValue("backup_retention_days", int(days))
