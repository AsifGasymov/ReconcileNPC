"""Backup service — copy outputs to backup_dir/YYYY-MM-DD/stage_N/, prune old."""
from __future__ import annotations

import shutil
from datetime import date, datetime, timedelta
from pathlib import Path

from .settings import get_backup_dir, get_backup_retention_days


def archive_output(out_path: str | Path, stage_key: str) -> Path | None:
    """Copy `out_path` into the backup directory for today/stage. Returns the
    archived path or None if archiving failed (caller should not block on it)."""
    src = Path(out_path)
    if not src.exists():
        return None

    backup_root = get_backup_dir()
    today = date.today().isoformat()  # YYYY-MM-DD
    target_dir = backup_root / today / stage_key
    try:
        target_dir.mkdir(parents=True, exist_ok=True)
        target = target_dir / src.name
        # If a file with the same name already exists, append a timestamp.
        if target.exists():
            stamp = datetime.now().strftime("%H%M%S")
            target = target_dir / f"{src.stem}_{stamp}{src.suffix}"
        shutil.copy2(src, target)
        return target
    except OSError:
        return None


def prune_old_backups() -> int:
    """Delete date folders older than retention. Returns count removed."""
    backup_root = get_backup_dir()
    if not backup_root.exists():
        return 0

    cutoff = date.today() - timedelta(days=get_backup_retention_days())
    removed = 0
    for child in backup_root.iterdir():
        if not child.is_dir():
            continue
        try:
            folder_date = date.fromisoformat(child.name)
        except ValueError:
            continue
        if folder_date < cutoff:
            try:
                shutil.rmtree(child)
                removed += 1
            except OSError:
                pass
    return removed
