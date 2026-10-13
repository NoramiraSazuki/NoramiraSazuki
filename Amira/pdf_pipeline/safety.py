"""Path guards that keep all writes inside the workspace and away from source data."""

from __future__ import annotations

import os
from dataclasses import dataclass
from pathlib import Path

from . import WORKSPACE


class SafetyError(RuntimeError):
    """Raised when an operation would write somewhere it must not."""


def is_within(path: Path, parent: Path) -> bool:
    try:
        path.resolve().relative_to(parent.resolve())
        return True
    except ValueError:
        return False


def check_write_location(target: Path, source: Path) -> None:
    """Refuse write targets outside the workspace or inside/beside the source data."""
    target = target.resolve()
    source = source.resolve()
    if not is_within(target, WORKSPACE):
        raise SafetyError(f"Refusing to write outside the workspace: {target}")
    source_dir = source if source.is_dir() else source.parent
    if is_within(target, source_dir):
        raise SafetyError(f"Refusing to write inside the source folder: {target}")
    if is_within(source, target):
        raise SafetyError(f"Source data must not live inside the output folder: {target}")


def create_new_dir(path: Path) -> Path:
    """Create a directory that must not already exist (never reuse/overwrite outputs)."""
    if path.exists():
        raise SafetyError(f"Output folder already exists, not overwriting: {path}")
    path.mkdir(parents=True)
    return path


@dataclass(frozen=True)
class FileSnapshot:
    """Size + modification time, used to confirm a source file was not changed."""

    size: int
    mtime_ns: int

    @classmethod
    def of(cls, path: Path) -> "FileSnapshot":
        st = os.stat(path)
        return cls(st.st_size, st.st_mtime_ns)
