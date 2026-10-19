"""Logging without document content, and error messages that are safe to share.

Logs record ids, page numbers, status codes and exception types only. ``sanitize`` removes
file paths and quoted fragments from exception messages before they are stored or shown.
"""

from __future__ import annotations

import logging
import re
from logging.handlers import RotatingFileHandler
from pathlib import Path

_PATH_RE = re.compile(r"(?:~|/)[^\s:'\"]*(?:/[^\s:'\"]*)+|[A-Za-z]:\\[^\s'\"]+")
_QUOTED_RE = re.compile(r"'[^']{0,200}'|\"[^\"]{0,200}\"")


def sanitize(exc: BaseException | str, limit: int = 160) -> str:
    if isinstance(exc, BaseException):
        text = f"{type(exc).__name__}: {exc}"
    else:
        text = str(exc)
    text = _PATH_RE.sub("<path>", text)
    text = _QUOTED_RE.sub("<…>", text)
    return " ".join(text.split())[:limit]


def setup(logs_dir: Path, name: str, console: bool = True) -> logging.Logger:
    log = logging.getLogger("pilot")
    for h in list(log.handlers):
        log.removeHandler(h)
        h.close()
    log.setLevel(logging.INFO)
    log.propagate = False
    logs_dir.mkdir(parents=True, exist_ok=True)
    fh = RotatingFileHandler(logs_dir / f"{name}.log", maxBytes=5_000_000, backupCount=3, encoding="utf-8")
    fh.setFormatter(logging.Formatter("%(asctime)s %(levelname)s %(message)s"))
    log.addHandler(fh)
    if console:
        ch = logging.StreamHandler()
        ch.setFormatter(logging.Formatter("%(asctime)s %(message)s", "%H:%M:%S"))
        log.addHandler(ch)
    # Library warnings can quote document fragments; keep them out of the logs entirely.
    for noisy in ("pypdf", "pypdfium2"):
        logging.getLogger(noisy).setLevel(logging.ERROR)
        logging.getLogger(noisy).propagate = False
    return log
