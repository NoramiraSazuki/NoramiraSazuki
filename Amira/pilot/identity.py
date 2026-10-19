"""Source-file identity: size + mtime + full SHA-256 (streamed, read-only).

The hash is cached per (size, mtime) so an unchanged 700 MB file on the NAS is hashed once.
Results are reused only for the same SHA-256, so a changed file never inherits old results.
"""

from __future__ import annotations

import hashlib
import os
import time
from dataclasses import dataclass
from pathlib import Path


class SourceUnavailable(OSError):
    """The source file or its folder (e.g. NAS share) cannot be reached right now."""


@dataclass(frozen=True)
class Fingerprint:
    size: int
    mtime_ns: int
    sha256: str


def stat_source(path: Path) -> os.stat_result:
    try:
        return os.stat(path)
    except FileNotFoundError as exc:
        # Distinguish "NAS gone" (parent missing) from "this file is gone".
        if not path.parent.exists():
            raise SourceUnavailable("source folder not reachable") from exc
        raise
    except OSError as exc:
        raise SourceUnavailable(f"source not reachable ({type(exc).__name__})") from exc


def sha256_file(path: Path, chunk: int = 4 * 1024 * 1024, should_stop=None) -> str:
    h = hashlib.sha256()
    with open(path, "rb") as fh:  # read-only
        while True:
            block = fh.read(chunk)
            if not block:
                break
            h.update(block)
            if should_stop and should_stop():
                raise InterruptedError("hashing interrupted")
    return h.hexdigest()


def fingerprint(con, doc_id: int, path: Path, should_stop=None) -> Fingerprint:
    st = stat_source(path)
    row = con.execute("SELECT size, mtime_ns, sha256 FROM sources WHERE doc_id=?", (doc_id,)).fetchone()
    if row and row["sha256"] and row["size"] == st.st_size and row["mtime_ns"] == st.st_mtime_ns:
        return Fingerprint(st.st_size, st.st_mtime_ns, row["sha256"])
    try:
        digest = sha256_file(path, should_stop=should_stop)
    except FileNotFoundError as exc:
        raise SourceUnavailable("source disappeared while hashing") from exc
    after = stat_source(path)
    if (after.st_size, after.st_mtime_ns) != (st.st_size, st.st_mtime_ns):
        raise SourceUnavailable("source changed while it was being hashed")
    con.execute(
        """INSERT INTO sources (doc_id, size, mtime_ns, sha256, checked) VALUES (?,?,?,?,?)
           ON CONFLICT(doc_id) DO UPDATE SET size=excluded.size, mtime_ns=excluded.mtime_ns,
               sha256=excluded.sha256, page_count=NULL, encryption=NULL, pdf_meta=NULL,
               checked=excluded.checked""",
        (doc_id, st.st_size, st.st_mtime_ns, digest, time.time()))
    return Fingerprint(st.st_size, st.st_mtime_ns, digest)


def unchanged(path: Path, fp: Fingerprint) -> bool:
    """Cheap check between batches: size and mtime still match the hashed identity."""
    st = stat_source(path)
    return st.st_size == fp.size and st.st_mtime_ns == fp.mtime_ns
