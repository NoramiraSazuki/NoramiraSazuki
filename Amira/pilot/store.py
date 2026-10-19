"""SQLite state (WAL) plus immutable per-page JSON result files, both inside the data root.

Processing status lives here and is kept apart from the inventory's own ingestion status, which
is only copied for display. Page results are JSON files written atomically; SQLite indexes them.
"""

from __future__ import annotations

import json
import os
import sqlite3
import tempfile
import time
from contextlib import contextmanager
from pathlib import Path

SCHEMA = """
CREATE TABLE IF NOT EXISTS meta (key TEXT PRIMARY KEY, value TEXT);

CREATE TABLE IF NOT EXISTS documents (
    id INTEGER PRIMARY KEY,
    doc_key TEXT UNIQUE NOT NULL,          -- relative path (+ inventory row id if listed twice)
    rel_path TEXT NOT NULL,
    file_name TEXT NOT NULL,
    resolve_status TEXT NOT NULL,
    resolved_path TEXT,
    note TEXT,
    inv_row_id TEXT, inv_pages INTEGER, inv_size INTEGER, inv_ingest_status TEXT,
    inv_duplicate_group TEXT, inv_copies INTEGER, inv_etag TEXT,
    in_inventory INTEGER NOT NULL DEFAULT 1,
    title TEXT,                            -- catalogue title from the mapping workbook (not from the PDF)
    updated REAL
);

-- Identity of a source file as last seen by the worker.
CREATE TABLE IF NOT EXISTS sources (
    doc_id INTEGER PRIMARY KEY REFERENCES documents(id),
    size INTEGER, mtime_ns INTEGER, sha256 TEXT, page_count INTEGER, encryption TEXT,
    pdf_meta TEXT,                        -- embedded PDF metadata (JSON), reference only
    checked REAL
);

CREATE TABLE IF NOT EXISTS jobs (
    id INTEGER PRIMARY KEY,
    created REAL, mode TEXT NOT NULL,
    status TEXT NOT NULL,                 -- queued/running/paused/cancelled/waiting_for_source/done/done_with_errors
    control TEXT,                         -- pause/cancel requested by the UI
    reason TEXT,                          -- why paused/waiting (low_disk, source_unavailable, ...)
    retry_after REAL,
    method_override TEXT,                 -- force one engine for every page of this job (retry jobs)
    started REAL, finished REAL, heartbeat REAL
);

CREATE TABLE IF NOT EXISTS job_items (
    id INTEGER PRIMARY KEY,
    job_id INTEGER NOT NULL REFERENCES jobs(id),
    doc_id INTEGER NOT NULL REFERENCES documents(id),
    page_spec TEXT NOT NULL,              -- original page numbers, e.g. "1-20,45" or "all"
    status TEXT NOT NULL,                 -- pending/running/done/done_with_errors/failed/cancelled
    sha256 TEXT, page_count INTEGER,
    pages_total INTEGER DEFAULT 0, pages_done INTEGER DEFAULT 0, pages_reused INTEGER DEFAULT 0,
    pages_failed INTEGER DEFAULT 0,
    warning TEXT, error TEXT
);

CREATE TABLE IF NOT EXISTS batches (
    id INTEGER PRIMARY KEY,
    item_id INTEGER NOT NULL REFERENCES job_items(id),
    seq INTEGER NOT NULL,
    pages TEXT NOT NULL,                  -- JSON list of original page numbers
    status TEXT NOT NULL,                 -- pending/running/done/done_with_errors/cancelled
    attempts INTEGER DEFAULT 0,
    error TEXT, updated REAL
);

CREATE TABLE IF NOT EXISTS page_routes (
    sha256 TEXT NOT NULL, page_no INTEGER NOT NULL, routing_version TEXT NOT NULL,
    evidence TEXT NOT NULL, method TEXT NOT NULL, reason TEXT NOT NULL,
    PRIMARY KEY (sha256, page_no, routing_version)
);

CREATE TABLE IF NOT EXISTS page_results (
    id INTEGER PRIMARY KEY,
    sha256 TEXT NOT NULL, page_no INTEGER NOT NULL,
    engine TEXT NOT NULL, engine_version TEXT NOT NULL, settings_hash TEXT NOT NULL,
    status TEXT NOT NULL,                 -- ok / failed
    path TEXT, chars INTEGER, flags TEXT, error TEXT, created REAL,
    UNIQUE (sha256, page_no, engine, engine_version, settings_hash)
);

-- Per-page method chosen by the user (history kept; latest wins).
CREATE TABLE IF NOT EXISTS overrides (
    id INTEGER PRIMARY KEY, sha256 TEXT NOT NULL, page_no INTEGER NOT NULL, method TEXT NOT NULL,
    created REAL
);

-- Human review, append-only revisions; never overwrites page_results.
CREATE TABLE IF NOT EXISTS reviews (
    id INTEGER PRIMARY KEY, sha256 TEXT NOT NULL, page_no INTEGER NOT NULL, result_id INTEGER,
    status TEXT NOT NULL,                 -- unreviewed/approved/corrected/rejected
    corrected_text TEXT, note TEXT, created REAL
);

CREATE INDEX IF NOT EXISTS ix_items_job ON job_items(job_id);
CREATE INDEX IF NOT EXISTS ix_batches_item ON batches(item_id, status);
CREATE INDEX IF NOT EXISTS ix_results_page ON page_results(sha256, page_no);
CREATE INDEX IF NOT EXISTS ix_reviews_page ON reviews(sha256, page_no);
"""


def connect(db_path: Path) -> sqlite3.Connection:
    con = sqlite3.connect(db_path, timeout=30, isolation_level=None)  # autocommit; explicit tx below
    con.row_factory = sqlite3.Row
    con.execute("PRAGMA journal_mode=WAL")
    con.execute("PRAGMA synchronous=NORMAL")
    con.execute("PRAGMA foreign_keys=ON")
    con.execute("PRAGMA busy_timeout=30000")
    return con


def init_db(db_path: Path) -> sqlite3.Connection:
    db_path.parent.mkdir(parents=True, exist_ok=True)
    con = connect(db_path)
    con.executescript(SCHEMA)
    cols = {r[1] for r in con.execute("PRAGMA table_info(documents)")}
    if "title" not in cols:  # databases created before titles were added
        con.execute("ALTER TABLE documents ADD COLUMN title TEXT")
    return con


@contextmanager
def tx(con: sqlite3.Connection):
    con.execute("BEGIN IMMEDIATE")
    try:
        yield con
    except BaseException:
        con.execute("ROLLBACK")
        raise
    con.execute("COMMIT")


def write_json_atomic(path: Path, obj) -> None:
    """Write JSON via a temp file in the same folder + fsync + rename, so readers never see half a file."""
    path.parent.mkdir(parents=True, exist_ok=True)
    fd, tmp = tempfile.mkstemp(dir=path.parent, prefix=".tmp_", suffix=".json")
    try:
        with os.fdopen(fd, "w", encoding="utf-8") as fh:
            json.dump(obj, fh, ensure_ascii=False)
            fh.flush()
            os.fsync(fh.fileno())
        os.replace(tmp, path)
    except BaseException:
        Path(tmp).unlink(missing_ok=True)
        raise


def read_json(path: Path):
    with open(path, encoding="utf-8") as fh:
        return json.load(fh)


# ---------------------------------------------------------------- documents

def upsert_documents(con: sqlite3.Connection, inv, titles=None) -> None:
    """Refresh the catalogue from an InventoryResult. Rows are never deleted (job history refers to them)."""
    now = time.time()
    counts: dict[str, int] = {}
    for r in inv.rows:
        counts[r.rel_path] = counts.get(r.rel_path, 0) + 1
    with tx(con):
        con.execute("UPDATE documents SET in_inventory = 0")
        for r in inv.rows:
            key = r.rel_path if counts[r.rel_path] == 1 else f"{r.rel_path}#row{r.extra.get('inv_row_id')}"
            e = r.extra
            con.execute(
                """INSERT INTO documents (doc_key, rel_path, file_name, resolve_status, resolved_path, note,
                       inv_row_id, inv_pages, inv_size, inv_ingest_status, inv_duplicate_group, inv_copies,
                       inv_etag, in_inventory, updated)
                   VALUES (?,?,?,?,?,?,?,?,?,?,?,?,?,1,?)
                   ON CONFLICT(doc_key) DO UPDATE SET rel_path=excluded.rel_path, file_name=excluded.file_name,
                       resolve_status=excluded.resolve_status, resolved_path=excluded.resolved_path,
                       note=excluded.note, inv_row_id=excluded.inv_row_id, inv_pages=excluded.inv_pages,
                       inv_size=excluded.inv_size, inv_ingest_status=excluded.inv_ingest_status,
                       inv_duplicate_group=excluded.inv_duplicate_group, inv_copies=excluded.inv_copies,
                       inv_etag=excluded.inv_etag, in_inventory=1, updated=excluded.updated""",
                (key, r.rel_path, r.file_name, r.resolve_status, r.resolved_path or None, r.note,
                 _s(e.get("inv_row_id")), _i(e.get("inv_pages")), _i(e.get("inv_size")),
                 _s(e.get("inv_ingest_status")), _s(e.get("inv_duplicate_group")), _i(e.get("inv_copies")),
                 _s(e.get("inv_etag")), now))
        con.execute("INSERT OR REPLACE INTO meta VALUES ('unlisted', ?)", (json.dumps(inv.unlisted),))
        if titles is not None:
            con.execute("UPDATE documents SET title = NULL")
            for rel, title in titles.titles.items():
                con.execute("UPDATE documents SET title=? WHERE rel_path=?", (title, rel))
            con.execute("INSERT OR REPLACE INTO meta VALUES ('titles_report', ?)", (json.dumps({
                "rows": titles.rows, "with_title": len(titles.titles), "blank": titles.blank,
                "duplicates": titles.duplicates, "bad_location": titles.bad_location, "error": titles.error,
                "header_row": titles.header_row, "title_header": titles.title_header}),))
        con.execute("INSERT OR REPLACE INTO meta VALUES ('inventory_loaded', ?)", (str(now),))
        con.execute("INSERT OR REPLACE INTO meta VALUES ('source_available', ?)", (str(int(inv.source_available)),))


def _s(v):
    return None if v is None else str(v)


def _i(v):
    return v if isinstance(v, int) else None


def meta_get(con, key: str, default=None):
    row = con.execute("SELECT value FROM meta WHERE key=?", (key,)).fetchone()
    return row[0] if row else default


def meta_set(con, key: str, value: str) -> None:
    con.execute("INSERT OR REPLACE INTO meta VALUES (?, ?)", (key, value))


# ---------------------------------------------------------------- page results

def find_result(con, sha256: str, page_no: int, engine: str, version: str, shash: str):
    return con.execute(
        """SELECT * FROM page_results WHERE sha256=? AND page_no=? AND engine=? AND engine_version=?
           AND settings_hash=?""", (sha256, page_no, engine, version, shash)).fetchone()


def record_result(con, sha256, page_no, engine, version, shash, status, path, chars, flags, error) -> int:
    cur = con.execute(
        """INSERT INTO page_results (sha256, page_no, engine, engine_version, settings_hash, status, path, chars,
               flags, error, created) VALUES (?,?,?,?,?,?,?,?,?,?,?)
           ON CONFLICT(sha256, page_no, engine, engine_version, settings_hash) DO UPDATE SET
               status=excluded.status, path=excluded.path, chars=excluded.chars, flags=excluded.flags,
               error=excluded.error, created=excluded.created
           RETURNING id""",
        (sha256, page_no, engine, version, shash, status, path, chars, json.dumps(flags), error, time.time()))
    return cur.fetchone()[0]


def latest_override(con, sha256: str, page_no: int) -> str | None:
    row = con.execute("SELECT method FROM overrides WHERE sha256=? AND page_no=? ORDER BY id DESC LIMIT 1",
                      (sha256, page_no)).fetchone()
    return row[0] if row else None


def latest_review(con, sha256: str, page_no: int):
    return con.execute("SELECT * FROM reviews WHERE sha256=? AND page_no=? ORDER BY id DESC LIMIT 1",
                       (sha256, page_no)).fetchone()
