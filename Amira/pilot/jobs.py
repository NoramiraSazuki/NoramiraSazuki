"""Create and control jobs (used by the web UI and tests). The worker does the actual processing.

State changes on a job the worker is running are requested through ``jobs.control`` and applied
by the worker after the current page; idle jobs are changed directly with conditional updates.
"""

from __future__ import annotations

import json
import time

from . import store
from .engines import MODES
from .pages import PageSpecError, parse_page_spec, to_spec


class JobError(ValueError):
    pass


def create_job(con, selections: list[tuple[int, str]], mode: str = "text_metadata",
               method_override: str | None = None) -> int:
    modes = {m["id"]: m for m in MODES}
    if mode not in modes or not modes[mode]["implemented"]:
        raise JobError(f"mode '{mode}' is not implemented")
    if not selections:
        raise JobError("no documents selected")
    clean = []
    for doc_id, spec in selections:
        doc = con.execute("SELECT id, resolve_status FROM documents WHERE id=?", (doc_id,)).fetchone()
        if doc is None:
            raise JobError(f"unknown document id {doc_id}")
        spec = (spec or "all").strip() or "all"
        if spec.lower() != "all":
            try:
                spec = to_spec(parse_page_spec(spec))  # syntax check; bounds are checked by the worker
            except PageSpecError as exc:
                raise JobError(f"document {doc_id}: {exc}") from exc
        clean.append((doc_id, spec))
    with store.tx(con):
        job_id = con.execute("INSERT INTO jobs (created, mode, status, method_override) VALUES (?,?,?,?) RETURNING id",
                             (time.time(), mode, "queued", method_override)).fetchone()[0]
        for doc_id, spec in clean:
            con.execute("INSERT INTO job_items (job_id, doc_id, page_spec, status) VALUES (?,?,?,?)",
                        (job_id, doc_id, spec, "pending"))
    return job_id


def estimate_pages(con, selections: list[tuple[int, str]]) -> tuple[int, int]:
    """Pages a job would cover: (total, documents whose page count is unknown and counted as 0)."""
    total = unknown = 0
    for doc_id, spec in selections:
        spec = (spec or "all").strip() or "all"
        if spec.lower() != "all":
            try:
                total += len(parse_page_spec(spec))
            except PageSpecError:
                pass  # create_job reports the error
            continue
        row = con.execute("""SELECT COALESCE(s.page_count, d.inv_pages) n FROM documents d
                             LEFT JOIN sources s ON s.doc_id=d.id WHERE d.id=?""", (doc_id,)).fetchone()
        if row is None or row["n"] is None:
            unknown += 1
        else:
            total += row["n"]
    return total, unknown


def _status(con, job_id: int) -> str:
    row = con.execute("SELECT status FROM jobs WHERE id=?", (job_id,)).fetchone()
    if row is None:
        raise JobError(f"unknown job {job_id}")
    return row[0]


def pause(con, job_id: int) -> None:
    st = _status(con, job_id)
    if st == "running":
        con.execute("UPDATE jobs SET control='pause' WHERE id=? AND status='running'", (job_id,))
    elif st in ("queued", "waiting_for_source"):
        con.execute("UPDATE jobs SET status='paused', reason='paused by user' WHERE id=? AND status=?", (job_id, st))
    # A queued job the worker just picked up is now 'running'; request the pause instead.
    if _status(con, job_id) == "running":
        con.execute("UPDATE jobs SET control='pause' WHERE id=?", (job_id,))


def cancel(con, job_id: int) -> None:
    st = _status(con, job_id)
    if st == "running":
        con.execute("UPDATE jobs SET control='cancel' WHERE id=?", (job_id,))
        return
    if st in ("queued", "paused", "waiting_for_source"):
        with store.tx(con):
            con.execute("""UPDATE batches SET status='cancelled' WHERE status='pending'
                           AND item_id IN (SELECT id FROM job_items WHERE job_id=?)""", (job_id,))
            con.execute("UPDATE job_items SET status='cancelled' WHERE job_id=? AND status IN ('pending','running')",
                        (job_id,))
            con.execute("UPDATE jobs SET status='cancelled', finished=? WHERE id=?", (time.time(), job_id))


def resume(con, job_id: int) -> None:
    """Continue unfinished work; completed pages are reused, not redone."""
    st = _status(con, job_id)
    if st not in ("paused", "cancelled", "waiting_for_source"):
        raise JobError(f"job is {st}; only paused, cancelled or waiting jobs can be resumed")
    with store.tx(con):
        con.execute("""UPDATE batches SET status='pending' WHERE status='cancelled'
                       AND item_id IN (SELECT id FROM job_items WHERE job_id=?)""", (job_id,))
        con.execute("UPDATE job_items SET status='pending' WHERE job_id=? AND status='cancelled'", (job_id,))
        con.execute("UPDATE jobs SET status='queued', control=NULL, reason=NULL, retry_after=NULL, finished=NULL "
                    "WHERE id=?", (job_id,))


def retry(con, job_id: int) -> int:
    """Re-queue failed documents and batches with failed pages. Returns how many were re-queued."""
    st = _status(con, job_id)
    if st not in ("done", "done_with_errors", "paused", "cancelled", "waiting_for_source"):
        raise JobError(f"job is {st}; wait until it stops")
    with store.tx(con):
        n = con.execute("""UPDATE batches SET status='pending', error=NULL WHERE status IN ('done_with_errors','cancelled')
                           AND item_id IN (SELECT id FROM job_items WHERE job_id=?)""", (job_id,)).rowcount
        n += con.execute("""UPDATE job_items SET status='pending', error=NULL
                            WHERE job_id=? AND status IN ('failed','done_with_errors','cancelled')""",
                         (job_id,)).rowcount
        con.execute("UPDATE jobs SET status='queued', control=NULL, reason=NULL, retry_after=NULL, finished=NULL "
                    "WHERE id=?", (job_id,))
    return n


def retry_pages(con, doc_id: int, pages: list[int], method: str, known_methods: set[str]) -> int:
    """Record a per-page method override (history kept) and queue a job for just those pages."""
    if method not in known_methods:
        raise JobError(f"unknown method '{method}'")
    src = con.execute("SELECT sha256 FROM sources WHERE doc_id=?", (doc_id,)).fetchone()
    if not src or not src["sha256"]:
        raise JobError("document has not been processed yet")
    now = time.time()
    with store.tx(con):
        for p in pages:
            con.execute("INSERT INTO overrides (sha256, page_no, method, created) VALUES (?,?,?,?)",
                        (src["sha256"], p, method, now))
    return create_job(con, [(doc_id, to_spec(pages))], method_override=method)


def add_review(con, sha256: str, page_no: int, result_id: int | None, status: str,
               corrected_text: str | None = None, note: str | None = None) -> int:
    if status not in ("unreviewed", "approved", "corrected", "rejected"):
        raise JobError(f"unknown review status '{status}'")
    if status == "corrected" and not (corrected_text or "").strip():
        raise JobError("a correction needs the corrected text")
    return con.execute(
        """INSERT INTO reviews (sha256, page_no, result_id, status, corrected_text, note, created)
           VALUES (?,?,?,?,?,?,?) RETURNING id""",
        (sha256, page_no, result_id, status, corrected_text if status == "corrected" else None, note,
         time.time())).fetchone()[0]


def job_summary(con, job_id: int) -> dict:
    job = dict(con.execute("SELECT * FROM jobs WHERE id=?", (job_id,)).fetchone())
    items = [dict(r) for r in con.execute(
        """SELECT i.*, d.file_name, d.rel_path FROM job_items i JOIN documents d ON d.id=i.doc_id
           WHERE i.job_id=? ORDER BY i.id""", (job_id,))]
    total = sum(i["pages_total"] or 0 for i in items)
    done = sum(i["pages_done"] or 0 for i in items)
    job.update(items=items, pages_total=total, pages_done=done,
               pages_failed=sum(i["pages_failed"] or 0 for i in items),
               pages_reused=sum(i["pages_reused"] or 0 for i in items))
    eta = None
    if job["status"] == "running" and job["started"] and done and total > done:
        rate = done / max(1.0, time.time() - job["started"])
        eta = round((total - done) / rate) if rate > 0 else None
    job["eta_seconds"] = eta
    job["batches"] = {r["status"]: r["n"] for r in con.execute(
        """SELECT b.status, COUNT(*) n FROM batches b JOIN job_items i ON i.id=b.item_id
           WHERE i.job_id=? GROUP BY b.status""", (job_id,))}
    return job


def page_view(con, doc_id: int, page_no: int, engines: dict, routing_version: str) -> dict:
    """Everything the review screen shows for one original page."""
    src = con.execute("SELECT * FROM sources WHERE doc_id=?", (doc_id,)).fetchone()
    if not src or not src["sha256"]:
        raise JobError("document has not been processed yet")
    sha = src["sha256"]
    route = con.execute("SELECT * FROM page_routes WHERE sha256=? AND page_no=? AND routing_version=?",
                        (sha, page_no, routing_version)).fetchone()
    results = []
    for r in con.execute("SELECT * FROM page_results WHERE sha256=? AND page_no=? ORDER BY id", (sha, page_no)):
        eng = engines.get(r["engine"])
        current = bool(eng and eng.version == r["engine_version"] and eng.settings_hash == r["settings_hash"])
        item = {k: r[k] for k in ("id", "engine", "engine_version", "settings_hash", "status", "chars", "error")}
        item.update(flags=json.loads(r["flags"] or "[]"), current_settings=current)
        if r["status"] == "ok" and r["path"]:
            try:
                data = store.read_json(r["path"])
                item.update(text=data["text"], warnings=data.get("warnings", []), stats=data.get("stats", {}),
                            reading_order=data.get("reading_order"), seconds=data.get("seconds"))
            except OSError:
                item.update(status="missing_file")
        results.append(item)
    override = store.latest_override(con, sha, page_no)
    selected = override or (route["method"] if route else None)
    reviews = [dict(r) for r in con.execute("SELECT * FROM reviews WHERE sha256=? AND page_no=? ORDER BY id",
                                            (sha, page_no))]
    ev = json.loads(route["evidence"]) if route else {}
    return {
        "doc_id": doc_id, "page_no": page_no, "page_count": src["page_count"], "sha256": sha,
        "route": {"method": route["method"], "reason": route["reason"], "flags": ev.pop("_flags", []),
                  "evidence": ev} if route else None,
        "override": override, "selected_method": selected, "results": results,
        "review": reviews[-1] if reviews else None, "review_history": reviews,
    }
