"""Background worker: runs one job at a time, in small checkpointed batches of original pages.

    python -m pilot.worker            # loop forever (started by the app launcher)
    python -m pilot.worker --once     # process runnable jobs, then exit (tests / manual runs)

Each page result is saved (atomically) as soon as it is produced, so an interruption loses at
most the page in progress. Results are reused only when source SHA-256, engine name/version and
engine settings hash all match. A failed page never removes completed results.
"""

from __future__ import annotations

import argparse
import fcntl
import json
import logging
import os
import shutil
import sys
import time
from pathlib import Path

from pdf_pipeline.safety import is_within

from . import routing, store
from .engines import build_engines
from .engines.base import RESULT_SCHEMA, EngineCancelled, EngineTimeout, PageContext
from .identity import Fingerprint, SourceUnavailable, fingerprint, unchanged
from .logutil import sanitize, setup
from .pages import PageSpecError, chunk, parse_page_spec
from .pdfio import PdfOpenError, encryption_of, open_pdf, pdf_metadata
from .settings import Settings, load_settings

log = logging.getLogger("pilot")

RUNNABLE = ("queued", "running")
LOW_CONF = 60.0


class _Stop(Exception):
    """Pause/cancel/low-disk requested: leave the job in a resumable state."""

    def __init__(self, outcome: str):
        super().__init__(outcome)
        self.outcome = outcome


def validation_flags(engine: str, result, route_flags: list[str]) -> list[str]:
    """Automatic checks only. Passing them does not mean the extraction is correct."""
    flags = []
    chars = result.stats.get("chars", 0)
    if chars == 0:
        flags.append("empty")
    conf = result.stats.get("mean_conf")
    if conf is not None and conf < LOW_CONF:
        flags.append("low_ocr_confidence")
    if routing._garbage_ratio(result.text) > 0.3:
        flags.append("garbled_text")
    if engine in ("textlayer", "pypdf") and any("image-dominant" in f for f in route_flags):
        flags.append("text_layer_may_miss_image_content")
    return flags


class Worker:
    def __init__(self, settings: Settings, engines=None, rules_path: Path | None = None):
        self.s = settings
        self.s.ensure_dirs()
        self.con = store.init_db(self.s.db_path)
        self.engines = engines or build_engines(settings)
        self.rules = routing.load_rules(rules_path or routing.DEFAULT_RULES, set(self.engines))
        self._lock_fh = None

    # ------------------------------------------------------------ lifecycle
    def acquire_lock(self) -> None:
        """Only one worker per data root (one heavy job at a time)."""
        self._lock_fh = open(self.s.data_root / "worker.lock", "w")
        try:
            fcntl.flock(self._lock_fh, fcntl.LOCK_EX | fcntl.LOCK_NB)
        except BlockingIOError:
            raise SystemExit("Another worker is already running for this data folder.")
        self._lock_fh.write(str(os.getpid()))
        self._lock_fh.flush()

    def recover(self) -> None:
        """After a crash or kill: anything marked running goes back to the queue (completed pages are kept)."""
        with store.tx(self.con):
            self.con.execute("UPDATE batches SET status='pending' WHERE status='running'")
            self.con.execute("UPDATE job_items SET status='pending' WHERE status='running'")
            self.con.execute("UPDATE jobs SET status='queued' WHERE status='running'")
        self.clean_temp()

    def clean_temp(self) -> None:
        keep = self.s.keep_temp_days
        cutoff = time.time() - (keep or 0) * 86400
        for d in self.s.tmp_dir.iterdir():
            try:
                if keep is None or d.stat().st_mtime < cutoff:
                    shutil.rmtree(d, ignore_errors=True) if d.is_dir() else d.unlink(missing_ok=True)
            except OSError:
                pass

    def heartbeat(self, activity: dict | None = None) -> None:
        store.meta_set(self.con, "worker_heartbeat", str(time.time()))
        if activity is not None:
            store.meta_set(self.con, "worker_activity", json.dumps(activity))

    def next_job(self):
        return self.con.execute(
            """SELECT * FROM jobs WHERE status IN ('queued','running')
                  OR (status='waiting_for_source' AND COALESCE(retry_after,0) <= ?)
               ORDER BY id LIMIT 1""", (time.time(),)).fetchone()

    def run_forever(self, poll: float = 1.0) -> None:
        self.acquire_lock()
        self.recover()
        log.info("worker started (pid %d)", os.getpid())
        while True:
            self.heartbeat()
            job = self.next_job()
            if job is None:
                self.heartbeat({"state": "idle"})
                time.sleep(poll)
                continue
            self.run_job(job["id"])

    def run_once(self) -> None:
        """Process runnable jobs until none is left (jobs waiting for the source are not waited for)."""
        self.recover()
        while (job := self.next_job()) is not None:
            self.run_job(job["id"])

    # ------------------------------------------------------------ control
    def _control(self, job_id: int) -> str | None:
        row = self.con.execute("SELECT control FROM jobs WHERE id=?", (job_id,)).fetchone()
        return row["control"] if row else None

    def _check_stop(self, job_id: int) -> None:
        ctl = self._control(job_id)
        if ctl in ("pause", "cancel"):
            raise _Stop("paused" if ctl == "pause" else "cancelled")

    def _check_disk(self) -> None:
        free_gb = shutil.disk_usage(self.s.data_root).free / 1e9
        if free_gb < self.s.min_free_gb:
            raise _Stop("low_disk")

    def _set_job(self, job_id: int, **cols) -> None:
        sets = ", ".join(f"{k}=?" for k in cols)
        self.con.execute(f"UPDATE jobs SET {sets} WHERE id=?", (*cols.values(), job_id))

    def _set_item(self, item_id: int, **cols) -> None:
        sets = ", ".join(f"{k}=?" for k in cols)
        self.con.execute(f"UPDATE job_items SET {sets} WHERE id=?", (*cols.values(), item_id))

    # ------------------------------------------------------------ job
    def run_job(self, job_id: int) -> str:
        job = self.con.execute("SELECT * FROM jobs WHERE id=?", (job_id,)).fetchone()
        self._set_job(job_id, status="running", reason=None, started=job["started"] or time.time(),
                      heartbeat=time.time())
        log.info("job %d running", job_id)
        items = self.con.execute(
            "SELECT * FROM job_items WHERE job_id=? AND status IN ('pending','running') ORDER BY id",
            (job_id,)).fetchall()
        try:
            for item in items:
                self._check_stop(job_id)
                self.run_item(job, item)
        except _Stop as stop:
            return self._stopped(job_id, stop.outcome)
        except SourceUnavailable as exc:
            self._set_job(job_id, status="waiting_for_source", reason=sanitize(exc), retry_after=time.time() + 30)
            log.warning("job %d waiting for source: %s", job_id, sanitize(exc))
            return "waiting_for_source"
        bad = self.con.execute(
            "SELECT COUNT(*) FROM job_items WHERE job_id=? AND status IN ('failed','done_with_errors')",
            (job_id,)).fetchone()[0]
        status = "done_with_errors" if bad else "done"
        self._set_job(job_id, status=status, control=None, finished=time.time())
        self.heartbeat({"state": "idle"})
        log.info("job %d %s", job_id, status)
        return status

    def _stopped(self, job_id: int, outcome: str) -> str:
        if outcome == "cancelled":
            with store.tx(self.con):
                self.con.execute(
                    """UPDATE batches SET status='cancelled' WHERE status IN ('pending','running')
                       AND item_id IN (SELECT id FROM job_items WHERE job_id=?)""", (job_id,))
                self.con.execute("UPDATE job_items SET status='cancelled' WHERE job_id=? AND status IN "
                                 "('pending','running')", (job_id,))
                self._set_job(job_id, status="cancelled", control=None, finished=time.time())
        else:
            self.con.execute(
                """UPDATE batches SET status='pending' WHERE status='running'
                   AND item_id IN (SELECT id FROM job_items WHERE job_id=?)""", (job_id,))
            self.con.execute("UPDATE job_items SET status='pending' WHERE job_id=? AND status='running'", (job_id,))
            reason = "free disk space below min_free_gb" if outcome == "low_disk" else "paused by user"
            self._set_job(job_id, status="paused", control=None, reason=reason)
        self.heartbeat({"state": "idle"})
        log.info("job %d %s", job_id, outcome)
        return outcome

    # ------------------------------------------------------------ item (one document)
    def run_item(self, job, item) -> None:
        doc = self.con.execute("SELECT * FROM documents WHERE id=?", (item["doc_id"],)).fetchone()
        if doc["resolve_status"] != "OK" or not doc["resolved_path"]:
            self._set_item(item["id"], status="failed", error=f"document not resolved: {doc['resolve_status']}")
            return
        path = Path(doc["resolved_path"])
        if not is_within(path, self.s.source_root):
            self._set_item(item["id"], status="failed", error="resolved path is outside the configured source folder")
            return
        self._set_item(item["id"], status="running")
        self.heartbeat({"state": "hashing", "job": job["id"], "item": item["id"], "doc": doc["id"]})

        prev = self.con.execute("SELECT sha256 FROM sources WHERE doc_id=?", (doc["id"],)).fetchone()
        try:
            fp = fingerprint(self.con, doc["id"], path, should_stop=lambda: self._control(job["id"]) is not None)
        except InterruptedError:
            self._check_stop(job["id"])
            raise
        except FileNotFoundError:
            self._set_item(item["id"], status="failed", error="source file not found")
            return
        warnings = []
        if prev and prev["sha256"] and prev["sha256"] != fp.sha256:
            warnings.append("source content changed since it was last processed; old results are not reused")
        if item["sha256"] and item["sha256"] != fp.sha256:
            # The file changed between runs of this job: completed batches belong to the old version.
            with store.tx(self.con):
                self.con.execute("UPDATE batches SET status='pending', error=NULL WHERE item_id=?", (item["id"],))
                self._set_item(item["id"], pages_done=0, pages_reused=0, pages_failed=0)

        try:
            pdf = open_pdf(path)
        except PdfOpenError as exc:
            enc = encryption_of(path) if exc.code == "encrypted" else ""
            self.con.execute("UPDATE sources SET encryption=? WHERE doc_id=?", (enc, doc["id"]))
            self._set_item(item["id"], status="failed", sha256=fp.sha256, error=str(exc))
            return
        except FileNotFoundError as exc:
            raise SourceUnavailable("source not reachable when opening") from exc
        try:
            page_count = len(pdf)
        finally:
            pdf.close()
        src = self.con.execute("SELECT encryption, pdf_meta FROM sources WHERE doc_id=?", (doc["id"],)).fetchone()
        enc = src["encryption"] if src and src["encryption"] is not None else encryption_of(path)
        meta = src["pdf_meta"] if src and src["pdf_meta"] is not None else json.dumps(pdf_metadata(path))
        self.con.execute("UPDATE sources SET page_count=?, encryption=?, pdf_meta=? WHERE doc_id=?",
                         (page_count, enc, meta, doc["id"]))

        if not self.con.execute("SELECT 1 FROM batches WHERE item_id=? LIMIT 1", (item["id"],)).fetchone():
            try:
                pages = parse_page_spec(item["page_spec"], page_count)
            except PageSpecError as exc:
                self._set_item(item["id"], status="failed", sha256=fp.sha256, page_count=page_count, error=str(exc))
                return
            with store.tx(self.con):
                for seq, part in enumerate(chunk(pages, self.s.batch_pages)):
                    self.con.execute("INSERT INTO batches (item_id, seq, pages, status, updated) VALUES (?,?,?,?,?)",
                                     (item["id"], seq, json.dumps(part), "pending", time.time()))
                self._set_item(item["id"], pages_total=len(pages))
        self._set_item(item["id"], sha256=fp.sha256, page_count=page_count,
                       warning="; ".join(warnings) or item["warning"])

        for batch in self.con.execute("SELECT * FROM batches WHERE item_id=? AND status='pending' ORDER BY seq",
                                      (item["id"],)).fetchall():
            self._check_stop(job["id"])
            self._check_disk()
            if not unchanged(path, fp):
                self._set_item(item["id"], status="failed",
                               error="source changed during processing; use Retry to process the new version")
                return
            self.run_batch(job, item, doc, batch, path, fp)

        failed = self.con.execute("SELECT COUNT(*) FROM batches WHERE item_id=? AND status='done_with_errors'",
                                  (item["id"],)).fetchone()[0]
        self._set_item(item["id"], status="done_with_errors" if failed else "done", error=None)

    # ------------------------------------------------------------ batch
    def run_batch(self, job, item, doc, batch, path: Path, fp: Fingerprint) -> None:
        pages = json.loads(batch["pages"])
        self.con.execute("UPDATE batches SET status='running', attempts=attempts+1, updated=? WHERE id=?",
                         (time.time(), batch["id"]))
        tmp = self.s.tmp_dir / f"job{job['id']}_batch{batch['id']}"
        keep_temp = self.s.keep_temp_days is not None
        cache: dict = {}
        any_failed, reused = False, 0
        try:
            try:
                pdf = open_pdf(path)
            except (FileNotFoundError, PdfOpenError) as exc:
                raise SourceUnavailable("source not reachable when opening batch") from exc
            try:
                for page_no in pages:
                    self._check_stop(job["id"])
                    self.heartbeat({"state": "processing", "job": job["id"], "item": item["id"], "doc": doc["id"],
                                    "page": page_no})
                    ok, was_reused = self.process_page(job, pdf, page_no, path, fp, tmp, cache, keep_temp)
                    any_failed |= not ok
                    reused += was_reused
            finally:
                for fh in cache.get("_close", []):
                    fh.close()
                pdf.close()
        except (SourceUnavailable, _Stop):
            self._refresh_counts(item["id"])
            raise
        finally:
            if not keep_temp:
                shutil.rmtree(tmp, ignore_errors=True)
        # Counted once, when the batch completes (pages finished before an interruption count as reused).
        self.con.execute("UPDATE job_items SET pages_reused=pages_reused+? WHERE id=?", (reused, item["id"]))
        self.con.execute("UPDATE batches SET status=?, error=NULL, updated=? WHERE id=?",
                         ("done_with_errors" if any_failed else "done", time.time(), batch["id"]))
        self._refresh_counts(item["id"])
        log.info("job %d item %d batch %d %s (%d pages)", job["id"], item["id"], batch["seq"],
                 "done_with_errors" if any_failed else "done", len(pages))

    def _refresh_counts(self, item_id: int) -> None:
        """Recompute item progress from the page-result index (the source of truth)."""
        item = self.con.execute("SELECT * FROM job_items WHERE id=?", (item_id,)).fetchone()
        rows = self.con.execute("SELECT pages, status FROM batches WHERE item_id=?", (item_id,)).fetchall()
        done_pages = [p for r in rows if r["status"] in ("done", "done_with_errors") for p in json.loads(r["pages"])]
        failed = 0
        if done_pages and item["sha256"]:
            q = ",".join("?" * len(done_pages))
            failed = self.con.execute(
                f"""SELECT COUNT(DISTINCT page_no) FROM page_results WHERE sha256=? AND status='failed'
                    AND page_no IN ({q})""", (item["sha256"], *done_pages)).fetchone()[0]
        self._set_item(item_id, pages_done=len(done_pages), pages_failed=failed)

    # ------------------------------------------------------------ page
    def route(self, page, sha256: str, page_no: int) -> tuple[str, str, list[str], dict]:
        row = self.con.execute("SELECT * FROM page_routes WHERE sha256=? AND page_no=? AND routing_version=?",
                               (sha256, page_no, self.rules.version)).fetchone()
        if row:
            ev = json.loads(row["evidence"])
            return row["method"], row["reason"], ev.pop("_flags", []), ev
        ev = routing.page_evidence(page)
        method, reason, flags = routing.decide(ev, self.rules)
        self.con.execute("INSERT OR REPLACE INTO page_routes VALUES (?,?,?,?,?,?)",
                         (sha256, page_no, self.rules.version, json.dumps({**ev, "_flags": flags}), method, reason))
        return method, reason, flags, ev

    def process_page(self, job, pdf, page_no: int, path: Path, fp: Fingerprint, tmp: Path, cache: dict,
                     keep_temp: bool) -> tuple[bool, bool]:
        """Run the routed (or overridden) engine on one page. Returns (chosen engine succeeded, all reused)."""
        page = pdf[page_no - 1]
        try:
            auto, reason, flags, _ = self.route(page, fp.sha256, page_no)
            chosen = job["method_override"] or store.latest_override(self.con, fp.sha256, page_no) or auto
            if chosen not in self.engines:
                chosen = auto
            # The digital text layer is always kept as well (it is cheap and preserves original text).
            methods = ["textlayer"] + ([chosen] if chosen != "textlayer" else [])
            ok, all_reused = True, True
            for m in methods:
                status = self._run_engine(m, page, page_no, path, fp, tmp, cache, keep_temp, flags, job["id"])
                all_reused &= status == "reused"
                if m == chosen and status == "failed":
                    ok = False
            return ok, all_reused
        finally:
            page.close()

    def _run_engine(self, name, page, page_no, path, fp, tmp, cache, keep_temp, route_flags, job_id) -> str:
        """Returns 'reused', 'ok' or 'failed'."""
        eng = self.engines[name]
        existing = store.find_result(self.con, fp.sha256, page_no, name, eng.version, eng.settings_hash)
        if existing and existing["status"] == "ok":
            return "reused"  # same source bytes, engine version and settings
        out = self.s.results_dir / fp.sha256 / name / eng.settings_hash / f"p{page_no:06d}.json"
        ctx = PageContext(page=page, page_no=page_no, tmp_dir=tmp, pdf_path=path, cache=cache,
                          should_stop=lambda: self._control(job_id) is not None, keep_temp=keep_temp)
        t0 = time.monotonic()
        try:
            avail, why = eng.available()
            if not avail:
                raise RuntimeError(why)
            res = eng.extract_page(ctx)
        except EngineCancelled:
            self._check_stop(job_id)
            raise
        except (EngineTimeout, Exception) as exc:  # one page must never stop the batch
            if not path.parent.exists():
                raise SourceUnavailable("source folder disappeared during processing") from exc
            err = "timeout" if isinstance(exc, EngineTimeout) else sanitize(exc)
            store.record_result(self.con, fp.sha256, page_no, name, eng.version, eng.settings_hash, "failed",
                                None, 0, ["engine_failed"], err)
            log.warning("page %d engine %s failed: %s", page_no, name, err)
            return "failed"
        vflags = validation_flags(name, res, route_flags)
        store.write_json_atomic(out, {
            "schema": RESULT_SCHEMA, "sha256": fp.sha256, "page_no": page_no,
            "engine": name, "engine_version": eng.version, "settings_hash": eng.settings_hash,
            "engine_settings": eng.settings(), "created": time.time(),
            "seconds": round(time.monotonic() - t0, 3),
            "text": res.text, "blocks": [b.to_dict() for b in res.blocks], "bbox_space": res.bbox_space,
            "reading_order": res.reading_order, "stats": res.stats, "warnings": res.warnings,
            "validation_flags": vflags,
        })
        store.record_result(self.con, fp.sha256, page_no, name, eng.version, eng.settings_hash, "ok", str(out),
                            res.stats.get("chars", 0), vflags, None)
        return "ok"


def main(argv=None) -> int:
    ap = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    ap.add_argument("--once", action="store_true", help="process runnable jobs, then exit")
    ap.add_argument("--settings", type=Path, help="settings.toml (default: $AMIRA_PILOT_SETTINGS or ~/AmiraPilotData)")
    args = ap.parse_args(argv)
    s = load_settings(args.settings)
    setup(s.logs_dir, "worker", console=True)
    w = Worker(s)
    if args.once:
        w.acquire_lock()
        w.run_once()
    else:
        w.run_forever()
    return 0


if __name__ == "__main__":
    sys.exit(main())
