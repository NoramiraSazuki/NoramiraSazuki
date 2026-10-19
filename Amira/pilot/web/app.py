"""Local web interface (FastAPI). Binds to 127.0.0.1 only; see pilot.__main__.

Protection for a localhost app: the Host header must be the local address (blocks DNS
rebinding), state-changing requests need a per-launch token that only the served page knows,
API docs routes are off, and no external scripts, fonts or analytics are loaded.
"""

from __future__ import annotations

import json
import secrets
import shutil
import threading
import time
from concurrent.futures import ThreadPoolExecutor
from concurrent.futures import TimeoutError as FutTimeout
from pathlib import Path

from fastapi import Body, FastAPI, HTTPException, Request
from fastapi.responses import FileResponse, HTMLResponse, JSONResponse, PlainTextResponse

from .. import inventory, jobs, routing, store, titles
from ..engines import MODES, build_engines
from ..exports import export_job
from ..identity import stat_source
from ..logutil import sanitize
from ..pages import PageSpecError, parse_page_spec
from ..pdfio import open_pdf, render_page

STATIC = Path(__file__).resolve().parent / "static"
_render_lock = threading.Lock()  # pdfium is not thread-safe
_io_pool = ThreadPoolExecutor(max_workers=2)


def _with_timeout(fn, seconds: float, default):
    """Filesystem checks on a dropped NAS can hang; never let them block the UI."""
    fut = _io_pool.submit(fn)
    try:
        return fut.result(timeout=seconds)
    except (FutTimeout, OSError):
        return default


def create_app(settings, allowed_hosts: set[str] | None = None, on_shutdown=None) -> FastAPI:
    settings.ensure_dirs()
    app = FastAPI(docs_url=None, redoc_url=None, openapi_url=None)
    token = secrets.token_urlsafe(24)
    app.state.token = token
    hosts = allowed_hosts or {f"127.0.0.1:{settings.port}", f"localhost:{settings.port}"}
    engines = build_engines(settings)
    rules = routing.load_rules(routing.DEFAULT_RULES, set(engines))
    local = threading.local()

    def db():
        if not hasattr(local, "con"):
            local.con = store.init_db(settings.db_path)
        return local.con

    @app.middleware("http")
    async def guard(request: Request, call_next):
        if request.headers.get("host", "") not in hosts:
            return PlainTextResponse("Forbidden host", status_code=403)
        if request.method not in ("GET", "HEAD") and request.headers.get("x-pilot-token") != token:
            return PlainTextResponse("Missing or wrong token", status_code=403)
        resp = await call_next(request)
        resp.headers["Cache-Control"] = "no-store"
        resp.headers["X-Content-Type-Options"] = "nosniff"
        resp.headers["Content-Security-Policy"] = ("default-src 'self'; img-src 'self' data:; "
                                                   "style-src 'self'; script-src 'self'; connect-src 'self'")
        return resp

    def err(exc: Exception, code: int = 400):
        raise HTTPException(code, sanitize(exc) if not isinstance(exc, (jobs.JobError, inventory.InventoryError))
                            else str(exc))

    # ------------------------------------------------------------ pages
    @app.get("/", response_class=HTMLResponse)
    def index():
        html = (STATIC / "index.html").read_text(encoding="utf-8")
        return html.replace("{{TOKEN}}", token)

    @app.get("/static/{name}")
    def static(name: str):
        path = (STATIC / name).resolve()
        if path.parent != STATIC or not path.is_file():
            raise HTTPException(404)
        return FileResponse(path)

    # ------------------------------------------------------------ status
    @app.get("/api/state")
    def state():
        con = db()
        hb = float(store.meta_get(con, "worker_heartbeat", "0") or 0)
        activity = json.loads(store.meta_get(con, "worker_activity", "{}") or "{}")
        free_gb = shutil.disk_usage(settings.data_root).free / 1e9
        return {
            "worker_alive": time.time() - hb < 15, "worker_activity": activity,
            "disk_free_gb": round(free_gb, 1), "min_free_gb": settings.min_free_gb,
            "source_available": _with_timeout(settings.source_root.is_dir, 3, False),
            "source_root": str(settings.source_root), "data_root": str(settings.data_root),
            "inventory_loaded": store.meta_get(con, "inventory_loaded"),
            "modes": MODES, "methods": {n: e.label for n, e in engines.items()},
            "routing_version": rules.version,
        }

    # ------------------------------------------------------------ catalogue
    @app.post("/api/inventory/reload")
    def reload_inventory():
        try:
            inv = _with_timeout(lambda: inventory.load(settings), 120, None)
        except inventory.InventoryError as exc:
            err(exc)
        if inv is None:
            raise HTTPException(504, "inventory or source folder did not respond in time")
        store.upsert_documents(db(), inv, _with_timeout(lambda: titles.load(settings), 120, None))
        return inventory_report()

    @app.get("/api/inventory/report")
    def inventory_report():
        con = db()
        counts = {r[0]: r[1] for r in con.execute(
            "SELECT resolve_status, COUNT(*) FROM documents WHERE in_inventory=1 GROUP BY resolve_status")}
        totals = con.execute("SELECT COUNT(*) n, SUM(inv_pages) pages, SUM(inv_size) size FROM documents "
                             "WHERE in_inventory=1").fetchone()
        return {"counts": counts, "documents": totals["n"], "pages": totals["pages"], "bytes": totals["size"],
                "unlisted": json.loads(store.meta_get(con, "unlisted", "[]")),
                "titles": json.loads(store.meta_get(con, "titles_report", "{}") or "{}"),
                "with_title": con.execute("SELECT COUNT(*) FROM documents WHERE in_inventory=1 AND title IS NOT NULL")
                .fetchone()[0],
                "source_available": store.meta_get(con, "source_available") == "1"}

    @app.get("/api/documents")
    def documents():
        rows = db().execute(
            """SELECT d.id, d.rel_path, d.file_name, d.title, d.resolve_status, d.note, d.inv_row_id, d.inv_pages, d.inv_size,
                      d.inv_ingest_status, d.inv_duplicate_group, d.inv_copies, s.page_count, s.sha256 IS NOT NULL
                      AS processed_before,
                      (SELECT i.status FROM job_items i WHERE i.doc_id=d.id ORDER BY i.id DESC LIMIT 1) AS last_status,
                      (SELECT i.job_id FROM job_items i WHERE i.doc_id=d.id ORDER BY i.id DESC LIMIT 1) AS last_job
               FROM documents d LEFT JOIN sources s ON s.doc_id=d.id WHERE d.in_inventory=1
               ORDER BY d.rel_path""").fetchall()
        return [dict(r) for r in rows]

    # ------------------------------------------------------------ jobs
    @app.post("/api/jobs")
    def create_job(payload: dict = Body(...)):
        try:
            sel = [(int(i["doc_id"]), str(i.get("pages") or "all")) for i in payload.get("items", [])]
            job_id = jobs.create_job(db(), sel, payload.get("mode", "text_metadata"))
        except (jobs.JobError, KeyError, ValueError, TypeError) as exc:
            err(exc)
        return {"job_id": job_id}

    @app.get("/api/jobs")
    def list_jobs():
        ids = [r[0] for r in db().execute("SELECT id FROM jobs ORDER BY id DESC LIMIT 50")]
        return [jobs.job_summary(db(), i) for i in ids]

    @app.get("/api/jobs/{job_id}")
    def get_job(job_id: int):
        try:
            return jobs.job_summary(db(), job_id)
        except TypeError:
            raise HTTPException(404, "no such job")

    @app.post("/api/jobs/{job_id}/{action}")
    def job_action(job_id: int, action: str):
        con = db()
        try:
            if action == "pause":
                jobs.pause(con, job_id)
            elif action == "resume":
                jobs.resume(con, job_id)
            elif action == "cancel":
                jobs.cancel(con, job_id)
            elif action == "retry":
                return {"requeued": jobs.retry(con, job_id)}
            elif action == "export":
                return {"export_dir": str(export_job(con, settings, job_id, engines, rules.version))}
            else:
                raise HTTPException(404, "unknown action")
        except (jobs.JobError, OSError, ValueError) as exc:
            err(exc)
        return {"ok": True}

    @app.get("/api/jobs/{job_id}/sanitized-report", response_class=PlainTextResponse)
    def sanitized_report(job_id: int):
        """Safe to paste into a support chat: ids, statuses and sanitised errors only (no names or text)."""
        try:
            j = jobs.job_summary(db(), job_id)
        except TypeError:
            raise HTTPException(404, "no such job")
        lines = [f"job {job_id}: status={j['status']} reason={sanitize(j['reason'] or '-')} "
                 f"pages={j['pages_done']}/{j['pages_total']} failed={j['pages_failed']} batches={j['batches']}"]
        for it in j["items"]:
            lines.append(f"  item {it['id']} doc {it['doc_id']}: {it['status']} pages={it['pages_done']}/"
                         f"{it['pages_total']} failed={it['pages_failed']} "
                         f"error={sanitize(it['error'] or '-')} warning={sanitize(it['warning'] or '-')}")
        errs = db().execute(
            """SELECT r.page_no, r.engine, r.error FROM page_results r JOIN job_items i ON i.sha256=r.sha256
               WHERE i.job_id=? AND r.status='failed' LIMIT 50""", (job_id,)).fetchall()
        for e in errs:
            lines.append(f"  page {e['page_no']} {e['engine']}: {sanitize(e['error'] or '-')}")
        lines.append("engines: " + ", ".join(f"{n}={e.version}" for n, e in engines.items()))
        return "\n".join(lines)

    # ------------------------------------------------------------ review
    @app.get("/api/docs/{doc_id}/pages/{page_no}")
    def page(doc_id: int, page_no: int):
        try:
            return jobs.page_view(db(), doc_id, page_no, engines, rules.version)
        except jobs.JobError as exc:
            err(exc, 404)

    @app.get("/api/docs/{doc_id}/pages/{page_no}/image")
    def page_image(doc_id: int, page_no: int):
        con = db()
        row = con.execute("""SELECT d.resolved_path, s.sha256, s.size, s.mtime_ns, s.page_count FROM documents d
                             JOIN sources s ON s.doc_id=d.id WHERE d.id=?""", (doc_id,)).fetchone()
        if not row or not row["sha256"]:
            raise HTTPException(404, "document has not been processed yet")
        if not (1 <= page_no <= (row["page_count"] or 0)):
            raise HTTPException(404, "page out of range")
        cache = settings.preview_dir / row["sha256"] / f"p{page_no:06d}_{settings.preview_dpi}.png"
        if not cache.is_file():
            path = Path(row["resolved_path"])
            st = _with_timeout(lambda: stat_source(path), 5, None)
            if st is None:
                raise HTTPException(503, "source PDF not reachable (NAS?)")
            if (st.st_size, st.st_mtime_ns) != (row["size"], row["mtime_ns"]):
                raise HTTPException(409, "source PDF changed since it was processed; re-run the job")
            with _render_lock:
                pdf = open_pdf(path)
                try:
                    pg = pdf[page_no - 1]
                    img, _ = render_page(pg, settings.preview_dpi, 12)
                    pg.close()
                finally:
                    pdf.close()
            cache.parent.mkdir(parents=True, exist_ok=True)
            tmp = cache.with_suffix(".tmp")
            img.save(tmp, format="PNG")
            tmp.replace(cache)
        return FileResponse(cache, media_type="image/png")

    @app.post("/api/docs/{doc_id}/pages/{page_no}/retry")
    def retry_page(doc_id: int, page_no: int, payload: dict = Body(...)):
        try:
            pages = parse_page_spec(payload["pages"]) if payload.get("pages") else [page_no]
            return {"job_id": jobs.retry_pages(db(), doc_id, pages, payload.get("method", ""), set(engines))}
        except (jobs.JobError, PageSpecError) as exc:
            err(exc)

    @app.post("/api/docs/{doc_id}/pages/{page_no}/review")
    def review(doc_id: int, page_no: int, payload: dict = Body(...)):
        src = db().execute("SELECT sha256 FROM sources WHERE doc_id=?", (doc_id,)).fetchone()
        if not src:
            raise HTTPException(404, "document has not been processed yet")
        try:
            rid = jobs.add_review(db(), src["sha256"], page_no, payload.get("result_id"), payload.get("status", ""),
                                  payload.get("corrected_text"), payload.get("note"))
        except jobs.JobError as exc:
            err(exc)
        return {"review_id": rid}

    @app.get("/api/processed-documents")
    def processed():
        rows = db().execute("""SELECT d.id, d.file_name, d.title, d.rel_path, s.page_count FROM documents d
                               JOIN sources s ON s.doc_id=d.id WHERE s.sha256 IS NOT NULL ORDER BY d.rel_path""")
        return [dict(r) for r in rows]

    @app.get("/api/docs/{doc_id}/processed-pages")
    def processed_pages(doc_id: int):
        src = db().execute("SELECT sha256 FROM sources WHERE doc_id=?", (doc_id,)).fetchone()
        if not src:
            return []
        return [r[0] for r in db().execute("SELECT DISTINCT page_no FROM page_results WHERE sha256=? ORDER BY 1",
                                           (src["sha256"],))]

    @app.post("/api/shutdown")
    def shutdown():
        if on_shutdown:
            threading.Timer(0.3, on_shutdown).start()
        return JSONResponse({"ok": True, "message": "Shutting down. Running work resumes on next start."})

    @app.exception_handler(Exception)
    async def unhandled(request: Request, exc: Exception):  # never echo document text or paths
        return JSONResponse({"detail": sanitize(exc)}, status_code=500)

    app.state.engines = engines
    app.state.rules = rules
    return app
