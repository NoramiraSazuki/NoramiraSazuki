"""Background catalogue refresh: read the inventory + titles, list the NAS, update the database.

Listing the whole share can take minutes over SMB, so it always runs in a background thread;
progress is kept in the ``meta`` table (key ``scan_state``) for the UI. Only one scan runs at a time.
"""

from __future__ import annotations

import json
import logging
import threading
import time

from . import inventory, store, titles
from .logutil import sanitize

log = logging.getLogger("pilot")
_lock = threading.Lock()


def _set(con, **state) -> None:
    store.meta_set(con, "scan_state", json.dumps(state))


def scan_state(con) -> dict:
    return json.loads(store.meta_get(con, "scan_state", "{}") or "{}")


def refresh(settings) -> dict:
    """Run one refresh in the calling thread. Returns the final scan state.

    Phase 1 (seconds): every inventory PDF goes into the catalogue at once, as NOT_CHECKED unless
    it was already confirmed. Phase 2 (minutes): the NAS is listed folder by folder and each file's
    status is updated as soon as its folder has been read.
    """
    if not _lock.acquire(blocking=False):
        return {"running": True, "message": "a scan is already running"}
    con = store.init_db(settings.db_path)
    started = time.time()
    interrupted = bool(scan_state(con).get("running"))  # app was closed during the previous scan
    last = [0.0]
    checked = [0]

    def progress(folders: int, files: int) -> None:
        if time.time() - last[0] > 1.0:  # at most one database write per second
            last[0] = time.time()
            _set(con, running=True, started=started, folders=folders, files=files, checked=checked[0],
                 phase="checking NAS", previous_interrupted=interrupted)

    def on_checked(rows) -> None:
        store.update_resolution(con, rows)
        checked[0] += len(rows)

    try:
        _set(con, running=True, started=started, phase="reading inventory", previous_interrupted=interrupted)
        raw, header_row = inventory.read_inventory(settings.inventory_workbook, settings.sheet,
                                                   settings.filter_column, settings.filter_value,
                                                   tuple(settings.extensions))
        rows = inventory.prepare(raw, settings.source_root, settings.path_prefix)
        t = titles.load(settings)
        previous_folders = set(json.loads(store.meta_get(con, "folders", "[]") or "[]"))  # keep empty NAS folders
        first = inventory.InventoryResult(rows, json.loads(store.meta_get(con, "unlisted", "[]") or "[]"),
                                          True, header_row, sorted(inventory.inventory_folders(rows) | previous_folders))
        result = store.upsert_documents(con, first, t, settings.db_path)
        _set(con, running=True, started=started, folders=0, files=0, checked=0, phase="checking NAS",
             previous_interrupted=interrupted)

        inv = inventory.check_on_disk(rows, settings.source_root, tuple(settings.extensions), progress,
                                      settings.scan_folder_timeout, on_checked)
        state = {"running": False, "started": started, "finished": time.time(), "folders": len(inv.folders),
                 "documents": len(rows), "source_available": inv.source_available, "checked": checked[0],
                 "rekeyed": result["rekeyed"], "failed_folders": len(inv.failed_folders),
                 "previous_interrupted": interrupted}
        if not inv.source_available:
            state["error"] = "source folder not reachable - files listed from the inventory are NOT CHECKED yet"
        else:
            store.set_folders(con, inv.folders, inv.unlisted)
            if inv.failed_folders:
                state["error"] = (f"{len(inv.failed_folders)} folder(s) could not be listed (NAS error); "
                                  "their files were not re-checked - use Reload inventory to try again")
        store.meta_set(con, "source_available", str(int(inv.source_available)))
        store.meta_set(con, "inventory_loaded", str(time.time()))
        _set(con, **state)
        log.info("catalogue refreshed: %d documents, %d folders, %d checked, %d re-keyed, %.0fs", len(rows),
                 len(inv.folders), checked[0], state["rekeyed"], time.time() - started)
        return state
    except Exception as exc:  # reported in the UI; the previous catalogue stays usable
        state = {"running": False, "started": started, "finished": time.time(), "error": sanitize(exc)}
        _set(con, **state)
        log.warning("catalogue refresh failed: %s", sanitize(exc))
        return state
    finally:
        con.close()
        _lock.release()


def start_background(settings) -> bool:
    """Start a refresh thread unless one is running. Returns False if one was already running."""
    if _lock.locked():
        return False
    threading.Thread(target=refresh, args=(settings,), daemon=True, name="catalogue-scan").start()
    return True
