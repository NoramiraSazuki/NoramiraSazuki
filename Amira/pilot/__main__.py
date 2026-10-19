"""Start the local pilot: background worker + web interface on http://127.0.0.1:<port>/

    cd ~/Documents/GitHub/NoramiraSazuki/Amira
    caffeinate -i ../.venv/bin/python -m pilot          (or double-click start_pilot.command)

Stop with Ctrl+C in this window or the "Shut down" button. Unfinished work resumes next time.
Closing the browser does NOT stop processing.
"""

from __future__ import annotations

import argparse
import signal
import socket
import subprocess
import sys
import threading
import time
import webbrowser
from pathlib import Path

from . import inventory, store, titles
from .logutil import sanitize, setup
from .settings import load_settings

PACKAGE_PARENT = Path(__file__).resolve().parent.parent


class WorkerSupervisor:
    """Runs ``python -m pilot.worker`` as a child process and restarts it if it dies unexpectedly."""

    def __init__(self, settings_file: Path, log):
        self.settings_file, self.log = settings_file, log
        self.proc: subprocess.Popen | None = None
        self.stopping = False
        self.restarts = 0

    def start(self) -> None:
        self.proc = subprocess.Popen([sys.executable, "-m", "pilot.worker", "--settings", str(self.settings_file)],
                                     cwd=PACKAGE_PARENT, stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL)
        threading.Thread(target=self._watch, daemon=True).start()

    def _watch(self) -> None:
        while not self.stopping:
            code = self.proc.wait()
            if self.stopping:
                return
            self.restarts += 1
            self.log.warning("worker exited (code %s); restart %d", code, self.restarts)
            if self.restarts > 5:
                self.log.error("worker keeps exiting; not restarting. See logs/worker.log")
                return
            time.sleep(3)
            self.proc = subprocess.Popen([sys.executable, "-m", "pilot.worker", "--settings", str(self.settings_file)],
                                         cwd=PACKAGE_PARENT, stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL)

    def stop(self) -> None:
        self.stopping = True
        if self.proc and self.proc.poll() is None:
            self.proc.terminate()  # checkpoints mean at most the current page is redone next time
            try:
                self.proc.wait(timeout=10)
            except subprocess.TimeoutExpired:
                self.proc.kill()


def _port_free(port: int) -> bool:
    with socket.socket() as s:
        return s.connect_ex(("127.0.0.1", port)) != 0


def main(argv=None) -> int:
    ap = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    ap.add_argument("--settings", type=Path)
    ap.add_argument("--no-browser", action="store_true")
    args = ap.parse_args(argv)

    try:
        s = load_settings(args.settings)
    except (OSError, ValueError, KeyError) as exc:
        print(f"Settings error: {exc}", file=sys.stderr)
        return 2
    s.ensure_dirs()
    log = setup(s.logs_dir, "app")
    if not _port_free(s.port):
        print(f"Port {s.port} is in use - is the pilot already running? Open http://127.0.0.1:{s.port}/",
              file=sys.stderr)
        return 1

    con = store.init_db(s.db_path)
    try:
        store.upsert_documents(con, inventory.load(s), titles.load(s))
        log.info("inventory loaded")
    except Exception as exc:  # the UI can still start; reload from the Catalogue tab
        log.warning("inventory not loaded: %s", sanitize(exc))
    con.close()

    import uvicorn

    from .web.app import create_app

    worker = WorkerSupervisor(s.settings_file, log)
    server: uvicorn.Server | None = None

    def shutdown():
        if server:
            server.should_exit = True

    app = create_app(s, on_shutdown=shutdown)
    server = uvicorn.Server(uvicorn.Config(app, host="127.0.0.1", port=s.port, log_level="warning",
                                           access_log=False, server_header=False))
    worker.start()
    url = f"http://127.0.0.1:{s.port}/"
    print(f"\nAmira PDF pilot running at {url}\nData folder: {s.data_root}\n"
          "Keep this window open. Stop with Ctrl+C (work resumes next start).\n")
    if not args.no_browser:
        threading.Timer(1.0, lambda: webbrowser.open(url)).start()
    signal.signal(signal.SIGTERM, lambda *_: shutdown())
    try:
        server.run()
    finally:
        worker.stop()
        log.info("stopped")
    return 0


if __name__ == "__main__":
    sys.exit(main())
