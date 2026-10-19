"""Local browser-based PDF extraction pilot.

The extraction engine (inventory, store, engines, routing, worker, exports) is independent of
the web interface in ``pilot.web``. Everything runs on this Mac; runtime data lives in the
data root from ``settings.toml`` (default ``~/AmiraPilotData``), outside the source repository.
"""

__version__ = "0.1.0"
