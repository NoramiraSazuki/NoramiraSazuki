"""Pilot settings, read from a TOML file kept in the data root (outside the source repository).

Lookup order: ``$AMIRA_PILOT_SETTINGS`` if set, else ``~/AmiraPilotData/settings.toml``.
A commented default file is written there on first start.
"""

from __future__ import annotations

import hashlib
import json
import os
import tomllib
from dataclasses import asdict, dataclass, field
from pathlib import Path

from pdf_pipeline.safety import is_within

REPO_DIR = Path(__file__).resolve().parent.parent.parent  # .../NoramiraSazuki (git working tree)
DEFAULT_DATA_ROOT = Path.home() / "AmiraPilotData"

DEFAULT_TOML = """\
# Dulang PDF Extraction settings. Edit, save, then restart the app.
# Runtime data (database, page results, previews, exports, logs) is kept under data_root,
# which must be OUTSIDE the git repository.

[paths]
data_root = "{data_root}"
inventory_workbook = "{repo}/07_DulangOSS2026_Github/Dulang_OSS_All_Batches_File_Inventory.xlsx"
source_root = "/Volumes/homes/public/DulangOSS"   # all batch folders (20260925, 20261006, ...)

[inventory]
sheet = "File Inventory"
filter_column = ""           # e.g. "Sub-folder" to limit the catalogue to one folder ...
filter_value = ""            # ... e.g. "Seismic Report"
path_prefix = ""             # inventory File Path is relative to source_root (starts with the batch folder)
extensions = ["pdf"]

[titles]
# Catalogue titles from a mapping workbook (read-only), joined by relative path in the location column.
title_workbook = "/Users/noramirasazuki/Library/CloudStorage/OneDrive-MRCHURROSF&BSDN.BHD/Working/08_STE/07_DulangOSS2026/02_Data Mapping/Dulang_DIRt_v1.xlsx"
title_sheet = "*"            # "*" = every sheet that has the location column
title_column = "H"
title_location_header = "File Location"

[processing]
batch_pages = 20             # pages per checkpointed batch
ocr_lang = "eng+msa"         # Tesseract languages
ocr_dpi = 300                # OCR render resolution (reduced automatically for huge pages)
max_render_megapixels = 40   # cap for any single rendered page image
ocr_page_timeout = 180       # seconds per page before OCR is abandoned for that page
min_free_gb = 15             # jobs pause when free disk space drops below this
large_job_pages = 1000       # jobs above this many pages need a second confirmation
scan_folder_timeout = 60     # seconds before a stalled NAS folder listing is skipped and reported
temp_retention = "delete_after_page"   # or "keep_days:3" to keep OCR images for troubleshooting
preview_dpi = 110            # review-view page images
preview_cache_days = 7       # rendered previews older than this are deleted at start-up

[server]
port = 8765
"""


@dataclass
class Settings:
    data_root: Path
    inventory_workbook: Path
    source_root: Path
    sheet: str = "File Inventory"
    filter_column: str = ""
    filter_value: str = ""
    path_prefix: str = ""
    extensions: list = field(default_factory=lambda: ["pdf"])
    large_job_pages: int = 1000
    scan_folder_timeout: float = 60
    title_workbook: str = "/Users/noramirasazuki/Library/CloudStorage/OneDrive-MRCHURROSF&BSDN.BHD/Working/08_STE/07_DulangOSS2026/02_Data Mapping/Dulang_DIRt_v1.xlsx"
    title_sheet: str = "*"
    title_column: str = "H"
    title_location_header: str = "File Location"
    batch_pages: int = 20
    ocr_lang: str = "eng+msa"
    ocr_dpi: int = 300
    max_render_megapixels: float = 40
    ocr_page_timeout: int = 180
    min_free_gb: float = 15
    temp_retention: str = "delete_after_page"
    preview_dpi: int = 110
    preview_cache_days: int = 7
    port: int = 8765
    settings_file: Path | None = field(default=None, compare=False)

    # ---- derived locations (all inside data_root)
    @property
    def db_path(self) -> Path:
        return self.data_root / "pilot.sqlite3"

    @property
    def results_dir(self) -> Path:
        return self.data_root / "results"

    @property
    def tmp_dir(self) -> Path:
        return self.data_root / "tmp"

    @property
    def preview_dir(self) -> Path:
        return self.data_root / "cache" / "previews"

    @property
    def exports_dir(self) -> Path:
        return self.data_root / "exports"

    @property
    def logs_dir(self) -> Path:
        return self.data_root / "logs"

    @property
    def keep_temp_days(self) -> int | None:
        """None = delete each page image right after OCR."""
        if self.temp_retention.startswith("keep_days:"):
            return int(self.temp_retention.split(":", 1)[1])
        return None

    def validate(self) -> None:
        if self.batch_pages < 1:
            raise ValueError("batch_pages must be >= 1")
        if not (72 <= self.ocr_dpi <= 600):
            raise ValueError("ocr_dpi must be between 72 and 600")
        if self.temp_retention != "delete_after_page" and self.keep_temp_days is None:
            raise ValueError("temp_retention must be 'delete_after_page' or 'keep_days:N'")
        if is_within(self.data_root, REPO_DIR):
            raise ValueError(f"data_root must be outside the source repository ({REPO_DIR})")
        if is_within(self.data_root, self.source_root):
            raise ValueError("data_root must not be inside the source PDF folder")

    def ensure_dirs(self) -> None:
        for d in (self.data_root, self.results_dir, self.tmp_dir, self.preview_dir, self.exports_dir,
                  self.logs_dir):
            d.mkdir(parents=True, exist_ok=True)

    def public_dict(self) -> dict:
        d = asdict(self)
        return {k: str(v) if isinstance(v, Path) else v for k, v in d.items()}


def settings_path() -> Path:
    env = os.environ.get("AMIRA_PILOT_SETTINGS")
    return Path(env).expanduser() if env else DEFAULT_DATA_ROOT / "settings.toml"


def write_default(path: Path, data_root: Path = DEFAULT_DATA_ROOT) -> None:
    path.parent.mkdir(parents=True, exist_ok=True)
    with open(path, "x", encoding="utf-8") as fh:  # never overwrite an edited settings file
        fh.write(DEFAULT_TOML.format(data_root=data_root, repo=REPO_DIR))


def load_settings(path: Path | None = None) -> Settings:
    path = path or settings_path()
    if not path.is_file():
        write_default(path)
    with open(path, "rb") as fh:
        cfg = tomllib.load(fh)
    p, inv, proc, srv, tit = (cfg.get(k, {}) for k in ("paths", "inventory", "processing", "server", "titles"))
    s = Settings(
        data_root=Path(p.get("data_root", DEFAULT_DATA_ROOT)).expanduser(),
        inventory_workbook=Path(p["inventory_workbook"]).expanduser(),
        source_root=Path(p["source_root"]).expanduser(),
        **{k: v for k, v in {**inv, **proc, **srv, **tit}.items() if k in Settings.__dataclass_fields__},
        settings_file=path,
    )
    s.validate()
    return s


def settings_hash(values: dict) -> str:
    """Stable short hash of the settings that affect an engine's output."""
    blob = json.dumps(values, sort_keys=True, default=str).encode()
    return hashlib.sha256(blob).hexdigest()[:16]
