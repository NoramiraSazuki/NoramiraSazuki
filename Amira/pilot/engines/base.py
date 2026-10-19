"""Engine interface: one page in, one JSON-serialisable page result out.

Every result records the engine name, version and the settings hash that produced it, so a
result is reused only when all three (and the source SHA-256) still match.
"""

from __future__ import annotations

from dataclasses import dataclass, field
from pathlib import Path
from typing import Callable

from ..settings import settings_hash

RESULT_SCHEMA = "pilot.page_result/1"


class EngineTimeout(RuntimeError):
    pass


class EngineCancelled(RuntimeError):
    pass


@dataclass
class PageContext:
    page: object                     # pypdfium2 PdfPage (already loaded)
    page_no: int                     # ORIGINAL 1-based page number in the source PDF
    tmp_dir: Path                    # per-batch scratch folder inside the data root
    pdf_path: Path | None = None
    cache: dict = field(default_factory=dict)   # per-batch handles an engine may reuse (e.g. a pypdf reader)
    should_stop: Callable[[], bool] = lambda: False
    keep_temp: bool = False


@dataclass
class Block:
    text: str
    bbox: list[float] | None         # [x0, y0, x1, y1] in points, see PageResult.bbox_space
    kind: str = "line"               # line / word / table / figure / unknown
    conf: float | None = None        # engine-reported confidence (NOT proof of correctness)

    def to_dict(self) -> dict:
        return {"text": self.text, "bbox": self.bbox, "kind": self.kind, "conf": self.conf}


@dataclass
class PageResult:
    text: str
    blocks: list[Block]
    bbox_space: str                  # describes the coordinate system of block bboxes
    reading_order: str               # how block order was obtained (and whether it is verified)
    stats: dict = field(default_factory=dict)
    warnings: list[str] = field(default_factory=list)


class Engine:
    name: str = ""
    label: str = ""                  # shown in the UI
    implemented: bool = True

    @property
    def version(self) -> str:
        raise NotImplementedError

    def settings(self) -> dict:
        """Only the settings that change this engine's output (they form the settings hash)."""
        return {}

    @property
    def settings_hash(self) -> str:
        return settings_hash({"engine": self.name, **self.settings()})

    def available(self) -> tuple[bool, str]:
        return True, ""

    def extract_page(self, ctx: PageContext) -> PageResult:
        raise NotImplementedError
