"""Engine registry. Modes that are not implemented are listed so the UI can show them disabled."""

from __future__ import annotations

from .base import Engine
from .tesseract import TesseractEngine
from .textlayer import PdfiumTextEngine, PypdfTextEngine

# Extraction modes shown in the UI. Only "implemented" ones can be selected.
MODES = [
    {"id": "text_metadata", "label": "Text and report metadata", "implemented": True,
     "help": "Digital text where reliable, local OCR where the page is image-like; report metadata derived "
             "from the text. Processing success is not verified accuracy - review pages."},
    {"id": "tables", "label": "Table extraction", "implemented": False,
     "help": "Not implemented yet (Stage 3). Tables are kept as plain text for now."},
    {"id": "multimodal", "label": "Local multimodal retry / comparison", "implemented": False,
     "help": "Not implemented yet (Stage 3)."},
]

# Test hooks may register extra engines here (e.g. a fake engine that fails on purpose).
EXTRA_ENGINES: dict[str, Engine] = {}


def build_engines(settings) -> dict[str, Engine]:
    engines: dict[str, Engine] = {
        "textlayer": PdfiumTextEngine(),
        "pypdf": PypdfTextEngine(),
        "tesseract": TesseractEngine(settings.ocr_lang, settings.ocr_dpi, settings.max_render_megapixels,
                                     settings.ocr_page_timeout),
    }
    engines.update(EXTRA_ENGINES)
    return engines
