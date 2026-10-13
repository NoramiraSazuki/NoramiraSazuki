"""Local PDF OCR, metadata extraction and CSV export pipeline.

Everything runs on this machine. Source PDFs are opened read-only and never
modified; OCR works on a temporary copy inside the workspace ``tmp/`` folder.
"""

from pathlib import Path

WORKSPACE = Path(__file__).resolve().parent.parent
