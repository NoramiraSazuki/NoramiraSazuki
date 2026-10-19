"""Digital text layer engines: pdfium (default, with positions) and pypdf (same reader as the CLI)."""

from __future__ import annotations

import pypdf

from ..pdfio import PDFIUM_VERSION
from .base import Block, Engine, PageContext, PageResult


class PdfiumTextEngine(Engine):
    name = "textlayer"
    label = "Digital text layer (pdfium)"

    @property
    def version(self) -> str:
        return PDFIUM_VERSION

    def extract_page(self, ctx: PageContext) -> PageResult:
        tp = ctx.page.get_textpage()
        try:
            full = tp.get_text_range()
            blocks = []
            n = tp.count_rects()  # must be called before get_rect()
            for i in range(n):
                left, bottom, right, top = tp.get_rect(i)
                text = tp.get_text_bounded(left, bottom, right, top).strip()
                if text:
                    blocks.append(Block(text, [round(v, 2) for v in (left, bottom, right, top)], "line"))
        finally:
            tp.close()
        warnings = []
        if full.strip() and not blocks:
            warnings.append("text found but no positioned segments; positions unavailable")
            blocks = [Block(full.strip(), None, "unknown")]
        return PageResult(
            text=full.replace("\r\n", "\n").replace("\r", "\n"),
            blocks=blocks,
            bbox_space="pdf_user_space_points_origin_bottom_left_unrotated",
            reading_order="content-stream order (not verified; multi-column order may be wrong)",
            stats={"chars": sum(1 for c in full if not c.isspace()), "segments": len(blocks)},
            warnings=warnings,
        )


class PypdfTextEngine(Engine):
    name = "pypdf"
    label = "Digital text layer (pypdf, as used by the CLI)"

    @property
    def version(self) -> str:
        return f"pypdf-{pypdf.__version__}"

    def extract_page(self, ctx: PageContext) -> PageResult:
        from pdf_pipeline.extract import open_reader

        reader = ctx.cache.get("pypdf_reader")
        if reader is None:
            fh = open(ctx.pdf_path, "rb")  # read-only; closed by the worker at batch end
            ctx.cache.setdefault("_close", []).append(fh)
            reader = ctx.cache["pypdf_reader"] = open_reader(fh)
        text = reader.pages[ctx.page_no - 1].extract_text() or ""
        return PageResult(
            text=text,
            blocks=[Block(text.strip(), None, "unknown")] if text.strip() else [],
            bbox_space="none",
            reading_order="pypdf layout order (not verified)",
            stats={"chars": sum(1 for c in text if not c.isspace())},
        )
