"""Downstream results computed from stored page results only (never re-runs OCR or reads the PDF).

Changing extraction rules (dictionaries.json) and re-exporting is therefore cheap. Original
extracted text, human corrections and derived values are kept in separate fields.
"""

from __future__ import annotations

import json
import re
from dataclasses import dataclass

from pdf_pipeline.config import Rules
from pdf_pipeline.extract import PageText
from pdf_pipeline.fields import Evidence, extract_fields
from pdf_pipeline.run_pipeline import _empty_column_evidence, _evidence_row, _fill_row
from pdf_pipeline.schema import (CONFIG, FILESYSTEM, NEVER_FILLED, PDF_METADATA, PDF_STRUCTURE, RESULT_COLUMNS)
from pdf_pipeline.sections import filter_references

from . import store

STATUS = "EXTRACTION STATUS"
OCR_ENGINES = {"tesseract"}
_HEADING_RE = re.compile(r"^(?:\d+(?:\.\d+)*\.?\s+\S.{0,80}|[A-Z][A-Z0-9 ,&/\-]{3,80})$")


@dataclass
class SelectedPage:
    page_no: int
    method: str | None            # engine whose result is used
    auto_method: str | None
    override: str | None
    result_id: int | None
    status: str                   # ok / failed / not_processed / missing_file
    text: str                     # ORIGINAL extracted text of the selected result ('' if none)
    flags: list[str]
    warnings: list[str]
    stats: dict
    reading_order: str | None
    bbox_space: str | None
    blocks: list[dict]
    review_status: str            # unreviewed / approved / corrected / rejected
    corrected_text: str | None    # human correction (kept apart from the original text)
    alternatives: list[str]       # other engines with an ok result for this page

    @property
    def best_text(self) -> str:
        """Corrected text when a reviewer supplied one, else the original extraction."""
        return self.corrected_text if self.review_status == "corrected" and self.corrected_text else self.text


def select_page(con, sha256: str, page_no: int, engines: dict, routing_version: str) -> SelectedPage:
    route = con.execute("SELECT method FROM page_routes WHERE sha256=? AND page_no=? AND routing_version=?",
                        (sha256, page_no, routing_version)).fetchone()
    auto = route["method"] if route else None
    override = store.latest_override(con, sha256, page_no)
    method = override or auto
    rows = {r["engine"]: r for r in con.execute(
        "SELECT * FROM page_results WHERE sha256=? AND page_no=? ORDER BY id", (sha256, page_no))
        if r["engine"] in engines and r["engine_version"] == engines[r["engine"]].version
        and r["settings_hash"] == engines[r["engine"]].settings_hash}
    review = store.latest_review(con, sha256, page_no)
    sel = SelectedPage(page_no, method, auto, override, None, "not_processed", "", [], [], {}, None, None, [],
                       review["status"] if review else "unreviewed",
                       review["corrected_text"] if review else None,
                       [e for e, r in rows.items() if r["status"] == "ok" and e != method])
    r = rows.get(method) if method else None
    if r is None:
        return sel
    sel.result_id, sel.status, sel.flags = r["id"], r["status"], json.loads(r["flags"] or "[]")
    tl = rows.get("textlayer")
    if method != "textlayer" and tl is not None and tl["status"] == "ok" and (tl["chars"] or 0) > (r["chars"] or 0):
        # e.g. OCR missed a small digital caption: never drop content silently.
        sel.flags.append("text_layer_has_more_text_than_selected_result")
    if r["status"] == "ok":
        try:
            data = store.read_json(r["path"])
        except OSError:
            sel.status = "missing_file"
            return sel
        sel.text, sel.warnings, sel.stats = data["text"], data.get("warnings", []), data.get("stats", {})
        sel.reading_order, sel.bbox_space, sel.blocks = data.get("reading_order"), data.get("bbox_space"), \
            data.get("blocks", [])
    return sel


def chunks_for_page(doc: dict, sp: SelectedPage, max_chars: int = 1500) -> list[dict]:
    """Paragraph chunks with provenance, ready for a later embedding/vector-search step."""
    text = sp.best_text
    if not text.strip():
        return []
    paras = [p.strip() for p in re.split(r"\n\s*\n", text) if p.strip()]
    if len(paras) <= 1:
        paras = [text.strip()]
    heading, pieces = None, []
    for p in paras:
        while len(p) > max_chars:
            cut = p.rfind("\n", 0, max_chars)
            cut = cut if cut > max_chars // 3 else max_chars
            pieces.append(p[:cut])
            p = p[cut:].lstrip()
        pieces.append(p)
    out = []
    for i, piece in enumerate(pieces):
        first = piece.splitlines()[0].strip()
        if _HEADING_RE.match(first) and len(first) < 90:
            heading = first
        out.append({
            "chunk_id": f"{doc['sha256'][:12]}-p{sp.page_no}-c{i}",
            "doc_id": doc["id"], "rel_path": doc["rel_path"], "file_name": doc["file_name"],
            "sha256": doc["sha256"], "page_no": sp.page_no, "chunk_index": i,
            "heading_guess": heading, "text": piece,
            "text_origin": "human_corrected" if sp.review_status == "corrected" and sp.corrected_text else "extracted",
            "method": sp.method, "review_status": sp.review_status, "validation_flags": sp.flags,
            "reading_order": sp.reading_order,
        })
    return out


def metadata_row(doc: dict, src: dict, pages: list[SelectedPage], rules: Rules) -> tuple[dict, list[dict]]:
    """Report metadata with the CLI's rules, from the selected (or human-corrected) page texts."""
    row = {c: "" for c in RESULT_COLUMNS}
    name = doc["file_name"]
    row.update({"File Name": name, "ORIGINAL_FILE_NAME": name, "File Size": src.get("size") or "",
                "DIGITAL_SIZE_UOM": "B", "File Format": "PDF", "PAGE_COUNT": src.get("page_count") or "",
                "DECRYPTION_TYPE": src.get("encryption") or ""})
    texts = [PageText(p.page_no, p.best_text,
                      "reviewed" if p.review_status == "corrected" and p.corrected_text
                      else ("ocr" if p.method in OCR_ENGINES else "text"))
             for p in pages if p.status == "ok" or p.review_status == "corrected"]
    meta = json.loads(src.get("pdf_meta") or "{}")
    filtered = filter_references(texts, rules)
    ex = extract_fields(filtered.pages, rules, name, meta.get("title", ""))
    for attr in ("File Name", "ORIGINAL_FILE_NAME", "File Size", "DIGITAL_SIZE_UOM", "File Format"):
        ex.evidence.insert(0, Evidence(attr, str(row[attr]), [], "filesystem", "file system attribute", FILESYSTEM))
    _fill_row(row, ex)
    ex.evidence.append(Evidence("PAGE_COUNT", str(row["PAGE_COUNT"]), [], "pdf_structure", "pdfium page count",
                                PDF_STRUCTURE))
    for label, key in (("PDF Title", "title"), ("PDF Author", "author"), ("PDF Creator", "creator"),
                       ("PDF Producer", "producer"), ("PDF Creation Date", "creation_date")):
        if meta.get(key):
            ex.evidence.append(Evidence(label, meta[key], [], "pdf_metadata", "embedded metadata", PDF_METADATA))
    for col, value in rules.constant_columns.items():
        if not row[col]:
            row[col] = value
            ex.evidence.append(Evidence(col, value, [], "config", "constant_columns", CONFIG))

    processed = {p.page_no for p in pages if p.status == "ok"}
    notes = [f"pages processed {len(processed)}/{row['PAGE_COUNT'] or '?'}"]
    cover = set(range(1, rules.cover_pages + 1))
    if src.get("page_count"):
        cover &= set(range(1, src["page_count"] + 1))
    if not cover <= processed:
        notes.append("cover pages not all processed (title/parties/date may be missing)")
    n_ocr = sum(1 for p in pages if p.status == "ok" and p.method in OCR_ENGINES)
    if n_ocr:
        notes.append(f"OCR {n_ocr} pages")
    failed = [p.page_no for p in pages if p.status in ("failed", "missing_file", "not_processed")]
    if failed:
        notes.append(f"no result for {len(failed)} page(s)")
    missing = [c for c in rules.required_columns if row[c] in ("", None)]
    status = "NO TEXT" if not any(t.text.strip() for t in texts) else ("PARTIAL" if missing else "OK")
    if missing:
        notes.append("missing: " + ", ".join(missing))
    review = ex.review + ex.flag_multiple()
    if review:
        notes.append("review: " + "; ".join(review))
    notes.append("not verified by a reviewer" if not any(p.review_status in ("approved", "corrected")
                                                         for p in pages) else "some pages reviewed")
    for col in NEVER_FILLED:
        row[col] = ""
    row[STATUS] = " | ".join([status] + notes)
    _empty_column_evidence(row, ex)
    return row, [_evidence_row(name, e) for e in ex.evidence]
