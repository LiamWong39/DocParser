from __future__ import annotations

from pathlib import Path

from docx import Document

from .analyze import StyleResolver, body_font_size
from .cleanup import cleanup
from .config import load_config
from .headings import detect_and_apply_headings
from .pagebreaks import apply_page_breaks
from .report import Report
from .toc import find_and_remove_toc, insert_toc, refresh_with_libreoffice
from .xmlutil import block_list


def normalize(doc, cfg: dict | None = None, report: Report | None = None) -> Report:
    """Normalize a python-docx Document in place. Does not save or refresh the TOC."""
    cfg = cfg or load_config()
    report = report or Report()
    body = doc.element.body
    res = StyleResolver(doc)

    cleanup(doc, cfg, report)

    anchor, hints = None, []
    if cfg["toc"]["enabled"]:
        anchor, hints = find_and_remove_toc(body, cfg, res, report)

    blocks = block_list(body)
    report.body_font_size = body_size = body_font_size(blocks, res)
    toc_index = next((i for i, b in enumerate(blocks) if b is anchor), None)
    levels = detect_and_apply_headings(doc, blocks, res, body_size, hints, toc_index, cfg, report)

    if cfg["toc"]["enabled"]:
        insert_toc(doc, anchor, levels, cfg, res, report)
    elif anchor is not None:
        anchor.getparent().remove(anchor)

    apply_page_breaks(body, levels, cfg, report)

    if not report.accepted:
        report.warnings.append("No headings detected. Check the near misses and loosen the [headings] config.")
    return report


def normalize_file(input_path: str | Path, output_path: str | Path | None, cfg: dict | None = None,
                   dry_run: bool = False) -> Report:
    cfg = cfg or load_config()
    report = Report(input=str(input_path), output=str(output_path) if output_path else "(dry run)")
    doc = Document(str(input_path))
    normalize(doc, cfg, report)
    if dry_run:
        report.output = "(dry run)"
        return report
    doc.save(str(output_path))
    if report.toc_inserted and cfg["toc"]["refresh_with_libreoffice"]:
        refresh_with_libreoffice(output_path, cfg, report)
    return report
