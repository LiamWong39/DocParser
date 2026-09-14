"""Start every top-level section and every table on a new page."""

from __future__ import annotations

import re

from docx.oxml.ns import qn

from .report import Report
from .xmlutil import (
    PPR_SEQ, TRPR_SEQ, block_list, get_or_add_ppr, has_ppr_flag, has_section_break, insert_ordered, is_blank,
    local, make, para_text, set_ppr_flag,
)


def apply_page_breaks(body, levels: dict, cfg: dict, report: Report):
    pb = cfg["page_breaks"]
    max_words = cfg["headings"]["max_words"]
    caption_re = re.compile(pb["caption_regex"], re.I)
    blocks = block_list(body)

    def at_top_of_page(i: int) -> bool:
        """Nothing but blank paragraphs since the document start or a new-page section break."""
        j = i - 1
        while j >= 0 and is_blank(blocks[j]):
            j -= 1
        return j < 0 or (local(blocks[j]) == "p" and has_section_break(blocks[j]))

    limit = pb["section_break_level"]
    for i, b in enumerate(blocks):
        level = levels.get(b)
        if level is not None and limit and level <= limit and not at_top_of_page(i):
            set_ppr_flag(b, "pageBreakBefore")
            report.section_breaks.append(" ".join(para_text(b).split()))

    if not pb["table_page_break"]:
        return
    for i, tbl in enumerate(blocks):
        if local(tbl) != "tbl":
            continue
        rows = tbl.findall(qn("w:tr"))
        if pb["repeat_table_header_row"] and len(rows) > 1:
            _repeat_header_row(rows[0])
        if len(rows) < pb["table_min_rows"]:
            continue

        # Lines that belong with the table: its heading, caption, or an "as follows:" lead-in.
        chain, top_idx, j = [], None, i - 1
        while j >= 0 and len(chain) < pb["table_leadin_max"]:
            prev = blocks[j]
            if is_blank(prev):
                j -= 1
                continue
            if local(prev) != "p":
                break
            text = " ".join(para_text(prev).split())
            if prev in levels or caption_re.match(text) or (text.endswith(":") and len(text.split()) <= 2 * max_words):
                chain.append(prev)
                top_idx = j
                j -= 1
            else:
                break
        for p in chain:
            set_ppr_flag(p, "keepNext")

        if chain:
            top = chain[-1]
            if has_ppr_flag(top, "pageBreakBefore") or at_top_of_page(top_idx):
                continue
            set_ppr_flag(top, "pageBreakBefore")
        else:
            if at_top_of_page(i):
                continue
            prev = tbl.getprevious()
            if prev is not None and is_blank(prev):
                brk = prev
            else:
                brk = make("p")
                tbl.addprevious(brk)
            _make_break_paragraph(brk)
        report.table_breaks.append(_table_label(tbl, rows))


def _make_break_paragraph(p):
    """A near-invisible paragraph that only starts a new page."""
    ppr = get_or_add_ppr(p)
    insert_ordered(ppr, make("pageBreakBefore"), PPR_SEQ)
    insert_ordered(ppr, make("spacing", before=0, after=0, line=20, lineRule="exact"), PPR_SEQ)
    rpr = make("rPr")
    rpr.append(make("sz", val=2))
    insert_ordered(ppr, rpr, PPR_SEQ)


def _repeat_header_row(tr):
    trpr = tr.find(qn("w:trPr"))
    if trpr is None:
        trpr = make("trPr")
        first = tr[0] if len(tr) else None
        tr.insert(1 if first is not None and local(first) == "tblPrEx" else 0, trpr)
    insert_ordered(trpr, make("tblHeader"), TRPR_SEQ)


def _table_label(tbl, rows) -> str:
    first = rows[0] if rows else None
    cells = [" ".join(para_text(tc).split()) for tc in first.findall(qn("w:tc"))] if first is not None else []
    label = " | ".join(c for c in cells if c)[:80] or "(empty first row)"
    return f"{label} ({len(rows)} rows)"
