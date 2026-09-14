"""Remove junk left behind by PDF -> Word conversion and manual editing."""

from __future__ import annotations

import re

from docx.oxml.ns import qn
from lxml import etree

from .report import Report
from .xmlutil import (
    P, PPR, R, RPR, T, XML_SPACE, block_list, has_content_other_than_text, is_blank, local,
    new_paragraph, para_text, remove, remove_children, style_id,
)

MC_FALLBACK = "{http://schemas.openxmlformats.org/markup-compatibility/2006}Fallback"
TXBX = qn("w:txbxContent")
PIC_TAGS = (
    "{http://schemas.openxmlformats.org/drawingml/2006/picture}pic",
    "{http://schemas.openxmlformats.org/drawingml/2006/main}blip",
    "{urn:schemas-microsoft-com:vml}imagedata",
)


def cleanup(doc, cfg: dict, report: Report):
    c = cfg["cleanup"]
    body = doc.element.body

    if c["unwrap_text_boxes"]:
        unwrap_text_boxes(body, report)
    if c["remove_frames"]:
        n = sum(remove_children(ppr, "framePr") for ppr in body.iter(PPR))
        if n:
            report.count("frames removed", n)
    if c["unfloat_tables"]:
        n = sum(remove_children(tp, "tblpPr", "tblOverlap") for tp in body.iter(qn("w:tblPr")))
        if n:
            report.count("floating tables made inline", n)

    # Pure noise: layout hints from the last render and spell-check markers.
    for el in list(body.iter(qn("w:lastRenderedPageBreak"), qn("w:proofErr"))):
        remove(el)

    break_only: set = set()
    if c["remove_manual_page_breaks"]:
        break_only = remove_manual_page_breaks(body, report)
    if c["merge_runs"]:
        merge_runs(body, report)
    if c["max_consecutive_empty"] >= 0:
        remove_empty_paragraphs(body, c["max_consecutive_empty"], break_only, report)
    if c["join_split_paragraphs"]:
        join_split_paragraphs(body, cfg, report)


def _inside_txbx(el) -> bool:
    return any(a.tag == TXBX for a in el.iterancestors())


def _outer_text_boxes(p) -> list:
    """Text boxes in p that aren't nested inside another text box in p."""
    out = []
    for tb in p.iter(TXBX):
        for a in tb.iterancestors():
            if a is p:
                out.append(tb)
                break
            if a.tag == TXBX:
                break
    return out


def _host_run(tb, p):
    """The outermost w:r below paragraph p that contains the text box."""
    run = None
    for a in tb.iterancestors():
        if a is p:
            break
        if a.tag == R:
            run = a
    return run


def unwrap_text_boxes(body, report: Report):
    """Move paragraphs out of text boxes / shapes into the normal flow, right after their anchor paragraph."""
    for _ in range(5):  # nested text boxes surface one level per pass
        hosts = [p for p in body.iter(P) if next(p.iter(TXBX), None) is not None and not _inside_txbx(p)]
        if not hosts:
            return
        progressed = False
        for p in hosts:
            # mc:Fallback holds a VML copy of the same text box; keep only the mc:Choice version.
            for fb in list(p.iter(MC_FALLBACK)):
                remove(fb)
            moved, runs = [], []
            for tb in _outer_text_boxes(p):
                run = _host_run(tb, p)
                if run is None:
                    continue
                if next(run.iter(*PIC_TAGS), None) is not None:
                    report.count("text boxes left alone (grouped with images)")
                    continue
                moved.extend(child for child in tb if local(child) in ("p", "tbl"))
                if run not in runs:
                    runs.append(run)
            anchor = p
            for block in moved:
                anchor.addnext(block)
                anchor = block
            for run in runs:
                remove(run)
            if runs:
                progressed = True
                report.count("text boxes unwrapped", len(runs))
        if not progressed:
            return


def remove_manual_page_breaks(body, report: Report) -> set:
    """Strip page-break characters and pageBreakBefore. Returns paragraphs that existed only to break a page."""
    break_only = set()
    for p in body.iter(P):
        has_break = any(br.get(qn("w:type")) == "page" for br in p.iter(qn("w:br")))
        ppr = p.find(PPR)
        if ppr is not None and ppr.find(qn("w:pageBreakBefore")) is not None:
            has_break = True
        if has_break and not para_text(p).strip() and not has_content_other_than_text(p):
            break_only.add(p)

    n = 0
    for br in list(body.iter(qn("w:br"))):
        if br.get(qn("w:type")) == "page":
            run = br.getparent()
            run.remove(br)
            n += 1
            if run.tag == R and all(child.tag == RPR for child in run):
                remove(run)
    n += sum(remove_children(ppr, "pageBreakBefore") for ppr in body.iter(PPR))
    if n:
        report.count("manual page breaks removed (re-added consistently)", n)
    return break_only


def _run_key(r):
    kids = [k for k in r if k.tag != RPR]
    if not kids or any(k.tag != T for k in kids):
        return None
    rpr = r.find(RPR)
    return etree.tostring(rpr) if rpr is not None else b""


def merge_runs(body, report: Report):
    n = 0
    containers = list(body.iter(P, qn("w:hyperlink"), qn("w:smartTag")))
    for container in containers:
        prev, prev_key = None, None
        for child in list(container):
            key = _run_key(child) if child.tag == R else None
            if key is not None and prev is not None and key == prev_key:
                last_t = prev.findall(T)[-1]
                last_t.text = (last_t.text or "") + "".join(t.text or "" for t in child.findall(T))
                last_t.set(XML_SPACE, "preserve")
                container.remove(child)
                n += 1
            elif key is not None:
                prev, prev_key = child, key
            else:
                prev, prev_key = None, None
    if n:
        report.count("runs merged", n)


def remove_empty_paragraphs(body, max_consecutive: int, break_only: set, report: Report):
    n = 0
    containers = [body, *body.iter(qn("w:tc")), *body.iter(qn("w:sdtContent"))]
    for container in containers:
        streak = 0
        for el in list(container):
            if local(el) != "p" or not is_blank(el):
                streak = 0
                continue
            prev, nxt = el.getprevious(), el.getnext()
            next_to_table = (prev is not None and local(prev) == "tbl") or (nxt is not None and local(nxt) == "tbl")
            # A table cell must end with a paragraph; two tables need a paragraph between them.
            required = (local(container) == "tc" and nxt is None) or (
                prev is not None and local(prev) == "tbl" and (nxt is None or local(nxt) in ("tbl", "sectPr"))
            )
            if required:
                continue
            if el in break_only:
                remove(el)
                n += 1
                continue
            streak += 1
            if streak > max_consecutive and not next_to_table:
                remove(el)
                n += 1
    if n:
        report.count("empty paragraphs removed", n)


_SENTENCE_END = re.compile(r"[.:;!?)\]\"”’]$")


def join_split_paragraphs(body, cfg: dict, report: Report):
    """Join a paragraph into the previous one when a PDF line wrap split a sentence."""
    heading_re = re.compile(cfg["headings"]["heading_regex"])
    n = 0
    prev = None
    for el in block_list(body):
        text = para_text(el).strip() if local(el) == "p" else ""
        if not text:
            prev = None
            continue
        if prev is not None and style_id(prev) == style_id(el):
            prev_text = para_text(prev).strip()
            if not _SENTENCE_END.search(prev_text) and text[0].islower() and not heading_re.match(prev_text):
                prev.append(new_paragraph(" ").find(R))
                for child in list(el):
                    if child.tag != PPR:
                        prev.append(child)
                remove(el)
                n += 1
                continue
        prev = el
    if n:
        report.count("split paragraphs joined", n)
