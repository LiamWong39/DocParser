"""Detect section headings that aren't styled as headings, and apply real Heading N styles."""

from __future__ import annotations

import re
from collections import Counter, defaultdict
from dataclasses import dataclass

from docx.opc.constants import RELATIONSHIP_TYPE as RT
from docx.oxml.ns import qn

from .analyze import ParaInfo, StyleResolver, ensure_heading_style, para_info
from .report import HeadingDecision, Report
from .toc import match_hint
from .xmlutil import (
    P, PPR, R, RPR, VAL, XML_SPACE, get_or_add_ppr, local, make, para_text, remove_children,
    set_ppr_flag, set_style_id, style_id,
)


class Numbering:
    """Looks up whether a paragraph's Word list numbering is decimal (1, 1.1, ...) and at which level."""

    def __init__(self, doc):
        self.num_to_abs: dict[str, str] = {}
        self.abstract: dict[str, object] = {}
        el = next((rel.target_part.element for rel in doc.part.rels.values() if rel.reltype == RT.NUMBERING), None)
        if el is None:
            return
        for a in el.findall(qn("w:abstractNum")):
            self.abstract[a.get(qn("w:abstractNumId"))] = a
        for n in el.findall(qn("w:num")):
            abs_id = n.find(qn("w:abstractNumId"))
            if abs_id is not None:
                self.num_to_abs[n.get(qn("w:numId"))] = abs_id.get(VAL)

    def decimal_level(self, p, res: StyleResolver) -> int | None:
        num_pr = None
        ppr = p.find(PPR)
        if ppr is not None:
            num_pr = ppr.find(qn("w:numPr"))
        if num_pr is None:
            for s in res.chain(style_id(p)):
                sp = s.find(PPR)
                if sp is not None and sp.find(qn("w:numPr")) is not None:
                    num_pr = sp.find(qn("w:numPr"))
                    break
        if num_pr is None:
            return None
        num_id = num_pr.find(qn("w:numId"))
        ilvl_el = num_pr.find(qn("w:ilvl"))
        ilvl = ilvl_el.get(VAL) if ilvl_el is not None else "0"
        abstract = self.abstract.get(self.num_to_abs.get(num_id.get(VAL) if num_id is not None else ""))
        if abstract is None:
            return None
        for lvl in abstract.findall(qn("w:lvl")):
            if lvl.get(qn("w:ilvl")) == ilvl:
                fmt, text = lvl.find(qn("w:numFmt")), lvl.find(qn("w:lvlText"))
                if fmt is not None and fmt.get(VAL) == "decimal" and text is not None and "%" in (text.get(VAL) or ""):
                    return int(ilvl) + 1
        return None


@dataclass
class Candidate:
    el: object
    text: str
    info: ParaInfo
    signals: list[str]
    level: int | None
    previous_style: str | None


def detect_and_apply_headings(doc, blocks, res: StyleResolver, body_size: float, hints: list[dict],
                              toc_index: int | None, cfg: dict, report: Report) -> dict:
    """Returns {paragraph element: heading level} for every heading in the document."""
    h = cfg["headings"]
    heading_re = re.compile(h["heading_regex"])
    caption_re = re.compile(cfg["page_breaks"]["caption_regex"], re.I)
    ignore = {s.lower() for s in h["ignore_styles"]}
    numbering = Numbering(doc)

    accepted: list[Candidate] = []
    for idx, el in enumerate(blocks):
        if local(el) != "p":
            continue
        if h["skip_before_toc"] and toc_index is not None and idx < toc_index:
            continue
        text = " ".join(para_text(el).split())
        if not text:
            continue
        sid = style_id(el)
        sname = res.name(sid)
        if (sid or "").lower() in ignore or (sname or "").lower() in ignore:
            continue
        existing = res.heading_level(sid)
        if not existing and caption_re.match(text):
            continue

        info = para_info(el, res)
        words = len(text.split())
        short = words <= h["max_words"]
        clean = short and text[-1] not in h["sentence_end_chars"]
        signals: list[str] = []

        num_level, title = None, None
        m = heading_re.match(text)
        if m and int(m.group("num").split(".")[0]) <= h["max_section_number"]:
            num_level, title = len(m.group("num").split(".")), m.group("title")
            signals.append(f"numbered {m.group('num')}")
        elif h["use_auto_numbering"]:
            num_level = numbering.decimal_level(el, res)
            if num_level:
                signals.append("auto-numbered")
        bold = info.all_bold
        bigger = info.size >= body_size + h["size_delta"]
        if bold:
            signals.append("bold")
        if bigger:
            signals.append(f"{info.size:g}pt")
        hint = match_hint(text, title, hints, h["toc_fuzzy_ratio"])
        if hint:
            signals.append("in old TOC")
        if existing:
            signals.append(f"style {sname}")

        level, accept, reason = None, False, ""
        if existing:
            accept, level = True, existing
            if h["fix_existing_levels"] and num_level and clean and num_level != existing:
                signals.append(f"level fixed H{existing}->H{num_level}")
                level = num_level
        elif num_level and clean and (bold or bigger or hint):
            accept, level = True, num_level
        elif not num_level and clean and bold and (bigger or h["accept_bold_only_unnumbered"]):
            accept = True  # level assigned below
        elif hint and short:
            accept, level = True, hint["level"] or num_level
        else:
            if not (bold or bigger or hint or (num_level and short)):
                continue
            if not short:
                reason = f"too long ({words} words)"
            elif not clean:
                reason = "ends like a sentence"
            elif num_level:
                reason = "numbered, but not bold, bigger or in the old TOC"
            elif bold:
                reason = "bold, but not bigger than body text and unnumbered"
            else:
                reason = "bigger, but not bold and unnumbered"
            report.headings.append(HeadingDecision(text, None, False, signals, sname, reason))
            continue
        accepted.append(Candidate(el, text, info, signals, level, sname))

    _assign_unnumbered_levels(accepted, h["unnumbered_level_strategy"])

    levels: dict = {}
    for c in accepted:
        c.level = max(1, min(c.level, h["max_level"]))
        _apply_heading(doc, res, c, h)
        levels[c.el] = c.level
        report.headings.append(HeadingDecision(c.text, c.level, True, c.signals, c.previous_style))

    # A direct outline level on body text would pull it into the TOC (\u switch).
    n = sum(remove_children(p.find(PPR), "outlineLvl") for p in doc.element.body.iter(P) if p not in levels)
    if n:
        report.count("outline levels removed from body text", n)
    return levels


def _assign_unnumbered_levels(cands: list[Candidate], strategy: str):
    unassigned = [c for c in cands if c.level is None]
    if not unassigned:
        return
    if strategy == "previous_plus_one":
        last = 0
        for c in cands:
            if c.level is None:
                c.level = last + 1
            else:
                last = c.level
        return

    def bucket(size: float) -> float:
        return round(size * 2) / 2

    by_size: dict[float, Counter] = defaultdict(Counter)
    for c in cands:
        if c.level is not None:
            by_size[bucket(c.info.size)][c.level] += 1
    known = [(size, counts.most_common(1)[0][0]) for size, counts in by_size.items()]

    if not known:
        sizes = sorted({bucket(c.info.size) for c in unassigned}, reverse=True)
        for c in unassigned:
            c.level = sizes.index(bucket(c.info.size)) + 1
        return
    for c in unassigned:
        s = c.info.size
        close = sorted((abs(ks - s), lv) for ks, lv in known if abs(ks - s) <= 0.75)
        if close:
            c.level = close[0][1]
            continue
        larger = sorted((ks, lv) for ks, lv in known if ks > s)
        c.level = larger[0][1] + 1 if larger else 1


def _apply_heading(doc, res: StyleResolver, c: Candidate, h: dict):
    set_style_id(c.el, ensure_heading_style(doc, res, c.level))
    for r in c.el.iter(R):
        rpr = r.find(RPR)
        if rpr is not None:
            remove_children(rpr, *h["strip_run_formatting"])
            if len(rpr) == 0:
                r.remove(rpr)
    remove_children(get_or_add_ppr(c.el), *h["strip_paragraph_formatting"], "outlineLvl")
    # Manual line breaks inside a heading end up in the TOC; make them spaces.
    for br in list(c.el.iter(qn("w:br"), qn("w:cr"))):
        if br.get(qn("w:type")) in (None, "textWrapping"):
            t = make("t")
            t.text = " "
            t.set(XML_SPACE, "preserve")
            br.getparent().replace(br, t)
    set_ppr_flag(c.el, "keepNext")
