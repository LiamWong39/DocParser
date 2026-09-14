"""Style resolution (effective font size / bold) and style creation."""

from __future__ import annotations

import re
from collections import Counter
from dataclasses import dataclass

from docx.oxml import parse_xml
from docx.oxml.ns import nsdecls, qn

from .xmlutil import PPR, PPR_SEQ, R, RPR, T, VAL, insert_ordered, is_on, local, make, style_id

# Word's size when neither the run, its styles, nor docDefaults say anything.
DEFAULT_SIZE_PT = 10.0


class StyleResolver:
    """Resolves effective run properties through rStyle -> pStyle chain -> docDefaults."""

    def __init__(self, doc):
        self.styles_el = doc.styles.element
        self.by_id: dict[str, object] = {}
        self.default_para: str | None = None
        for s in self.styles_el.findall(qn("w:style")):
            self.add(s)
        self.default_rpr = self.styles_el.find(f"{qn('w:docDefaults')}/{qn('w:rPrDefault')}/{qn('w:rPr')}")

    def add(self, s):
        sid = s.get(qn("w:styleId"))
        self.by_id[sid] = s
        if s.get(qn("w:type")) == "paragraph" and s.get(qn("w:default")) in ("1", "true", "on"):
            self.default_para = sid

    def _effective_para_style(self, sid: str | None) -> str | None:
        return sid if sid in self.by_id else self.default_para

    def name(self, sid: str | None) -> str | None:
        s = self.by_id.get(self._effective_para_style(sid))
        if s is None:
            return None
        n = s.find(qn("w:name"))
        return n.get(VAL) if n is not None else None

    def find_by_name(self, name: str, type_: str = "paragraph") -> str | None:
        for sid, s in self.by_id.items():
            n = s.find(qn("w:name"))
            if s.get(qn("w:type")) == type_ and n is not None and (n.get(VAL) or "").lower() == name.lower():
                return sid
        return None

    def heading_level(self, sid: str | None) -> int | None:
        if sid is None:
            return None
        m = re.fullmatch(r"heading\s*(\d)", (self.name(sid) or "").strip().lower())
        return int(m.group(1)) if m else None

    def chain(self, sid: str | None):
        seen = set()
        while sid and sid in self.by_id and sid not in seen:
            seen.add(sid)
            s = self.by_id[sid]
            yield s
            based = s.find(qn("w:basedOn"))
            sid = based.get(VAL) if based is not None else None

    @staticmethod
    def _style_rpr_prop(s, q):
        rpr = s.find(RPR)
        return rpr.find(q) if rpr is not None else None

    def run_prop(self, r, p, tag: str):
        q = qn(tag)
        rpr = r.find(RPR)
        if rpr is not None:
            el = rpr.find(q)
            if el is not None:
                return el
            rstyle = rpr.find(qn("w:rStyle"))
            if rstyle is not None:
                for s in self.chain(rstyle.get(VAL)):
                    el = self._style_rpr_prop(s, q)
                    if el is not None:
                        return el
        return self.para_prop(p, tag)

    def para_prop(self, p, tag: str):
        q = qn(tag)
        for s in self.chain(self._effective_para_style(style_id(p))):
            el = self._style_rpr_prop(s, q)
            if el is not None:
                return el
        return self.default_rpr.find(q) if self.default_rpr is not None else None

    @staticmethod
    def _size(el) -> float:
        try:
            return int(el.get(VAL)) / 2
        except (AttributeError, TypeError, ValueError):
            return DEFAULT_SIZE_PT

    def font_size(self, r, p) -> float:
        return self._size(self.run_prop(r, p, "w:sz"))

    def para_size(self, p) -> float:
        return self._size(self.para_prop(p, "w:sz"))

    def is_bold(self, r, p) -> bool:
        return is_on(self.run_prop(r, p, "w:b"))


@dataclass
class ParaInfo:
    size: float          # size of the majority of the characters
    all_bold: bool


def _run_text(r) -> str:
    return "".join(t.text or "" for t in r.findall(T))


def para_info(p, res: StyleResolver) -> ParaInfo:
    sizes: Counter[float] = Counter()
    total = bold = 0
    for r in p.iter(R):
        n = len(_run_text(r).strip())
        if not n:
            continue
        total += n
        sizes[res.font_size(r, p)] += n
        if res.is_bold(r, p):
            bold += n
    size = sizes.most_common(1)[0][0] if sizes else res.para_size(p)
    return ParaInfo(size=size, all_bold=total > 0 and bold == total)


def body_font_size(blocks, res: StyleResolver) -> float:
    """Most common font size across top-level paragraphs, weighted by characters."""
    sizes: Counter[float] = Counter()
    for p in blocks:
        if local(p) != "p":
            continue
        for r in p.iter(R):
            n = len(_run_text(r).strip())
            if n:
                sizes[res.font_size(r, p)] += n
    return sizes.most_common(1)[0][0] if sizes else DEFAULT_SIZE_PT


_STYLE_CHILDREN_AFTER_PPR = {"rPr", "tblPr", "trPr", "tcPr", "tblStylePr"}


def _ensure_outline_level(style_el, outline: int):
    ppr = style_el.find(PPR)
    if ppr is None:
        ppr = make("pPr")
        for i, child in enumerate(style_el):
            if local(child) in _STYLE_CHILDREN_AFTER_PPR:
                style_el.insert(i, ppr)
                break
        else:
            style_el.append(ppr)
    if ppr.find(qn("w:outlineLvl")) is None:
        insert_ordered(ppr, make("outlineLvl", val=outline), PPR_SEQ)


def _add_style(doc, res: StyleResolver, style_id_: str, name: str, ppr: str, rpr: str) -> str:
    sid = style_id_
    while sid in res.by_id:
        sid += "X"
    base = res.default_para
    based = f'<w:basedOn w:val="{base}"/><w:next w:val="{base}"/>' if base else ""
    el = parse_xml(
        f'<w:style {nsdecls("w")} w:type="paragraph" w:styleId="{sid}">'
        f'<w:name w:val="{name}"/>{based}<w:uiPriority w:val="9"/><w:qFormat/>'
        f"<w:pPr>{ppr}</w:pPr><w:rPr>{rpr}</w:rPr></w:style>"
    )
    res.styles_el.append(el)
    res.add(el)
    return sid


def ensure_heading_style(doc, res: StyleResolver, level: int) -> str:
    """Style id of 'heading N', creating it if the document doesn't define it."""
    sid = res.find_by_name(f"heading {level}")
    if sid is not None:
        _ensure_outline_level(res.by_id[sid], level - 1)
        return sid
    size = {1: 32, 2: 28, 3: 24}.get(level, 22)
    return _add_style(
        doc, res, f"Heading{level}", f"heading {level}",
        f'<w:keepNext/><w:keepLines/><w:spacing w:before="240" w:after="120"/><w:outlineLvl w:val="{level - 1}"/>',
        f'<w:b/><w:bCs/><w:sz w:val="{size}"/><w:szCs w:val="{size}"/>',
    )


def ensure_toc_heading_style(doc, res: StyleResolver) -> str:
    sid = res.find_by_name("TOC Heading")
    if sid is not None:
        return sid
    return _add_style(
        doc, res, "TOCHeading", "TOC Heading",
        '<w:keepNext/><w:spacing w:before="240" w:after="240"/><w:outlineLvl w:val="9"/>',
        '<w:b/><w:bCs/><w:sz w:val="32"/><w:szCs w:val="32"/>',
    )
