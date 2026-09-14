"""Low-level OOXML helpers shared by all pipeline steps.

Everything here works on raw lxml elements (w:p, w:tbl, ...) so the steps can
handle paragraphs python-docx doesn't wrap (inside text boxes, content controls).
"""

from __future__ import annotations

import copy
import re
from typing import Iterator

from docx.oxml import parse_xml
from docx.oxml.ns import nsdecls, qn

W_NS = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"

P = qn("w:p")
TBL = qn("w:tbl")
R = qn("w:r")
T = qn("w:t")
PPR = qn("w:pPr")
RPR = qn("w:rPr")
VAL = qn("w:val")

# Schema order of w:pPr children. Word rejects files whose children are out of order.
PPR_SEQ = [
    "pStyle", "keepNext", "keepLines", "pageBreakBefore", "framePr", "widowControl",
    "numPr", "suppressLineNumbers", "pBdr", "shd", "tabs", "suppressAutoHyphens",
    "kinsoku", "wordWrap", "overflowPunct", "topLinePunct", "autoSpaceDE", "autoSpaceDN",
    "bidi", "adjustRightInd", "snapToGrid", "spacing", "ind", "contextualSpacing",
    "mirrorIndents", "suppressOverlap", "jc", "textDirection", "textAlignment",
    "textboxTightWrap", "outlineLvl", "divId", "cnfStyle", "rPr", "sectPr", "pPrChange",
]
TRPR_SEQ = [
    "cnfStyle", "divId", "gridBefore", "gridAfter", "wBefore", "wAfter", "cantSplit",
    "trHeight", "tblHeader", "tblCellSpacing", "jc", "hidden", "ins", "del", "trPrChange",
]
# Children of w:settings that must come AFTER w:updateFields.
SETTINGS_AFTER_UPDATEFIELDS = [
    "hdrShapeDefaults", "footnotePr", "endnotePr", "compat", "docVars", "rsids", "mathPr",
    "attachedSchema", "themeFontLang", "clrSchemeMapping", "doNotIncludeSubdocsInStats",
    "doNotAutoCompressPictures", "forceUpgrade", "captions", "readModeInkLockDown",
    "smartTagType", "schemaLibrary", "shapeDefaults", "doNotEmbedSmartTags",
    "decimalSymbol", "listSeparator",
]


def local(el) -> str:
    tag = el.tag
    return tag.rsplit("}", 1)[-1] if isinstance(tag, str) else ""


def is_on(el) -> bool:
    """Evaluate an OOXML on/off element (<w:b/>, <w:b w:val="0"/>, ...)."""
    if el is None:
        return False
    return el.get(VAL) not in ("0", "false", "off")


def insert_ordered(parent, child, seq: list[str]):
    """Insert `child` into `parent` respecting the schema sequence `seq` (local names)."""
    name = local(child)
    existing = parent.find(qn(f"w:{name}"))
    if existing is not None:
        parent.replace(existing, child)
        return child
    idx = seq.index(name)
    later = set(seq[idx + 1:])
    for i, sib in enumerate(parent):
        if local(sib) in later:
            parent.insert(i, child)
            return child
    parent.append(child)
    return child


def make(tag: str, **attrs):
    """Create a w: element, e.g. make("pageBreakBefore") or make("outlineLvl", val="0")."""
    el = parse_xml(f"<w:{tag} {nsdecls('w')}/>")
    for k, v in attrs.items():
        el.set(qn(f"w:{k}"), str(v))
    return el


def get_or_add_ppr(p):
    ppr = p.find(PPR)
    if ppr is None:
        ppr = make("pPr")
        p.insert(0, ppr)
    return ppr


def set_ppr_flag(p, name: str, on: bool = True):
    ppr = get_or_add_ppr(p)
    if on:
        insert_ordered(ppr, make(name), PPR_SEQ)
    else:
        remove_children(ppr, name)


def has_ppr_flag(p, name: str) -> bool:
    ppr = p.find(PPR)
    return ppr is not None and is_on(ppr.find(qn(f"w:{name}")))


def remove_children(parent, *names: str) -> int:
    n = 0
    if parent is None:
        return 0
    for name in names:
        for el in parent.findall(qn(f"w:{name}")):
            parent.remove(el)
            n += 1
    return n


def remove(el):
    parent = el.getparent()
    if parent is not None:
        parent.remove(el)


def para_text(p) -> str:
    """Visible text of a paragraph: w:t text, tabs as '\\t', line breaks as '\\n'.

    Field instructions (w:instrText) and deleted text are not included.
    """
    out: list[str] = []
    for el in p.iter(T, qn("w:tab"), qn("w:br"), qn("w:cr"), qn("w:noBreakHyphen")):
        name = local(el)
        if name == "t":
            out.append(el.text or "")
        elif name == "tab":
            # w:tab inside w:tabs (pPr) is a tab stop definition, not a tab character.
            if local(el.getparent()) != "tabs":
                out.append("\t")
        elif name == "noBreakHyphen":
            out.append("-")
        elif el.get(qn("w:type")) in (None, "textWrapping"):
            out.append("\n")
    return "".join(out)


def style_id(p) -> str | None:
    ppr = p.find(PPR)
    if ppr is None:
        return None
    ps = ppr.find(qn("w:pStyle"))
    return ps.get(VAL) if ps is not None else None


def set_style_id(p, sid: str):
    insert_ordered(get_or_add_ppr(p), make("pStyle", val=sid), PPR_SEQ)


def iter_blocks(container) -> Iterator:
    """Yield top-level w:p / w:tbl in document order, descending into content controls."""
    for child in container:
        name = local(child)
        if name in ("p", "tbl"):
            yield child
        elif name == "sdt":
            content = child.find(qn("w:sdtContent"))
            if content is not None:
                yield from iter_blocks(content)
        elif name in ("customXml", "sdtContent"):
            yield from iter_blocks(child)


def block_list(body) -> list:
    return list(iter_blocks(body))


def has_section_break(p) -> bool:
    """True if the paragraph ends a section whose break starts a new page."""
    ppr = p.find(PPR)
    if ppr is None:
        return False
    sect = ppr.find(qn("w:sectPr"))
    if sect is None:
        return False
    typ = sect.find(qn("w:type"))
    return typ is None or typ.get(VAL) in ("nextPage", "evenPage", "oddPage")


M_NS = "http://schemas.openxmlformats.org/officeDocument/2006/math"
_NON_TEXT_CONTENT = tuple(qn(t) for t in ("w:drawing", "w:pict", "w:object", "w:fldSimple", "w:fldChar")) + (
    f"{{{M_NS}}}oMath",
    f"{{{M_NS}}}oMathPara",
)
XML_SPACE = "{http://www.w3.org/XML/1998/namespace}space"


def has_content_other_than_text(p) -> bool:
    """Drawings, objects, fields, section breaks: things we must not delete with an 'empty' paragraph."""
    if next(p.iter(*_NON_TEXT_CONTENT), None) is not None:
        return True
    ppr = p.find(PPR)
    return ppr is not None and ppr.find(qn("w:sectPr")) is not None


def is_blank(block) -> bool:
    """An empty paragraph with nothing worth keeping. Tables are never blank."""
    return local(block) == "p" and not para_text(block).strip() and not has_content_other_than_text(block)


def norm_text(s: str) -> str:
    """Lowercase, punctuation -> single spaces. Used to compare headings with TOC entries."""
    return re.sub(r"[\W_]+", " ", s.lower()).strip()


def new_paragraph(text: str | None = None, style: str | None = None):
    p = make("p")
    if style:
        set_style_id(p, style)
    if text:
        r = make("r")
        t = make("t")
        t.text = text
        t.set(XML_SPACE, "preserve")
        r.append(t)
        p.append(r)
    return p


def clone(el):
    return copy.deepcopy(el)
