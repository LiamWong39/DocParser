"""Find and remove the old (field and/or hand-typed) TOC, insert a real TOC field, refresh it."""

from __future__ import annotations

import os
import re
import shutil
import subprocess
from importlib import resources
from pathlib import Path

from docx.oxml import parse_xml
from docx.oxml.ns import nsdecls, qn

from .analyze import StyleResolver, ensure_toc_heading_style
from .report import Report
from .xmlutil import (
    PPR, SETTINGS_AFTER_UPDATEFIELDS, VAL, block_list, has_content_other_than_text, has_section_break,
    is_blank, iter_blocks, local, make, new_paragraph, norm_text, para_text, remove, set_ppr_flag, style_id,
)

TOC_INSTR_RE = re.compile(r"(?<![\w])TOC(?![\w])")

TOC_FIELD_XML = (
    f"<w:p {nsdecls('w')}>"
    '<w:r><w:fldChar w:fldCharType="begin" w:dirty="true"/></w:r>'
    '<w:r><w:instrText xml:space="preserve"> TOC \\o "__LEVELS__" \\h \\z \\u </w:instrText></w:r>'
    '<w:r><w:fldChar w:fldCharType="separate"/></w:r>'
    "<w:r><w:t>Right-click and choose Update Field to build the table of contents.</w:t></w:r>"
    '<w:r><w:fldChar w:fldCharType="end"/></w:r>'
    "</w:p>"
)


def _instr_text(p) -> str:
    parts = [i.text or "" for i in p.iter(qn("w:instrText"))]
    parts += [f.get(qn("w:instr")) or "" for f in p.iter(qn("w:fldSimple"))]
    return " ".join(parts)


def _toc_sdt_blocks(body) -> list:
    out = []
    path = f"{qn('w:sdtPr')}/{qn('w:docPartObj')}/{qn('w:docPartGallery')}"
    for sdt in body.iter(qn("w:sdt")):
        gallery = sdt.find(path)
        if gallery is not None and "table of contents" in (gallery.get(VAL) or "").lower():
            content = sdt.find(qn("w:sdtContent"))
            if content is not None:
                out.extend(iter_blocks(content))
    return out


def find_and_remove_toc(body, cfg: dict, res: StyleResolver, report: Report):
    """Remove the existing TOC region. Returns (anchor paragraph or None, entry hints for heading detection)."""
    tc = cfg["toc"]
    title_re = re.compile(tc["title_regex"], re.I)
    entry_re = re.compile(tc["entry_regex"], re.I)
    max_words = cfg["headings"]["max_words"]

    blocks = block_list(body)
    raw = [para_text(b).replace("\n", " ").strip() if local(b) == "p" else None for b in blocks]

    def toc_style_level(i):
        if raw[i] is None:
            return None
        m = re.fullmatch(r"toc\s*(\d)", (res.name(style_id(blocks[i])) or "").strip(), re.I)
        return int(m.group(1)) if m else None

    def is_entry(i):
        return bool(raw[i]) and (entry_re.match(raw[i]) is not None or toc_style_level(i) is not None)

    def is_title(i):
        return bool(raw[i]) and title_re.match(raw[i]) is not None

    def next_nonblank(i):
        while i < len(blocks) and is_blank(blocks[i]):
            i += 1
        return i if i < len(blocks) else None

    # 1. Generated TOC: content control and/or TOC field (may span many paragraphs).
    core: set[int] = set()
    sdt_ids = {id(b) for b in _toc_sdt_blocks(body)}
    in_field, depth = False, 0
    for i, b in enumerate(blocks):
        if id(b) in sdt_ids:
            core.add(i)
        if local(b) != "p":
            if in_field:
                core.add(i)
            continue
        if not in_field and TOC_INSTR_RE.search(_instr_text(b)):
            in_field, depth = True, 0
        if in_field:
            core.add(i)
            for fc in b.iter(qn("w:fldChar")):
                kind = fc.get(qn("w:fldCharType"))
                depth += 1 if kind == "begin" else -1 if kind == "end" else 0
            if depth <= 0:
                in_field = False

    kinds = set()
    if core:
        s, e = min(core), max(core)
        kinds.add("field")
    else:
        s = e = None
        # 2. A "Contents" title followed by entry lines.
        for i in range(len(blocks)):
            if is_title(i):
                j = next_nonblank(i + 1)
                if j is not None and is_entry(j):
                    s = e = i
                    break
        # 3. No title: the first run of 3+ entry lines.
        if s is None:
            run: list[int] = []
            for i in range(len(blocks)):
                if is_entry(i):
                    run.append(i)
                elif is_blank(blocks[i]) and run:
                    continue
                elif len(run) >= 3:
                    break
                else:
                    run = []
            if len(run) >= 3:
                s, e = run[0], run[-1]
        if s is None:
            report.toc_found = "none"
            return None, []

    # Grow the region: hand-typed lines right after / before the generated part.
    j = e + 1
    while j < len(blocks):
        if is_entry(j) or j in core:
            e = j
            j += 1
        elif is_blank(blocks[j]):
            j += 1
        elif raw[j] and len(raw[j].split()) <= max_words:
            # An entry wrapped over two lines: "4.2 A long title that" / "wraps ....... 12"
            k = next_nonblank(j + 1)
            if k is not None and is_entry(k) and not has_content_other_than_text(blocks[j]):
                e = k
                j = k + 1
            else:
                break
        else:
            break
    j = s - 1
    while j >= 0:
        if is_title(j):
            s = j
            break
        if is_entry(j):
            s = j
        elif not is_blank(blocks[j]):
            break
        j -= 1

    # Hints (titles + levels) for heading detection, and the removal log.
    hints, pending = [], ""
    for i in range(s, e + 1):
        if not raw[i] or is_title(i):
            continue
        if i not in core and is_entry(i):
            kinds.add("hand-typed")
        m = entry_re.match(raw[i])
        if m:
            num = (m.group("num") or "").rstrip(".")
            title = f"{pending} {m.group('title')}".strip()
        elif toc_style_level(i) is not None:
            num, title = "", re.sub(r"\s*\d+\s*$", "", raw[i])
            num_m = re.match(r"^(\d+(?:\.\d+)*)\.?\s+(.*)$", title)
            if num_m:
                num, title = num_m.group(1), f"{pending} {num_m.group(2)}".strip()
        else:
            pending = raw[i]
            report.toc_removed_lines.append(raw[i])
            continue
        pending = ""
        level = len(num.split(".")) if num else toc_style_level(i)
        hints.append({
            "num": num, "title": title, "level": level,
            "norm_title": norm_text(title), "norm_full": norm_text(f"{num} {title}"),
        })
        report.toc_removed_lines.append(" ".join(raw[i].split()))

    anchor = new_paragraph()
    blocks[s].addprevious(anchor)
    for b in blocks[s:e + 1]:
        ppr = b.find(PPR) if local(b) == "p" else None
        if ppr is not None and ppr.find(qn("w:sectPr")) is not None:
            # Keep section breaks (TOC pages often end with one); just drop the text.
            for child in list(b):
                if child.tag != PPR:
                    b.remove(child)
            continue
        remove(b)
    _drop_empty_block_sdts(body, anchor)

    report.toc_found = "+".join(sorted(kinds)) or "field"
    return anchor, hints


def _drop_empty_block_sdts(body, anchor):
    parent = anchor.getparent()
    while local(parent) == "sdtContent":
        sdt = parent.getparent()
        if any(b is not anchor for b in iter_blocks(parent)):
            break
        sdt.addprevious(anchor)
        remove(sdt)
        parent = anchor.getparent()
    for sdt in list(body.iter(qn("w:sdt"))):
        content = sdt.find(qn("w:sdtContent"))
        if local(sdt.getparent()) in ("body", "sdtContent", "tc") and content is not None and not list(iter_blocks(content)):
            remove(sdt)


def match_hint(text: str, title: str | None, hints: list[dict], ratio: float) -> dict | None:
    if not hints:
        return None
    full = norm_text(text)
    title_n = norm_text(title) if title else full
    for h in hints:
        if full == h["norm_full"] or (title_n and title_n == h["norm_title"]):
            return h
    if len(full) > 120:
        return None
    from difflib import SequenceMatcher

    for h in hints:
        if SequenceMatcher(None, full, h["norm_full"]).ratio() >= ratio:
            return h
    return None


def _set_update_fields(doc):
    settings = doc.settings.element
    existing = settings.find(qn("w:updateFields"))
    if existing is not None:
        existing.set(VAL, "true")
        return
    el = make("updateFields", val="true")
    for i, child in enumerate(settings):
        if local(child) in SETTINGS_AFTER_UPDATEFIELDS:
            settings.insert(i, el)
            return
    settings.append(el)


def insert_toc(doc, anchor, levels: dict, cfg: dict, res: StyleResolver, report: Report):
    tc = cfg["toc"]
    body = doc.element.body
    if anchor is None:
        if not tc["insert_if_missing"]:
            return
        first = next((b for b in block_list(body) if b in levels), None)
        if first is None:
            report.warnings.append("No headings found, so no TOC was inserted.")
            return
        anchor = new_paragraph()
        first.addprevious(anchor)

    title = new_paragraph(tc["title"], style=ensure_toc_heading_style(doc, res))
    field = parse_xml(TOC_FIELD_XML.replace("__LEVELS__", tc["levels"]))
    anchor.addprevious(title)
    anchor.addprevious(field)
    remove(anchor)
    set_ppr_flag(title, "keepNext")

    blocks = block_list(body)
    i = blocks.index(title)
    before = blocks[:i]
    prev_content = next((b for b in reversed(before) if not is_blank(b)), None)
    if prev_content is not None and not (local(prev_content) == "p" and has_section_break(prev_content)):
        set_ppr_flag(title, "pageBreakBefore")
    # Whatever follows the TOC starts on a new page.
    for b in blocks[i + 2:]:
        if local(b) == "p" and has_section_break(b):
            break
        if is_blank(b):
            continue
        if local(b) == "p":
            set_ppr_flag(b, "pageBreakBefore")
        break  # tables get their break from the page-break step

    _set_update_fields(doc)
    report.toc_inserted = True


def refresh_with_libreoffice(path: str | Path, cfg: dict, report: Report):
    """Open the saved file in headless LibreOffice, update all indexes, save it back as .docx."""
    tc = cfg["toc"]
    path = Path(path)
    if shutil.which(tc["soffice"]) is None:
        report.toc_refreshed = "skipped (LibreOffice not found)"
        report.warnings.append("LibreOffice not found: the TOC stays empty until someone updates fields in Word (Word will ask on open).")
        return
    tmp = path.with_name(path.stem + ".lo-tmp.docx")
    with resources.as_file(resources.files("docx_normalizer").joinpath("lo_refresh.py")) as script:
        try:
            proc = subprocess.run(
                [tc["uno_python"], str(script), str(path), str(tmp), tc["soffice"]],
                capture_output=True, text=True, timeout=tc["refresh_timeout"],
            )
        except (FileNotFoundError, subprocess.TimeoutExpired) as exc:
            report.toc_refreshed = f"failed ({exc.__class__.__name__})"
            report.warnings.append(f"TOC refresh failed: {exc}. The TOC will be filled when the file is opened in Word.")
            tmp.unlink(missing_ok=True)
            return
    if proc.returncode == 0 and tmp.exists():
        os.replace(tmp, path)
        report.toc_refreshed = f"yes ({proc.stdout.strip()})"
    else:
        tmp.unlink(missing_ok=True)
        tail = (proc.stderr or proc.stdout).strip().splitlines()[-3:]
        report.toc_refreshed = "failed"
        report.warnings.append("TOC refresh with LibreOffice failed: " + " / ".join(tail))
