"""Build synthetic "broken" Word documents that imitate a PDF -> Word -> hand-edited file.

    uv run python tests/make_fixtures.py tests/fixtures
"""

from __future__ import annotations

import sys
from pathlib import Path

from docx import Document
from docx.enum.text import WD_BREAK
from docx.oxml import parse_xml
from docx.shared import Pt

NS = (
    'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" '
    'xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006" '
    'xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing" '
    'xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" '
    'xmlns:wps="http://schemas.microsoft.com/office/word/2010/wordprocessingShape" '
    'xmlns:v="urn:schemas-microsoft-com:vml"'
)

TEXT_BOX_HEADING = f"""
<w:p {NS}><w:r><mc:AlternateContent>
  <mc:Choice Requires="wps"><w:drawing>
    <wp:inline distT="0" distB="0" distL="0" distR="0">
      <wp:extent cx="3000000" cy="400000"/><wp:docPr id="100" name="Text Box 1"/>
      <a:graphic><a:graphicData uri="http://schemas.microsoft.com/office/word/2010/wordprocessingShape">
        <wps:wsp><wps:spPr/><wps:txbx><w:txbxContent>
          <w:p><w:r><w:rPr><w:b/><w:sz w:val="28"/></w:rPr><w:t>3 Fees</w:t></w:r></w:p>
        </w:txbxContent></wps:txbx><wps:bodyPr/></wps:wsp>
      </a:graphicData></a:graphic>
    </wp:inline>
  </w:drawing></mc:Choice>
  <mc:Fallback><w:pict><v:shape><v:textbox><w:txbxContent>
    <w:p><w:r><w:t>3 Fees</w:t></w:r></w:p>
  </w:txbxContent></v:textbox></v:shape></w:pict></mc:Fallback>
</mc:AlternateContent></w:r></w:p>
"""

W = 'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"'
FIELD_TOC = [
    f'<w:p {W}><w:r><w:fldChar w:fldCharType="begin"/></w:r><w:r><w:instrText xml:space="preserve"> TOC \\o "1-3" \\h \\z \\u </w:instrText></w:r>'
    '<w:r><w:fldChar w:fldCharType="separate"/></w:r><w:r><w:t>1 Introduction</w:t></w:r><w:r><w:tab/></w:r><w:r><w:t>3</w:t></w:r></w:p>',
    f'<w:p {W}><w:r><w:t>1.1 Purpose</w:t></w:r><w:r><w:tab/></w:r><w:r><w:t>3</w:t></w:r></w:p>',
    f'<w:p {W}><w:r><w:t>2 Scope</w:t></w:r><w:r><w:tab/></w:r><w:r><w:t>4</w:t></w:r><w:r><w:fldChar w:fldCharType="end"/></w:r></w:p>',
]


def _append_xml(doc, xml: str):
    doc.element.body.sectPr.addprevious(parse_xml(xml))


def _run(doc, text, bold=False, size=None):
    p = doc.add_paragraph()
    r = p.add_run(text)
    r.bold = bold or None
    if size:
        r.font.size = Pt(size)
    return p


def _table(doc, rows):
    t = doc.add_table(rows=len(rows), cols=len(rows[0]))
    for i, row in enumerate(rows):
        for j, value in enumerate(row):
            t.cell(i, j).text = value
    return t


def build_broken(path: str | Path, toc_kind: str = "hand") -> Path:
    """toc_kind: "hand" = fully hand-typed TOC; "mixed" = generated TOC field + hand-typed extra lines."""
    doc = Document()
    doc.styles["Normal"].font.size = Pt(11)

    _run(doc, "ACME Supplier Policy", bold=True, size=20)  # cover title: must not become a heading
    if toc_kind == "hand":
        doc.add_paragraph("Table of Contents")
        doc.add_paragraph("1 Introduction ........ 3")
        doc.add_paragraph("1.1 Purpose\t3")
        doc.add_paragraph("2 Scope ........ 4")
    else:
        doc.add_paragraph("Contents")
        for xml in FIELD_TOC:
            _append_xml(doc, xml)
    doc.add_paragraph("2.1 In scope ...... 4")
    doc.add_paragraph("3 Fees ........ 5")
    doc.add_paragraph("Appendix A Glossary ...... 9")
    for _ in range(3):
        doc.add_paragraph()

    doc.add_paragraph("1 Introduction", style="Heading 1")
    doc.add_paragraph("This document describes how suppliers are selected and paid.")
    _run(doc, "1.1 Purpose", bold=True, size=13)  # real heading, styled Normal
    doc.add_paragraph("The purpose of this policy is to keep purchasing consistent.")
    _run(doc, "This is a bold sentence that should stay body text.", bold=True)
    doc.add_paragraph().add_run().add_break(WD_BREAK.PAGE)  # manual page break paragraph
    _run(doc, "2 Scope", bold=True)  # numbered + bold, same size as body, same page as previous section
    doc.add_paragraph("This policy applies to all departments.")
    doc.add_paragraph("2.1 In scope", style="Heading 1")  # wrong level
    doc.add_paragraph("1. The supplier shall deliver all goods within thirty days of receiving a valid purchase order from us")
    doc.add_paragraph("2. Short item")

    _append_xml(doc, TEXT_BOX_HEADING)  # heading trapped in a text box
    doc.add_paragraph("Fees are listed below:")
    _table(doc, [["Code", "Description", "Fee"], ["1.1 foo", "Setup", "100"], ["1.2 bar", "Monthly", "20"]])
    doc.add_paragraph("More text about fees.")
    doc.add_paragraph("Table 2: Contacts")
    _table(doc, [["Name", "Email"], ["Ann", "ann@example.com"]])

    _run(doc, "Appendix A Glossary", bold=True, size=16)  # unnumbered heading
    doc.add_paragraph("Some glossary intro text here.")
    _table(doc, [["Term", "Meaning"], ["PO", "Purchase order"]])
    doc.add_paragraph("End of document.")

    path = Path(path)
    path.parent.mkdir(parents=True, exist_ok=True)
    doc.save(str(path))
    return path


if __name__ == "__main__":
    out = Path(sys.argv[1] if len(sys.argv) > 1 else "tests/fixtures")
    print(build_broken(out / "broken_hand_toc.docx", "hand"))
    print(build_broken(out / "broken_mixed_toc.docx", "mixed"))
