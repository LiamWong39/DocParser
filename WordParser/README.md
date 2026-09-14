# docx-normalizer

Cleans up messy Word documents (PDF → Word → hand-edited) before they are chunked
into TOON files, so every document has the same structure:

- **Real headings.** Section titles that are just bold / bigger / numbered text
  (`4.2 Payment terms`) get real `Heading 1/2/3…` styles. The level comes from the number
  (`4` → H1, `4.2` → H2). Headings with the wrong level (`2.1 X` styled Heading 1) are fixed.
- **A real TOC.** The old TOC is removed, whether it's generated, hand-typed, or a mix
  (a TOC field plus hand-typed lines). A proper TOC field replaces it and is filled in
  with headless LibreOffice. Lines from the old TOC are also used as hints to find headings.
- **Every top-level section starts on a new page** (`pageBreakBefore` on Heading 1).
- **Every table starts on a new page**, together with its heading, caption
  (`Table 3: …`) or lead-in line (`… as follows:`). The first row repeats when a table
  runs over several pages.
- **PDF conversion junk removed.** Text boxes are unwrapped into normal paragraphs, frames and
  floating tables are made inline, stray manual page breaks and piles of empty paragraphs are
  removed, and split runs are merged.
- **A report** (`*.report.md`) lists every heading it found and why, plus the **near misses**
  (lines that almost looked like headings) so you can check them quickly.

## Setup

```bash
uv sync
```

To fill in the TOC automatically you need LibreOffice (`soffice`) and a Python that can
`import uno` (on Fedora/Ubuntu that's the system `python3` with LibreOffice installed). Without
them, the TOC is left for Word to fill in: Word asks to update fields when the file is opened.

## Usage

```bash
uv run docx-normalize messy.docx clean.docx
#   -> clean.docx + clean.report.md

uv run docx-normalize messy.docx --dry-run          # only the report, to see what it would do
uv run docx-normalize messy.docx clean.docx --no-refresh-toc
uv run docx-normalize messy.docx clean.docx --config my.toml --json report.json
```

From Python (e.g. inside the agent pipeline):

```python
from docx_normalizer import normalize_file
from docx_normalizer.config import load_config

report = normalize_file("messy.docx", "clean.docx", load_config("my.toml"))
print([ (h.level, h.text) for h in report.accepted ])
```

## Suggested workflow for a new document

1. `docx-normalize messy.docx clean.docx`
2. Open `clean.report.md` and check the **Headings (outline)** table. Does the outline look like the document?
3. Check the **Near misses** table. If real headings are listed there, tune the config and run again:
   - bold headings the same size as body text → `accept_bold_only_unnumbered = true`
   - headings only slightly bigger → lower `size_delta`
   - long headings → raise `max_words`
   - different numbering (`Section 4`, `A.1`) → change `heading_regex` (it must keep the `num` and `title` groups)
4. If body lines were wrongly promoted, tighten those same settings.
5. Feed `clean.docx` to the chunking agent.

`uv run docx-normalize --print-config > my.toml` prints every setting with comments. Your
file only needs the keys you change:

```toml
[headings]
size_delta = 1.0
accept_bold_only_unnumbered = true

[page_breaks]
table_min_rows = 3   # don't push tiny layout tables onto their own page
```

## How headings are decided

A non-empty paragraph outside a table (and after the TOC) becomes a heading if **any** of these is true:

| Rule | Level |
|---|---|
| already styled `Heading N` | N, or the number's level if they disagree |
| numbered (`4.2 Title`), ≤ `max_words`, doesn't end in `.,;` **and** (bold **or** bigger **or** listed in the old TOC) | parts in the number |
| unnumbered, bold **and** bigger, short | matched to the size of other headings |
| listed in the old TOC and short | from the TOC entry |

Captions (`Table 3`, `Figure 2`), `Title`/`Subtitle` styles, and anything before the TOC
(cover page) are never headings.

## Limitations

- The TOC refresh saves the file through LibreOffice. The content is the same, but LibreOffice
  rewrites the file's internal XML. If that causes problems for a consumer, use `--no-refresh-toc`.
- Text boxes grouped with images are left in place (reported under cleanup).
- Very large tables still flow over several pages; they just always *start* on a fresh page.
- Headers and footers are not touched.

## Development

```bash
uv run pytest                                          # includes a LibreOffice round trip if available
uv run python tests/make_fixtures.py tests/fixtures    # build the synthetic broken docs to look at
```
