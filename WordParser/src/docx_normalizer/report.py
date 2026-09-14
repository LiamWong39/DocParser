from __future__ import annotations

import json
from dataclasses import asdict, dataclass, field


@dataclass
class HeadingDecision:
    text: str
    level: int | None
    accepted: bool
    signals: list[str]
    previous_style: str | None = None
    reason: str = ""


@dataclass
class Report:
    input: str = ""
    output: str = ""
    body_font_size: float | None = None
    cleanup: dict[str, int] = field(default_factory=dict)
    toc_found: str = "none"          # "field", "hand-typed", "field+hand-typed", "none"
    toc_removed_lines: list[str] = field(default_factory=list)
    toc_inserted: bool = False
    toc_refreshed: str = "not attempted"
    headings: list[HeadingDecision] = field(default_factory=list)
    section_breaks: list[str] = field(default_factory=list)
    table_breaks: list[str] = field(default_factory=list)
    warnings: list[str] = field(default_factory=list)

    def count(self, key: str, n: int = 1):
        self.cleanup[key] = self.cleanup.get(key, 0) + n

    @property
    def accepted(self) -> list[HeadingDecision]:
        return [h for h in self.headings if h.accepted]

    @property
    def near_misses(self) -> list[HeadingDecision]:
        return [h for h in self.headings if not h.accepted]

    def to_json(self) -> str:
        return json.dumps(asdict(self), indent=2, ensure_ascii=False)

    def to_markdown(self) -> str:
        def esc(s: str) -> str:
            s = " ".join(s.split())
            s = s if len(s) <= 90 else s[:87] + "..."
            return s.replace("|", "\\|")

        lines = [f"# docx-normalizer report", ""]
        lines += [f"- Input: `{self.input}`", f"- Output: `{self.output}`"]
        if self.body_font_size:
            lines.append(f"- Body text size: {self.body_font_size:g} pt")
        lines.append(f"- Headings: {len(self.accepted)} accepted, {len(self.near_misses)} near misses to review")
        lines.append(f"- Section page breaks: {len(self.section_breaks)}; table page breaks: {len(self.table_breaks)}")
        lines.append(f"- Old TOC: {self.toc_found}; new TOC inserted: {'yes' if self.toc_inserted else 'no'}; refreshed: {self.toc_refreshed}")
        lines.append("")

        if self.warnings:
            lines += ["## Warnings", ""] + [f"- {w}" for w in self.warnings] + [""]

        lines += ["## Cleanup", ""]
        if self.cleanup:
            lines += [f"- {k}: {v}" for k, v in sorted(self.cleanup.items())]
        else:
            lines.append("- nothing to clean")
        lines.append("")

        lines += ["## Headings (outline)", "", "| Level | Text | Signals | Was |", "|---|---|---|---|"]
        for h in self.accepted:
            indent = "&nbsp;&nbsp;" * ((h.level or 1) - 1)
            lines.append(f"| H{h.level} | {indent}{esc(h.text)} | {', '.join(h.signals)} | {h.previous_style or '-'} |")
        lines.append("")

        lines += [
            "## Near misses (not changed, check these)", "",
            "Lines that look a bit like headings but didn't pass the rules. If some are real headings,",
            "adjust the config (e.g. `size_delta`, `max_words`, `accept_bold_only_unnumbered`) or style them in Word.", "",
            "| Text | Signals | Why rejected |", "|---|---|---|",
        ]
        for h in self.near_misses:
            lines.append(f"| {esc(h.text)} | {', '.join(h.signals)} | {h.reason} |")
        lines.append("")

        if self.toc_removed_lines:
            lines += ["## Removed old TOC lines", ""] + [f"- {esc(t)}" for t in self.toc_removed_lines] + [""]
        if self.table_breaks:
            lines += ["## Tables moved to a new page", ""] + [f"- {esc(t)}" for t in self.table_breaks] + [""]
        return "\n".join(lines)
