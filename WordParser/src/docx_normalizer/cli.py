from __future__ import annotations

import argparse
import sys
from pathlib import Path

from .config import default_config_text, load_config
from .pipeline import normalize_file


def main(argv: list[str] | None = None) -> int:
    ap = argparse.ArgumentParser(
        prog="docx-normalize",
        description="Fix messy Word documents: real headings, a real TOC, page breaks before sections and tables.",
    )
    ap.add_argument("input", nargs="?", help="messy .docx")
    ap.add_argument("output", nargs="?", help="normalized .docx (default: <input>.normalized.docx)")
    ap.add_argument("--config", help="TOML file overriding the defaults (see --print-config)")
    ap.add_argument("--report", help="markdown report path (default: <output>.report.md)")
    ap.add_argument("--json", help="also write the report as JSON to this path")
    ap.add_argument("--dry-run", action="store_true", help="only write the report, not the document")
    ap.add_argument("--no-refresh-toc", action="store_true", help="don't fill in the TOC with LibreOffice")
    ap.add_argument("--print-config", action="store_true", help="print the default config and exit")
    args = ap.parse_args(argv)

    if args.print_config:
        print(default_config_text())
        return 0
    if not args.input:
        ap.error("input is required")

    src = Path(args.input)
    if not src.is_file():
        ap.error(f"{src} not found")
    dst = Path(args.output) if args.output else src.with_name(src.stem + ".normalized.docx")
    if dst.resolve() == src.resolve():
        ap.error("output must be a different file from input")

    try:
        cfg = load_config(args.config)
    except (OSError, ValueError) as exc:
        ap.error(f"bad config: {exc}")
    if args.no_refresh_toc:
        cfg["toc"]["refresh_with_libreoffice"] = False

    report = normalize_file(src, None if args.dry_run else dst, cfg, dry_run=args.dry_run)

    report_path = Path(args.report) if args.report else dst.with_name(dst.stem + ".report.md")
    report_path.write_text(report.to_markdown(), encoding="utf-8")
    if args.json:
        Path(args.json).write_text(report.to_json(), encoding="utf-8")

    print(f"Headings: {len(report.accepted)} ({len(report.near_misses)} near misses to review)")
    print(f"Page breaks: {len(report.section_breaks)} sections, {len(report.table_breaks)} tables")
    print(f"Old TOC: {report.toc_found}; new TOC: {'inserted' if report.toc_inserted else 'no'}; refreshed: {report.toc_refreshed}")
    for w in report.warnings:
        print(f"warning: {w}", file=sys.stderr)
    if not args.dry_run:
        print(f"Wrote {dst}")
    print(f"Report: {report_path}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
