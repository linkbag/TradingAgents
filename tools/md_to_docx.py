"""Markdown → Word (.docx) converter for agent reports.

Why
---
TradingAgents emits its analyst/debate/decision output as Markdown
(per-agent ``.md`` files plus a consolidated ``report.md``). Markdown pipe
tables render as raw, misaligned text in email clients and text editors, and
CJK text has no proper East-Asian font binding. This tool converts that
Markdown into a Word document: headings, real bordered tables, bold runs,
bullets and numbered items, with a Chinese East-Asian font so reports read
cleanly in Word/WPS.

Scope (matched to what the agents actually emit): ATX headings, pipe tables,
``---`` horizontal rules, ``-``/``*`` bullets, ``1.`` numbered items and
``**bold**`` inline runs. Code fences and blockquotes pass through as plain
paragraphs.

Usage
-----
As a library::

    from tools.md_to_docx import md_to_docx, convert_report_file, md_to_docx_bytes

    md_to_docx(md_text, "report.docx")          # from a string
    convert_report_file("report.md")            # writes report.docx alongside
    md_to_docx_bytes(md_text)                   # bytes, for email attachments

From the command line::

    python tools/md_to_docx.py report.md                    # -> report.docx
    python tools/md_to_docx.py report.md out.docx           # explicit output
    python tools/md_to_docx.py a.md b.md c.md               # batch, alongside
    type report.md | python tools/md_to_docx.py - out.docx  # stdin

Requires ``python-docx`` (``pip install python-docx``). This file is kept
dependency-light and self-contained so it can be copied out of the repo
without dragging the TradingAgents package with it.
"""

from __future__ import annotations

import io
import logging
import re
import sys
from pathlib import Path
from typing import Optional

from docx import Document
from docx.oxml.ns import qn
from docx.shared import Pt, RGBColor

logger = logging.getLogger("md_to_docx")

#: Font for Chinese text (East-Asian run property). 微软雅黑 renders cleanly
#: in Word/WPS on every Windows box; the Latin font stays Calibri.
EA_FONT = "微软雅黑"
LATIN_FONT = "Calibri"

#: Header-row shading for tables (light grey) — via w:shd fill.
_HEADER_SHADING = "F2F2F2"

_HEADING_RE = re.compile(r"^(#{1,6})\s+(.*)$")
_HR_RE = re.compile(r"^\s*(?:-{3,}|\*{3,}|_{3,})\s*$")
_BULLET_RE = re.compile(r"^(\s*)[-*]\s+(.*)$")
_NUMBERED_RE = re.compile(r"^(\s*)(\d+)[.)]\s+(.*)$")
_TABLE_ROW_RE = re.compile(r"^\s*\|.*\|\s*$")
# GFM separators allow single-dash cells (e.g. "|:-:|---|" — one dash in the
# first cell), so the per-cell run is "-+" rather than "-{3,}".
_TABLE_SEP_RE = re.compile(r"^\s*\|?\s*:?-+:?\s*(\|\s*:?-+:?\s*)*\|?\s*$")
_BOLD_SPLIT_RE = re.compile(r"\*\*")


def _split_bold(text: str) -> list[tuple[str, bool]]:
    """Split ``text`` on ``**bold**`` markers → [(segment, is_bold), ...]."""
    parts = _BOLD_SPLIT_RE.split(text)
    out: list[tuple[str, bool]] = []
    for idx, part in enumerate(parts):
        if part == "":
            continue
        out.append((part, idx % 2 == 1))
    return out


def _set_ea_font(style) -> None:
    """Set the East-Asian font on a style so Chinese glyphs use 微软雅黑."""
    style.font.name = LATIN_FONT
    rpr = style.element.get_or_add_rPr()
    rfonts = rpr.get_or_add_rFonts()
    rfonts.set(qn("w:eastAsia"), EA_FONT)


def _add_runs(paragraph, text: str) -> None:
    """Add runs to ``paragraph`` honouring ``**bold**`` spans."""
    for segment, bold in _split_bold(text):
        run = paragraph.add_run(segment)
        run.bold = bold or None


def _add_hr(doc) -> None:
    """Render a Markdown horizontal rule as a paragraph with a bottom border."""
    p = doc.add_paragraph()
    ppr = p._p.get_or_add_pPr()
    pbdr = ppr.makeelement(qn("w:pBdr"), {})
    bottom = pbdr.makeelement(qn("w:bottom"), {
        qn("w:val"): "single",
        qn("w:sz"): "6",
        qn("w:space"): "1",
        qn("w:color"): "BFBFBF",
    })
    pbdr.append(bottom)
    ppr.append(pbdr)


def _shade(cell, fill: str) -> None:
    tcpr = cell._tc.get_or_add_tcPr()
    shd = tcpr.makeelement(qn("w:shd"), {
        qn("w:val"): "clear",
        qn("w:color"): "auto",
        qn("w:fill"): fill,
    })
    tcpr.append(shd)


def _add_table(doc, rows: list[list[str]]) -> None:
    """Add a bordered table; the first row is the bold shaded header."""
    if not rows:
        return
    ncols = max(len(r) for r in rows)
    table = doc.add_table(rows=len(rows), cols=ncols)
    table.style = "Table Grid"
    table.autofit = True
    for r_idx, row in enumerate(rows):
        for c_idx in range(ncols):
            cell = table.cell(r_idx, c_idx)
            cell.text = ""
            p = cell.paragraphs[0]
            text = row[c_idx] if c_idx < len(row) else ""
            _add_runs(p, text)
            for run in p.runs:
                run.font.size = Pt(9)
            if r_idx == 0:
                for run in p.runs:
                    run.bold = True
                _shade(cell, _HEADER_SHADING)


def _parse_table_row(line: str) -> list[str]:
    """Split a Markdown table row into cells (outer pipes trimmed)."""
    stripped = line.strip()
    if stripped.startswith("|"):
        stripped = stripped[1:]
    if stripped.endswith("|"):
        stripped = stripped[:-1]
    return [cell.strip() for cell in stripped.split("|")]


def md_to_docx(md_text: str, output_path: str | Path) -> Path:
    """Convert Markdown text to a Word document.

    Args:
        md_text: The Markdown source (agent report).
        output_path: Where to write the .docx.

    Returns:
        Path to the written document.
    """
    doc = _build_document(md_text)
    out = Path(output_path)
    out.parent.mkdir(parents=True, exist_ok=True)
    doc.save(str(out))
    logger.info("Markdown converted to Word: %s", out)
    return out


def _build_document(md_text: str):
    """Shared conversion body for :func:`md_to_docx` / :func:`md_to_docx_bytes`."""
    doc = Document()
    _set_ea_font(doc.styles["Normal"])
    doc.styles["Normal"].font.size = Pt(10.5)

    lines = (md_text or "").splitlines()
    i = 0
    n = len(lines)
    while i < n:
        line = lines[i]

        heading = _HEADING_RE.match(line)
        if heading:
            level = min(len(heading.group(1)), 4)
            text = heading.group(2).strip()
            if text:
                h = doc.add_heading("", level=level)
                _add_runs(h, text)
                for run in h.runs:
                    run.font.color.rgb = RGBColor(0x1F, 0x38, 0x64)
            i += 1
            continue

        if _TABLE_ROW_RE.match(line) and i + 1 < n and _TABLE_SEP_RE.match(lines[i + 1]):
            rows: list[list[str]] = [_parse_table_row(line)]
            i += 2
            while i < n and _TABLE_ROW_RE.match(lines[i]):
                rows.append(_parse_table_row(lines[i]))
                i += 1
            _add_table(doc, rows)
            continue

        if _HR_RE.match(line):
            _add_hr(doc)
            i += 1
            continue

        bullet = _BULLET_RE.match(line)
        if bullet:
            p = doc.add_paragraph(style="List Bullet")
            _add_runs(p, bullet.group(2))
            i += 1
            continue

        numbered = _NUMBERED_RE.match(line)
        if numbered:
            # Keep the literal number: Word list numbering across separate
            # sequences is fragile, and the agents' own numbering is meaningful.
            p = doc.add_paragraph()
            p.paragraph_format.left_indent = Pt(18)
            _add_runs(p, f"{numbered.group(2)}. {numbered.group(3)}")
            i += 1
            continue

        if line.strip() == "":
            i += 1
            continue

        p = doc.add_paragraph()
        _add_runs(p, line.strip())
        i += 1

    return doc


def convert_report_file(md_path: str | Path, output_path: Optional[str | Path] = None) -> Path:
    """Convert a report ``.md`` file to ``.docx``.

    ``output_path`` defaults to the same stem with a ``.docx`` suffix, written
    alongside the source (e.g. ``CF_2026-09-18.md`` → ``CF_2026-09-18.docx``).
    """
    src = Path(md_path)
    text = src.read_text(encoding="utf-8-sig", errors="replace")
    out = Path(output_path) if output_path else src.with_suffix(".docx")
    return md_to_docx(text, out)


def md_to_docx_bytes(md_text: str) -> bytes:
    """Convert Markdown text to ``.docx`` bytes (for email attachments).

    Same conversion as :func:`md_to_docx`, returned in memory so the email
    layer can attach the document without temp files.
    """
    doc = _build_document(md_text)
    buf = io.BytesIO()
    doc.save(buf)
    return buf.getvalue()


def _main(argv: Optional[list[str]] = None) -> int:
    argv = list(sys.argv[1:] if argv is None else argv)
    if not argv or argv in (["-h"], ["--help"]):
        print(__doc__)
        return 0 if argv else 2

    paths = [Path(a) for a in argv if not a.startswith("-")]
    explicit_out = None
    if "-o" in argv:
        idx = argv.index("-o")
        if len(argv) < idx + 2:
            print("error: -o needs an output path", file=sys.stderr)
            return 2
        explicit_out = Path(argv[idx + 1])
        paths = [Path(a) for a in argv[:idx] + argv[idx + 2:] if not a.startswith("-")]

    if not paths:
        print("error: no input file (use '-' for stdin)", file=sys.stderr)
        return 2
    if explicit_out and len(paths) > 1:
        print("error: -o only valid with a single input", file=sys.stderr)
        return 2
    # Two-positional form: ``md_to_docx.py report.md out.docx``.
    if explicit_out is None and len(paths) == 2 and paths[1].suffix.lower() == ".docx":
        explicit_out = paths[1]
        paths = paths[:1]

    for path in paths:
        try:
            if str(path) == "-":
                text = sys.stdin.read()
                if explicit_out is None:
                    print("error: stdin input needs -o OUT.docx", file=sys.stderr)
                    return 2
                md_to_docx(text, explicit_out)
            else:
                out = explicit_out or path.with_suffix(".docx")
                convert_report_file(path, out)
                print(f"{path} -> {out}")
        except Exception as exc:  # noqa: BLE001 - CLI boundary
            print(f"error: {path}: {exc}", file=sys.stderr)
            return 1
    return 0


if __name__ == "__main__":
    raise SystemExit(_main())
