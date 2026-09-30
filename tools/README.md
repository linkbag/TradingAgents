# tools/

Standalone helper scripts for running TradingAgents in production. They are
kept out of the `tradingagents/` package on purpose: no package imports, no
test fixtures — safe to copy onto any machine with a plain Python install.

## md_to_docx.py — Markdown → Word (.docx) report converter

TradingAgents emits its reports as Markdown (per-agent `.md` files plus the
consolidated report). Pipe tables render as raw misaligned text in email
clients and editors, and CJK text needs an explicit East-Asian font binding.
This converter produces a real Word document: ATX headings, bordered tables
with a shaded header row, `**bold**` runs, bullets, numbered items and
horizontal rules, with 微软雅黑 bound as the East-Asian font so Chinese
reports read cleanly in Word/WPS.

Requires `python-docx` (`pip install python-docx`).

```bash
python tools/md_to_docx.py report.md              # -> report.docx alongside
python tools/md_to_docx.py report.md out.docx     # explicit output path
python tools/md_to_docx.py a.md b.md c.md         # batch conversion
type report.md | python tools/md_to_docx.py - out.docx   # stdin (Windows)
```

Library use (same API as the StockSelector v3 `reporting/md_to_docx.py`
module this was extracted from):

```python
import sys; sys.path.insert(0, "tools")
from md_to_docx import md_to_docx, convert_report_file, md_to_docx_bytes

md_to_docx(md_text, "report.docx")     # string -> file
convert_report_file("report.md")       # file -> .docx alongside
md_to_docx_bytes(md_text)              # -> bytes, for email attachments
```

## update_from_upstream.ps1 — weekly upstream sync

Fetches the latest [TauricResearch/TradingAgents](https://github.com/TauricResearch/TradingAgents)
`main`, mirrors it to this fork's `main`, and fast-forwards/rebases the
`stockselector-integration` branch (which carries the GLM `reasoning_effort`
plumbing, the FRED series-id guard, and this converter) onto it. A rebase
that would conflict is aborted automatically — the working tree is left
untouched and the log says a manual merge is needed, so unattended runs are
never left broken. Log: `logs/upstream_update_YYYY-MM-DD.log`.
