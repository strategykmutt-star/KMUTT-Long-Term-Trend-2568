# Google Sheets Data Source — Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** Replace the PPTX-based JSON generator with a Google Sheets-driven flow, so the Strategy Office data collector can update the 20 dashboard charts end-to-end via a "Publish" button — no Git or CLI required.

**Architecture:** Google Sheets workbook (one tab per chart) is the source of truth. An Apps Script "📤 Publish" menu triggers a GitHub Action via `repository_dispatch`. The action runs a Python script that reads the Sheet via the Sheets API, validates, writes `web/src/data/*.json`, commits — then an inline deploy job (same workflow run) builds the React app and publishes to GitHub Pages. The standalone `deploy.yml` is preserved for manual deploys and non-sync pushes.

**Tech Stack:** Python 3.11+, pytest, gspread, google-auth, Google Apps Script, GitHub Actions, TypeScript (React app unchanged except small type fix).

**Spec:** [`docs/superpowers/specs/2026-05-25-google-sheets-data-source-design.md`](../specs/2026-05-25-google-sheets-data-source-design.md) — read this first.

---

## File Structure

**New files:**

```
scripts/
├── __init__.py
├── lib/
│   ├── __init__.py
│   ├── types.py              # ChartData TypedDict, StyleCharts, StyleSeries types
│   ├── parsers.py            # parse_chart_tab(rows) → ChartData; parse_style_*
│   ├── validators.py         # validate(charts, style) → list[Error]
│   ├── writer.py             # write_json + idempotent diff
│   └── sheets_client.py      # gspread wrapper (reads only)
├── sync_from_sheets.py       # main entry: orchestrate parse → validate → write
└── bootstrap_sheets.py       # one-time: 20 JSON files → fresh Sheets workbook

tests/
├── __init__.py
├── fixtures/
│   ├── students-all-rows.json    # mock 2D array as returned by Sheets API
│   ├── style-charts-rows.json
│   ├── style-series-rows.json
│   └── students-all-expected.json # parsed ChartData output
├── test_parsers.py
├── test_validators.py
├── test_writer.py
├── test_sheets_client.py
└── test_sync.py

apps_script/
├── Code.gs                   # Server-side: menu, dispatch, polling RPCs
└── PublishModal.html         # Client-side modal: live polling, result rendering

.github/workflows/
└── sync-from-sheets.yml

docs/
├── data-collector-guide-th.md  # Thai-language manual for data collector
└── architecture.md             # English dev-facing architecture doc

# Project-root files
requirements-dev.txt           # pytest, gspread, google-auth
pyproject.toml                 # pytest config
```

**Modified files:**

- `web/src/types.ts` — drop `slide: number`, tighten `subtitle: Bilingual | null` → `subtitle: Bilingual` (Phase 0.5)
- `web/src/components/ChartCard.tsx` — remove now-unnecessary null-check on subtitle (Phase 0.5)
- `web/src/data/*.json` (all 20) — strip `slide` key in Phase 0.5 so first dry-run shows zero diff
- `build_chart_json.py` — add `DEPRECATED` header (Phase 6)
- `README.md` — replace "Updating data from a new PPT" section with Sheets workflow (Phase 6)

---

## Phase 0: Project Setup

### Task 0.1: Python project scaffolding

**Files:**
- Create: `scripts/__init__.py`, `scripts/lib/__init__.py`, `tests/__init__.py`
- Create: `requirements-dev.txt`
- Create: `pyproject.toml`

- [ ] **Step 1: Create empty package init files**

```bash
mkdir -p scripts/lib tests/fixtures
touch scripts/__init__.py scripts/lib/__init__.py tests/__init__.py
```

- [ ] **Step 2: Write requirements-dev.txt**

```
gspread==6.1.2
google-auth==2.35.0
pytest==8.3.3
pytest-mock==3.14.0
```

- [ ] **Step 3: Write pyproject.toml (minimal pytest config)**

```toml
[tool.pytest.ini_options]
testpaths = ["tests"]
python_files = ["test_*.py"]
pythonpath = ["."]
```

- [ ] **Step 4: Install dependencies and verify pytest runs**

```bash
pip install -r requirements-dev.txt
pytest --collect-only
```
Expected: "no tests ran" (no test files yet) — not an error.

- [ ] **Step 5: Commit**

```bash
git add scripts/ tests/ requirements-dev.txt pyproject.toml
git commit -m "chore: scaffold Python project for Sheets sync"
```

---

## Phase 0.5: Schema-cleanup migration (one-time)

This is the migration commit referenced as **step 0** of the spec's Initial
Migration section. Doing it now (before any new code runs) means the first
dry-run after Sheets bootstrap will report 0 changed files — the proof that
the new round-trip is byte-stable.

### Task 0.5.1: Strip `slide` from JSON, tighten types

**Files:**
- Modify: all 20 files in `web/src/data/*.json`
- Modify: `web/src/types.ts`
- Modify: `web/src/components/ChartCard.tsx`

- [ ] **Step 1: Create a one-shot strip script (works on Windows + Unix)**

Create `scripts/_strip_slide.py`:

```python
"""One-time: strip the legacy `slide` field from all chart JSON files
ahead of the Sheets-driven sync. Runs identically on PowerShell, cmd,
bash, etc. — no shell heredoc required.

Idempotent: re-running on already-stripped files is a no-op."""
import json
from pathlib import Path

DATA_DIR = Path("web/src/data")

def main():
    stripped = 0
    for p in sorted(DATA_DIR.glob("*.json")):
        obj = json.loads(p.read_text(encoding="utf-8"))
        if "slide" in obj:
            obj.pop("slide")
            p.write_text(
                json.dumps(obj, ensure_ascii=False, indent=2) + "\n",
                encoding="utf-8",
            )
            stripped += 1
            print(f"stripped slide from {p.name}")
        else:
            print(f"  ({p.name} already clean)")
    print(f"Done: {stripped} file(s) modified.")

if __name__ == "__main__":
    main()
```

Run it:

```bash
python scripts/_strip_slide.py
```

- [ ] **Step 2: Verify no `slide` keys remain**

Run a tiny Python check (works in any shell):

```bash
python -c "import json,glob,sys; bad=[p for p in glob.glob('web/src/data/*.json') if 'slide' in json.load(open(p,encoding='utf-8'))]; print('LEFTOVERS:',bad) if bad else print('all clean'); sys.exit(1 if bad else 0)"
```
Expected: `all clean`.

- [ ] **Step 3: Update web/src/types.ts**

Replace the full file with:

```typescript
export type Lang = 'th' | 'en'

export interface Bilingual {
  th: string
  en: string
}

export interface ChartSeries {
  key: string
  name: Bilingual
  color: string
  values: (number | null)[]
  emphasis?: boolean
  exclude_from_stack?: boolean
  is_cumulative?: boolean
}

export type ChartType = 'line' | 'stacked-bar' | 'clustered-bar'

export interface ChartData {
  id: string
  section: 'education' | 'personnel' | 'research' | 'finance'
  chart_type: ChartType
  title: Bilingual
  subtitle: Bilingual
  categories_buddhist: string[]
  series: ChartSeries[]
  methodology: Bilingual
  source: Bilingual
}
```

- [ ] **Step 4: Update web/src/components/ChartCard.tsx**

Change:

```tsx
{data.subtitle && (
  <p className="mt-1 text-sm text-slate-500">{data.subtitle[lang]}</p>
)}
```

to:

```tsx
<p className="mt-1 text-sm text-slate-500">{data.subtitle[lang]}</p>
```

- [ ] **Step 5: Verify the React app still builds**

```bash
cd web && npm run build
```
Expected: build succeeds with no errors.

- [ ] **Step 6: Commit**

```bash
git add web/src/data/ web/src/types.ts web/src/components/ChartCard.tsx scripts/_strip_slide.py
```

Then commit using a HEREDOC (bash/Git Bash) or PowerShell here-string
(PowerShell). Both work because the message is multi-line.

Bash / Git Bash:
```bash
git commit -m "$(cat <<'EOF'
refactor: drop legacy slide field, require subtitle, reformat JSON

The strip script re-emits every JSON file with indent=2 and a trailing
newline — the diff is large by line count but the only semantic change
is removing the `slide` field from every chart. The writer in
sync_from_sheets.py uses the same indent + trailing newline, so the
first dry-run after bootstrap reports zero changed files.
EOF
)"
```

PowerShell:
```powershell
git commit -m @'
refactor: drop legacy slide field, require subtitle, reformat JSON

The strip script re-emits every JSON file with indent=2 and a trailing
newline — the diff is large by line count but the only semantic change
is removing the `slide` field from every chart. The writer in
sync_from_sheets.py uses the same indent + trailing newline, so the
first dry-run after bootstrap reports zero changed files.
'@
```

---

## Phase 1: Core Library (TDD)

### Task 1.1: Type definitions

**Files:**
- Create: `scripts/lib/types.py`

- [ ] **Step 1: Write types.py**

```python
"""Type definitions for the Sheets-sync pipeline.

Mirrors the JSON schema defined in
docs/superpowers/specs/2026-05-25-google-sheets-data-source-design.md
"""
from typing import TypedDict, Literal


class Bilingual(TypedDict):
    th: str
    en: str


class SeriesData(TypedDict, total=False):
    key: str               # required
    name: Bilingual        # required
    color: str             # required (from STYLE-series)
    values: list[float | None]  # required
    emphasis: bool         # optional
    exclude_from_stack: bool  # optional
    is_cumulative: bool    # optional


Section = Literal["education", "personnel", "research", "finance"]
ChartType = Literal["line", "stacked-bar", "clustered-bar"]


class ChartData(TypedDict):
    id: str
    section: Section
    chart_type: ChartType
    title: Bilingual
    subtitle: Bilingual
    categories_buddhist: list[str]
    series: list[SeriesData]
    methodology: Bilingual
    source: Bilingual


class StyleChart(TypedDict):
    section: Section
    chart_type: ChartType
    kpi_series_key: str  # series key the React KpiCard highlights


class StyleSeries(TypedDict):
    color: str
    flags: list[str]  # subset of ["emphasis", "exclude_from_stack", "is_cumulative"]


class ValidationError(TypedDict):
    tab: str | None
    field: str | None
    message_th: str
    message_en: str
```

- [ ] **Step 2: Commit**

```bash
git add scripts/lib/types.py
git commit -m "feat(sync): add type definitions for chart data and style"
```

### Task 1.2: Test fixtures

**Files:**
- Create: `tests/fixtures/students-all-rows.json`
- Create: `tests/fixtures/style-charts-rows.json`
- Create: `tests/fixtures/style-series-rows.json`
- Create: `tests/fixtures/students-all-expected.json`

- [ ] **Step 1: Create chart-tab fixture (`students-all-rows.json`)**

Build a 2D-array JSON that mirrors what `gspread`'s `worksheet.get_all_values()` returns for a chart tab populated with `students-all` data. Take values from the existing `web/src/data/students-all.json`. The structure should follow the per-chart tab schema in the spec (rows 1–11 metadata, rows 14–16 series header, rows 18+ data).

Example skeleton (truncate the year list for brevity in fixtures — use 5 years not 33):

```json
[
  ["Chart ID", "students-all"],
  ["", ""],
  ["TITLE (TH)", "จำนวนนักศึกษาทั้งหมด"],
  ["TITLE (EN)", "Total Enrolled Students"],
  ["SUBTITLE (TH)", "จำแนกตามระดับการศึกษา (ภาคที่ 2)"],
  ["SUBTITLE (EN)", "By education level (semester 2)"],
  ["", ""],
  ["METHODOLOGY (TH)", "ทุกปีเป็นข้อมูล ณ ภาคการศึกษาที่ 2..."],
  ["METHODOLOGY (EN)", "Data as of semester 2..."],
  ["SOURCE (TH)", "หนังสือรายงานประจำปี..."],
  ["SOURCE (EN)", "Annual reports..."],
  ["", ""],
  ["— data table ↓ —", "", "", ""],
  ["series_key →", "bachelor", "graduate", "total"],
  ["Name TH", "ปริญญาตรี", "บัณฑิตศึกษา", "รวม"],
  ["Name EN", "Bachelor's", "Graduate (M+P)", "Total"],
  ["Year (พ.ศ.)", "", "", ""],
  ["2564", "11541", "2930", "14471"],
  ["2565", "11731", "2751", "14482"],
  ["2566", "11905", "2577", "14482"],
  ["2567", "12084", "2566", "14650"],
  ["2568", "12253", "2472", "14725"]
]
```

- [ ] **Step 2: Create STYLE-charts fixture**

```json
[
  ["chart_id", "section", "chart_type", "kpi_series_key"],
  ["students-all", "education", "line", "total"],
  ["patents", "research", "clustered-bar", "patent_filed"]
]
```

- [ ] **Step 3: Create STYLE-series fixture**

```json
[
  ["chart_id", "series_key", "color", "flags"],
  ["students-all", "bachelor", "#f29400", ""],
  ["students-all", "graduate", "#1e6091", ""],
  ["students-all", "total", "#0f172a", "emphasis"]
]
```

- [ ] **Step 4: Create expected parsed output**

`students-all-expected.json` — what `parse_chart_tab` should return after joining with STYLE:

```json
{
  "id": "students-all",
  "section": "education",
  "chart_type": "line",
  "title": {"th": "จำนวนนักศึกษาทั้งหมด", "en": "Total Enrolled Students"},
  "subtitle": {"th": "จำแนกตามระดับการศึกษา (ภาคที่ 2)", "en": "By education level (semester 2)"},
  "categories_buddhist": ["2564", "2565", "2566", "2567", "2568"],
  "series": [
    {"key": "bachelor", "name": {"th": "ปริญญาตรี", "en": "Bachelor's"}, "color": "#f29400", "values": [11541, 11731, 11905, 12084, 12253]},
    {"key": "graduate", "name": {"th": "บัณฑิตศึกษา", "en": "Graduate (M+P)"}, "color": "#1e6091", "values": [2930, 2751, 2577, 2566, 2472]},
    {"key": "total", "name": {"th": "รวม", "en": "Total"}, "color": "#0f172a", "values": [14471, 14482, 14482, 14650, 14725], "emphasis": true}
  ],
  "methodology": {"th": "ทุกปีเป็นข้อมูล ณ ภาคการศึกษาที่ 2...", "en": "Data as of semester 2..."},
  "source": {"th": "หนังสือรายงานประจำปี...", "en": "Annual reports..."}
}
```

- [ ] **Step 5: Commit**

```bash
git add tests/fixtures/
git commit -m "test(sync): add fixtures for parser tests"
```

### Task 1.3: Parser — STYLE tabs

**Files:**
- Create: `scripts/lib/parsers.py`
- Create: `tests/test_parsers.py`

- [ ] **Step 1: Write failing tests for parse_style_charts and parse_style_series**

```python
# tests/test_parsers.py
import json
from pathlib import Path
import pytest
from scripts.lib.parsers import parse_style_charts, parse_style_series

FIXTURES = Path(__file__).parent / "fixtures"

def load(name):
    return json.loads((FIXTURES / name).read_text(encoding="utf-8"))

def test_parse_style_charts_returns_dict_keyed_by_chart_id():
    rows = load("style-charts-rows.json")
    result = parse_style_charts(rows)
    assert result == {
        "students-all": {"section": "education", "chart_type": "line", "kpi_series_key": "total"},
        "patents": {"section": "research", "chart_type": "clustered-bar", "kpi_series_key": "patent_filed"},
    }

def test_parse_style_charts_skips_blank_rows():
    rows = [
        ["chart_id", "section", "chart_type", "kpi_series_key"],
        ["", "", "", ""],
        ["students-all", "education", "line", "total"],
    ]
    assert parse_style_charts(rows) == {
        "students-all": {"section": "education", "chart_type": "line", "kpi_series_key": "total"},
    }

def test_parse_style_series_returns_dict_keyed_by_tuple():
    rows = load("style-series-rows.json")
    result = parse_style_series(rows)
    assert result[("students-all", "bachelor")] == {"color": "#f29400", "flags": []}
    assert result[("students-all", "total")] == {"color": "#0f172a", "flags": ["emphasis"]}

def test_parse_style_series_handles_multiple_flags():
    rows = [
        ["chart_id", "series_key", "color", "flags"],
        ["x", "y", "#000000", "emphasis,is_cumulative"],
    ]
    assert parse_style_series(rows)[("x", "y")]["flags"] == ["emphasis", "is_cumulative"]
```

- [ ] **Step 2: Run tests, verify they fail**

```bash
pytest tests/test_parsers.py -v
```
Expected: ImportError or AttributeError.

- [ ] **Step 3: Implement parse_style_charts and parse_style_series**

```python
# scripts/lib/parsers.py
"""Parsers convert Sheets API row data (list of lists) into typed dicts."""
from .types import StyleChart, StyleSeries, ChartData


def parse_style_charts(rows: list[list[str]]) -> dict[str, StyleChart]:
    """Parse STYLE-charts tab. First row is header, skip blank rows.

    Schema: chart_id | section | chart_type | kpi_series_key
    """
    out: dict[str, StyleChart] = {}
    for row in rows[1:]:
        if not row or not row[0].strip():
            continue
        chart_id, section, chart_type, kpi_key = (row + ["", "", "", ""])[:4]
        out[chart_id.strip()] = {
            "section": section.strip(),  # type: ignore[typeddict-item]
            "chart_type": chart_type.strip(),  # type: ignore[typeddict-item]
            "kpi_series_key": kpi_key.strip(),
        }
    return out


def parse_style_series(rows: list[list[str]]) -> dict[tuple[str, str], StyleSeries]:
    """Parse STYLE-series tab. Returns dict keyed by (chart_id, series_key)."""
    out: dict[tuple[str, str], StyleSeries] = {}
    for row in rows[1:]:
        if not row or not row[0].strip():
            continue
        chart_id, series_key, color, flags_raw = (row + ["", "", "", ""])[:4]
        flags = [f.strip() for f in flags_raw.split(",") if f.strip()]
        out[(chart_id.strip(), series_key.strip())] = {
            "color": color.strip(),
            "flags": flags,
        }
    return out
```

- [ ] **Step 4: Run tests, verify pass**

```bash
pytest tests/test_parsers.py -v
```

- [ ] **Step 5: Commit**

```bash
git add scripts/lib/parsers.py tests/test_parsers.py
git commit -m "feat(sync): parse STYLE-charts and STYLE-series tabs"
```

### Task 1.4: Parser — chart tab

**Files:**
- Modify: `scripts/lib/parsers.py`
- Modify: `tests/test_parsers.py`

- [ ] **Step 1: Add failing test for parse_chart_tab**

```python
# Append to tests/test_parsers.py
from scripts.lib.parsers import parse_chart_tab

def test_parse_chart_tab_returns_expected_chartdata():
    rows = load("students-all-rows.json")
    style_charts = parse_style_charts(load("style-charts-rows.json"))
    style_series = parse_style_series(load("style-series-rows.json"))
    expected = load("students-all-expected.json")

    result = parse_chart_tab(rows, style_charts, style_series)
    assert result == expected

def test_parse_chart_tab_treats_empty_value_cells_as_none():
    rows = load("students-all-rows.json")
    # Fixture data rows (0-indexed): 17=2564, 18=2565, 19=2566, 20=2567, 21=2568
    # Blank the bachelor 2566 cell → values index 2.
    rows[19][1] = ""
    style_charts = parse_style_charts(load("style-charts-rows.json"))
    style_series = parse_style_series(load("style-series-rows.json"))

    result = parse_chart_tab(rows, style_charts, style_series)
    bachelor = next(s for s in result["series"] if s["key"] == "bachelor")
    assert bachelor["values"][2] is None

def test_parse_chart_tab_preserves_blank_series_key_position():
    """Critical: a blank cell in row 14 must NOT cause data columns to shift."""
    rows = load("students-all-rows.json")
    # Blank the "graduate" key in column C of row 14 (0-idx 13)
    rows[13][2] = ""
    style_charts = parse_style_charts(load("style-charts-rows.json"))
    style_series = parse_style_series(load("style-series-rows.json"))

    result = parse_chart_tab(rows, style_charts, style_series)
    # Series list still has 3 entries; the middle one's key is empty
    assert len(result["series"]) == 3
    assert result["series"][0]["key"] == "bachelor"
    assert result["series"][1]["key"] == ""  # the blank
    assert result["series"][2]["key"] == "total"
    # Crucially: bachelor data did NOT absorb the middle column's values
    assert result["series"][0]["values"] == [11541, 11731, 11905, 12084, 12253]
    # The blank-key series still carries its column's data — validator will reject it
    assert result["series"][1]["values"] == [2930, 2751, 2577, 2566, 2472]

def test_parse_chart_tab_rejects_blank_year_with_data():
    """Critical: blanking the year cell while data remains in B+ must raise,
    not silently drop the row (which would erase a year's data from the
    dashboard without any warning)."""
    rows = load("students-all-rows.json")
    # Blank the year cell for 2566 (0-idx 19, column A) but leave bachelor/etc intact
    rows[19][0] = ""
    style_charts = parse_style_charts(load("style-charts-rows.json"))
    style_series = parse_style_series(load("style-series-rows.json"))

    with pytest.raises(ValueError, match="year cell is blank but value columns are not"):
        parse_chart_tab(rows, style_charts, style_series)

def test_parse_chart_tab_skips_completely_blank_rows():
    """A row that is entirely blank (year + every value cell) is skipped
    silently — this is normal and lets the data collector leave trailing
    blank rows in the sheet."""
    rows = load("students-all-rows.json")
    # Wipe the entire last row (2568 + all values)
    rows[21] = ["", "", "", ""]
    style_charts = parse_style_charts(load("style-charts-rows.json"))
    style_series = parse_style_series(load("style-series-rows.json"))

    result = parse_chart_tab(rows, style_charts, style_series)
    # Only 4 years now (2564, 2565, 2566, 2567)
    assert result["categories_buddhist"] == ["2564", "2565", "2566", "2567"]

def test_parse_chart_tab_detects_deleted_metadata_row():
    """Critical: if a user deletes a metadata row (e.g. Name TH), every
    row below shifts up by one. The parser would otherwise silently read
    Name EN as data values. Sentinel check at A17 must catch this."""
    rows = load("students-all-rows.json")
    # Simulate "Name TH" row (0-idx 14) being deleted
    del rows[14]
    style_charts = parse_style_charts(load("style-charts-rows.json"))
    style_series = parse_style_series(load("style-series-rows.json"))

    with pytest.raises(ValueError, match="expected 'Year"):
        parse_chart_tab(rows, style_charts, style_series)

def test_parse_chart_tab_accepts_year_range_strings():
    """The 3yr charts use NNNN-NNNN year ranges. Parser must accept them
    as strings (validator checks the format)."""
    rows = load("students-all-rows.json")
    # Replace all year cells with range strings
    for i, y in enumerate(["2536-2538", "2539-2541", "2542-2544", "2545-2547", "2548-2550"]):
        rows[17 + i][0] = y
    style_charts = parse_style_charts(load("style-charts-rows.json"))
    style_series = parse_style_series(load("style-series-rows.json"))

    result = parse_chart_tab(rows, style_charts, style_series)
    assert result["categories_buddhist"] == [
        "2536-2538", "2539-2541", "2542-2544", "2545-2547", "2548-2550"
    ]
```

- [ ] **Step 2: Verify tests fail**

- [ ] **Step 3: Implement parse_chart_tab**

**Key design point:** the parser MUST preserve column positions in the
series_key row. If row 14 has `[bachelor, "", graduate, total]`, the
parser emits a 4-entry series list with one empty-key entry. The validator
catches that and fails. If the parser collapsed to 3 entries, the data
columns would silently shift onto the wrong series. STYLE-series lookup
for empty keys is skipped (those are reported by the validator).

```python
# Append to scripts/lib/parsers.py
DATA_START_ROW = 17  # 0-indexed: row 18 in spec = index 17
YEAR_HEADER_ROW = 16  # 0-indexed: row 17 in spec = "Year (พ.ศ.)"


def parse_chart_tab(
    rows: list[list[str]],
    style_charts: dict,
    style_series: dict,
) -> ChartData:
    """Parse a single chart tab. Joins with STYLE config for color/section/chart_type.

    Preserves column positions in the series_key row so the validator can
    detect blank/duplicate keys instead of silently shifting data columns.
    Verifies a sentinel cell (A17 = "Year ...") to catch the case where
    the data collector deleted an entire metadata row (which would shift
    all rows below and silently misalign the parser).
    """
    chart_id = rows[0][1].strip()
    style_c = style_charts.get(chart_id)
    if style_c is None:
        raise KeyError(f"Chart id '{chart_id}' missing from STYLE-charts")

    # Sanity check: the row that should hold "Year (พ.ศ.)" must actually do so.
    # If not, rows were deleted/shifted; refuse to read further so we don't
    # silently publish garbage (e.g. reading Name EN as data values).
    if YEAR_HEADER_ROW >= len(rows):
        raise ValueError(f"chart tab has fewer than {YEAR_HEADER_ROW + 1} rows — "
                         f"required metadata rows are missing")
    year_header = rows[YEAR_HEADER_ROW][0].strip() if rows[YEAR_HEADER_ROW] else ""
    if not year_header.lower().startswith("year"):
        raise ValueError(
            f"expected 'Year (พ.ศ.)' header at A17, got '{year_header}'. "
            f"This usually means a row was accidentally deleted; please "
            f"undo and try again."
        )

    def cell(r: int) -> str:
        return rows[r][1].strip() if len(rows[r]) > 1 else ""

    # Series header at rows 13,14,15 (0-indexed). Determine the
    # rightmost non-blank column in the entire header block; truncate
    # trailing all-blank columns but keep mid-block blanks.
    raw_keys = list(rows[13][1:]) if len(rows[13]) > 1 else []
    raw_names_th = list(rows[14][1:]) if len(rows[14]) > 1 else []
    raw_names_en = list(rows[15][1:]) if len(rows[15]) > 1 else []
    width = max(len(raw_keys), len(raw_names_th), len(raw_names_en))
    # Right-trim columns that are blank in ALL three header rows
    while width > 0:
        i = width - 1
        last_key = raw_keys[i].strip() if i < len(raw_keys) else ""
        last_th = raw_names_th[i].strip() if i < len(raw_names_th) else ""
        last_en = raw_names_en[i].strip() if i < len(raw_names_en) else ""
        if last_key or last_th or last_en:
            break
        width -= 1

    def at(lst: list, i: int) -> str:
        return lst[i].strip() if i < len(lst) else ""

    keys = [at(raw_keys, i) for i in range(width)]
    names_th = [at(raw_names_th, i) for i in range(width)]
    names_en = [at(raw_names_en, i) for i in range(width)]

    # Data table from row 17 onward (0-indexed).
    #
    # Skipping rule: only skip rows that are ENTIRELY blank (year + every
    # value column). If the year cell is blank but any data cell has a
    # value, that's a corrupted row — the data collector likely deleted
    # the year but forgot the data. Raising forces a validation error
    # rather than silently losing the whole row's data.
    years: list[str] = []
    values_by_col: list[list[float | None]] = [[] for _ in range(width)]
    for row_idx, row in enumerate(rows[DATA_START_ROW:], start=DATA_START_ROW):
        if not row:
            continue
        year_cell = row[0].strip() if len(row) > 0 and isinstance(row[0], str) else str(row[0] or "").strip()
        data_cells = []
        for i in range(width):
            v = row[i + 1] if len(row) > i + 1 else ""
            v = v.strip() if isinstance(v, str) else v
            data_cells.append(v)
        any_data = any(c not in ("", None) for c in data_cells)
        if not year_cell:
            if any_data:
                # Sheet row number = row_idx + 1 (Sheets is 1-indexed)
                raise ValueError(
                    f"row {row_idx + 1}: year cell is blank but value columns are not — "
                    f"refusing to silently drop the row"
                )
            continue  # truly empty row, ignore
        years.append(year_cell)
        for i in range(width):
            cell_val = data_cells[i]
            if cell_val == "" or cell_val is None:
                values_by_col[i].append(None)
            else:
                values_by_col[i].append(float(cell_val))

    series = []
    for i, key in enumerate(keys):
        entry: dict = {
            "key": key,
            "name": {"th": names_th[i], "en": names_en[i]},
            "values": values_by_col[i],
        }
        # Only join with STYLE-series for non-blank keys; blanks are
        # caught by the validator.
        if key:
            style_s = style_series.get((chart_id, key))
            if style_s is not None:
                entry["color"] = style_s["color"]
                for flag in style_s["flags"]:
                    entry[flag] = True
            else:
                entry["color"] = ""  # validator will report missing STYLE-series
        else:
            entry["color"] = ""
        series.append(entry)

    # Convert floats with no fractional part back to int for cleaner JSON
    for s in series:
        s["values"] = [int(v) if isinstance(v, float) and v.is_integer() else v
                       for v in s["values"]]

    return {
        "id": chart_id,
        "section": style_c["section"],
        "chart_type": style_c["chart_type"],
        "title": {"th": cell(2), "en": cell(3)},
        "subtitle": {"th": cell(4), "en": cell(5)},
        "categories_buddhist": years,
        "series": series,
        "methodology": {"th": cell(7), "en": cell(8)},
        "source": {"th": cell(9), "en": cell(10)},
    }
```

- [ ] **Step 4: Verify tests pass**

- [ ] **Step 5: Commit**

```bash
git add scripts/lib/parsers.py tests/test_parsers.py
git commit -m "feat(sync): parse chart tab into ChartData"
```

### Task 1.5: Validator

**Files:**
- Create: `scripts/lib/validators.py`
- Create: `tests/test_validators.py`

- [ ] **Step 1: Write failing tests covering each validation rule**

```python
# tests/test_validators.py
from scripts.lib.validators import validate

GOOD_CHART = {
    "id": "students-all", "section": "education", "chart_type": "line",
    "title": {"th": "x", "en": "y"}, "subtitle": {"th": "x", "en": "y"},
    "categories_buddhist": ["2566", "2567", "2568"],
    "series": [
        {"key": "bachelor", "name": {"th": "x", "en": "y"}, "color": "#000000",
         "values": [1, 2, 3]},
    ],
    "methodology": {"th": "x", "en": "y"},
    "source": {"th": "x", "en": "y"},
}
STYLE_CHARTS = {"students-all": {"section": "education", "chart_type": "line",
                                  "kpi_series_key": "bachelor"}}
STYLE_SERIES = {("students-all", "bachelor"): {"color": "#000000", "flags": []}}

def test_valid_chart_produces_no_errors():
    errors = validate({"EDU-students-all": GOOD_CHART}, STYLE_CHARTS, STYLE_SERIES)
    assert errors == []

def test_duplicate_years_flagged():
    bad = {**GOOD_CHART, "categories_buddhist": ["2566", "2566", "2568"]}
    errors = validate({"EDU-students-all": bad}, STYLE_CHARTS, STYLE_SERIES)
    assert any("duplicate" in e["message_en"].lower() for e in errors)

def test_unsorted_years_flagged():
    bad = {**GOOD_CHART, "categories_buddhist": ["2568", "2566", "2567"]}
    errors = validate({"EDU-students-all": bad}, STYLE_CHARTS, STYLE_SERIES)
    assert any("sorted" in e["message_en"].lower() for e in errors)

def test_value_count_mismatch_flagged():
    bad_series = {**GOOD_CHART["series"][0], "values": [1, 2]}  # only 2 of 3 years
    bad = {**GOOD_CHART, "series": [bad_series]}
    errors = validate({"EDU-students-all": bad}, STYLE_CHARTS, STYLE_SERIES)
    assert any("value count" in e["message_en"].lower() for e in errors)

def test_missing_text_field_flagged():
    bad = {**GOOD_CHART, "methodology": {"th": "", "en": "x"}}
    errors = validate({"EDU-students-all": bad}, STYLE_CHARTS, STYLE_SERIES)
    assert any("methodology" in e["message_en"].lower() for e in errors)

def test_chart_id_not_in_style_flagged():
    bad = {**GOOD_CHART, "id": "ghost"}
    errors = validate({"EDU-ghost": bad}, STYLE_CHARTS, STYLE_SERIES)
    assert any("ghost" in e["message_en"] and "STYLE-charts" in e["message_en"] for e in errors)

def test_series_key_not_in_style_flagged():
    bad_series = {**GOOD_CHART["series"][0], "key": "phantom"}
    bad = {**GOOD_CHART, "series": [bad_series]}
    errors = validate({"EDU-students-all": bad}, STYLE_CHARTS, STYLE_SERIES)
    assert any("phantom" in e["message_en"] and "STYLE-series" in e["message_en"] for e in errors)

def test_duplicate_series_key_flagged():
    s = GOOD_CHART["series"][0]
    bad = {**GOOD_CHART, "series": [s, s]}
    errors = validate({"EDU-students-all": bad}, STYLE_CHARTS, STYLE_SERIES)
    assert any("duplicate series_key" in e["message_en"].lower() for e in errors)

def test_missing_chart_tab_for_declared_chart_id():
    errors = validate({}, STYLE_CHARTS, STYLE_SERIES)
    assert any("missing chart tab" in e["message_en"].lower()
               and "students-all" in e["message_en"] for e in errors)

def test_year_out_of_range_flagged():
    bad = {**GOOD_CHART, "categories_buddhist": ["2566", "2567", "9999"]}
    errors = validate({"EDU-students-all": bad}, STYLE_CHARTS, STYLE_SERIES)
    assert any("9999" in e["message_en"] and "outside" in e["message_en"] for e in errors)

def test_non_integer_year_flagged():
    bad = {**GOOD_CHART, "categories_buddhist": ["2566", "abc", "2568"]}
    errors = validate({"EDU-students-all": bad}, STYLE_CHARTS, STYLE_SERIES)
    assert any("'abc'" in e["message_en"] and "integer" in e["message_en"] for e in errors)

def test_blank_series_name_flagged():
    bad_series = {**GOOD_CHART["series"][0], "name": {"th": "", "en": "Bachelor's"}}
    bad = {**GOOD_CHART, "series": [bad_series]}
    errors = validate({"EDU-students-all": bad}, STYLE_CHARTS, STYLE_SERIES)
    assert any("name (TH)" in e["message_en"] for e in errors)

def test_blank_series_key_in_middle_flagged():
    """Critical: parser keeps blank-key entries; validator must surface them."""
    blank = {"key": "", "name": {"th": "", "en": ""}, "color": "", "values": [1, 2, 3]}
    bad = {**GOOD_CHART, "series": [GOOD_CHART["series"][0], blank]}
    errors = validate({"EDU-students-all": bad}, STYLE_CHARTS, STYLE_SERIES)
    assert any("empty series_key" in e["message_en"] for e in errors)

def test_invalid_hex_color_in_style_flagged():
    bad_style = {("students-all", "bachelor"): {"color": "red", "flags": []}}
    errors = validate({"EDU-students-all": GOOD_CHART}, STYLE_CHARTS, bad_style)
    assert any("invalid hex color" in e["message_en"] for e in errors)

def test_unknown_flag_in_style_flagged():
    bad_style = {("students-all", "bachelor"): {"color": "#000000", "flags": ["bogus"]}}
    errors = validate({"EDU-students-all": GOOD_CHART}, STYLE_CHARTS, bad_style)
    assert any("unknown flag" in e["message_en"] and "bogus" in e["message_en"] for e in errors)

def test_series_declared_in_style_but_missing_from_tab_flagged():
    """Critical: prevents data collector from deleting a header row and
    crashing the React KpiCard which does `series.values` directly."""
    # STYLE-series declares two series; chart only has 'bachelor'
    style = {
        ("students-all", "bachelor"): {"color": "#000000", "flags": []},
        ("students-all", "extra"):    {"color": "#111111", "flags": []},
    }
    errors = validate({"EDU-students-all": GOOD_CHART}, STYLE_CHARTS, style)
    assert any("'extra'" in e["message_en"] and "missing from chart tab" in e["message_en"]
               for e in errors)

def test_chart_with_zero_series_flagged():
    """Critical: catches the case where user blanks BOTH STYLE-series
    rows AND chart-tab headers — previous cross-check would pass with
    empty expected_by_chart, but publishing series:[] crashes React."""
    bad = {**GOOD_CHART, "series": []}
    # Empty STYLE-series so cross-check passes
    errors = validate({"EDU-students-all": bad}, STYLE_CHARTS, {})
    assert any("zero series" in e["message_en"] for e in errors)

def test_kpi_series_key_not_in_chart_flagged():
    """Catches a STYLE-charts kpi_series_key that doesn't match any
    series in the chart tab — KpiCard would silently fall back to the
    wrong series."""
    # STYLE-charts says KPI is 'phantom' but chart only has 'bachelor'
    style_charts = {"students-all": {"section": "education", "chart_type": "line",
                                      "kpi_series_key": "phantom"}}
    style_series = {("students-all", "bachelor"): {"color": "#000000", "flags": []}}
    errors = validate({"EDU-students-all": GOOD_CHART}, style_charts, style_series)
    assert any("kpi_series_key 'phantom'" in e["message_en"] for e in errors)

def test_blank_kpi_series_key_flagged():
    """Blank STYLE-charts.kpi_series_key must be rejected explicitly —
    falling back to series[0] in KpiCard would silently publish the
    wrong KPI without any signal at validation time."""
    style_charts = {"students-all": {"section": "education", "chart_type": "line",
                                      "kpi_series_key": ""}}
    errors = validate({"EDU-students-all": GOOD_CHART}, style_charts, STYLE_SERIES)
    assert any("must declare kpi_series_key" in e["message_en"] for e in errors)

def test_year_range_string_accepted():
    """The *-3yr charts use NNNN-NNNN year ranges. These must validate."""
    chart = {**GOOD_CHART,
             "categories_buddhist": ["2536-2538", "2539-2541", "2542-2544"]}
    # Adjust series values to match year count
    chart["series"] = [{**chart["series"][0], "values": [1, 2, 3]}]
    errors = validate({"EDU-students-all": chart}, STYLE_CHARTS, STYLE_SERIES)
    assert errors == []

def test_year_range_with_endpoint_out_of_range_flagged():
    chart = {**GOOD_CHART, "categories_buddhist": ["2536-9999"]}
    chart["series"] = [{**chart["series"][0], "values": [1]}]
    errors = validate({"EDU-students-all": chart}, STYLE_CHARTS, STYLE_SERIES)
    assert any("outside" in e["message_en"] for e in errors)

def test_year_range_start_greater_than_end_flagged():
    chart = {**GOOD_CHART, "categories_buddhist": ["2566-2564"]}
    chart["series"] = [{**chart["series"][0], "values": [1]}]
    errors = validate({"EDU-students-all": chart}, STYLE_CHARTS, STYLE_SERIES)
    assert any("start > end" in e["message_en"] for e in errors)

def test_duplicate_chart_id_across_tabs_flagged():
    """Critical: prevents silent overwrite when user duplicates a tab.
    Two different tabs both claim chart_id='students-all' — write_charts
    would silently overwrite one with the other."""
    errors = validate(
        {"EDU-students-all": GOOD_CHART, "EDU-students-all-copy": GOOD_CHART},
        STYLE_CHARTS, STYLE_SERIES,
    )
    assert any("duplicate chart_id" in e["message_en"] and "students-all" in e["message_en"]
               for e in errors)
```

- [ ] **Step 2: Verify tests fail**

- [ ] **Step 3: Implement validators.py**

```python
# scripts/lib/validators.py
"""Validation rules for parsed chart data.

All errors are returned in both Thai (for the data collector's UI modal)
and English (for the GitHub Action log).
"""
import re
from .types import ChartData, StyleChart, StyleSeries, ValidationError

REQUIRED_TEXT_FIELDS = [
    ("title", "th"), ("title", "en"),
    ("subtitle", "th"), ("subtitle", "en"),
    ("methodology", "th"), ("methodology", "en"),
    ("source", "th"), ("source", "en"),
]
ALLOWED_FLAGS = {"emphasis", "exclude_from_stack", "is_cumulative"}
HEX_COLOR_RE = re.compile(r"^#[0-9a-fA-F]{6}$")
# Year cell may be a plain 4-digit BE year OR a NNNN-NNNN BE range
# (the *-3yr charts and patents row 1 use range format).
YEAR_RE = re.compile(r"^(\d{4})(?:-(\d{4}))?$")
YEAR_MIN, YEAR_MAX = 2500, 2700


def year_sort_key(y: str) -> int:
    """Sort years by their start year so ranges and plain years interleave."""
    m = YEAR_RE.match(y)
    return int(m.group(1)) if m else -1


def _err(tab: str | None, field: str | None, th: str, en: str) -> ValidationError:
    return {"tab": tab, "field": field, "message_th": th, "message_en": en}


def validate(
    parsed_charts: dict[str, ChartData],
    style_charts: dict[str, StyleChart],
    style_series: dict[tuple[str, str], StyleSeries],
) -> list[ValidationError]:
    errors: list[ValidationError] = []

    for tab, data in parsed_charts.items():
        cid = data["id"]

        if cid not in style_charts:
            errors.append(_err(tab, "chart_id",
                f"chart_id '{cid}' ไม่อยู่ใน STYLE-charts",
                f"chart_id '{cid}' missing from STYLE-charts"))

        years = data["categories_buddhist"]
        if len(set(years)) != len(years):
            errors.append(_err(tab, "categories_buddhist",
                "ปีในตารางซ้ำกัน", "duplicate years in data table"))
        if years != sorted(years, key=year_sort_key):
            errors.append(_err(tab, "categories_buddhist",
                "ปีไม่ได้เรียงจากน้อยไปมาก (ตาม start-year)",
                "years not sorted ascending (by start-year)"))
        for y in years:
            m = YEAR_RE.match(y)
            if not m:
                errors.append(_err(tab, "categories_buddhist",
                    f"ปี '{y}' ไม่ใช่ปี 4 หลัก หรือช่วงปี NNNN-NNNN",
                    f"year '{y}' not a 4-digit BE year or NNNN-NNNN range"))
                continue
            start = int(m.group(1))
            end = int(m.group(2)) if m.group(2) else start
            if not (YEAR_MIN <= start <= YEAR_MAX and YEAR_MIN <= end <= YEAR_MAX):
                errors.append(_err(tab, "categories_buddhist",
                    f"ปี '{y}' อยู่นอกช่วง {YEAR_MIN}-{YEAR_MAX}",
                    f"year '{y}' outside {YEAR_MIN}-{YEAR_MAX}"))
            if m.group(2) and start > end:
                errors.append(_err(tab, "categories_buddhist",
                    f"ช่วงปี '{y}' มีต้น > ปลาย",
                    f"year range '{y}' has start > end"))

        # Column-position-preserving series_key check: blanks ARE reported
        # (the parser keeps them as empty strings; we surface them here).
        keys = [s["key"] for s in data["series"]]
        for i, k in enumerate(keys):
            if not k.strip():
                errors.append(_err(tab, "series_key",
                    f"ช่อง series_key ว่างเปล่า (แถว 14, คอลัมน์ {chr(ord('B') + i)})",
                    f"empty series_key in row 14, column {chr(ord('B') + i)}"))
        nonblank = [k for k in keys if k.strip()]
        if len(set(nonblank)) != len(nonblank):
            errors.append(_err(tab, "series_key",
                "series_key ซ้ำกัน (แถว 14)", "duplicate series_key in row 14"))

        for s in data["series"]:
            if not s["key"].strip():
                continue  # already reported above
            if (cid, s["key"]) not in style_series:
                errors.append(_err(tab, f"series.{s['key']}",
                    f"series '{s['key']}' ไม่อยู่ใน STYLE-series",
                    f"series '{s['key']}' missing from STYLE-series"))
            if len(s["values"]) != len(years):
                errors.append(_err(tab, f"series.{s['key']}",
                    f"จำนวนค่า '{s['key']}' ไม่เท่ากับจำนวนปี",
                    f"series '{s['key']}' value count != year count"))
            if not s["name"]["th"].strip():
                errors.append(_err(tab, f"series.{s['key']}.name.th",
                    f"ขาดชื่อ TH ของ series '{s['key']}'",
                    f"missing series name (TH) for '{s['key']}'"))
            if not s["name"]["en"].strip():
                errors.append(_err(tab, f"series.{s['key']}.name.en",
                    f"ขาดชื่อ EN ของ series '{s['key']}'",
                    f"missing series name (EN) for '{s['key']}'"))

        for path in REQUIRED_TEXT_FIELDS:
            value = data.get(path[0], {}).get(path[1], "")  # type: ignore[call-overload]
            if not value or not value.strip():
                field_name = ".".join(path)
                errors.append(_err(tab, field_name,
                    f"ขาดข้อมูล {field_name}", f"missing {field_name}"))

    # STYLE-series sanity checks (run once)
    for (cid, sk), style in style_series.items():
        if not HEX_COLOR_RE.match(style["color"]):
            errors.append(_err(f"🎨 STYLE-series", f"{cid}/{sk}/color",
                f"สี '{style['color']}' ของ {cid}/{sk} ไม่ใช่ hex 6 หลัก",
                f"invalid hex color '{style['color']}' for {cid}/{sk}"))
        for flag in style["flags"]:
            if flag not in ALLOWED_FLAGS:
                errors.append(_err(f"🎨 STYLE-series", f"{cid}/{sk}/flags",
                    f"flag '{flag}' ของ {cid}/{sk} ไม่รู้จัก",
                    f"unknown flag '{flag}' for {cid}/{sk}"))

    # Cross-check: every chart_id in STYLE-charts must have a matching tab
    parsed_ids = {d["id"] for d in parsed_charts.values()}
    for cid in style_charts:
        if cid not in parsed_ids:
            errors.append(_err(None, "tab",
                f"ไม่พบ tab สำหรับกราฟ '{cid}' (อยู่ใน STYLE-charts)",
                f"missing chart tab for '{cid}' (declared in STYLE-charts)"))

    # Cross-check: no two chart tabs may have the same chart_id (e.g. data
    # collector duplicated a tab in Sheets and renamed it). Without this,
    # write_charts() would silently overwrite one chart's JSON with another.
    seen_ids: dict[str, str] = {}
    for tab, data in parsed_charts.items():
        cid = data["id"]
        if cid in seen_ids:
            errors.append(_err(tab, "chart_id",
                f"chart_id '{cid}' ซ้ำกับ tab '{seen_ids[cid]}'",
                f"duplicate chart_id '{cid}' (also in tab '{seen_ids[cid]}')"))
        else:
            seen_ids[cid] = tab

    # Cross-check: every series declared in STYLE-series must exist in its
    # chart tab. Without this, blanking the entire row-14 header would
    # publish series:[] and crash the React KpiCard at runtime.
    expected_by_chart: dict[str, set[str]] = {}
    for (cid, sk) in style_series:
        expected_by_chart.setdefault(cid, set()).add(sk)
    for tab, data in parsed_charts.items():
        cid = data["id"]
        actual_keys = {s["key"] for s in data["series"] if s["key"].strip()}
        for sk in expected_by_chart.get(cid, set()) - actual_keys:
            errors.append(_err(tab, f"series.{sk}",
                f"ขาด series '{sk}' (ระบุไว้ใน STYLE-series)",
                f"series '{sk}' declared in STYLE-series but missing from chart tab"))

    # Every chart must have at least one non-blank series. Catches the
    # case where the user blanked BOTH STYLE-series rows AND the chart
    # tab's row 14 — the previous cross-check would silently pass.
    # Publishing series:[] would crash the React KpiCard.
    for tab, data in parsed_charts.items():
        if not any(s["key"].strip() for s in data["series"]):
            errors.append(_err(tab, "series",
                "กราฟต้องมี series อย่างน้อย 1 ตัว",
                "chart has zero series — at least one is required"))

    # The kpi_series_key declared in STYLE-charts must exist in the
    # chart's actual series list. Without this, KpiCard's fallback to
    # data.series[0] silently shows the wrong KPI. A BLANK kpi_series_key
    # is itself a contract violation — STYLE-charts is documented as the
    # authoritative source for which series the KpiCard highlights, so
    # we reject blanks explicitly rather than silently letting KpiCard
    # fall back to series[0].
    for tab, data in parsed_charts.items():
        cid = data["id"]
        kpi_key = style_charts.get(cid, {}).get("kpi_series_key", "").strip()
        if not kpi_key:
            errors.append(_err(tab, "kpi_series_key",
                "STYLE-charts ต้องระบุ kpi_series_key (กำหนดว่า KPI card แสดง series ใด)",
                "STYLE-charts must declare kpi_series_key (drives which series the KPI card shows)"))
            continue
        actual_keys = {s["key"] for s in data["series"] if s["key"].strip()}
        if kpi_key not in actual_keys:
            errors.append(_err(tab, "kpi_series_key",
                f"kpi_series_key '{kpi_key}' (STYLE-charts) ไม่พบใน series",
                f"kpi_series_key '{kpi_key}' (STYLE-charts) not found in chart series"))

    return errors
```

- [ ] **Step 4: Verify all tests pass**

- [ ] **Step 5: Commit**

```bash
git add scripts/lib/validators.py tests/test_validators.py
git commit -m "feat(sync): validate parsed charts against STYLE config"
```

### Task 1.6: JSON writer with idempotent diff

**Files:**
- Create: `scripts/lib/writer.py`
- Create: `tests/test_writer.py`

- [ ] **Step 1: Write failing tests**

```python
# tests/test_writer.py
import json
from scripts.lib.writer import write_charts, json_equal

CHART = {
    "id": "students-all", "section": "education", "chart_type": "line",
    "title": {"th": "x", "en": "y"}, "subtitle": {"th": "x", "en": "y"},
    "categories_buddhist": ["2566", "2567"],
    "series": [{"key": "bachelor", "name": {"th": "x", "en": "y"},
                "color": "#000000", "values": [1, 2]}],
    "methodology": {"th": "x", "en": "y"},
    "source": {"th": "x", "en": "y"},
}

def test_write_creates_new_files(tmp_path):
    changed = write_charts([CHART], tmp_path)
    assert changed == ["students-all.json"]
    written = json.loads((tmp_path / "students-all.json").read_text(encoding="utf-8"))
    assert written["id"] == "students-all"
    assert written["categories_buddhist"] == ["2566", "2567"]

def test_write_skips_unchanged_files(tmp_path):
    write_charts([CHART], tmp_path)
    changed = write_charts([CHART], tmp_path)
    assert changed == []

def test_write_reports_modified_files(tmp_path):
    write_charts([CHART], tmp_path)
    modified = {**CHART, "categories_buddhist": ["2566", "2567", "2568"],
                "series": [{**CHART["series"][0], "values": [1, 2, 3]}]}
    changed = write_charts([modified], tmp_path)
    assert changed == ["students-all.json"]

def test_json_equal_ignores_key_order(tmp_path):
    a = {"a": 1, "b": 2}
    b = {"b": 2, "a": 1}
    assert json_equal(a, b)
```

- [ ] **Step 2: Verify tests fail**

- [ ] **Step 3: Implement writer.py**

```python
# scripts/lib/writer.py
"""Write ChartData dicts as JSON, idempotent on no-change."""
import json
from pathlib import Path
from .types import ChartData

INDENT = 2


def _canonical(obj) -> str:
    return json.dumps(obj, ensure_ascii=False, indent=INDENT, sort_keys=True)


def json_equal(a, b) -> bool:
    return _canonical(a) == _canonical(b)


def write_charts(charts: list[ChartData], out_dir: Path) -> list[str]:
    """Write one JSON file per chart. Returns list of changed filenames."""
    out_dir.mkdir(parents=True, exist_ok=True)
    changed: list[str] = []
    for chart in charts:
        filename = f"{chart['id']}.json"
        target = out_dir / filename
        if target.exists():
            existing = json.loads(target.read_text(encoding="utf-8"))
            if json_equal(existing, chart):
                continue
        # Write WITHOUT sort_keys so the file matches the natural order
        # callers prefer; equality check uses sort_keys for stability.
        # Append "\n" to match the Phase 0.5 strip-script output — without
        # this, the first sync after schema-cleanup would diff every file
        # purely on trailing-newline mismatch.
        target.write_text(
            json.dumps(chart, ensure_ascii=False, indent=INDENT) + "\n",
            encoding="utf-8",
        )
        changed.append(filename)
    return sorted(changed)
```

- [ ] **Step 4: Verify tests pass**

- [ ] **Step 5: Commit**

```bash
git add scripts/lib/writer.py tests/test_writer.py
git commit -m "feat(sync): write JSON with idempotent diff"
```

### Task 1.7: Sheets client wrapper

**Files:**
- Create: `scripts/lib/sheets_client.py`
- Create: `tests/test_sheets_client.py`

- [ ] **Step 1: Write tests using pytest-mock**

```python
# tests/test_sheets_client.py
from unittest.mock import MagicMock
from scripts.lib.sheets_client import SheetsClient, classify_tab

def test_classify_tab_recognises_section_prefixes():
    assert classify_tab("EDU-students-all") == "chart"
    assert classify_tab("PER-staff-total") == "chart"
    assert classify_tab("RES-patents") == "chart"
    assert classify_tab("FIN-income-expense") == "chart"

def test_classify_tab_recognises_admin_tabs():
    assert classify_tab("📋 INDEX") == "index"
    assert classify_tab("🎨 STYLE-charts") == "style-charts"
    assert classify_tab("🎨 STYLE-series") == "style-series"

def test_classify_tab_ignores_underscore_prefix():
    assert classify_tab("_scratch") == "ignore"

def test_classify_tab_unknown_returns_unknown():
    assert classify_tab("random-tab") == "unknown"

def test_sheets_client_get_all_chart_tabs(mocker):
    fake_gc = mocker.MagicMock()
    fake_sheet = mocker.MagicMock()
    fake_gc.open_by_key.return_value = fake_sheet
    tab1 = mocker.MagicMock(); tab1.title = "EDU-students-all"
    tab1.get_all_values.return_value = [["a"]]
    tab2 = mocker.MagicMock(); tab2.title = "🎨 STYLE-charts"
    tab3 = mocker.MagicMock(); tab3.title = "_scratch"
    fake_sheet.worksheets.return_value = [tab1, tab2, tab3]

    client = SheetsClient(gc=fake_gc, sheet_id="abc")
    chart_tabs = client.get_chart_tabs()
    assert list(chart_tabs.keys()) == ["EDU-students-all"]
```

- [ ] **Step 2: Verify tests fail**

- [ ] **Step 3: Implement sheets_client.py**

```python
# scripts/lib/sheets_client.py
"""Thin wrapper around gspread. Read-only by design."""
from typing import Literal

import gspread
from google.oauth2.service_account import Credentials

TabKind = Literal["chart", "index", "style-charts", "style-series", "ignore", "unknown"]

SECTION_PREFIXES = ("EDU-", "PER-", "RES-", "FIN-")


def classify_tab(name: str) -> TabKind:
    if name.startswith("_"):
        return "ignore"
    if name.startswith(SECTION_PREFIXES):
        return "chart"
    if "INDEX" in name:
        return "index"
    if "STYLE-charts" in name:
        return "style-charts"
    if "STYLE-series" in name:
        return "style-series"
    return "unknown"


class SheetsClient:
    def __init__(self, gc=None, sheet_id: str = "", credentials_path: str | None = None):
        if gc is None:
            scopes = ["https://www.googleapis.com/auth/spreadsheets.readonly"]
            creds = Credentials.from_service_account_file(credentials_path, scopes=scopes)
            gc = gspread.authorize(creds)
        self._gc = gc
        self._sheet = gc.open_by_key(sheet_id)

    def get_chart_tabs(self) -> dict[str, list[list[str]]]:
        out: dict[str, list[list[str]]] = {}
        for ws in self._sheet.worksheets():
            if classify_tab(ws.title) == "chart":
                out[ws.title] = ws.get_all_values()
        return out

    def get_style_charts(self) -> list[list[str]]:
        for ws in self._sheet.worksheets():
            if classify_tab(ws.title) == "style-charts":
                return ws.get_all_values()
        raise RuntimeError("STYLE-charts tab not found")

    def get_style_series(self) -> list[list[str]]:
        for ws in self._sheet.worksheets():
            if classify_tab(ws.title) == "style-series":
                return ws.get_all_values()
        raise RuntimeError("STYLE-series tab not found")
```

- [ ] **Step 4: Verify tests pass**

- [ ] **Step 5: Commit**

```bash
git add scripts/lib/sheets_client.py tests/test_sheets_client.py
git commit -m "feat(sync): add gspread wrapper with tab classification"
```

---

## Phase 2: Main sync script

### Task 2.1: sync_from_sheets.py orchestrator

**Files:**
- Create: `scripts/sync_from_sheets.py`
- Create: `tests/test_sync.py`

- [ ] **Step 1: Write tests for the orchestrator (mocking SheetsClient)**

```python
# tests/test_sync.py
import json
from pathlib import Path
from unittest.mock import MagicMock
from scripts.sync_from_sheets import run_sync

FIX = Path(__file__).parent / "fixtures"

def make_fake_client():
    """Orchestrator test client: declares ONLY students-all in STYLE-charts.

    The shared style-charts-rows.json fixture also lists 'patents' for
    parser/validator tests, but the orchestrator test only provides a
    students-all chart tab — using the full fixture would (correctly)
    trigger the validator's 'missing chart tab for patents' cross-check
    and fail every happy-path test."""
    client = MagicMock()
    client.get_chart_tabs.return_value = {
        "EDU-students-all": json.loads((FIX / "students-all-rows.json").read_text(encoding="utf-8"))
    }
    client.get_style_charts.return_value = [
        ["chart_id", "section", "chart_type", "kpi_series_key"],
        ["students-all", "education", "line", "total"],
    ]
    # style-series fixture already only has students-all entries
    client.get_style_series.return_value = json.loads((FIX / "style-series-rows.json").read_text(encoding="utf-8"))
    return client

def test_run_sync_writes_json_and_returns_changed_list(tmp_path):
    client = make_fake_client()
    result = run_sync(client, out_dir=tmp_path, dry_run=False)
    assert result["status"] == "success"
    assert "students-all.json" in result["changed_files"]
    assert (tmp_path / "students-all.json").exists()

def test_run_sync_dry_run_does_not_write(tmp_path):
    client = make_fake_client()
    result = run_sync(client, out_dir=tmp_path, dry_run=True)
    assert result["status"] == "success"
    assert "students-all.json" in result["changed_files"]
    assert not (tmp_path / "students-all.json").exists()

def test_run_sync_returns_errors_on_validation_failure(tmp_path):
    client = make_fake_client()
    # Corrupt STYLE-charts so chart_id won't match
    client.get_style_charts.return_value = [
        ["chart_id", "section", "chart_type"],
        ["wrong-id", "education", "line"],
    ]
    result = run_sync(client, out_dir=tmp_path, dry_run=False)
    assert result["status"] == "validation_failed"
    assert len(result["errors"]) > 0
```

- [ ] **Step 2: Verify tests fail**

- [ ] **Step 3: Implement sync_from_sheets.py**

```python
# scripts/sync_from_sheets.py
"""Main sync entry point.

Reads a Google Sheets workbook, validates, writes web/src/data/*.json,
and emits a structured result + errors file for the GitHub Action /
Apps Script to consume.
"""
import argparse
import json
import sys
from pathlib import Path

from scripts.lib.sheets_client import SheetsClient
from scripts.lib.parsers import parse_chart_tab, parse_style_charts, parse_style_series
from scripts.lib.validators import validate
from scripts.lib.writer import write_charts


def run_sync(client, out_dir: Path, dry_run: bool) -> dict:
    style_charts = parse_style_charts(client.get_style_charts())
    style_series = parse_style_series(client.get_style_series())
    tab_rows = client.get_chart_tabs()

    parsed = {}
    parse_errors = []
    for tab_name, rows in tab_rows.items():
        try:
            parsed[tab_name] = parse_chart_tab(rows, style_charts, style_series)
        except (KeyError, IndexError, ValueError) as e:
            parse_errors.append({
                "tab": tab_name, "field": None,
                "message_th": f"แปลข้อมูล tab ไม่ได้: {e}",
                "message_en": f"failed to parse tab: {e}",
            })

    errors = parse_errors + validate(parsed, style_charts, style_series)
    if errors:
        return {"status": "validation_failed", "errors": errors,
                "changed_files": []}

    charts = list(parsed.values())
    if dry_run:
        # Compute diff against existing files without writing
        from scripts.lib.writer import json_equal
        changed = []
        for c in charts:
            target = out_dir / f"{c['id']}.json"
            if not target.exists():
                changed.append(f"{c['id']}.json")
            else:
                existing = json.loads(target.read_text(encoding="utf-8"))
                if not json_equal(existing, c):
                    changed.append(f"{c['id']}.json")
        return {"status": "success", "dry_run": True,
                "errors": [], "changed_files": sorted(changed)}

    changed = write_charts(charts, out_dir)
    return {"status": "success", "dry_run": False,
            "errors": [], "changed_files": changed}


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--sheet-id", required=True)
    ap.add_argument("--dry-run", action="store_true")
    ap.add_argument("--out-dir", default="web/src/data")
    ap.add_argument("--result-out", default="result.json")
    ap.add_argument("--errors-out", default="errors.json")
    ap.add_argument("--credentials", default=None,
                    help="Path to service-account JSON. Defaults to "
                         "GOOGLE_APPLICATION_CREDENTIALS env var.")
    args = ap.parse_args()

    client = SheetsClient(sheet_id=args.sheet_id, credentials_path=args.credentials)
    result = run_sync(client, out_dir=Path(args.out_dir), dry_run=args.dry_run)

    if result["status"] == "validation_failed":
        Path(args.errors_out).write_text(
            json.dumps(result["errors"], ensure_ascii=False, indent=2),
            encoding="utf-8")
        print(f"Validation failed with {len(result['errors'])} errors. See {args.errors_out}.")
        sys.exit(1)

    Path(args.result_out).write_text(
        json.dumps(result, ensure_ascii=False, indent=2),
        encoding="utf-8")
    print(f"Sync OK. {len(result['changed_files'])} file(s) changed.")


if __name__ == "__main__":
    main()
```

- [ ] **Step 4: Verify tests pass**

```bash
pytest tests/ -v
```

- [ ] **Step 5: Commit**

```bash
git add scripts/sync_from_sheets.py tests/test_sync.py
git commit -m "feat(sync): orchestrator that ties parser + validator + writer"
```

---

## Phase 3: GitHub Action workflow

### Task 3.1: sync-from-sheets.yml

**Files:**
- Create: `.github/workflows/sync-from-sheets.yml`

- [ ] **Step 1: Write the workflow**

```yaml
# .github/workflows/sync-from-sheets.yml
name: Sync from Sheets

run-name: "Sync from Sheets [${{ github.event.client_payload.correlation_id }}]"

on:
  repository_dispatch:
    types: [sync-sheets]

# pages + id-token are required by the deploy job;
# contents:write is required by the sync job to commit JSON;
# actions:read is required so jobs can read each other's outputs/artifacts.
permissions:
  contents: write
  pages: write
  id-token: write
  actions: read

# Avoid two concurrent publishes racing each other into Pages.
concurrency:
  group: sync-from-sheets
  cancel-in-progress: false

jobs:
  sync:
    runs-on: ubuntu-latest
    outputs:
      did_commit: ${{ steps.commit.outputs.did_commit }}
    steps:
      - uses: actions/checkout@v4
        with:
          token: ${{ secrets.GITHUB_TOKEN }}

      - uses: actions/setup-python@v5
        with:
          python-version: '3.11'

      - run: pip install -r requirements-dev.txt

      - name: Preflight — required config present
        env:
          SHEET_ID: ${{ vars.KMUTT_TRENDS_SHEET_ID }}
        run: |
          if [ -z "$SHEET_ID" ]; then
            # Write a structured errors.json so the Apps Script modal can
            # render a meaningful message instead of "ไม่ทราบสาเหตุ".
            cat > errors.json <<'EOF'
          [{"tab":null,"field":"config","message_th":"GitHub Variable KMUTT_TRENDS_SHEET_ID ยังไม่ได้ตั้งค่า — ติดต่อ dev","message_en":"GitHub repository Variable KMUTT_TRENDS_SHEET_ID is not set — contact the developer"}]
          EOF
            echo "::error::KMUTT_TRENDS_SHEET_ID is unset. Set it in repo Settings → Variables."
            exit 1
          fi

      - name: Write service account credentials
        env:
          GCP_KEY: ${{ secrets.GOOGLE_SERVICE_ACCOUNT_JSON }}
        run: |
          # Use printf via env-var (NOT echo via ${{ … }}) so single-quotes
          # or backslashes inside the JSON cannot break the shell or be
          # interpreted as escape sequences.
          printf '%s' "$GCP_KEY" > /tmp/gcp-key.json

      # SHEET_ID and USER_EMAIL go through `env:` rather than direct
      # `${{ ... }}` interpolation in shell. This prevents script
      # injection from a malicious `repository_dispatch` payload: with
      # env vars, a payload value like `abc"; rm -rf` becomes a literal
      # string assigned to the env var, not shell tokens. See:
      # https://docs.github.com/en/actions/security-for-github-actions/security-guides/security-hardening-for-github-actions#understanding-the-risk-of-script-injections
      #
      # Defense-in-depth: SHEET_ID is read from a GitHub repository
      # Variable, NOT the dispatch payload, so an attacker with PAT
      # cannot point the workflow at a different sheet they control.

      - name: Run sync script (publish)
        if: github.event.client_payload.dry_run != true
        env:
          SHEET_ID: ${{ vars.KMUTT_TRENDS_SHEET_ID }}
        run: |
          python scripts/sync_from_sheets.py \
            --sheet-id "$SHEET_ID" \
            --credentials /tmp/gcp-key.json \
            --result-out result.json \
            --errors-out errors.json

      - name: Run sync script (dry-run)
        if: github.event.client_payload.dry_run == true
        env:
          SHEET_ID: ${{ vars.KMUTT_TRENDS_SHEET_ID }}
        run: |
          python scripts/sync_from_sheets.py \
            --sheet-id "$SHEET_ID" \
            --credentials /tmp/gcp-key.json \
            --result-out result.json \
            --errors-out errors.json \
            --dry-run

      - name: Upload sync-result artifact
        if: success()
        uses: actions/upload-artifact@v4
        with:
          name: sync-result
          path: result.json
          retention-days: 7

      - name: Upload sync-errors artifact
        if: failure() && hashFiles('errors.json') != ''
        uses: actions/upload-artifact@v4
        with:
          name: sync-errors
          path: errors.json
          retention-days: 7

      - name: Commit and push if changes
        id: commit
        if: success() && github.event.client_payload.dry_run != true
        env:
          # USER_EMAIL also goes through env to avoid shell injection in
          # the commit message. Double-quoting `"$USER_EMAIL"` makes
          # bash treat the value as a literal string.
          USER_EMAIL: ${{ github.event.client_payload.user_email }}
        run: |
          if [ -z "$(git status --porcelain web/src/data/)" ]; then
            echo "No JSON changes — skipping commit (idempotent no-op)."
            echo "did_commit=false" >> "$GITHUB_OUTPUT"
            exit 0
          fi
          git config user.name "kmutt-trends-bot"
          git config user.email "bot@users.noreply.github.com"
          git add web/src/data/
          git commit -m "Sync from Sheets by $USER_EMAIL at $(date -u +%FT%TZ)"
          git push
          echo "did_commit=true" >> "$GITHUB_OUTPUT"

  # The deploy job replicates the existing deploy.yml steps and ALWAYS runs
  # after a successful non-dry-run sync — even when no JSON changed. This
  # makes "click Publish again" a valid recovery action when Pages got out
  # of sync with main (e.g. a previous deploy failed). It also ensures the
  # Apps Script modal only reports "success" after the dashboard has
  # actually been redeployed.
  deploy:
    needs: sync
    if: success() && github.event.client_payload.dry_run != true
    runs-on: ubuntu-latest
    # MUST share the 'pages' concurrency group with the existing deploy.yml
    # so a sync-triggered deploy can't race a manual deploy or push-triggered
    # deploy and overwrite the dashboard with a stale artifact. Workflow-
    # level group 'sync-from-sheets' already serializes sync runs; this
    # job-level group also serializes ALL Pages deployments across both
    # workflows.
    concurrency:
      group: pages
      cancel-in-progress: false
    environment:
      name: github-pages
      url: ${{ steps.deployment.outputs.page_url }}
    defaults:
      run:
        working-directory: ./web
    steps:
      # Always check out main HEAD; if sync committed, this includes the
      # new JSON. If sync was a no-op, this is just main.
      - uses: actions/checkout@v4
        with:
          ref: main
      - uses: actions/setup-node@v4
        with:
          node-version: 22
          cache: npm
          cache-dependency-path: web/package-lock.json
      - run: npm ci
      - run: npm run build
        env:
          VITE_BASE: /${{ github.event.repository.name }}/
      - uses: actions/configure-pages@v5
        with:
          enablement: true
      - uses: actions/upload-pages-artifact@v3
        with:
          path: web/dist
      - id: deployment
        uses: actions/deploy-pages@v4
```

- [ ] **Step 2: Sanity check YAML syntax**

```bash
python -c "import yaml; yaml.safe_load(open('.github/workflows/sync-from-sheets.yml'))"
```
Expected: no output (parsed successfully).

- [ ] **Step 3: Commit**

```bash
git add .github/workflows/sync-from-sheets.yml
git commit -m "feat(ci): workflow to sync Sheets → JSON on repository_dispatch"
```

---

## Phase 4: Apps Script

### Task 4.1: Code.gs + PublishModal.html — async publish UX

**Files:**
- Create: `apps_script/Code.gs`
- Create: `apps_script/PublishModal.html`

**Architecture note.** Apps Script server-side functions are capped at
6 minutes per invocation. The earlier draft did synchronous `Utilities.sleep`
polling in a single function call, which blocked the spreadsheet and would
have hit the runtime cap on slow workflows. This revision uses
[HtmlService](https://developers.google.com/apps-script/guides/html) with
client-side polling via `google.script.run` — each server RPC is short
(< 1 second), the modal updates live in the user's browser, and there's
no execution-cap risk.

Apps Script has limited automated-testing infrastructure; this task relies
on careful writing + manual integration testing in Phase 7.

- [ ] **Step 1: Write `apps_script/Code.gs` (server side)**

```javascript
// apps_script/Code.gs
//
// KMUTT Trends Dashboard — publish button (server side).
// Setup: set Script Properties GITHUB_PAT, SHEET_ID, REPO, HELP_URL.
//
// All long-running work lives client-side in PublishModal.html, which calls
// these server functions via google.script.run. Each call returns quickly.
//

function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('📤 Publish to Dashboard')
    .addItem('Check what will change (dry-run)', 'runDryRun')
    .addItem('Publish all changes', 'runPublish')
    .addSeparator()
    .addItem('📖 Help (Thai guide)', 'openHelp')
    .addToUi();
}

function openHelp() {
  const url = PropertiesService.getScriptProperties().getProperty('HELP_URL')
    || 'https://github.com/<org>/<repo>/blob/main/docs/data-collector-guide-th.md';
  const html = HtmlService.createHtmlOutput(
    `<p>เปิดคู่มือผู้รวบรวมข้อมูล:</p>
     <p><a href="${url}" target="_blank">${url}</a></p>`
  ).setWidth(420).setHeight(120);
  SpreadsheetApp.getUi().showModalDialog(html, '📖 คู่มือ');
}

function runDryRun() { _openPublishModal(true); }
function runPublish() { _openPublishModal(false); }

function _openPublishModal(dryRun) {
  // Dispatch first (fast). Then open the modal that polls.
  const dispatch = _dispatchSync(dryRun);
  if (!dispatch.ok) {
    SpreadsheetApp.getUi().alert(`❌ ${dispatch.error}`);
    return;
  }
  const template = HtmlService.createTemplateFromFile('PublishModal');
  template.correlationId = dispatch.correlationId;
  template.dryRun = dryRun;
  template.repo = _props().getProperty('REPO');
  const html = template.evaluate().setWidth(560).setHeight(440);
  SpreadsheetApp.getUi().showModalDialog(
    html, dryRun ? '🔍 ตรวจสอบความเปลี่ยนแปลง' : '📤 Publish to Dashboard'
  );
}

// ---------- functions invoked from the modal via google.script.run ----------

function dispatchSync(dryRun) { return _dispatchSync(dryRun); }

function pollStatus(correlationId) {
  const props = _props();
  const repo = props.getProperty('REPO');
  const pat = props.getProperty('GITHUB_PAT');
  // per_page=100 is GitHub's hard cap. We need this many because
  // `event=repository_dispatch` filter cannot also filter by event_type;
  // other dispatch types or bursty publishes could push our run off the
  // first page with the default per_page=30.
  const r = UrlFetchApp.fetch(
    `https://api.github.com/repos/${repo}/actions/runs?event=repository_dispatch&per_page=100`,
    { headers: _ghHeaders(pat), muteHttpExceptions: true }
  );
  if (r.getResponseCode() !== 200) {
    return { found: false, error: `GitHub API ตอบ HTTP ${r.getResponseCode()}` };
  }
  const runs = JSON.parse(r.getContentText()).workflow_runs || [];
  // IMPORTANT: filter on display_title (which reflects `run-name:`), NOT name
  // (which is the workflow's static `name:` field).
  const match = runs.find(run =>
    run.display_title && run.display_title.indexOf(correlationId) >= 0
  );
  if (!match) return { found: false };
  return {
    found: true,
    runId: match.id,
    status: match.status,            // queued | in_progress | completed
    conclusion: match.conclusion,    // null until status==completed
    htmlUrl: match.html_url,
  };
}

function fetchResult(runId, kind) {
  // kind is 'sync-result' or 'sync-errors'
  const fileName = kind === 'sync-errors' ? 'errors.json' : 'result.json';
  return _downloadArtifactJson(runId, kind, fileName);
}

// ---------- internals ----------

function _props() { return PropertiesService.getScriptProperties(); }

function _ghHeaders(pat) {
  return { Authorization: `Bearer ${pat}`, Accept: 'application/vnd.github+json' };
}

function _dispatchSync(dryRun) {
  const props = _props();
  const pat = props.getProperty('GITHUB_PAT');
  const repo = props.getProperty('REPO');
  if (!pat || !repo) {
    return { ok: false, error: 'Script Properties ไม่ครบ (GITHUB_PAT, REPO)' };
  }
  // NOTE: sheet_id is NOT sent in client_payload. The workflow reads the
  // sheet ID from a GitHub repository Variable (KMUTT_TRENDS_SHEET_ID)
  // instead. This is defense-in-depth: even if the PAT leaks, an
  // attacker cannot redirect the workflow to a sheet they control.
  // SHEET_ID is still set as a Script Property for the Help menu / future
  // use, but the dispatch path does not depend on it.
  const correlationId = Utilities.getUuid().replace(/-/g, '').slice(0, 16);
  const userEmail = Session.getActiveUser().getEmail() || 'unknown';
  const resp = UrlFetchApp.fetch(
    `https://api.github.com/repos/${repo}/dispatches`,
    {
      method: 'post',
      headers: _ghHeaders(pat),
      contentType: 'application/json',
      payload: JSON.stringify({
        event_type: 'sync-sheets',
        client_payload: {
          user_email: userEmail,
          dry_run: dryRun,
          correlation_id: correlationId,
        },
      }),
      muteHttpExceptions: true,
    }
  );
  if (resp.getResponseCode() !== 204) {
    return { ok: false, error: `dispatch failed (HTTP ${resp.getResponseCode()}): ${resp.getContentText()}` };
  }
  return { ok: true, correlationId };
}

function _downloadArtifactJson(runId, artifactName, fileInsideZip) {
  const props = _props();
  const repo = props.getProperty('REPO');
  const pat = props.getProperty('GITHUB_PAT');
  const listResp = UrlFetchApp.fetch(
    `https://api.github.com/repos/${repo}/actions/runs/${runId}/artifacts`,
    { headers: _ghHeaders(pat), muteHttpExceptions: true }
  );
  if (listResp.getResponseCode() !== 200) return null;
  const artifact = (JSON.parse(listResp.getContentText()).artifacts || [])
    .find(a => a.name === artifactName);
  if (!artifact) return null;
  const zipResp = UrlFetchApp.fetch(artifact.archive_download_url, {
    headers: _ghHeaders(pat), followRedirects: true,
  });
  const blobs = Utilities.unzip(zipResp.getBlob().setContentType('application/zip'));
  const file = blobs.find(b => b.getName() === fileInsideZip);
  return file ? JSON.parse(file.getDataAsString('UTF-8')) : null;
}
```

- [ ] **Step 2: Write `apps_script/PublishModal.html` (client side)**

```html
<!-- apps_script/PublishModal.html -->
<!DOCTYPE html>
<html>
<head>
<base target="_top">
<style>
  body { font-family: -apple-system, "Segoe UI", Roboto, sans-serif; padding: 16px; color: #0f172a; }
  #status { font-size: 14px; margin-bottom: 12px; min-height: 24px; }
  ul { margin: 4px 0 12px 18px; padding: 0; }
  li { margin-bottom: 4px; line-height: 1.4; }
  .err-tab { font-weight: 600; margin-top: 12px; }
  .err-field { color: #6b7280; font-family: monospace; font-size: 12px; }
  .footer { margin-top: 16px; font-size: 12px; color: #475569; }
  .footer a { color: #1e6091; }
  .spinner { display: inline-block; width: 12px; height: 12px;
             border: 2px solid #cbd5e1; border-top-color: #1e6091;
             border-radius: 50%; animation: spin 1s linear infinite;
             vertical-align: middle; margin-right: 6px; }
  @keyframes spin { to { transform: rotate(360deg); } }
</style>
</head>
<body>
  <div id="status"><span class="spinner"></span> ส่งคำสั่งไป GitHub แล้ว — กำลังรอเริ่มงาน…</div>
  <div id="result"></div>
  <div class="footer">
    <a href="https://github.com/<?= repo ?>/actions" target="_blank">ดูสถานะ build บน GitHub →</a>
  </div>
<script>
  const CORRELATION_ID = <?= JSON.stringify(correlationId) ?>;
  const DRY_RUN = <?= dryRun ? 'true' : 'false' ?>;
  const POLL_INTERVAL_MS = 5000;
  const POLL_TIMEOUT_MS = 5 * 60 * 1000;
  const startedAt = Date.now();

  function setStatus(html) {
    document.getElementById('status').innerHTML = html;
  }
  function renderResultPanel(html) {
    document.getElementById('result').innerHTML = html;
    document.getElementById('status').innerHTML = '';
  }

  function poll() {
    if (Date.now() - startedAt > POLL_TIMEOUT_MS) {
      setStatus('⏳ ใช้เวลานานกว่าปกติ — ดูสถานะที่ลิงก์ด้านล่าง');
      return;
    }
    google.script.run
      .withSuccessHandler(handlePollResult)
      .withFailureHandler(handlePollError)
      .pollStatus(CORRELATION_ID);
  }

  function handlePollError(err) {
    setStatus('⚠️ poll error: ' + (err && err.message) + ' — กำลังลองใหม่');
    setTimeout(poll, POLL_INTERVAL_MS);
  }

  function handlePollResult(res) {
    if (!res.found) {
      setStatus('<span class="spinner"></span> รอ GitHub Action เริ่มงาน…');
      setTimeout(poll, POLL_INTERVAL_MS);
      return;
    }
    if (res.status !== 'completed') {
      setStatus('<span class="spinner"></span> กำลังรันบน GitHub (' + res.status + ')…');
      setTimeout(poll, POLL_INTERVAL_MS);
      return;
    }
    if (res.conclusion === 'success') {
      google.script.run
        .withSuccessHandler(d => renderSuccess(d, res.htmlUrl))
        .fetchResult(res.runId, 'sync-result');
    } else {
      google.script.run
        .withSuccessHandler(e => renderError(e, res.htmlUrl))
        .fetchResult(res.runId, 'sync-errors');
    }
  }

  function renderSuccess(data, runUrl) {
    if (!data || !data.changed_files || data.changed_files.length === 0) {
      renderResultPanel('<p>✅ ' + (DRY_RUN ? 'ไม่มีอะไรจะเปลี่ยน — Sheets ตรงกับ dashboard แล้ว'
                                            : 'Publish สำเร็จ (แต่ไม่มีไฟล์ที่ต้องอัปเดต)') + '</p>');
      return;
    }
    const items = data.changed_files.map(f => '<li>' + f + '</li>').join('');
    const verb = DRY_RUN ? 'จะมีการเปลี่ยนแปลง' : '✅ Publish สำเร็จ — เปลี่ยน';
    renderResultPanel('<p>' + verb + ' ' + data.changed_files.length + ' ไฟล์:</p><ul>' + items + '</ul>');
  }

  function renderError(errors, runUrl) {
    if (!errors || errors.length === 0) {
      renderResultPanel('<p>❌ Publish ไม่สำเร็จ (ไม่ทราบสาเหตุ) — ดูรายละเอียดบน GitHub</p>');
      return;
    }
    const grouped = {};
    errors.forEach(e => {
      const k = e.tab || '(ทั่วไป)';
      (grouped[k] = grouped[k] || []).push(e);
    });
    let html = '<p>❌ พบข้อผิดพลาด ' + errors.length + ' จุด:</p>';
    Object.entries(grouped).forEach(([tab, errs]) => {
      html += '<div class="err-tab">▸ ' + tab + '</div><ul>';
      errs.forEach(e => {
        const f = e.field ? '<span class="err-field">' + e.field + ':</span> ' : '';
        html += '<li>' + f + (e.message_th || e.message_en) + '</li>';
      });
      html += '</ul>';
    });
    html += '<p>กรุณาแก้ไขใน Sheets แล้วลอง Publish ใหม่</p>';
    renderResultPanel(html);
  }

  setTimeout(poll, 2000); // wait briefly so the run can register
</script>
</body>
</html>
```

- [ ] **Step 3: Lint-check by reading through carefully**

Verify all Script Properties (`GITHUB_PAT`, `SHEET_ID`, `REPO`, `HELP_URL`)
are documented in the setup runbook (Task 7.1). Verify the `display_title`
filter (not `name`). Verify the modal opens BEFORE polling starts so the
"View on GitHub" fallback link is always available.

- [ ] **Step 4: Commit**

```bash
git add apps_script/Code.gs apps_script/PublishModal.html
git commit -m "feat(sheets): async Apps Script publish UX

Server side opens the modal immediately after dispatching the GitHub
workflow; the modal polls client-side via google.script.run, avoiding
the Apps Script 6-minute execution cap and providing live progress.
Polls filter workflow runs by display_title (which reflects run-name),
not name."
```

---

## Phase 5: Bootstrap script + manual integration test

*(Schema migration was completed in Phase 0.5 — see top of plan.)*

### Task 5.1: bootstrap_sheets.py

**Files:**
- Create: `scripts/bootstrap_sheets.py`
- Create: `tests/test_bootstrap.py` (basic API-call assertions; full integration is manual)

This script is run **once** by the developer to create the initial Sheets workbook from the existing JSON files. It writes data, sets data validation, conditional formatting, cell protection, and tab colours via Sheets API batch updates.

- [ ] **Step 1: Write the script**

```python
# scripts/bootstrap_sheets.py
"""One-shot: populate a fresh Google Sheets workbook from web/src/data/*.json.

Usage:
  python scripts/bootstrap_sheets.py \\
    --sheet-id <id> --credentials <path> --dev-email <your-google-email>

Pre-conditions:
  - A blank Google Sheets workbook exists.
  - The service account at <path> has Editor access to the sheet
    (Viewer is NOT enough — script creates/deletes tabs and runs batchUpdate).
  - <your-google-email> is the human account that should be allowed to edit
    protected ranges. Without this, you will be blocked from editing
    locked cells alongside the data collector.
"""
import argparse
import json
import sys
from pathlib import Path

import gspread
from google.oauth2.service_account import Credentials

SECTION_PREFIX = {
    "education": "EDU",
    "personnel": "PER",
    "research": "RES",
    "finance": "FIN",
}
SECTION_COLOR = {
    "education": {"red": 0.85, "green": 0.95, "blue": 0.85},
    "personnel": {"red": 1.0, "green": 0.95, "blue": 0.80},
    "research": {"red": 0.90, "green": 0.92, "blue": 1.0},
    "finance": {"red": 0.95, "green": 0.90, "blue": 0.95},
}


def load_charts(data_dir: Path) -> list[dict]:
    return [json.loads(p.read_text(encoding="utf-8")) for p in sorted(data_dir.glob("*.json"))]


def tab_name(chart: dict) -> str:
    return f"{SECTION_PREFIX[chart['section']]}-{chart['id']}"


def build_chart_tab_values(chart: dict) -> list[list]:
    """Return 2D array matching the per-chart-tab schema in the spec."""
    rows = [
        ["Chart ID", chart["id"]],
        ["", ""],
        ["TITLE (TH)", chart["title"]["th"]],
        ["TITLE (EN)", chart["title"]["en"]],
        ["SUBTITLE (TH)", chart["subtitle"]["th"]],
        ["SUBTITLE (EN)", chart["subtitle"]["en"]],
        ["", ""],
        ["METHODOLOGY (TH)", chart["methodology"]["th"]],
        ["METHODOLOGY (EN)", chart["methodology"]["en"]],
        ["SOURCE (TH)", chart["source"]["th"]],
        ["SOURCE (EN)", chart["source"]["en"]],
        ["", ""],
        ["— data table ↓ —"],
        ["series_key →"] + [s["key"] for s in chart["series"]],
        ["Name TH"] + [s["name"]["th"] for s in chart["series"]],
        ["Name EN"] + [s["name"]["en"] for s in chart["series"]],
        ["Year (พ.ศ.)"],
    ]
    for i, year in enumerate(chart["categories_buddhist"]):
        rows.append([year] + [s["values"][i] if s["values"][i] is not None else "" for s in chart["series"]])
    return rows


# Pinned KPI-series-key mapping. Source of truth is the kpiSeriesKey
# props in web/src/App.tsx — if App.tsx changes, update this table AND
# re-run bootstrap. The sync-time validator enforces that the chart tab
# actually contains the named series, so KPI-card crashes are caught.
KPI_SERIES_KEY = {
    "students-new": "total", "students-all": "total",
    "programs": "thai", "graduates": "total",
    "employment-bachelor": "employed", "employment-graduate": "employed",
    "staff-total": "total", "faculty-degree": "total",
    "staff-academic-support": "academic",
    "research-funding": "total", "research-funding-3yr": "total",
    "research-per-staff": "per_active_researcher",
    "research-per-staff-3yr": "per_active_researcher",
    "research-per-academic-3yr": "per_academic",
    "publications": "total", "publications-3yr": "total",
    "publications-per-academic": "per_academic",
    "patents": "patent_filed",
    "income-expense": "revenue", "income-expense-3yr": "revenue",
}


def build_style_charts_rows(charts: list[dict]) -> list[list]:
    rows = [["chart_id", "section", "chart_type", "kpi_series_key"]]
    for c in charts:
        rows.append([c["id"], c["section"], c["chart_type"],
                     KPI_SERIES_KEY.get(c["id"], "")])
    return rows


def build_style_series_rows(charts: list[dict]) -> list[list]:
    rows = [["chart_id", "series_key", "color", "flags"]]
    for c in charts:
        for s in c["series"]:
            flags = ",".join(
                f for f in ("emphasis", "exclude_from_stack", "is_cumulative") if s.get(f)
            )
            rows.append([c["id"], s["key"], s["color"], flags])
    return rows


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--sheet-id", required=True)
    ap.add_argument("--credentials", required=True)
    ap.add_argument("--dev-email", required=True,
                    help="Google account email allowed to edit locked ranges. "
                         "Data collector edits are blocked by editors-allowlist.")
    ap.add_argument("--data-dir", default="web/src/data")
    args = ap.parse_args()

    scopes = ["https://www.googleapis.com/auth/spreadsheets"]
    creds = Credentials.from_service_account_file(args.credentials, scopes=scopes)
    gc = gspread.authorize(creds)
    sh = gc.open_by_key(args.sheet_id)

    charts = load_charts(Path(args.data_dir))
    print(f"Loaded {len(charts)} charts")

    # Self-check: every KPI_SERIES_KEY mapping must point to a series that
    # actually exists in that chart's JSON. A typo here would create a
    # workbook that fails validation from Day 1 with no easy fix path
    # (STYLE-charts is protected → only dev can correct via re-bootstrap).
    kpi_problems = []
    for c in charts:
        kpi = KPI_SERIES_KEY.get(c["id"])
        if not kpi:
            kpi_problems.append(f"missing KPI_SERIES_KEY for chart '{c['id']}'")
            continue
        if not any(s["key"] == kpi for s in c["series"]):
            kpi_problems.append(
                f"KPI_SERIES_KEY['{c['id']}']='{kpi}' but no series with that "
                f"key in {c['id']}.json (available: "
                f"{[s['key'] for s in c['series']]})"
            )
    if kpi_problems:
        print("BUG in KPI_SERIES_KEY mapping (must match web/src/App.tsx):")
        for p in kpi_problems:
            print(f"  - {p}")
        sys.exit(1)

    # 1) Reset to a clean slate. Bootstrap is idempotent: re-runs are
    # supported. We:
    #   a) Drop all existing protected ranges and conditional format rules
    #      (otherwise leftover protections from a prior run will block
    #      subsequent writes, even from the service account).
    #   b) Create a NEW placeholder with a unique timestamp suffix FIRST,
    #      BEFORE deleting anything else. The Sheets API rejects deleting
    #      the workbook's last remaining sheet — if a previous bootstrap
    #      failed in the middle of step (d) below and the only surviving
    #      tab is a leftover `_bootstrap_placeholder_*`, deleting it before
    #      adding a replacement would hit that constraint and make rerun
    #      recovery impossible. Adding first guarantees at least 2 sheets
    #      exist before any delete fires.
    #   c) Delete every tab except the new placeholder (this also sweeps
    #      any leftover `_bootstrap_placeholder_*` tabs from prior runs).
    #   d) At the end of the script, delete the placeholder.
    _reset_workbook_state(sh)
    import time
    placeholder_name = f"_bootstrap_placeholder_{int(time.time())}"
    placeholder = sh.add_worksheet(title=placeholder_name, rows=1, cols=1)
    for ws in list(sh.worksheets()):
        if ws.id != placeholder.id:
            sh.del_worksheet(ws)

    # 2) Create STYLE-charts tab
    ws = sh.add_worksheet(title="🎨 STYLE-charts", rows=50, cols=5)
    ws.update("A1", build_style_charts_rows(charts), value_input_option="RAW")

    # 3) Create STYLE-series tab
    rows = build_style_series_rows(charts)
    ws = sh.add_worksheet(title="🎨 STYLE-series", rows=max(50, len(rows) + 5), cols=5)
    ws.update("A1", rows, value_input_option="RAW")

    # 4) Create one tab per chart, capturing each tab's gid so the INDEX
    #    HYPERLINKs can target the right tab.
    chart_gid: dict[str, int] = {}
    for c in charts:
        values = build_chart_tab_values(c)
        ws = sh.add_worksheet(title=tab_name(c), rows=max(60, len(values) + 10),
                              cols=1 + len(c["series"]) + 1)
        ws.update("A1", values, value_input_option="RAW")
        ws.update_tab_color(SECTION_COLOR[c["section"]])
        ws.freeze(rows=17)
        chart_gid[c["id"]] = ws.id  # gspread Worksheet.id == Sheets gid

    # 5) Create INDEX tab with HYPERLINKs that target each tab's real gid.
    # Spec round 4: # years / # series columns dropped — they would go stale
    # because INDEX is bootstrap-static (service account is Viewer at runtime).
    index_rows = [["Chart ID", "Section", "Title TH", "Title EN", "Jump"]]
    for c in charts:
        gid = chart_gid[c["id"]]
        index_rows.append([
            c["id"], c["section"], c["title"]["th"], c["title"]["en"],
            f'=HYPERLINK("#gid={gid}","→ {tab_name(c)}")',
        ])
    ws = sh.add_worksheet(title="📋 INDEX", rows=30, cols=6)
    ws.update("A1", index_rows, value_input_option="USER_ENTERED")
    # Move INDEX to first position
    sh.reorder_worksheets([ws] + [w for w in sh.worksheets() if w.id != ws.id])

    # 6) Delete the placeholder (safe — 23 other tabs exist by now)
    sh.del_worksheet(placeholder)

    # 7) Apply protections + data validation + conditional formatting in
    #    one batchUpdate per chart tab. Hard protection (editors-allowlist)
    #    not warning-only — the data collector is a workbook Editor, so
    #    warningOnly wouldn't block them.
    #
    # Editors list MUST include the service account email alongside the
    # dev. Sheets API enforces editor membership for protected-range
    # deletion: a general workbook Editor cannot remove a protected range
    # whose `editors.users` list doesn't include them. Without the service
    # account on this list, _reset_workbook_state() on a re-run would
    # error out trying to drop these ranges. The service account's email
    # lives in the credentials JSON as `client_email`.
    import json as _json
    sa_email = _json.loads(Path(args.credentials).read_text())["client_email"]
    allowed_editors = [args.dev_email, sa_email]
    _apply_chart_tab_guardrails(sh, charts, chart_gid, allowed_editors)
    _apply_style_tab_protections(sh, allowed_editors)

    print(f"Bootstrap complete: {len(charts) + 3} tabs created (INDEX + 2 STYLE + {len(charts)} charts).")
    print("Protections, data validation, and conditional formatting applied programmatically.")
    print("Re-runs are supported: this script drops existing protected ranges + conditional")
    print("formats before rebuilding, so you can re-run safely as long as the service")
    print("account still has Editor on the sheet.")


REQUIRED_TEXT_ROWS = (2, 3, 4, 5, 7, 8, 9, 10)  # 0-indexed for rows 3-6 + 8-11
PINK = {"red": 0.99, "green": 0.85, "blue": 0.85}


def _reset_workbook_state(sh):
    """Drop all protected ranges and conditional format rules in the
    workbook. Run before deleting tabs so we don't fight stale rules
    on a re-run.

    Uses fetch_sheet_metadata() to enumerate existing rules. We delete
    conditional format rules from highest index to 0 to avoid the index
    shifting underneath us.
    """
    meta = sh.fetch_sheet_metadata()
    requests = []
    for sheet in meta.get("sheets", []):
        gid = sheet["properties"]["sheetId"]
        for pr in sheet.get("protectedRanges", []):
            requests.append({"deleteProtectedRange":
                {"protectedRangeId": pr["protectedRangeId"]}})
        rules = sheet.get("conditionalFormats", [])
        for i in range(len(rules) - 1, -1, -1):
            requests.append({"deleteConditionalFormatRule":
                {"sheetId": gid, "index": i}})
    if requests:
        sh.batch_update({"requests": requests})


def _apply_chart_tab_guardrails(sh, charts, chart_gid, allowed_editors):
    """For every chart tab:
    - HARD protect rows 1 (chart_id), 13 (table divider), 14 (series_key)
      with editors-allowlist (warningOnly would NOT block workbook editors)
    - Data validation: year column A from row 18 = integer 2500-2700
    - Data validation: value columns from row 18 = numeric OR empty
    - Conditional format: highlight duplicate year rows
    - Conditional format: highlight blank required-text cells (rows 3-6, 8-11)
    """
    requests = []
    for c in charts:
        gid = chart_gid[c["id"]]
        num_series = len(c["series"])

        # Hard-protect rows 1, 13, 14 (0-indexed: 0, 12, 13) with allowlist
        for r in (0, 12, 13):
            requests.append({
                "addProtectedRange": {
                    "protectedRange": {
                        "range": {"sheetId": gid, "startRowIndex": r, "endRowIndex": r + 1},
                        "description": "Locked — only dev can edit",
                        "editors": {"users": allowed_editors},
                        # NOTE: omit warningOnly → defaults to hard protection
                    }
                }
            })

        # Year column (A) from row 18 onward.
        # Accept either a plain BE year in [2500, 2700] OR a NNNN-NNNN BE
        # range with both endpoints in [2500, 2700]. The *-3yr charts and
        # patents row 1 use range format ("2542-2544" etc), so a naive
        # NUMBER_BETWEEN rule would reject 6 of the 20 charts on Day 1.
        #
        # Tightening (round 7):
        # - Scalar branch adds `A18=INT(A18)` so fractional values like
        #   `2500.5` are rejected at entry time (Layer 2 catches them too
        #   via re.match(r"\d{4}"), but Layer 1 is supposed to be the
        #   entry-time gate).
        # - Range branch adds a start<=end ordering check so reversed
        #   ranges like `2568-2566` cannot be saved (Layer 2 also rejects,
        #   but Sheets data validation is the first line of defence).
        year_formula = (
            '=OR('
              'AND(ISNUMBER(A18),A18=INT(A18),A18>=2500,A18<=2700),'
              'AND('
                'REGEXMATCH(A18&"","^\\d{4}-\\d{4}$"),'
                'VALUE(LEFT(A18,4))>=2500,VALUE(LEFT(A18,4))<=2700,'
                'VALUE(RIGHT(A18,4))>=2500,VALUE(RIGHT(A18,4))<=2700,'
                'VALUE(LEFT(A18,4))<=VALUE(RIGHT(A18,4))'
              ')'
            ')'
        )
        requests.append({
            "setDataValidation": {
                "range": {"sheetId": gid, "startRowIndex": 17, "startColumnIndex": 0, "endColumnIndex": 1},
                "rule": {
                    "condition": {"type": "CUSTOM_FORMULA",
                                  "values": [{"userEnteredValue": year_formula}]},
                    "inputMessage": "ปีต้องเป็น 4 หลัก (2500-2700) หรือช่วงปี NNNN-NNNN",
                    "strict": True,
                }
            }
        })
        # Value columns from row 18 — numeric OR blank
        requests.append({
            "setDataValidation": {
                "range": {"sheetId": gid, "startRowIndex": 17,
                          "startColumnIndex": 1, "endColumnIndex": 1 + num_series},
                "rule": {
                    "condition": {"type": "CUSTOM_FORMULA",
                                  "values": [{"userEnteredValue": "=OR(ISNUMBER(B18),B18=\"\")"}]},
                    "inputMessage": "ค่าต้องเป็นตัวเลขหรือว่างเปล่า",
                    "strict": True,
                }
            }
        })
        # Conditional format: duplicate year rows.
        # IMPORTANT: the AND(A18<>"") guard is necessary. A naked
        # COUNTIF(A:A,A18)>1 treats blank cells as duplicates of each
        # other, so every empty future-year row would highlight red the
        # instant bootstrap finishes.
        requests.append({
            "addConditionalFormatRule": {
                "rule": {
                    "ranges": [{"sheetId": gid, "startRowIndex": 17, "startColumnIndex": 0, "endColumnIndex": 1}],
                    "booleanRule": {
                        "condition": {"type": "CUSTOM_FORMULA",
                                      "values": [{"userEnteredValue": "=AND(A18<>\"\",COUNTIF(A:A,A18)>1)"}]},
                        "format": {"backgroundColor": PINK},
                    }
                },
                "index": 0,
            }
        })
        # Conditional format: required-text fields must not be blank.
        # Each row 3-6 and 8-11 needs B<row> non-empty.
        for r0 in REQUIRED_TEXT_ROWS:
            requests.append({
                "addConditionalFormatRule": {
                    "rule": {
                        "ranges": [{"sheetId": gid,
                                    "startRowIndex": r0, "endRowIndex": r0 + 1,
                                    "startColumnIndex": 1, "endColumnIndex": 2}],
                        "booleanRule": {
                            "condition": {"type": "BLANK"},
                            "format": {"backgroundColor": PINK},
                        }
                    },
                    "index": 0,
                }
            })
    if requests:
        sh.batch_update({"requests": requests})


def _apply_style_tab_protections(sh, allowed_editors):
    """Hard-protect the entire STYLE-charts and STYLE-series tabs."""
    requests = []
    for ws in sh.worksheets():
        if ws.title.startswith("🎨 STYLE"):
            requests.append({
                "addProtectedRange": {
                    "protectedRange": {
                        "range": {"sheetId": ws.id},
                        "description": "STYLE tab — dev-only",
                        "editors": {"users": allowed_editors},
                    }
                }
            })
    if requests:
        sh.batch_update({"requests": requests})


if __name__ == "__main__":
    main()
```

> **Note for plan executor:** This bootstrap script handles everything
> programmatically — populating data, creating tabs, tab colours, freezing
> rows, AND all guardrails (hard-protected ranges with editors-allowlist,
> data validation rules for year + value cells, conditional formatting for
> duplicate years and required-text fields). Re-running is supported via
> the `_reset_workbook_state` helper. No manual Sheets-UI follow-up is
> required after a successful run.

- [ ] **Step 2: Write minimal smoke test**

```python
# tests/test_bootstrap.py
from scripts.bootstrap_sheets import build_chart_tab_values, build_style_charts_rows, build_style_series_rows

SAMPLE = {
    "id": "students-all", "section": "education", "chart_type": "line",
    "title": {"th": "x", "en": "y"}, "subtitle": {"th": "x", "en": "y"},
    "categories_buddhist": ["2566", "2567"],
    "series": [
        {"key": "bachelor", "name": {"th": "ป.ตรี", "en": "B"}, "color": "#000000", "values": [1, 2]},
        {"key": "total", "name": {"th": "รวม", "en": "T"}, "color": "#fff", "values": [3, 4], "emphasis": True},
    ],
    "methodology": {"th": "x", "en": "y"},
    "source": {"th": "x", "en": "y"},
}

def test_chart_tab_values_metadata_layout():
    rows = build_chart_tab_values(SAMPLE)
    assert rows[0] == ["Chart ID", "students-all"]
    assert rows[2] == ["TITLE (TH)", "x"]
    assert rows[13] == ["series_key →", "bachelor", "total"]
    assert rows[17] == ["2566", 1, 3]
    assert rows[18] == ["2567", 2, 4]

def test_style_series_emits_flags_csv():
    rows = build_style_series_rows([SAMPLE])
    total_row = next(r for r in rows if r[1] == "total")
    assert total_row[3] == "emphasis"
```

- [ ] **Step 3: Run tests**

```bash
pytest tests/test_bootstrap.py -v
```

- [ ] **Step 4: Commit**

```bash
git add scripts/bootstrap_sheets.py tests/test_bootstrap.py
git commit -m "feat(sync): one-shot bootstrap script for fresh Sheets workbook"
```

---

## Phase 6: Documentation

### Task 6.1: Thai data-collector guide

**Files:**
- Create: `docs/data-collector-guide-th.md`

- [ ] **Step 1: Write the guide**

Content checklist (full Thai, screenshots can be added later by the developer):

1. ภาพรวม: Sheets คืออะไร, Publish คืออะไร, dashboard อยู่ที่ไหน
2. การเข้าถึง: เปิด Sheets ยังไง, login บัญชีไหน
3. โครงสร้าง workbook: 📋 INDEX, 🎨 STYLE tabs (อย่าแก้), tab ของแต่ละกราฟ
4. งานปกติ: อัพเดตปีใหม่ (step-by-step)
5. งานบ่อยพอควร: แก้ตัวเลขย้อนหลัง, แก้ methodology หรือ source
6. เมนู `📤 Publish to Dashboard`:
   - `Check what will change (dry-run)` — เช็คก่อน
   - `Publish all changes` — กดเมื่อมั่นใจ
7. การอ่าน error modal: ตัวอย่างทุก error message + วิธีแก้
8. สิ่งที่แก้ไม่ได้ (cells สีเทา, STYLE tabs) — ติดต่อ dev
9. FAQ: ลบ tab ผิด, อยากเพิ่มกราฟใหม่, dashboard ไม่อัพเดตหลัง publish
10. ติดต่อใคร: ใส่ชื่อ dev + ช่องทางติดต่อ

- [ ] **Step 2: Commit**

```bash
git add docs/data-collector-guide-th.md
git commit -m "docs: Thai guide for data collector"
```

### Task 6.2: Architecture doc + README update + deprecation

**Files:**
- Create: `docs/architecture.md`
- Modify: `README.md`
- Modify: `build_chart_json.py`

- [ ] **Step 1: Write architecture.md**

Cover: data-flow diagram, key files (link to spec), deploy chain, debugging tips, credential rotation, where to look when things break.

- [ ] **Step 2: Update README.md — replace the "Updating data from a new PPT" section**

Find the heading `## Updating data from a new PPT` in `README.md` and
replace that entire section (down to the next `##` heading) with a new
section titled `## Updating data` that describes the Sheets flow at a high
level, links to the data-collector guide for non-devs, and the architecture
doc for devs.

- [ ] **Step 3: Add deprecation header to build_chart_json.py**

Prepend:

```python
"""DEPRECATED — Replaced by scripts/sync_from_sheets.py.

Kept temporarily for reference / rollback. Will be removed after one
successful annual cycle of the Sheets-based flow.

See docs/superpowers/specs/2026-05-25-google-sheets-data-source-design.md
"""
```

- [ ] **Step 4: Commit**

```bash
git add docs/architecture.md README.md build_chart_json.py
git commit -m "docs: architecture doc, README update, deprecate PPTX script"
```

---

## Phase 7: Initial migration + e2e test (manual)

### Task 7.1: Hand-off checklist

**Files:**
- Create: `docs/runbook-initial-setup.md`

This task is documentation of the manual one-time steps from the spec's "Initial Migration" section. The plan does not automate the GCP / GitHub PAT setup because they involve out-of-band approvals.

- [ ] **Step 1: Write runbook**

Content (numbered to match spec Initial Migration steps 0–5):

1. **Step 0 — done in Phase 0.5 of this plan.** Verify the schema-cleanup
   commit is on `main` before proceeding (`grep '"slide"' web/src/data/*.json`
   should return nothing).
2. Create empty Google Sheets workbook → note Sheet ID.
3. Create GCP project + service account + JSON key. Share the Sheets file
   with the service account email — **Editor role initially** (bootstrap
   needs write access for batchUpdate). Also share with the data collector
   (Editor) and any dev maintainers (Editor).
4. Generate fine-grained GitHub PAT — **Contents: write, Actions: read**.
   (Contents: read is NOT enough; `repository_dispatch` requires write.
   Actions: read is needed for the Apps Script to poll workflow runs.)
5. Add GitHub Secret `GOOGLE_SERVICE_ACCOUNT_JSON`. Also add GitHub
   **Variable** (not secret — it's not sensitive) `KMUTT_TRENDS_SHEET_ID`
   with the same Sheet ID from step 2. The workflow reads from this
   Variable instead of trusting `repository_dispatch` payload — defense
   against PAT leak.
6. Run `python scripts/bootstrap_sheets.py --sheet-id <id> --credentials <path> --dev-email <your-google-account>`.
   The `--dev-email` is the Google account that should be allowed to edit
   protected ranges (rows 1, 13, 14 and the STYLE tabs). Without this you
   yourself will be blocked from editing those cells.
7. **(Optional hardening)** Downgrade service account from Editor to
   Viewer in Sheets share dialog.
   - Future bootstrap re-runs require: restore service account to
     Editor, then re-run `bootstrap_sheets.py`. The script is idempotent
     — it drops existing protected ranges + conditional formats, deletes
     all tabs, and rebuilds from scratch.
   - After a successful re-run, downgrade back to Viewer.
8. Open Sheet → Extensions → Apps Script → paste BOTH `apps_script/Code.gs`
   AND `apps_script/PublishModal.html`.
9. Set Script Properties: `GITHUB_PAT`, `SHEET_ID`, `REPO` (e.g. `org/repo`),
   `HELP_URL` (link to the published Thai guide).
10. Reload Sheet, verify `📤 Publish to Dashboard` menu appears.
11. Run "Check what will change (dry-run)" — expect "no changes" (data
    matches JSON because of Phase 0.5).
12. Edit one cell, dry-run again — expect 1 file diff.
13. Click "Publish all changes" — verify: sync job commits, deploy job
    runs (inline in same workflow run), modal reports success only AFTER
    deploy finishes, dashboard updates on GitHub Pages.
14. **Test recovery:** click Publish a second time with no edits. Expected:
    sync no-op (no commit), deploy job still runs, dashboard re-publishes
    identically, modal reports success. This proves the "repo updated but
    Pages stuck" scenario is recoverable.
15. **Test hard protection:** as a non-dev Google account (or in an
    incognito Sheets session as the data collector), try to edit a cell in
    row 14 of any chart tab. Expected: "You don't have permission" error
    (NOT a warning with an OK button).
16. Read through the credential-trust-boundary section of the spec and
    confirm you accept it.
17. Hand the Sheet URL and the Thai guide URL to the data collector.

- [ ] **Step 2: Commit**

```bash
git add docs/runbook-initial-setup.md
git commit -m "docs: runbook for initial Sheets bootstrap"
```

---

## Phase 8: Final sweep

### Task 8.1: Full test run + lint

- [ ] **Step 1: Run all tests**

```bash
pytest tests/ -v
```
Expected: all green.

- [ ] **Step 2: Build the React app**

```bash
cd web && npm run build
```
Expected: build succeeds.

- [ ] **Step 3: If any failures, fix and re-test before declaring done.**

- [ ] **Step 4: Sanity-check workflow YAML**

```bash
python -c "import yaml; print(yaml.safe_load(open('.github/workflows/sync-from-sheets.yml')))"
```

- [ ] **Step 5: Commit anything outstanding, push, and tag a milestone**

```bash
git status   # should be clean
git push
```

---

## Notes on what's deferred

- **`build_chart_json.py` deletion** is intentionally deferred to after one
  successful production cycle of the new flow.
- **Apps Script unit tests** are not part of this plan (Apps Script test
  infrastructure is awkward). The integration test in Task 7.1 (steps 10-12)
  covers the critical paths: dispatch, polling, success modal, error modal.
- **Adding new charts** (vs editing existing ones) remains a dev task per
  the spec's non-goals — covered by future plans if/when needed.
- **GitHub App / Cloud Function alternative for the PAT.** Documented in
  the spec's Authentication section as a future migration path if the
  trust assumption ("single internal data collector, public data") ever
  changes. Not implemented now.
- **INDEX dynamic refresh on sync** was deliberately dropped — the service
  account stays Viewer-only. INDEX is bootstrap-static; dev re-runs
  bootstrap if a chart title changes and they want INDEX to reflect it.
