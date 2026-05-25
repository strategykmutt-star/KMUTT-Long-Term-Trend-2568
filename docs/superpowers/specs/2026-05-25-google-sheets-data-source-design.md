# Google Sheets as Data Source for KMUTT Trends Dashboard

**Date:** 2026-05-25
**Status:** Design (pending implementation plan)
**Owner:** Strategy Office, KMUTT

## Problem

The KMUTT Long-Term Trends dashboard is generated from an annual PowerPoint
deck via `build_chart_json.py`, then deployed to GitHub Pages. Today every
update requires a developer: drop a new PPTX, run a Python script, inspect
the JSON diff, `git commit`, `git push`.

The data collector — one person in the Strategy Office who already produces
the annual PPTX — knows Excel and Google Sheets but does not use Git or the
command line. They want to update the dashboard 1–2 times per year without
filing a ticket with engineering.

## Goal

Let the data collector update **numbers, year ranges, and all TH/EN text**
(title, subtitle, methodology, source, series names) for the existing 20
charts, end-to-end, by themselves — using only Google Sheets and a single
"Publish" button.

## Non-Goals

- **Adding entirely new charts** stays a developer task. New charts are rare
  (≪ 1 per year) and require code-level decisions (chart type, section,
  colour palette, layout). Forcing those decisions into a spreadsheet schema
  would balloon complexity for a use case that may never occur.
- **Real-time / continuous sync.** Updates are deliberately batched and
  triggered. 1–2 times per year, not on every cell edit.
- **Multiple concurrent editors / approval workflow.** A single person owns
  the Sheet. If that changes, revisit.
- **Replacing the PPTX.** The annual PPTX continues as a separate deliverable;
  the dashboard simply stops depending on it as its source of truth.

## Architecture

```
┌────────────────────────┐
│   Google Sheets        │  Data collector edits here
│   1 workbook,          │
│   23 tabs:             │
│     📋 INDEX           │
│     🎨 STYLE-charts    │
│     🎨 STYLE-series    │
│     20 chart tabs      │
│                        │
│   Custom menu:         │
│     "📤 Publish to     │
│      Dashboard"        │
└───────────┬────────────┘
            │ Apps Script POSTs to GitHub API
            │ (repository_dispatch event)
            ▼
┌────────────────────────┐
│   GitHub Action:       │
│   sync-from-sheets.yml │
│   - reads Sheets via   │
│     service account    │
│   - validates          │
│   - writes JSON        │
│   - commits + pushes   │
└───────────┬────────────┘
            │ push to main triggers existing
            ▼
┌────────────────────────┐
│   deploy.yml (existing)│  Builds + deploys to GitHub Pages
└───────────┬────────────┘
            ▼
       Live dashboard
```

**What is preserved:**

- React/Vite/TypeScript app structure
- `web/src/data/*.json` schema (consumer code unchanged)
- Existing `deploy.yml` (push-to-main → GitHub Pages)
- `build_chart_json.py` (kept for reference, deprecated)
- The annual PPTX (continues as a separate deliverable)

**What is added:**

| Component | Purpose |
|---|---|
| Google Sheets workbook | Source of truth |
| Apps Script (`apps_script/Code.gs`) | Publish/dry-run menu, status modals, error reporting |
| `.github/workflows/sync-from-sheets.yml` | Reacts to repository_dispatch from Apps Script |
| `scripts/sync_from_sheets.py` | Reads Sheets → validates → writes JSON |
| `scripts/bootstrap_sheets.py` | One-time: 20 JSON files → Sheets |
| `scripts/lib/{sheets_client,parsers,validators}.py` | Shared helpers |
| `docs/data-collector-guide-th.md` | Thai-language manual for the data collector |
| GitHub Secret `GOOGLE_SERVICE_ACCOUNT_JSON` | GitHub Action → Sheets API auth |
| Apps Script property `GITHUB_PAT` | Apps Script → GitHub API auth |

## Sheet Schema

### Canonical chart list (20 charts)

This is the authoritative mapping that drives both the bootstrap script and
validation. Tab prefixes group charts visually; tab background colour mirrors
the section.

| chart_id | section | chart_type | tab name |
|---|---|---|---|
| programs | education | stacked-bar | `EDU-programs` |
| students-new | education | line | `EDU-students-new` |
| students-all | education | line | `EDU-students-all` |
| graduates | education | line | `EDU-graduates` |
| employment-bachelor | education | stacked-bar | `EDU-employment-bachelor` |
| employment-graduate | education | stacked-bar | `EDU-employment-graduate` |
| staff-total | personnel | stacked-bar | `PER-staff-total` |
| faculty-degree | personnel | stacked-bar | `PER-faculty-degree` |
| staff-academic-support | personnel | stacked-bar | `PER-staff-academic-support` |
| research-funding | research | stacked-bar | `RES-research-funding` |
| research-funding-3yr | research | stacked-bar | `RES-research-funding-3yr` |
| research-per-staff | research | line | `RES-research-per-staff` |
| research-per-staff-3yr | research | line | `RES-research-per-staff-3yr` |
| research-per-academic-3yr | research | line | `RES-research-per-academic-3yr` |
| publications | research | stacked-bar | `RES-publications` |
| publications-3yr | research | stacked-bar | `RES-publications-3yr` |
| publications-per-academic | research | clustered-bar | `RES-publications-per-academic` |
| patents | research | clustered-bar | `RES-patents` |
| income-expense | finance | line | `FIN-income-expense` |
| income-expense-3yr | finance | line | `FIN-income-expense-3yr` |

Prefix → section: `EDU-` → education, `PER-` → personnel, `RES-` → research,
`FIN-` → finance.

### Workbook layout

23 tabs in one Google Sheets workbook:

```
📋 INDEX            (auto-generated navigation, read-only)
🎨 STYLE-charts     (per-chart config: section, chart_type — locked)
🎨 STYLE-series     (per-series config: color, flags — locked)
EDU-programs        (20 chart tabs — see canonical list above)
EDU-students-new
... (18 more)
FIN-income-expense-3yr
```

### Per-chart tab schema

| Row | Col A | Col B | Col C | Col D | Col E |
|---|---|---|---|---|---|
| 1 | 🔒 Chart ID | students-all | | | |
| 2 | *(blank)* | | | | |
| 3 | 📝 TITLE (TH) | จำนวนนักศึกษาทั้งหมด | | | |
| 4 | 📝 TITLE (EN) | Total Enrolled Students | | | |
| 5 | 📝 SUBTITLE (TH) | จำแนกตามระดับการศึกษา (ภาคที่ 2) | | | |
| 6 | 📝 SUBTITLE (EN) | By education level (semester 2) | | | |
| 7 | *(blank)* | | | | |
| 8 | 📝 METHODOLOGY (TH) | ทุกปีเป็นข้อมูล ณ ภาคการศึกษาที่ 2... | | | |
| 9 | 📝 METHODOLOGY (EN) | Data as of semester 2 each year... | | | |
| 10 | 📝 SOURCE (TH) | หนังสือรายงานประจำปี... | | | |
| 11 | 📝 SOURCE (EN) | Annual reports... | | | |
| 12 | *(blank)* | | | | |
| 13 | **— data table ↓ —** | | | | |
| 14 | 🔒 series_key → | bachelor | graduate | total | |
| 15 | 📝 Name TH | ปริญญาตรี | บัณฑิตศึกษา | รวม | |
| 16 | 📝 Name EN | Bachelor's | Graduate (M+P) | Total | |
| 17 | Year (พ.ศ.) | | | | |
| 18 | 2536 | 3097 | 551 | 3648 | |
| ... | ... | ... | ... | ... | |
| 50 | 2568 | 12253 | 2472 | 14725 | |
| 51 | *[add new year here]* | | | | |

**Conventions:**

- 🔒 = locked (Sheets cell protection). Data collector cannot edit. Editing
  STYLE-tab values that are derived from STYLE is also prohibited.
- 📝 = editable, white background.
- Locked cells have a light-grey background.
- Freeze panes at row 17 so the header stays visible while scrolling years.
- Adding a year = new row at the bottom. Removing a year = delete row.

### STYLE-charts tab schema (dev-only)

| chart_id | section | chart_type |
|---|---|---|
| students-all | education | line |
| students-new | education | line |
| ... | ... | ... |

- One row per chart, exactly 20 rows.
- This tab is the authoritative source for `section` and `chart_type`.
- Entire tab is cell-protected; only developers can edit.

### STYLE-series tab schema (dev-only)

| chart_id | series_key | color | flags |
|---|---|---|---|
| students-all | bachelor | #f29400 | |
| students-all | graduate | #1e6091 | |
| students-all | total | #0f172a | emphasis |
| students-new | bachelor | #f29400 | |
| ... | ... | ... | ... |

- One row per (chart × series).
- `color` is a 7-character hex code (`#RRGGBB`).
- `flags` accepts a comma-separated subset of `emphasis`,
  `exclude_from_stack`, `is_cumulative`. Empty cell = no flags.
- Entire tab is cell-protected; only developers can edit.

### JSON output schema (consumer contract)

The sync script produces per-chart files at `web/src/data/<chart_id>.json`
matching the existing schema **with one change**: the `slide` field (a
legacy reference to PPTX slide numbers, present in every current JSON file
but **unused by the React app**) is dropped. The implementation plan must
also remove `slide: number` from `web/src/types.ts`.

Internal Python key names used by the parser, validator, and JSON writer
(must agree across all three):

```python
ChartData = {
    "id": str,                   # from chart tab cell B1
    "section": str,              # from STYLE-charts
    "chart_type": str,           # from STYLE-charts
    "title": {"th": str, "en": str},       # rows 3-4
    "subtitle": {"th": str, "en": str},    # rows 5-6
    "categories_buddhist": [str, ...],     # column A, rows 18+
    "series": [
        {
            "key": str,                    # row 14
            "name": {"th": str, "en": str},  # rows 15-16
            "color": str,                  # from STYLE-series
            "values": [float | None, ...], # rows 18+ for that series column
            "emphasis": bool,              # only present if flag set
            "exclude_from_stack": bool,    # only present if flag set
            "is_cumulative": bool,         # only present if flag set
        },
        ...
    ],
    "methodology": {"th": str, "en": str}, # rows 8-9
    "source": {"th": str, "en": str},      # rows 10-11
}
```

### INDEX tab schema (auto-generated, read-only)

| Chart ID | Section | Title TH | Title EN | # years | # series | Jump |
|---|---|---|---|---|---|---|
| students-all | education | จำนวนนักศึกษาทั้งหมด | Total Enrolled Students | 33 | 3 | → click |
| ... | ... | ... | ... | ... | ... | ... |

- Generated by the bootstrap script. Updated by `sync_from_sheets.py` whenever
  a chart tab changes.
- "Jump" column uses `HYPERLINK("#gid=…")` to jump to that chart tab.

## Publish Flow

### What the data collector sees

1. Open Sheets → INDEX tab → click the chart they want to update.
2. Edit cells (numbers, text, add a new year row).
3. Open menu `📤 Publish to Dashboard`:
   - `Check what will change (dry-run)` — runs validation + diff without
     committing
   - `Publish all changes` — commits + deploys

Modal during publish:

```
📤 กำลัง Publish ไปยัง Dashboard...
  ✅ ส่งข้อมูลไป GitHub
  ✅ ตรวจสอบความถูกต้องข้อมูล
  ⏳ สร้างไฟล์ JSON ใหม่
  ⏳ Deploy ขึ้น dashboard
คาดว่าเสร็จใน: ~1-2 นาที
[ดูสถานะ build บน GitHub]
```

Success:

```
✅ Publish สำเร็จ!
มีการเปลี่ยนแปลง 3 กราฟ:
  • students-all (เพิ่มปี 2569)
  • patents (แก้ตัวเลขปี 2568)
  • income-expense (แก้ methodology TH)
[เปิด dashboard]
```

Failure (validation or workflow error):

```
❌ Publish ไม่สำเร็จ — มีข้อผิดพลาด 3 จุด
▸ Tab "EDU-students-all"
  • บรรทัด 51: ปี "2569" ซ้ำกับบรรทัด 50
  • บรรทัด 52: คอลัมน์ "graduate" ว่าง
▸ Tab "RES-patents"
  • methodology (TH) ว่างเปล่า
กรุณาแก้ไขใน Sheets แล้วลอง Publish ใหม่
[เปิด tab ที่มีปัญหา] [ปิด]
```

### What happens behind the scenes

1. **Apps Script** (in Sheets)
   - Reads `SHEET_ID` and `GITHUB_PAT` from Script Properties (set once at
     install; not hard-coded so the same script works across forks).
   - Reads `repo` (e.g. `kmutt-strategy/trends-dashboard`) from a third
     Script Property.
   - Generates a 16-character hex `correlation_id` (UUID-derived).
   - POST `/repos/{repo}/dispatches` with body:
     ```json
     {
       "event_type": "sync-sheets",
       "client_payload": {
         "sheet_id": "...",
         "user_email": "...",
         "dry_run": false,
         "correlation_id": "abc123…"
       }
     }
     ```
   - Polls `/repos/{repo}/actions/runs?event=repository_dispatch&per_page=20`
     every 5 seconds, filtering for a run whose `name` field contains the
     `correlation_id` (set by the workflow's `run-name`).
   - When the matched run reaches `status=completed`:
     - On `conclusion=success` → download the `sync-result` artifact (a
       single JSON file) and render the appropriate modal (publish success
       summary, or dry-run diff).
     - On `conclusion=failure` → download the `sync-errors` artifact (JSON
       list of validation errors) and render the error modal.
   - Times out after 5 minutes; modal shows a "View run on GitHub" link as
     fallback.

2. **GitHub Action `sync-from-sheets.yml`**
   - Trigger: `repository_dispatch` with `event_type == sync-sheets`.
   - Top-level `run-name: "Sync from Sheets [${{ github.event.client_payload.correlation_id }}]"`
     so Apps Script can locate the run.
   - Steps:
     1. checkout → setup Python → `pip install gspread google-auth`
     2. Write `GOOGLE_SERVICE_ACCOUNT_JSON` secret to a temp file
     3. `python scripts/sync_from_sheets.py --sheet-id "$SHEET_ID" [--dry-run] --result-out result.json --errors-out errors.json`
     4. If exit code 0: upload `result.json` as artifact `sync-result`. On
        non-dry-run runs with `result.changed_files` non-empty:
        `git add web/src/data/*.json`,
        `git commit -m "Sync from Sheets by $USER_EMAIL at $TIMESTAMP"`,
        `git push`. **Idempotency:** if `changed_files` is empty, skip the
        commit entirely (no empty commit on no-op publishes).
     5. If exit code 1: upload `errors.json` as artifact `sync-errors` and
        fail the workflow (exit 1) so Apps Script sees the failure
        conclusion.
   - The existing `deploy.yml` picks up the push and deploys.

3. **`scripts/sync_from_sheets.py`**
   - CLI: `--sheet-id`, `--dry-run` (no writes), `--result-out PATH`,
     `--errors-out PATH`.
   - Auth via `GOOGLE_APPLICATION_CREDENTIALS` env (path to service account
     JSON) → open the Sheet by ID.
   - Read STYLE-charts → `{chart_id: {section, chart_type}}`.
   - Read STYLE-series → `{(chart_id, series_key): {color, flags}}`.
   - Iterate tabs whose names match the canonical prefix list
     (`EDU-` / `PER-` / `RES-` / `FIN-`); ignore everything else.
   - For each chart tab:
     - Parse metadata (rows 1–11) into `ChartData` (see JSON Schema above).
     - Parse series header (rows 14–16) and join with STYLE-series.
     - Parse data table (rows 18+).
     - Validate (see Validation section).
   - If any validation errors: write structured error list to
     `errors.json` and exit 1.
   - If validation passes:
     - Compare each candidate JSON against existing
       `web/src/data/<chart_id>.json` (canonical sorted/indented JSON
       comparison to avoid spurious diffs).
     - In `--dry-run` mode: write the diff summary to `result.json` and
       exit 0 (no file writes).
     - Otherwise: write changed JSON files, refresh the INDEX tab via
       Sheets API, write a `result.json` summarising
       `changed_files: [...]` and `change_summaries: [...]`, exit 0.

4. **`deploy.yml` (existing)** — unchanged.

## Validation

Two layers. The first catches errors early; the second is the safety net.

### Layer 1 — Real-time in Sheets

Set up at bootstrap, using Google Sheets built-in Data Validation and
Conditional Formatting:

| Check | Mechanism | Effect on failure |
|---|---|---|
| Year column is 4-digit Buddhist year (2500–2700) | Data validation: number range | Reject input + tooltip |
| Value cells are numeric | Data validation: number | Reject input |
| No duplicate years | Conditional format: highlight duplicates | Red cell (warning, not blocked) |
| No empty value cells in a populated row | Conditional format: blank-in-row | Red cell border |
| Required text fields not empty (rows 3–11) | Conditional format: ISBLANK | Red cell border |

### Layer 2 — Python script before commit

`sync_from_sheets.py` re-validates everything before writing. **All-or-nothing**:
if any chart fails, no JSON is written for any chart. The dashboard never sees
a half-broken state.

```python
# pseudo (operates on ChartData dicts produced by the parser)
errors = []
for tab_name, data in parsed_charts.items():
    cid = data["id"]
    if cid not in STYLE_CHARTS:
        errors.append(f"{tab_name}: chart_id '{cid}' missing from STYLE-charts")
    years = data["categories_buddhist"]
    if len(set(years)) != len(years):
        errors.append(f"{tab_name}: duplicate years")
    if years != sorted(years):
        errors.append(f"{tab_name}: years not sorted ascending")
    for s in data["series"]:
        if (cid, s["key"]) not in STYLE_SERIES:
            errors.append(f"{tab_name}: series '{s['key']}' missing from STYLE-series")
        if len(s["values"]) != len(years):
            errors.append(f"{tab_name}/{s['key']}: value count != year count")
    for path in [("title", "th"), ("title", "en"),
                 ("subtitle", "th"), ("subtitle", "en"),
                 ("methodology", "th"), ("methodology", "en"),
                 ("source", "th"), ("source", "en")]:
        if not data[path[0]][path[1]].strip():
            errors.append(f"{tab_name}: missing {'.'.join(path)}")
# Cross-check: every chart_id in STYLE-charts must have a matching tab
for cid in STYLE_CHARTS:
    if cid not in {d["id"] for d in parsed_charts.values()}:
        errors.append(f"missing chart tab for '{cid}' (declared in STYLE-charts)")
if errors:
    write_errors_json(errors, args.errors_out)
    sys.exit(1)
```

### Edge cases

- **Deleted chart tab** → if a `chart_id` exists in STYLE but no tab matches,
  fail with `"missing tab for chart X"`. Prevents accidental loss.
- **New non-chart tab** → tabs prefixed with `_` (e.g. `_scratch`) are
  ignored. Anything else without a section prefix triggers a warning but does
  not fail.
- **Edits to STYLE** → tab is cell-protected. If protection is bypassed (only
  possible by the dev), changes appear in the next `git diff` for review.
- **Sheets API quota / timeout** → Apps Script retries the dispatch 3 times,
  then surfaces a clear error in the modal.

## Authentication

Two independent credentials, set up once.

**Apps Script → GitHub:** Fine-grained Personal Access Token, scoped to this
repo only, with `Contents: read` and `Actions: write`. Stored in Apps Script
Script Properties as `GITHUB_PAT`. Never committed.

**GitHub Action → Sheets:** Google Cloud service account (free tier). Enable
Sheets API in a Google Cloud project, create a service account, download the
JSON key, share the Sheets file with the service account's email (Viewer
role is sufficient since the action only reads). Store the full JSON in
GitHub Secrets as `GOOGLE_SERVICE_ACCOUNT_JSON`.

Both credentials are long-lived and do not require routine rotation.

## Initial Migration

One-time, run by the developer:

| Step | Time | Action |
|---|---|---|
| 1 | 5 min | Create empty Google Sheets named "KMUTT Trends — Data Source". Note Sheet ID. Create GCP project + service account, download JSON key, share Sheet with service account email. Generate fine-grained GitHub PAT. |
| 2 | 10 min | `python scripts/bootstrap_sheets.py --sheet-id <id> --credentials <path>` — script reads `web/src/data/*.json`, creates STYLE-charts + STYLE-series + 20 chart tabs with full schema (data validation, conditional formatting, cell protection, freeze panes, tab colours), and creates INDEX with HYPERLINKs. |
| 3 | 5 min | Open Sheet → Extensions → Apps Script → paste `apps_script/Code.gs` → set Script Properties `GITHUB_PAT`, `SHEET_ID`, `REPO` (e.g. `org/repo`). Reload Sheet; verify `📤 Publish` menu appears. |
| 4 | 5 min | In GitHub repo: add Secret `GOOGLE_SERVICE_ACCOUNT_JSON`, commit `.github/workflows/sync-from-sheets.yml` and `scripts/sync_from_sheets.py`. |
| 5 | 5 min | Sanity check: click "Check what will change (dry-run)" — should report no diff (data matches JSON). Edit one cell, dry-run again — should show diff. Click "Publish all changes" — verify commit + deploy + live dashboard update. |

After step 5, hand off to the data collector with `docs/data-collector-guide-th.md`.

## Open Questions

1. **Should INDEX tab's "last edited" column be populated?** Showing per-chart
   last-edit timestamps would help the data collector remember what they
   touched. Cost: an extra column update on every sync. Decide during
   implementation.
2. **Apps Script poll interval.** 5 seconds keeps the modal feeling alive but
   uses ~12 GitHub API calls per publish. If we hit secondary rate limits,
   back off to 10 s.
3. **Where to host the data-collector guide.** Markdown in `docs/` is
   discoverable for devs but harder for the data collector to find. The
   Apps Script will include a "📖 Help" menu item that opens the guide's
   URL — the actual hosting (GitHub raw, Notion, or a `/help` route under
   the dashboard) is deferred to handoff.
4. **Should `build_chart_json.py` be deleted or retained?** Retained for now
   (deprecated comment added in header). Delete after one successful cycle
   of the new flow in production.
5. **Multi-language error messages.** Error UI in Sheets modals is in Thai
   for the data collector. Validator log output in GitHub Actions stays
   English for the developer. The Python script's `errors.json` carries
   both: `{ "th": "...", "en": "..." }` per error. Confirm during
   implementation.

## Resolved decisions (from spec review)

- **Chart count is 20**, not 21 (one-time error in initial draft, corrected
  throughout).
- **`slide` field is dropped** from the JSON schema. It is a legacy PPTX
  artefact unused by the React app (confirmed: no reference in
  `web/src/{App,Chart,ChartCard,KpiCard}.tsx`). Implementation plan must
  also remove `slide: number` from `web/src/types.ts`.
- **STYLE tab is split** into STYLE-charts (per chart) and STYLE-series (per
  chart × series) to eliminate the duplicate-section problem flagged in
  review.
- **Run identification** uses a `correlation_id` injected into the
  workflow's `run-name`, so Apps Script can locate its dispatched run even
  under concurrent triggers.
- **Dry-run results and validation errors** are returned to Apps Script via
  GitHub workflow artifacts (`sync-result`, `sync-errors`) — not stdout,
  not commit comments.
- **Idempotency:** the workflow skips the commit if no JSON files changed,
  so repeated "Publish" clicks with no edits do not produce empty commits.

## Effort Estimate

~3–5 developer-days, including documentation and testing. Breakdown will be
worked out in the implementation plan.

## Risks

- **Service account credential leak.** Mitigated by storing only in GitHub
  Secrets and Apps Script Script Properties; never in code; fine-grained
  scope.
- **Sheets workbook accidentally deleted or made un-shared.** Mitigated by
  Google Drive's 30-day trash + an annual reminder for the dev to confirm
  the Sheet is still accessible to the service account.
- **Schema drift between STYLE tab and chart tabs.** Mitigated by the
  validator: any mismatch fails the publish.
- **One person, no backup.** If the data collector leaves, the Sheet
  workbook + the Thai guide must be enough for a successor to take over.
  Address through documentation quality, not infrastructure.
