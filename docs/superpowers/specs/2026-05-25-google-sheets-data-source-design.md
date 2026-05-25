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
(title, subtitle, methodology, source, series names) for the existing 21
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
│     🎨 STYLE           │
│     21 chart tabs      │
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
| `scripts/bootstrap_sheets.py` | One-time: 21 JSON files → Sheets |
| `scripts/lib/{sheets_client,parsers,validators}.py` | Shared helpers |
| `docs/data-collector-guide-th.md` | Thai-language manual for the data collector |
| GitHub Secret `GOOGLE_SERVICE_ACCOUNT_JSON` | GitHub Action → Sheets API auth |
| Apps Script property `GITHUB_PAT` | Apps Script → GitHub API auth |

## Sheet Schema

### Workbook layout

23 tabs in one Google Sheets workbook:

```
📋 INDEX            (auto-generated navigation, read-only)
🎨 STYLE            (chart_type, section, colour, flags — locked to data collector)
EDU-students-all    (1 tab per chart, prefix = section)
EDU-students-new
EDU-graduates
... (18 more chart tabs)
RES-patents
FIN-income-expense
```

Tab prefixes group charts visually (`EDU-` / `PER-` / `RES-` / `FIN-`).
Tab background colour mirrors the section.

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

### STYLE tab schema (dev-only)

| chart_id | chart_type | section | series_key | color | flags |
|---|---|---|---|---|---|
| students-all | line | education | bachelor | #f29400 | |
| students-all | line | education | graduate | #1e6091 | |
| students-all | line | education | total | #0f172a | emphasis |
| ... | ... | ... | ... | ... | ... |

- One row per (chart × series).
- Entire tab is protected; only developers (with edit permission on the
  sheet) can change it.
- `flags` column accepts a comma-separated subset of `emphasis`,
  `exclude_from_stack`, `is_cumulative`.

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
   - Reads Sheet ID + current user's email.
   - POST `/repos/<org>/<repo>/dispatches` to GitHub API with header
     `Authorization: Bearer <GITHUB_PAT>` and body
     `{ event_type: "sync-sheets", client_payload: { sheet_id, user_email, dry_run: bool } }`.
   - Polls workflow status every 5 seconds and updates the modal.

2. **GitHub Action `sync-from-sheets.yml`**
   - Trigger: `repository_dispatch` with `event_type == sync-sheets`.
   - Steps: checkout → setup Python → `pip install gspread google-auth` →
     write `GOOGLE_SERVICE_ACCOUNT_JSON` secret to a temp file →
     `python scripts/sync_from_sheets.py --sheet-id $SHEET_ID [--dry-run]`.
   - If non-dry-run and the script wrote changes:
     `git add web/src/data/*.json`,
     `git commit -m "Sync from Sheets by $USER_EMAIL at $TIMESTAMP"`,
     `git push`.
   - The existing `deploy.yml` picks up the push and deploys.

3. **`scripts/sync_from_sheets.py`**
   - Auth via service account JSON → open the Sheet by ID.
   - Read STYLE tab → build `{chart_id: {series_key: {color, flags, chart_type, section}}}` map.
   - Iterate non-admin tabs (those not starting with `📋` or `🎨` or `_`):
     - Parse metadata (rows 1–11).
     - Parse series header (rows 14–16) and join with STYLE config.
     - Parse data table (rows 18+).
     - Run validation (see below).
   - If validation passes: write each chart's JSON to `web/src/data/<chart_id>.json` and update INDEX tab.
   - If dry-run: print a unified diff to stdout (consumed by Apps Script) and exit 0 without writing.
   - If validation fails: print errors as structured Markdown and exit 1.

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
# pseudo
errors = []
for tab in chart_tabs:
    data = parse_tab(tab)
    if data.chart_id not in STYLE_CONFIG:
        errors.append(f"{tab}: chart_id '{data.chart_id}' missing from STYLE")
    years = data.years
    if len(set(years)) != len(years): errors.append(f"{tab}: duplicate years")
    if years != sorted(years): errors.append(f"{tab}: years not sorted")
    for sk in data.series_keys:
        if sk not in STYLE_CONFIG.get(data.chart_id, {}):
            errors.append(f"{tab}: series '{sk}' not in STYLE")
        if len(data.values[sk]) != len(years):
            errors.append(f"{tab}: series '{sk}' value count != year count")
    for f in ["title_th", "title_en", "methodology_th", "methodology_en", "source_th", "source_en"]:
        if not data.metadata[f].strip():
            errors.append(f"{tab}: missing {f}")
if errors:
    print_errors_as_markdown(errors)
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
| 2 | 10 min | `python scripts/bootstrap_sheets.py --sheet-id <id> --credentials <path>` — script reads `web/src/data/*.json`, creates STYLE + 21 chart tabs with full schema (data validation, conditional formatting, cell protection, freeze panes, tab colours), and creates INDEX with HYPERLINKs. |
| 3 | 5 min | Open Sheet → Extensions → Apps Script → paste `apps_script/Code.gs` → set Script Property `GITHUB_PAT`. Reload Sheet; verify `📤 Publish` menu appears. |
| 4 | 5 min | In GitHub repo: add Secret `GOOGLE_SERVICE_ACCOUNT_JSON`, commit `.github/workflows/sync-from-sheets.yml` and `scripts/sync_from_sheets.py`. |
| 5 | 5 min | Sanity check: click "Check what will change (dry-run)" — should report no diff (data matches JSON). Edit one cell, dry-run again — should show diff. Click "Publish all changes" — verify commit + deploy + live dashboard update. |

After step 5, hand off to the data collector with `docs/data-collector-guide-th.md`.

## Open Questions

1. **Should INDEX tab's "last edited" column be populated?** Showing per-chart
   last-edit timestamps would help the data collector remember what they
   touched. Cost: an extra column update on every sync. Decide during
   implementation.
2. **Apps Script poll interval.** 5 seconds keeps the modal feeling alive but
   uses ~24 GitHub API calls per publish. If we hit secondary rate limits,
   back off to 10 s.
3. **Where to host the data-collector guide.** Markdown in `docs/` is
   discoverable for devs but harder for the data collector to find. Options:
   embed a "Help" menu item in the Apps Script that opens the guide URL; or
   publish to a static page under the dashboard itself. Decide before
   handoff.
4. **Should `build_chart_json.py` be deleted or retained?** Retained for now
   (deprecated comment in header). Delete after one successful cycle of the
   new flow.
5. **Multi-language error messages.** All error UI is currently in Thai for
   the data collector. Validator log output in GitHub Actions stays English
   for the developer. Confirm this split is fine.

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
