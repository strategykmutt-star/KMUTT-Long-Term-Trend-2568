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
┌────────────────────────────────────┐
│   GitHub Action:                    │
│   sync-from-sheets.yml              │
│   Job 1 (sync):                     │
│     - reads Sheets via service acct │
│     - validates                     │
│     - writes JSON                   │
│     - commits + pushes (if changed) │
│   Job 2 (deploy):                   │
│     - npm ci + npm run build        │
│     - upload-pages-artifact +       │
│       deploy-pages                  │
│     - ALWAYS runs after a           │
│       non-dry-run sync (even on    │
│       no-op, for recovery)         │
└───────────┬────────────────────────┘
            ▼
       Live dashboard
       (GitHub Pages)

       The standalone deploy.yml workflow stays
       as-is for manual deploys + non-sync
       pushes to main.
```

> **Why deploy is inlined as a job, not chained via dispatch:** GitHub does
> not start a downstream workflow when commits are pushed using the default
> `GITHUB_TOKEN` (an anti-recursion safeguard, see
> [docs](https://docs.github.com/en/actions/security-for-github-actions/security-guides/automatic-token-authentication)).
> The first design tried to bridge this with `gh workflow run deploy.yml`,
> but that requires `actions: write` AND introduces a two-workflow split
> where the publish modal can falsely report success while the deploy run
> is still queued (or has failed). Inlining the deploy as a job in the
> sync workflow means: (a) one workflow run, one polling target, one
> truth; (b) the modal only reports success after build+deploy finish;
> (c) re-clicking Publish on a no-op edit still re-runs deploy, so a
> stuck "repo updated but Pages not deployed" state is recoverable by
> just clicking Publish again.

**What is preserved:**

- React/Vite/TypeScript app structure
- `web/src/data/*.json` schema (consumer code unchanged)
- Existing `deploy.yml` (kept unchanged — still serves manual deploys
  and non-sync pushes to main)
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

- 🔒 = **hard-protected via Sheets API protected range with explicit
  `editors` allowlist** = `[<dev-email>]`. The data collector is an
  Editor of the workbook, so warning-only protection wouldn't block them;
  the editors-allowlist mechanism does. Bootstrap collects the dev's
  email via `--dev-email` CLI flag and writes it into each protected
  range definition.
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

### INDEX tab schema (bootstrap-generated, static thereafter)

| Chart ID | Section | Title TH | Title EN | Jump |
|---|---|---|---|---|
| students-all | education | จำนวนนักศึกษาทั้งหมด | Total Enrolled Students | → click |
| ... | ... | ... | ... | ... |

- **Generated once** by `bootstrap_sheets.py`. Not updated by
  `sync_from_sheets.py` (the service account is Viewer-only — no write
  path to the workbook).
- "Jump" column uses `HYPERLINK("#gid=…")` with each chart tab's real gid,
  captured during bootstrap.
- **Trade-off:** if a chart title is renamed in a chart tab after bootstrap,
  the INDEX title becomes stale. This is acceptable because: titles rarely
  change, the INDEX is for navigation (the Title columns are informational),
  and a dev re-running bootstrap (or manually fixing one cell in INDEX) is
  cheap. The earlier draft of this spec promised dynamic INDEX refresh, but
  that required Editor permission on the workbook for the service account
  — a security/complexity cost not worth the convenience.
- Earlier "# years" and "# series" columns are dropped: they would have
  required dynamic refresh too.

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
   - Reads `SHEET_ID`, `GITHUB_PAT`, and `REPO` from Script Properties (set
     once at install; not hard-coded so the same script works across forks).
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
   - Opens an HtmlService modal **immediately** after the dispatch returns
     (the dispatch call itself is fast; no synchronous polling on the
     Apps Script server). The modal contains client-side JavaScript that
     calls back into the script via `google.script.run` to poll status
     every 5 seconds. This avoids hitting the Apps Script 6-minute
     execution-time cap, and the user sees live progress.
   - Per-poll, the server calls
     `GET /repos/{repo}/actions/runs?event=repository_dispatch&per_page=20`,
     and filters by **`display_title`** (NOT `name` — see note below).
     The workflow run's `display_title` is the field that reflects the
     workflow's `run-name:` value; `name` is the workflow's static
     `name:` field.
   - When the matched run reaches `status=completed`:
     - On `conclusion=success` → download the `sync-result` artifact (a
       single JSON file) and render the appropriate modal (publish success
       summary, or dry-run diff).
     - On `conclusion=failure` → download the `sync-errors` artifact (JSON
       list of validation errors) and render the error modal.
   - Modal shows a "View run on GitHub" link from the moment it opens, so
     the user has a fallback even if polling fails.

2. **GitHub Action `sync-from-sheets.yml`** — two jobs in one workflow run.

   **Top-level configuration:**
   - Trigger: `repository_dispatch` with `event_type == sync-sheets`
   - `run-name: "Sync from Sheets [${{ github.event.client_payload.correlation_id }}]"`
   - `permissions: contents: write, pages: write, id-token: write,
     actions: read` (the `pages` and `id-token` permissions are required
     by the deploy job)

   **Job 1 — `sync`:**
   1. checkout → setup Python → `pip install -r requirements-dev.txt`
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

   **Job 2 — `deploy`:**
   - `needs: sync`, `if: success() && github.event.client_payload.dry_run != true`
     — runs ALWAYS on a successful non-dry-run sync, even when no JSON
     files changed (no-op publish should still re-deploy, so the data
     collector can recover a "repo updated but Pages stale" state by just
     clicking Publish again).
   - Replicates the build/deploy steps from the existing `deploy.yml`
     (checkout → setup-node → npm ci → npm run build → upload-pages-artifact
     → deploy-pages). Kept inline rather than calling `deploy.yml` via
     `workflow_call` to avoid permission-inheritance ambiguity.

   The Apps Script modal polls the parent workflow run; the run's
   `status=completed` only fires after BOTH jobs finish. The modal's
   "success" message is therefore only shown when the Pages deploy has
   actually completed.

3. **`scripts/sync_from_sheets.py`**
   - CLI: `--sheet-id`, `--dry-run` (no writes), `--result-out PATH`,
     `--errors-out PATH`.
   - Auth via service account JSON (path passed as `--credentials`) → open
     the Sheet by ID. Read-only scope is sufficient.
   - Read STYLE-charts → `{chart_id: {section, chart_type}}`.
   - Read STYLE-series → `{(chart_id, series_key): {color, flags}}`.
   - Iterate tabs whose names match the canonical prefix list
     (`EDU-` / `PER-` / `RES-` / `FIN-`); ignore everything else.
   - For each chart tab:
     - Parse metadata (rows 1–11) into `ChartData` (see JSON Schema above).
     - Parse series header (rows 14–16) — **without filtering blank
       series_key cells**, so the validator can detect column-position
       drift (see Validation section).
     - Parse data table (rows 18+).
     - Validate.
   - If any validation errors: write structured error list to
     `errors.json` and exit 1.
   - If validation passes:
     - Compare each candidate JSON against existing
       `web/src/data/<chart_id>.json` (canonical sorted/indented JSON
       comparison to avoid spurious diffs).
     - In `--dry-run` mode: write the diff summary to `result.json` and
       exit 0 (no file writes).
     - Otherwise: write changed JSON files, write a `result.json`
       summarising `changed_files: [...]`, exit 0. **Does not touch the
       INDEX tab** — INDEX is static after bootstrap (the service account
       has Viewer-only access; no write path needed).

4. **`deploy.yml` (existing)** — unchanged. It continues to handle non-sync
   pushes to `main` and manual `workflow_dispatch` deploys. The sync
   workflow does NOT call it.

## Validation

Two layers. The first catches errors early; the second is the safety net.

### Layer 1 — Real-time in Sheets

Set up at bootstrap, using Google Sheets built-in Data Validation and
Conditional Formatting:

| Check | Mechanism | Effect on failure |
|---|---|---|
| Year column is 4-digit Buddhist year (2500–2700) | Data validation: number range | Reject input + tooltip |
| Value cells are numeric **or empty** | Data validation: `=OR(ISNUMBER(A1),A1="")` | Reject input |
| No duplicate years | Conditional format: highlight duplicates | Red cell (warning, not blocked) |
| Required text fields not empty (rows 3, 4, 5, 6, 8, 9, 10, 11 — title TH/EN, subtitle TH/EN, methodology TH/EN, source TH/EN) | Conditional format: ISBLANK | Red cell border |

**Note on intentional `null` values.** 5 of the 20 current charts (patents,
programs, staff-total, research-funding, research-funding-3yr) contain
intentionally empty cells where data is genuinely missing for some years.
The validation rules **must not flag these as errors** — a blank value
cell is valid input. The Python layer translates blank cells to `null`
in the output JSON. There is intentionally **no "no empty value cells"
rule**; the earlier draft of this spec had one and it would have
false-flagged every baseline chart.

**Scope of bootstrap-time formatting.** The bootstrap script applies the
data-validation rules above and the cell protections listed below
programmatically via Sheets API `batchUpdate`. The conditional-formatting
rules (duplicate years, blank-required-fields) are also set
programmatically. If the bootstrap fails partway and re-running creates
duplicate rules, the operator can clear formatting in Sheets UI and
re-run.

### Layer 2 — Python script before commit

`sync_from_sheets.py` re-validates everything before writing. **All-or-nothing**:
if any chart fails, no JSON is written for any chart. The dashboard never sees
a half-broken state.

**Important parser contract for the validator to work.** The parser
**must preserve column positions** of row 14 (series_key row). It must
NOT silently drop blank cells. If row 14 has `[bachelor, "", graduate,
total]` (with column C blank), the parser returns a series list with
four entries, one of which has an empty key; the validator then catches
the empty cell. If the parser collapsed the list to three entries, the
data columns would silently shift to align with the wrong series.

```python
# pseudo (operates on ChartData dicts produced by the parser)
ALLOWED_FLAGS = {"emphasis", "exclude_from_stack", "is_cumulative"}
HEX_COLOR_RE = re.compile(r"^#[0-9a-fA-F]{6}$")

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
    for y in years:
        try:
            yi = int(y)
            if yi < 2500 or yi > 2700:
                errors.append(f"{tab_name}: year '{y}' outside 2500-2700")
        except ValueError:
            errors.append(f"{tab_name}: year '{y}' is not an integer")

    # series_key row, preserved-position
    series_keys = [s["key"] for s in data["series"]]
    for i, k in enumerate(series_keys):
        if not k.strip():
            errors.append(f"{tab_name}: empty series_key in row 14, column {i+2}")
    nonblank = [k for k in series_keys if k.strip()]
    if len(set(nonblank)) != len(nonblank):
        errors.append(f"{tab_name}: duplicate series_key in row 14")

    for s in data["series"]:
        if not s["key"].strip():
            continue  # already reported above; skip downstream checks for this column
        if (cid, s["key"]) not in STYLE_SERIES:
            errors.append(f"{tab_name}: series '{s['key']}' missing from STYLE-series")
        if len(s["values"]) != len(years):
            errors.append(f"{tab_name}/{s['key']}: value count != year count")
        if not s["name"]["th"].strip():
            errors.append(f"{tab_name}/{s['key']}: missing series name (TH)")
        if not s["name"]["en"].strip():
            errors.append(f"{tab_name}/{s['key']}: missing series name (EN)")

    for path in [("title", "th"), ("title", "en"),
                 ("subtitle", "th"), ("subtitle", "en"),
                 ("methodology", "th"), ("methodology", "en"),
                 ("source", "th"), ("source", "en")]:
        if not data[path[0]][path[1]].strip():
            errors.append(f"{tab_name}: missing {'.'.join(path)}")

# STYLE-series sanity checks (run once, not per tab)
for (cid, sk), style in STYLE_SERIES.items():
    if not HEX_COLOR_RE.match(style["color"]):
        errors.append(f"STYLE-series ({cid}/{sk}): invalid color '{style['color']}'")
    for flag in style["flags"]:
        if flag not in ALLOWED_FLAGS:
            errors.append(f"STYLE-series ({cid}/{sk}): unknown flag '{flag}'")

# Cross-check: every chart_id in STYLE-charts must have a matching tab
for cid in STYLE_CHARTS:
    if cid not in {d["id"] for d in parsed_charts.values()}:
        errors.append(f"missing chart tab for '{cid}' (declared in STYLE-charts)")

# Cross-check: every (chart_id, series_key) declared in STYLE-series must
# appear in the corresponding chart tab. This prevents a data collector
# from accidentally deleting a series header (which would otherwise produce
# series:[...] missing the key, crashing the React app's KpiCard).
expected_by_chart = {}
for cid, sk in STYLE_SERIES:
    expected_by_chart.setdefault(cid, set()).add(sk)
for tab, data in parsed_charts.items():
    cid = data["id"]
    actual = {s["key"] for s in data["series"] if s["key"].strip()}
    for missing_sk in expected_by_chart.get(cid, set()) - actual:
        errors.append(f"{tab}: series '{missing_sk}' declared in STYLE-series "
                      f"but missing from chart tab")

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
repo only, with `Contents: write` and `Actions: read` permissions.

- `Contents: write` is required by the
  [`POST /repos/{owner}/{repo}/dispatches`](https://docs.github.com/en/rest/repos/repos#create-a-repository-dispatch-event)
  endpoint. (`read` is not enough — `repository_dispatch` is classified as
  a Contents-write operation.)
- `Actions: read` is required to poll workflow runs and download artifacts
  (sync-result, sync-errors).
- Stored in Apps Script Script Properties as `GITHUB_PAT`. Never committed.

**Credential trust boundary (important).** The Apps Script is a
[bound script](https://developers.google.com/apps-script/guides/bound)
attached to the Sheets workbook. Anyone with Editor access on the workbook
can open Extensions → Apps Script, read Script Properties, and exfiltrate
the PAT. Because the PAT has `Contents: write` on the repo, an editor of
the workbook is effectively trusted with write access to `web/src/data/`
and the ability to trigger workflows.

This trust boundary is **explicit and accepted** for this project:
- The data collector already needs Editor on the workbook to do their job.
- The dashboard is non-sensitive (public-facing data already).
- There is one data collector, internal to KMUTT.

If the trust assumption changes later (multiple editors, sensitive data,
adversarial environment), migrate to one of:
1. A GitHub App with `Contents: write` only on this repo (cleaner scope,
   revocable per-install).
2. A Cloud Function intermediary that holds the PAT and exposes only a
   `dispatch-sync` endpoint to the Apps Script.

**GitHub Action → Sheets:** Google Cloud service account (free tier). Enable
Sheets API in a Google Cloud project, create a service account, download the
JSON key, share the Sheets file with the service account's email.

- **Role: Editor** (initially required by `bootstrap_sheets.py`, which
  creates/deletes tabs, writes values, and applies protections/validation
  via `batchUpdate`). After bootstrap completes successfully, the operator
  **may downgrade the service account to Viewer** for least-privilege
  runtime, since `sync_from_sheets.py` only reads. The runbook documents
  this as an optional hardening step.
- The trade-off: leaving the service account as Editor means a leaked
  service-account key would allow writes to the workbook (vandalism is
  recoverable from Google Drive's 30-day trash). Downgrading to Viewer
  closes that gap at the cost of needing to re-grant Editor before any
  future bootstrap re-run. Acceptable either way; default is "downgrade
  after bootstrap" because it's a single click in Sheets UI.
- Store the full JSON in GitHub Secrets as `GOOGLE_SERVICE_ACCOUNT_JSON`.

Both credentials are long-lived and do not require routine rotation.

## Initial Migration

One-time, run by the developer. Step 0 is a **prerequisite**: without it,
the "first dry-run should show no diff" sanity check in step 5 cannot pass
because every JSON file still carries the legacy `slide` key.

| Step | Time | Action |
|---|---|---|
| **0** | **5 min** | **Schema-cleanup commit on `main`:** strip the `slide` key from all 20 `web/src/data/*.json` files (e.g. `jq 'del(.slide)' file.json` or a small Python one-liner). Update `web/src/types.ts` to drop `slide: number` and change `subtitle: Bilingual | null` to `subtitle: Bilingual`. Update `web/src/components/ChartCard.tsx` to remove the now-unnecessary `{data.subtitle && ...}` null check. Run `npm run build` to confirm. Commit and push. This is a no-op for end-users; it just makes the post-bootstrap JSON byte-identical to the pre-bootstrap JSON, so subsequent dry-runs report 0 changes (proving the round-trip works). |
| 1 | 5 min | Create empty Google Sheets named "KMUTT Trends — Data Source". Note Sheet ID. Create GCP project + service account, download JSON key, share Sheet with service account email — **Editor role** (bootstrap needs write). Also share with the data collector (Editor) and any other dev maintainers (Editor). Generate fine-grained GitHub PAT (Contents: write, Actions: read). |
| 2 | 10 min | `python scripts/bootstrap_sheets.py --sheet-id <id> --credentials <path> --dev-email <dev's google account>` — script reads `web/src/data/*.json`, creates STYLE-charts + STYLE-series + 20 chart tabs with full schema (data validation, conditional formatting for required text fields and duplicate years, **hard-protected ranges with `editors=[dev-email]`**, freeze panes, tab colours), and creates INDEX with HYPERLINKs. |
| 2b | 1 min | (Optional hardening) In the Sheets share dialog, **downgrade the service account from Editor to Viewer**. After this, future bootstrap re-runs require restoring Editor temporarily. |
| 3 | 5 min | Open Sheet → Extensions → Apps Script → paste `apps_script/Code.gs` AND `apps_script/PublishModal.html` → set Script Properties `GITHUB_PAT`, `SHEET_ID`, `REPO` (e.g. `org/repo`), `HELP_URL`. Reload Sheet; verify `📤 Publish` menu appears. |
| 4 | 5 min | In GitHub repo: add Secret `GOOGLE_SERVICE_ACCOUNT_JSON`, commit `.github/workflows/sync-from-sheets.yml` and `scripts/sync_from_sheets.py`. |
| 5 | 5 min | Sanity check: click "Check what will change (dry-run)" — should report no diff (data matches JSON because step 0 cleaned up the schema). Edit one cell, dry-run again — should show 1 file diff. Click "Publish all changes" — verify: sync job commits, deploy job builds + deploys, modal reports success only after BOTH finish, live dashboard updates. |
| 5b | 2 min | Test recovery path: click Publish a second time without editing. Modal should still complete (sync no-op, deploy re-runs), dashboard should re-publish identical content. |

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
- **`subtitle` becomes required** (was `Bilingual | null` in types.ts). All
  20 current JSON files already have a non-null subtitle, so this is a
  no-op for the data but tightens the contract. Implementation plan must
  also change `subtitle: Bilingual | null` to `subtitle: Bilingual` in
  `web/src/types.ts` and verify `Chart.tsx`/`ChartCard.tsx` no longer
  branch on null.
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

## Resolved decisions (round 3 — Codex external review)

- **Deploy is dispatched explicitly, not via push trigger.** GitHub's
  `GITHUB_TOKEN` does not start downstream workflows when used to push, so
  the sync workflow ends with an explicit `gh workflow run deploy.yml`
  step. `deploy.yml`'s `workflow_dispatch:` trigger (already present)
  becomes load-bearing.
- **PAT scope is `Contents: write` + `Actions: write`** (was incorrectly
  written as `Contents: read`). `repository_dispatch` is a Contents-write
  operation per GitHub's permission model.
- **Apps Script polls workflow runs via `display_title`, not `name`.**
  GitHub's run schema separates the workflow's static `name:` field from
  the per-run `display_title` set by `run-name:`. Polling on `name` would
  always miss.
- **Credential trust boundary is documented and accepted.** Anyone with
  Editor access to the Sheets workbook can exfiltrate the PAT via Apps
  Script Properties. Accepted because there's one trusted internal data
  collector and public-facing data. Future migration paths (GitHub App /
  Cloud Function proxy) are listed in the Authentication section.
- **Initial migration includes a `slide`-cleanup commit (step 0)** so the
  first dry-run after bootstrap shows zero diff. Without this step, the
  sanity check at step 5 would surface a 20-file diff that's purely the
  schema-cleanup migration.
- **INDEX tab is bootstrap-static, not dynamically refreshed by sync.**
  Service account stays Viewer-only; INDEX columns trimmed to fields that
  don't need refresh (Chart ID, Section, Title TH/EN, Jump).
- **Layer-1 validation removes the "no empty value cells" rule** because
  5 charts have legitimate `null` values for years where data is
  genuinely missing. Empty value cells are valid input.
- **Parser preserves column positions** of the series_key row (row 14);
  it does NOT silently drop blanks. The validator detects and reports
  blank/duplicate series_key cells so the column→series alignment can't
  silently drift.
- **Validator adds checks** for year integer-and-range, series name
  TH/EN presence, STYLE-series color hex format, STYLE-series flag
  allowlist.
- **Apps Script UX is async via HtmlService + `google.script.run`.**
  Server-side polling was incompatible with the Apps Script 6-minute
  execution cap; the new UX runs polling client-side in the modal's
  browser context, with short server-side RPCs per poll.

## Resolved decisions (round 4 — Codex external review)

- **Deploy is inlined as a job in the sync workflow**, not chained via
  `gh workflow run`. Two-workflow chain caused two problems: (a) needed
  `actions: write` (was set to `read`), and (b) the modal could report
  success while deploy was still queued or had failed. Single workflow
  with sync + deploy jobs makes the modal honest about end-to-end state.
- **Deploy job always runs on a non-dry-run sync**, even when no JSON
  files changed. This gives the data collector a recovery path for the
  "repo updated but Pages stale" stuck state — just click Publish again.
- **PAT scope corrected from `Actions: write` to `Actions: read`.** The
  Apps Script no longer dispatches anything; it only polls runs and
  downloads artifacts.
- **Service account starts as Editor** (bootstrap requires it). Optional
  step 2b in Initial Migration downgrades to Viewer after bootstrap.
- **Protected ranges use editors-allowlist, not warningOnly.** With the
  data collector as a workbook Editor, warning-only protections still
  let them edit after clicking through the warning. Hard protection with
  `editors=[dev-email]` actually blocks them. Bootstrap takes `--dev-email`.
- **Validator checks STYLE-series → chart-tab completeness.** For each
  (chart_id, series_key) declared in STYLE-series, the chart tab must
  contain that series. Without this, a data collector blanking the entire
  series header row would publish `series:[]` — and the React app's
  KpiCard does `series.values` directly, so the dashboard would crash.
- **Bootstrap implements the full Layer-1 conditional formatting set**
  promised by the spec — duplicate years AND blank-required-text rules
  for rows 3–6, 8–11 (title/subtitle/methodology/source TH+EN).
- **INDEX schema trimmed** to Chart ID, Section, Title TH/EN, Jump — the
  `# years` and `# series` columns are removed because they can't be
  refreshed (INDEX is bootstrap-static; service account is Viewer).
- **Test fixture color fixed** from `#000` to `#000000` so
  `test_valid_chart_produces_no_errors` actually passes the hex-format
  validator added in round 3.

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
