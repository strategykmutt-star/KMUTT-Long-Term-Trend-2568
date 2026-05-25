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
