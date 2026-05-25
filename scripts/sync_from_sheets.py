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
