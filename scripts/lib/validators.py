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
                    f"year '{y}' not a 4-digit BE year or NNNN-NNNN range (must be integer)"))
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
            errors.append(_err("🎨 STYLE-series", f"{cid}/{sk}/color",
                f"สี '{style['color']}' ของ {cid}/{sk} ไม่ใช่ hex 6 หลัก",
                f"invalid hex color '{style['color']}' for {cid}/{sk}"))
        for flag in style["flags"]:
            if flag not in ALLOWED_FLAGS:
                errors.append(_err("🎨 STYLE-series", f"{cid}/{sk}/flags",
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
    # tab's row 14 -- the previous cross-check would silently pass.
    # Publishing series:[] would crash the React KpiCard.
    for tab, data in parsed_charts.items():
        if not any(s["key"].strip() for s in data["series"]):
            errors.append(_err(tab, "series",
                "กราฟต้องมี series อย่างน้อย 1 ตัว",
                "chart has zero series -- at least one is required"))

    # The kpi_series_key declared in STYLE-charts must exist in the
    # chart's actual series list. Without this, KpiCard's fallback to
    # data.series[0] silently shows the wrong KPI. A BLANK kpi_series_key
    # is itself a contract violation -- STYLE-charts is documented as the
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
