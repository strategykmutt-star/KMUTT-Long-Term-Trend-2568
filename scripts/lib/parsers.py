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
