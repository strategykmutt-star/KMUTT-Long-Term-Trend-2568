# tests/test_parsers.py
import json
from pathlib import Path
import pytest
from scripts.lib.parsers import parse_style_charts, parse_style_series, parse_chart_tab

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
