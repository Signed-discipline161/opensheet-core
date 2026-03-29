import os
import pytest
import opensheet_core

FIXTURES = os.path.join(os.path.dirname(__file__), "fixtures")
BASIC = os.path.join(FIXTURES, "basic.xlsx")


def test_sheet_names():
    names = opensheet_core.sheet_names(BASIC)
    assert names == ["Basic", "Sparse", "Formulas"]


def test_read_xlsx_returns_all_sheets():
    sheets = opensheet_core.read_xlsx(BASIC)
    assert len(sheets) == 3
    assert sheets[0]["name"] == "Basic"
    assert sheets[1]["name"] == "Sparse"
    assert sheets[2]["name"] == "Formulas"


def test_read_xlsx_basic_types():
    sheets = opensheet_core.read_xlsx(BASIC)
    rows = sheets[0]["rows"]

    # Header row
    assert rows[0] == ["Name", "Age", "Active", "Score"]

    # Data rows with mixed types
    assert rows[1][0] == "Alice"
    assert rows[1][1] == 30
    assert rows[1][2] is True
    assert rows[1][3] == 95.5

    assert rows[2][0] == "Bob"
    assert rows[2][1] == 25
    assert rows[2][2] is False
    assert rows[2][3] == 87  # 87.0 -> int

    assert rows[3][0] == "Charlie"
    assert rows[3][1] == 35
    assert rows[3][2] is True
    assert rows[3][3] == 91.2


def test_read_xlsx_sparse_sheet():
    sheets = opensheet_core.read_xlsx(BASIC)
    rows = sheets[1]["rows"]

    # Row 1: A1="Header A", B1=empty, C1="Header C"
    assert rows[0][0] == "Header A"
    assert rows[0][1] is None
    assert rows[0][2] == "Header C"

    # Row 2: empty
    assert rows[1] == []

    # Row 3: numbers
    assert rows[2] == [100, 200, 300]


def test_read_sheet_by_name():
    rows = opensheet_core.read_sheet(BASIC, sheet_name="Basic")
    assert rows[0] == ["Name", "Age", "Active", "Score"]
    assert len(rows) == 4


def test_read_sheet_by_index():
    rows = opensheet_core.read_sheet(BASIC, sheet_index=1)
    assert rows[0][0] == "Header A"


def test_read_sheet_default_first():
    rows = opensheet_core.read_sheet(BASIC)
    assert rows[0] == ["Name", "Age", "Active", "Score"]


def test_read_sheet_not_found():
    with pytest.raises(ValueError, match="not found"):
        opensheet_core.read_sheet(BASIC, sheet_name="DoesNotExist")


def test_read_sheet_index_out_of_range():
    with pytest.raises(ValueError, match="out of range"):
        opensheet_core.read_sheet(BASIC, sheet_index=99)


def test_file_not_found():
    with pytest.raises(FileNotFoundError):
        opensheet_core.read_xlsx("/nonexistent/file.xlsx")


def test_integer_vs_float():
    """Whole numbers should be returned as int, fractional as float."""
    sheets = opensheet_core.read_xlsx(BASIC)
    rows = sheets[0]["rows"]

    # 30 is int
    assert isinstance(rows[1][1], int)
    # 95.5 is float
    assert isinstance(rows[1][3], float)
