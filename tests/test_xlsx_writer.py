import os
import tempfile
import pytest
import opensheet_core


@pytest.fixture
def tmp_xlsx(tmp_path):
    return str(tmp_path / "output.xlsx")


def test_basic_write_and_read(tmp_xlsx):
    writer = opensheet_core.XlsxWriter(tmp_xlsx)
    writer.add_sheet("Data")
    writer.write_row(["Name", "Value"])
    writer.write_row(["Alice", 42])
    writer.close()

    sheets = opensheet_core.read_xlsx(tmp_xlsx)
    assert len(sheets) == 1
    assert sheets[0]["name"] == "Data"
    assert sheets[0]["rows"][0] == ["Name", "Value"]
    assert sheets[0]["rows"][1] == ["Alice", 42]


def test_multiple_sheets(tmp_xlsx):
    writer = opensheet_core.XlsxWriter(tmp_xlsx)
    writer.add_sheet("Sheet1")
    writer.write_row(["a", "b"])
    writer.add_sheet("Sheet2")
    writer.write_row([1, 2, 3])
    writer.close()

    names = opensheet_core.sheet_names(tmp_xlsx)
    assert names == ["Sheet1", "Sheet2"]

    rows1 = opensheet_core.read_sheet(tmp_xlsx, sheet_name="Sheet1")
    assert rows1 == [["a", "b"]]

    rows2 = opensheet_core.read_sheet(tmp_xlsx, sheet_name="Sheet2")
    assert rows2 == [[1, 2, 3]]


def test_all_types(tmp_xlsx):
    writer = opensheet_core.XlsxWriter(tmp_xlsx)
    writer.add_sheet("Types")
    writer.write_row(["text", 42, 3.14, True, False, None])
    writer.close()

    rows = opensheet_core.read_sheet(tmp_xlsx)
    assert rows[0][0] == "text"
    assert rows[0][1] == 42
    assert rows[0][2] == 3.14
    assert rows[0][3] is True
    assert rows[0][4] is False
    # None (empty) cells are not written, so they won't appear at the end


def test_context_manager(tmp_xlsx):
    with opensheet_core.XlsxWriter(tmp_xlsx) as w:
        w.add_sheet("Auto")
        w.write_row(["closed", "automatically"])

    rows = opensheet_core.read_sheet(tmp_xlsx)
    assert rows == [["closed", "automatically"]]


def test_write_after_close_raises(tmp_xlsx):
    writer = opensheet_core.XlsxWriter(tmp_xlsx)
    writer.add_sheet("X")
    writer.write_row(["ok"])
    writer.close()

    with pytest.raises(RuntimeError, match="already closed"):
        writer.write_row(["fail"])


def test_write_without_sheet_raises(tmp_xlsx):
    writer = opensheet_core.XlsxWriter(tmp_xlsx)
    with pytest.raises(Exception):
        writer.write_row(["no sheet"])
    writer.close()


def test_special_characters(tmp_xlsx):
    writer = opensheet_core.XlsxWriter(tmp_xlsx)
    writer.add_sheet("Special")
    writer.write_row(["a & b", "<tag>", 'quote "here"', "it's fine"])
    writer.close()

    rows = opensheet_core.read_sheet(tmp_xlsx)
    assert rows[0][0] == "a & b"
    assert rows[0][1] == "<tag>"
    assert rows[0][2] == 'quote "here"'
    assert rows[0][3] == "it's fine"


def test_empty_rows(tmp_xlsx):
    writer = opensheet_core.XlsxWriter(tmp_xlsx)
    writer.add_sheet("Gaps")
    writer.write_row(["row1"])
    writer.write_row([])
    writer.write_row(["row3"])
    writer.close()

    rows = opensheet_core.read_sheet(tmp_xlsx)
    assert len(rows) == 3
    assert rows[0] == ["row1"]
    assert rows[2] == ["row3"]


def test_large_write(tmp_xlsx):
    """Write 10k rows to verify streaming doesn't blow up."""
    with opensheet_core.XlsxWriter(tmp_xlsx) as w:
        w.add_sheet("Big")
        for i in range(10000):
            w.write_row([f"row_{i}", i, i * 0.1])

    rows = opensheet_core.read_sheet(tmp_xlsx)
    assert len(rows) == 10000
    assert rows[0] == ["row_0", 0, 0.0]
    assert rows[9999][0] == "row_9999"


def test_formula_write_and_read(tmp_xlsx):
    """Write formulas and verify they round-trip."""
    from opensheet_core import Formula

    with opensheet_core.XlsxWriter(tmp_xlsx) as w:
        w.add_sheet("Formulas")
        w.write_row(["A", "B", "Sum"])
        w.write_row([10, 20, Formula("A2+B2", cached_value=30)])
        w.write_row([5, 15, Formula("A3+B3")])

    sheets = opensheet_core.read_xlsx(tmp_xlsx)
    rows = sheets[0]["rows"]

    assert rows[0] == ["A", "B", "Sum"]
    # Row with formula + cached value
    assert rows[1][0] == 10
    assert rows[1][1] == 20
    f1 = rows[1][2]
    assert isinstance(f1, Formula)
    assert f1.formula == "A2+B2"
    assert f1.cached_value == 30

    # Row with formula, no cached value
    f2 = rows[2][2]
    assert isinstance(f2, Formula)
    assert f2.formula == "A3+B3"
    assert f2.cached_value is None


def test_formula_class():
    """Test Formula class creation and attributes."""
    from opensheet_core import Formula

    f = Formula("SUM(A1:A10)")
    assert f.formula == "SUM(A1:A10)"
    assert f.cached_value is None
    assert "SUM(A1:A10)" in repr(f)

    f2 = Formula("A1*2", cached_value=42)
    assert f2.formula == "A1*2"
    assert f2.cached_value == 42


def test_formula_equality():
    """Test Formula __eq__ comparisons."""
    from opensheet_core import Formula

    assert Formula("A1+B1") == Formula("A1+B1")
    assert Formula("A1+B1", cached_value=10) == Formula("A1+B1", cached_value=10)
    assert Formula("A1+B1") != Formula("A1+B2")
    assert Formula("A1+B1", cached_value=10) != Formula("A1+B1", cached_value=20)
    assert Formula("A1+B1") != Formula("A1+B1", cached_value=10)


def test_read_openpyxl_file(tmp_xlsx):
    """Write with openpyxl, read with opensheet_core (interop validation)."""
    openpyxl = pytest.importorskip("openpyxl")

    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Interop"
    ws.append(["Name", "Age", "Score"])
    ws.append(["Alice", 30, 95.5])
    ws.append(["Bob", 25, 87])
    ws.append([None, None, None])  # empty row
    ws.append(["Charlie", 35, 91.2])
    wb.save(tmp_xlsx)

    rows = opensheet_core.read_sheet(tmp_xlsx)
    assert rows[0] == ["Name", "Age", "Score"]
    assert rows[1] == ["Alice", 30, 95.5]
    assert rows[2] == ["Bob", 25, 87]
    assert rows[3] == []  # empty row
    assert rows[4] == ["Charlie", 35, 91.2]


def test_date_write_and_read(tmp_xlsx):
    """Write dates and verify they round-trip."""
    import datetime

    with opensheet_core.XlsxWriter(tmp_xlsx) as w:
        w.add_sheet("Dates")
        w.write_row(["Event", "Date", "Timestamp"])
        w.write_row([
            "Launch",
            datetime.date(2025, 3, 15),
            datetime.datetime(2025, 3, 15, 14, 30, 0),
        ])
        w.write_row([
            "Update",
            datetime.date(2021, 1, 1),
            datetime.datetime(2021, 6, 15, 9, 0, 0),
        ])

    rows = opensheet_core.read_sheet(tmp_xlsx)

    assert rows[0] == ["Event", "Date", "Timestamp"]

    # Date cells
    assert rows[1][0] == "Launch"
    assert rows[1][1] == datetime.date(2025, 3, 15)
    assert isinstance(rows[1][1], datetime.date)

    # DateTime cells
    assert rows[1][2] == datetime.datetime(2025, 3, 15, 14, 30, 0)
    assert isinstance(rows[1][2], datetime.datetime)

    # Second row
    assert rows[2][1] == datetime.date(2021, 1, 1)
    assert rows[2][2] == datetime.datetime(2021, 6, 15, 9, 0, 0)


def test_date_read_openpyxl(tmp_xlsx):
    """Write dates with openpyxl, read with opensheet_core."""
    import datetime

    openpyxl = pytest.importorskip("openpyxl")

    wb = openpyxl.Workbook()
    ws = wb.active
    ws.append(["Date", "DateTime"])
    ws.append([datetime.date(2025, 1, 1), datetime.datetime(2025, 6, 15, 10, 30, 0)])
    wb.save(tmp_xlsx)

    rows = opensheet_core.read_sheet(tmp_xlsx)
    assert rows[0] == ["Date", "DateTime"]
    # openpyxl stores dates with format codes that our reader should detect
    assert rows[1][0] == datetime.date(2025, 1, 1) or rows[1][0] == datetime.datetime(2025, 1, 1, 0, 0, 0)
    assert rows[1][1] == datetime.datetime(2025, 6, 15, 10, 30, 0)
