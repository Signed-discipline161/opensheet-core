# Migrating from openpyxl to OpenSheet Core

This guide provides side-by-side code comparisons for every supported operation, explains key architectural differences, and covers what is not yet supported. If you are currently using openpyxl and want faster reads, faster writes, and lower memory usage, this document will walk you through the transition.

---

## Table of Contents

- [Installation](#installation)
- [Side-by-Side Code Comparisons](#side-by-side-code-comparisons)
  - [Reading a Workbook (All Sheets)](#reading-a-workbook-all-sheets)
  - [Reading a Single Sheet](#reading-a-single-sheet)
  - [Writing a Workbook](#writing-a-workbook)
  - [Formulas](#formulas)
  - [Merging Cells](#merging-cells)
  - [Column Widths and Row Heights](#column-widths-and-row-heights)
  - [Freeze Panes](#freeze-panes)
  - [Auto-Filter](#auto-filter)
  - [Dates](#dates)
  - [Number Formats](#number-formats)
  - [Cell Styling (Fonts, Fills, Borders, Alignment)](#cell-styling-fonts-fills-borders-alignment)
  - [Named Ranges / Defined Names](#named-ranges--defined-names)
  - [Comments](#comments)
  - [Hyperlinks](#hyperlinks)
  - [Sheet Protection](#sheet-protection)
  - [Structured Tables](#structured-tables)
  - [Data Validation](#data-validation)
  - [Document Properties](#document-properties)
  - [Pandas Integration](#pandas-integration)
  - [Sheet Visibility](#sheet-visibility)
- [Key Differences](#key-differences)
- [Performance Comparison](#performance-comparison)
- [What Is Not Supported Yet](#what-is-not-supported-yet)

---

## Installation

```bash
# openpyxl
pip install openpyxl

# OpenSheet Core
pip install opensheet-core

# OpenSheet Core with pandas support
pip install opensheet-core[pandas]
```

OpenSheet Core has zero Python dependencies. The single native extension includes everything needed for XLSX read/write.

---

## Side-by-Side Code Comparisons

### Reading a Workbook (All Sheets)

**openpyxl:**

```python
from openpyxl import load_workbook

wb = load_workbook("report.xlsx")
for ws in wb.worksheets:
    print(f"Sheet: {ws.title}")
    for row in ws.iter_rows(values_only=True):
        print(list(row))
wb.close()
```

**OpenSheet Core:**

```python
from opensheet_core import read_xlsx

sheets = read_xlsx("report.xlsx")
for sheet in sheets:
    print(f"Sheet: {sheet['name']}")
    for row in sheet["rows"]:
        print(row)
```

`read_xlsx()` returns a list of dicts. Each dict contains `"name"`, `"rows"`, `"merges"`, `"column_widths"`, `"row_heights"`, `"freeze_pane"`, `"auto_filter"`, `"state"`, `"comments"`, `"hyperlinks"`, `"protection"`, and `"tables"` -- all metadata is returned in one call.

---

### Reading a Single Sheet

**openpyxl:**

```python
from openpyxl import load_workbook

wb = load_workbook("report.xlsx")
ws = wb["Data"]
for row in ws.iter_rows(values_only=True):
    print(list(row))
wb.close()
```

**OpenSheet Core:**

```python
from opensheet_core import read_sheet

rows = read_sheet("report.xlsx", sheet_name="Data")
for row in rows:
    print(row)
```

You can also read by index:

```python
rows = read_sheet("report.xlsx", sheet_index=0)  # First sheet
```

To list sheet names without reading data:

```python
from opensheet_core import sheet_names

names = sheet_names("report.xlsx")  # ["Sheet1", "Data", "Summary"]
```

---

### Writing a Workbook

**openpyxl:**

```python
from openpyxl import Workbook

wb = Workbook()
ws = wb.active
ws.title = "Data"
ws.append(["Name", "Age", "Active"])
ws.append(["Alice", 30, True])
ws.append(["Bob", 25, False])
wb.save("output.xlsx")
wb.close()
```

**OpenSheet Core:**

```python
from opensheet_core import XlsxWriter

with XlsxWriter("output.xlsx") as writer:
    writer.add_sheet("Data")
    writer.write_row(["Name", "Age", "Active"])
    writer.write_row(["Alice", 30, True])
    writer.write_row(["Bob", 25, False])
```

The context manager (`with` statement) ensures the file is finalized and closed properly. You can also write multiple rows at once for better performance:

```python
with XlsxWriter("output.xlsx") as writer:
    writer.add_sheet("Data")
    writer.write_rows([
        ["Name", "Age", "Active"],
        ["Alice", 30, True],
        ["Bob", 25, False],
    ])
```

`write_rows()` crosses the Python-to-Rust boundary only once, reducing FFI overhead for bulk writes.

---

### Formulas

**openpyxl:**

```python
from openpyxl import Workbook

wb = Workbook()
ws = wb.active
ws.append(["Item", "Cost"])
ws.append(["Rent", 1200])
ws.append(["Food", 400])
ws["B4"] = "=SUM(B2:B3)"
wb.save("output.xlsx")
```

**OpenSheet Core:**

```python
from opensheet_core import XlsxWriter, Formula

with XlsxWriter("output.xlsx") as writer:
    writer.add_sheet("Budget")
    writer.write_row(["Item", "Cost"])
    writer.write_row(["Rent", 1200])
    writer.write_row(["Food", 400])
    writer.write_row(["Total", Formula("SUM(B2:B3)", cached_value=1600)])
```

Note: formulas do not include a leading `=` sign. The `cached_value` parameter is optional and provides a pre-computed result so spreadsheet viewers can display a value without recalculating.

**Reading formulas back:**

```python
# openpyxl
wb = load_workbook("output.xlsx")
cell = wb.active["B4"]
print(cell.value)  # "=SUM(B2:B3)"

# OpenSheet Core
rows = read_sheet("output.xlsx")
cell = rows[3][1]  # Row 4, column B (0-indexed)
print(cell.formula)        # "SUM(B2:B3)"
print(cell.cached_value)   # 1600
```

---

### Merging Cells

**openpyxl:**

```python
from openpyxl import Workbook

wb = Workbook()
ws = wb.active
ws.merge_cells("A1:C1")
ws["A1"] = "Title spanning three columns"
ws.append(["", "", ""])  # Row 1 already written via cell assignment
ws.append(["A", "B", "C"])
wb.save("output.xlsx")
```

**OpenSheet Core:**

```python
from opensheet_core import XlsxWriter

with XlsxWriter("output.xlsx") as writer:
    writer.add_sheet("Report")
    writer.write_row(["Title spanning three columns", "", ""])
    writer.write_row(["A", "B", "C"])
    writer.merge_cells("A1:C1")
```

**Reading merges back:**

```python
sheets = read_xlsx("output.xlsx")
print(sheets[0]["merges"])  # ["A1:C1"]
```

---

### Column Widths and Row Heights

**openpyxl:**

```python
from openpyxl import Workbook

wb = Workbook()
ws = wb.active
ws.column_dimensions["A"].width = 25.0
ws.column_dimensions["B"].width = 15.0
ws.row_dimensions[1].height = 30.0
ws.append(["Name", "Age"])
ws.append(["Alice", 30])
wb.save("output.xlsx")
```

**OpenSheet Core:**

```python
from opensheet_core import XlsxWriter

with XlsxWriter("output.xlsx") as writer:
    writer.add_sheet("Data")
    writer.set_column_width("A", 25.0)   # By letter
    writer.set_column_width(1, 15.0)     # By 0-based index (column B)
    writer.set_row_height(1, 30.0)       # Row 1, 1-based
    writer.write_row(["Name", "Age"])
    writer.write_row(["Alice", 30])
```

Column widths must be set after `add_sheet()` but before `write_row()` calls. Row heights can be set at any point.

**Reading back:**

```python
sheets = read_xlsx("output.xlsx")
print(sheets[0]["column_widths"])  # {0: 25.0, 1: 15.0}
print(sheets[0]["row_heights"])    # {0: 30.0}
```

---

### Freeze Panes

**openpyxl:**

```python
from openpyxl import Workbook

wb = Workbook()
ws = wb.active
ws.freeze_panes = "A2"  # Freeze top row
ws.append(["Header1", "Header2", "Header3"])
ws.append(["data", "data", "data"])
wb.save("output.xlsx")
```

**OpenSheet Core:**

```python
from opensheet_core import XlsxWriter

with XlsxWriter("output.xlsx") as writer:
    writer.add_sheet("Data")
    writer.freeze_panes(row=1, col=0)  # Freeze top row
    writer.write_row(["Header1", "Header2", "Header3"])
    writer.write_row(["data", "data", "data"])
```

openpyxl uses a cell reference string (e.g. `"A2"`) to indicate where the freeze starts. OpenSheet Core uses explicit `row` and `col` integer counts. `freeze_panes(row=1, col=0)` freezes the first row. `freeze_panes(row=1, col=1)` freezes the first row and first column.

**Reading back:**

```python
sheets = read_xlsx("output.xlsx")
print(sheets[0]["freeze_pane"])  # (1, 0)
```

---

### Auto-Filter

**openpyxl:**

```python
from openpyxl import Workbook

wb = Workbook()
ws = wb.active
ws.append(["Name", "Age", "City"])
ws.append(["Alice", 30, "NYC"])
ws.append(["Bob", 25, "LA"])
ws.auto_filter.ref = "A1:C3"
wb.save("output.xlsx")
```

**OpenSheet Core:**

```python
from opensheet_core import XlsxWriter

with XlsxWriter("output.xlsx") as writer:
    writer.add_sheet("Data")
    writer.write_row(["Name", "Age", "City"])
    writer.write_row(["Alice", 30, "NYC"])
    writer.write_row(["Bob", 25, "LA"])
    writer.auto_filter("A1:C1")
```

**Reading back:**

```python
sheets = read_xlsx("output.xlsx")
print(sheets[0]["auto_filter"])  # "A1:C1"
```

---

### Dates

**openpyxl:**

```python
import datetime
from openpyxl import Workbook

wb = Workbook()
ws = wb.active
ws.append(["Event", "Date", "Timestamp"])
ws.append(["Launch", datetime.date(2025, 3, 15), datetime.datetime(2025, 3, 15, 14, 30)])
# openpyxl requires manual number format for date display
ws["B2"].number_format = "YYYY-MM-DD"
ws["C2"].number_format = "YYYY-MM-DD HH:MM:SS"
wb.save("output.xlsx")
```

**OpenSheet Core:**

```python
import datetime
from opensheet_core import XlsxWriter

with XlsxWriter("output.xlsx") as writer:
    writer.add_sheet("Events")
    writer.write_row(["Event", "Date", "Timestamp"])
    writer.write_row([
        "Launch",
        datetime.date(2025, 3, 15),
        datetime.datetime(2025, 3, 15, 14, 30),
    ])
```

OpenSheet Core automatically applies appropriate date/datetime number formats when it detects `datetime.date` or `datetime.datetime` values. No manual number format assignment is needed.

**Reading dates back:**

Dates are returned as native `datetime.date` or `datetime.datetime` Python objects. There is no need to check cell types or parse serial numbers.

```python
rows = read_sheet("output.xlsx")
print(type(rows[1][1]))  # <class 'datetime.date'>
print(type(rows[1][2]))  # <class 'datetime.datetime'>
```

---

### Number Formats

**openpyxl:**

```python
from openpyxl import Workbook

wb = Workbook()
ws = wb.active
ws.append(["Item", "Price", "Tax Rate"])
ws["B2"] = 19.99
ws["B2"].number_format = "$#,##0.00"
ws["C2"] = 0.08
ws["C2"].number_format = "0.00%"
wb.save("output.xlsx")
```

**OpenSheet Core:**

```python
from opensheet_core import XlsxWriter, FormattedCell

with XlsxWriter("output.xlsx") as writer:
    writer.add_sheet("Finance")
    writer.write_row(["Item", "Price", "Tax Rate"])
    writer.write_row([
        "Widget",
        FormattedCell(19.99, "$#,##0.00"),   # Currency
        FormattedCell(0.08, "0.00%"),         # Percentage
    ])
```

In openpyxl, you write the value first, then set the `number_format` property on the cell object. In OpenSheet Core, value and format are bundled together in a `FormattedCell` since there is no random cell access -- everything is written row-by-row.

**Reading back:**

```python
rows = read_sheet("output.xlsx")
cell = rows[1][1]  # FormattedCell object
print(cell.value)          # 19.99
print(cell.number_format)  # "$#,##0.00"
```

---

### Cell Styling (Fonts, Fills, Borders, Alignment)

**openpyxl:**

```python
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Border, Side, Alignment

wb = Workbook()
ws = wb.active

# Header row with bold white text on blue background
header_font = Font(bold=True, color="FFFFFF", name="Arial", size=12)
header_fill = PatternFill(start_color="4472C4", end_color="4472C4", fill_type="solid")
for col, val in enumerate(["Name", "Score"], 1):
    cell = ws.cell(row=1, column=col, value=val)
    cell.font = header_font
    cell.fill = header_fill

# Data row with borders and alignment
thin_border = Border(
    left=Side(style="thin"),
    right=Side(style="thin"),
    top=Side(style="thin"),
    bottom=Side(style="thin"),
)
ws.cell(row=2, column=1, value="Alice").border = thin_border
ws.cell(row=2, column=1).alignment = Alignment(horizontal="left")
cell_b2 = ws.cell(row=2, column=2, value=95)
cell_b2.border = thin_border
cell_b2.number_format = "0.0"

wb.save("output.xlsx")
```

**OpenSheet Core:**

```python
from opensheet_core import XlsxWriter, CellStyle, StyledCell

with XlsxWriter("output.xlsx") as writer:
    writer.add_sheet("Report")

    # Header row with bold white text on blue background
    header = CellStyle(
        bold=True,
        font_color="FFFFFF",
        font_name="Arial",
        font_size=12.0,
        fill_color="4472C4",
    )
    writer.write_row([
        StyledCell("Name", header),
        StyledCell("Score", header),
    ])

    # Data row with borders and alignment
    writer.write_row([
        StyledCell("Alice", CellStyle(border="thin", horizontal_alignment="left")),
        StyledCell(95, CellStyle(border="thin", number_format="0.0")),
    ])
```

Key mapping from openpyxl style objects to `CellStyle` parameters:

| openpyxl | OpenSheet Core `CellStyle` |
|----------|---------------------------|
| `Font(bold=True)` | `bold=True` |
| `Font(italic=True)` | `italic=True` |
| `Font(underline="single")` | `underline=True` |
| `Font(name="Arial")` | `font_name="Arial"` |
| `Font(size=12)` | `font_size=12.0` |
| `Font(color="FF0000")` | `font_color="FF0000"` |
| `PatternFill(start_color="4472C4", fill_type="solid")` | `fill_color="4472C4"` |
| `Border(left=Side(style="thin"), ...)` | `border="thin"` (all sides) or `border_left="thin"`, etc. |
| `Side(color="000000")` | `border_color="000000"` |
| `Alignment(horizontal="center")` | `horizontal_alignment="center"` |
| `Alignment(vertical="top")` | `vertical_alignment="top"` |
| `Alignment(wrap_text=True)` | `wrap_text=True` |
| `Alignment(text_rotation=45)` | `text_rotation=45` |
| `cell.number_format = "0.00%"` | `number_format="0.00%"` (on `CellStyle`) |

The `border` shorthand sets all four sides at once. For per-side control, use `border_left`, `border_right`, `border_top`, `border_bottom`. Supported border styles: `"thin"`, `"medium"`, `"thick"`, `"dashed"`, `"dotted"`, `"double"`.

---

### Named Ranges / Defined Names

**openpyxl:**

```python
from openpyxl import Workbook
from openpyxl.workbook.defined_name import DefinedName

wb = Workbook()
ws = wb.active
ws.title = "Config"
ws.append(["Rate"])
ws.append([0.08])

# Workbook-scoped name
ref = DefinedName("TaxRate", attr_text="Config!$A$2")
wb.defined_names.add(ref)

# Sheet-scoped name
ref = DefinedName("LocalRate", attr_text="Config!$A$2")
ref.localSheetId = 0
wb.defined_names.add(ref)

wb.save("output.xlsx")
```

**OpenSheet Core:**

```python
from opensheet_core import XlsxWriter

with XlsxWriter("output.xlsx") as writer:
    writer.add_sheet("Config")
    writer.write_row(["Rate"])
    writer.write_row([0.08])
    writer.define_name("TaxRate", "Config!$A$2")                    # Workbook-scoped
    writer.define_name("LocalRate", "Config!$A$2", sheet_index=0)   # Sheet-scoped
```

**Reading named ranges:**

```python
# openpyxl
wb = load_workbook("output.xlsx")
for name in wb.defined_names.definedName:
    print(name.name, name.attr_text, name.localSheetId)

# OpenSheet Core
from opensheet_core import defined_names

names = defined_names("output.xlsx")
for n in names:
    print(f"{n['name']} -> {n['value']} (sheet_index={n['sheet_index']})")
```

---

### Comments

**openpyxl:**

```python
from openpyxl import Workbook
from openpyxl.comments import Comment

wb = Workbook()
ws = wb.active
ws.append(["Name", "Website"])
ws.append(["Alice", "https://example.com"])
ws["A1"].comment = Comment("Primary contact", "Admin")
wb.save("output.xlsx")
```

**OpenSheet Core:**

```python
from opensheet_core import XlsxWriter

with XlsxWriter("output.xlsx") as writer:
    writer.add_sheet("Data")
    writer.write_row(["Name", "Website"])
    writer.write_row(["Alice", "https://example.com"])
    writer.add_comment("A1", "Admin", "Primary contact")
```

Note the parameter order difference: openpyxl's `Comment(text, author)` takes text first, while OpenSheet Core's `add_comment(cell_ref, author, text)` takes author first.

**Reading comments back:**

```python
sheets = read_xlsx("output.xlsx")
for comment in sheets[0]["comments"]:
    print(f"Cell {comment['cell']}: {comment['author']} - {comment['text']}")
```

---

### Hyperlinks

**openpyxl:**

```python
from openpyxl import Workbook

wb = Workbook()
ws = wb.active
ws.append(["Name", "Website"])
ws.append(["Alice", "https://example.com"])
ws["B2"].hyperlink = "https://example.com"
ws["B2"].style = "Hyperlink"
wb.save("output.xlsx")
```

**OpenSheet Core:**

```python
from opensheet_core import XlsxWriter

with XlsxWriter("output.xlsx") as writer:
    writer.add_sheet("Data")
    writer.write_row(["Name", "Website"])
    writer.write_row(["Alice", "https://example.com"])
    writer.add_hyperlink("B2", "https://example.com", tooltip="Visit site")
```

**Reading hyperlinks back:**

```python
sheets = read_xlsx("output.xlsx")
for link in sheets[0]["hyperlinks"]:
    print(f"Cell {link['cell']}: {link['url']} ({link['tooltip']})")
```

---

### Sheet Protection

**openpyxl:**

```python
from openpyxl import Workbook

wb = Workbook()
ws = wb.active
ws.append(["Locked data"])
ws.protection.sheet = True
ws.protection.password = "secret"
ws.protection.sort = True
ws.protection.autoFilter = True
wb.save("output.xlsx")
```

**OpenSheet Core:**

```python
from opensheet_core import XlsxWriter

with XlsxWriter("output.xlsx") as writer:
    writer.add_sheet("Protected")
    writer.write_row(["Locked data"])
    writer.protect_sheet(
        password="secret",
        sheet=True,
        sort=True,
        auto_filter=True,
    )
```

OpenSheet Core supports 15+ permission flags on `protect_sheet()`: `sheet`, `objects`, `scenarios`, `format_cells`, `format_columns`, `format_rows`, `insert_columns`, `insert_rows`, `insert_hyperlinks`, `delete_columns`, `delete_rows`, `sort`, `auto_filter`, `pivot_tables`, `select_locked_cells`, `select_unlocked_cells`.

**Reading protection back:**

```python
sheets = read_xlsx("output.xlsx")
print(sheets[0]["protection"])  # dict of protection settings, or None
```

---

### Structured Tables

**openpyxl:**

```python
from openpyxl import Workbook
from openpyxl.worksheet.table import Table, TableStyleInfo

wb = Workbook()
ws = wb.active
ws.append(["Name", "Age", "City"])
ws.append(["Alice", 30, "NYC"])
ws.append(["Bob", 25, "LA"])

table = Table(displayName="People", ref="A1:C3")
table.tableStyleInfo = TableStyleInfo(name="TableStyleMedium2")
ws.add_table(table)
wb.save("output.xlsx")
```

**OpenSheet Core:**

```python
from opensheet_core import XlsxWriter

with XlsxWriter("output.xlsx") as writer:
    writer.add_sheet("Data")
    writer.write_row(["Name", "Age", "City"])
    writer.write_row(["Alice", 30, "NYC"])
    writer.write_row(["Bob", 25, "LA"])
    writer.add_table(
        "A1:C3",
        ["Name", "Age", "City"],
        name="People",
        style="TableStyleMedium2",
    )
```

**Reading tables back:**

```python
sheets = read_xlsx("output.xlsx")
for table in sheets[0]["tables"]:
    print(table)  # dict with table definition
```

---

### Data Validation

**openpyxl:**

```python
from openpyxl import Workbook
from openpyxl.worksheet.datavalidation import DataValidation

wb = Workbook()
ws = wb.active
ws.append(["Status"])

# Drop-down list validation
dv = DataValidation(type="list", formula1='"Open,Closed,Pending"', allow_blank=True)
dv.sqref = "A2:A100"
dv.prompt = "Choose a status"
dv.promptTitle = "Status"
dv.error = "Invalid status"
dv.errorTitle = "Error"
ws.add_data_validation(dv)
wb.save("output.xlsx")
```

**OpenSheet Core:**

```python
from opensheet_core import XlsxWriter

with XlsxWriter("output.xlsx") as writer:
    writer.add_sheet("Data")
    writer.write_row(["Status"])
    writer.add_data_validation(
        validation_type="list",
        sqref="A2:A100",
        formula1='"Open,Closed,Pending"',
        allow_blank=True,
        show_input_message=True,
        prompt_title="Status",
        prompt="Choose a status",
        show_error_message=True,
        error_title="Error",
        error_message="Invalid status",
    )
```

Supported validation types: `"list"`, `"whole"`, `"decimal"`, `"date"`, `"time"`, `"textLength"`, `"custom"`.

**Numeric range validation example:**

```python
# openpyxl
dv = DataValidation(type="whole", operator="between", formula1="1", formula2="100")
dv.sqref = "B2:B100"
ws.add_data_validation(dv)

# OpenSheet Core
writer.add_data_validation(
    validation_type="whole",
    sqref="B2:B100",
    operator="between",
    formula1="1",
    formula2="100",
    show_error_message=True,
    error_message="Enter a number between 1 and 100",
)
```

---

### Document Properties

**openpyxl:**

```python
from openpyxl import Workbook

wb = Workbook()
wb.properties.title = "Quarterly Report"
wb.properties.creator = "Finance Team"
wb.properties.subject = "Q1 2025"
wb.properties.category = "Reports"
# Custom properties require a different mechanism in openpyxl
wb.save("output.xlsx")
```

**OpenSheet Core:**

```python
from opensheet_core import XlsxWriter

with XlsxWriter("output.xlsx") as writer:
    writer.add_sheet("Data")
    writer.set_document_property("title", "Quarterly Report")
    writer.set_document_property("creator", "Finance Team")
    writer.set_document_property("subject", "Q1 2025")
    writer.set_document_property("category", "Reports")
    writer.set_custom_property("Department", "Finance")
    writer.set_custom_property("Reviewed", "Yes")
    writer.write_row(["Data goes here"])
```

Valid core property keys: `"title"`, `"subject"`, `"creator"`, `"keywords"`, `"description"`, `"last_modified_by"`, `"category"`.

**Reading properties back:**

```python
from opensheet_core import document_properties

props = document_properties("output.xlsx")
print(props["core"]["title"])     # "Quarterly Report"
print(props["core"]["creator"])   # "Finance Team"
for custom in props["custom"]:
    print(f"{custom['name']}: {custom['value']}")
```

---

### Pandas Integration

**openpyxl:**

```python
import pandas as pd

# Reading
df = pd.read_excel("data.xlsx", sheet_name="Sheet1", engine="openpyxl")

# Writing
df = pd.DataFrame({"Name": ["Alice", "Bob"], "Age": [30, 25]})
df.to_excel("output.xlsx", sheet_name="Results", index=False, engine="openpyxl")
```

**OpenSheet Core:**

```python
from opensheet_core import read_xlsx_df, to_xlsx

# Reading
df = read_xlsx_df("data.xlsx", sheet_name="Sheet1")

# Writing
import pandas as pd
df = pd.DataFrame({"Name": ["Alice", "Bob"], "Age": [30, 25]})
to_xlsx(df, "output.xlsx", sheet_name="Results")
```

Install with `pip install opensheet-core[pandas]` to enable pandas support.

Key differences:
- `read_xlsx_df()` automatically unwraps `Formula`, `FormattedCell`, and `StyledCell` values to plain Python types.
- `to_xlsx()` handles NumPy types (`int64`, `float64`, `bool_`), `NaN`/`NaT` (written as empty cells), and `Timestamp`/`datetime64` columns.
- Set `header=False` to skip column headers. Set `index=True` to include the DataFrame index.

---

### Sheet Visibility

**openpyxl:**

```python
from openpyxl import Workbook

wb = Workbook()
ws1 = wb.active
ws1.title = "Visible"
ws2 = wb.create_sheet("Hidden")
ws2.sheet_state = "hidden"
ws3 = wb.create_sheet("VeryHidden")
ws3.sheet_state = "veryHidden"
wb.save("output.xlsx")
```

**OpenSheet Core:**

```python
from opensheet_core import XlsxWriter

with XlsxWriter("output.xlsx") as writer:
    writer.add_sheet("Visible")
    writer.write_row(["This sheet is visible"])

    writer.add_sheet("Hidden")
    writer.set_sheet_state("hidden")
    writer.write_row(["This sheet is hidden"])

    writer.add_sheet("VeryHidden")
    writer.set_sheet_state("veryHidden")
    writer.write_row(["This sheet is very hidden"])
```

Valid states: `"visible"` (default), `"hidden"`, `"veryHidden"`.

**Reading visibility back:**

```python
from opensheet_core import read_xlsx

sheets = read_xlsx("output.xlsx")
for sheet in sheets:
    print(f"{sheet['name']}: {sheet['state']}")
# Visible: visible
# Hidden: hidden
# VeryHidden: veryHidden
```

---

## Key Differences

### Streaming by Default vs Opt-In

In openpyxl, you must explicitly enable streaming mode with `read_only=True` or `write_only=True`. The default mode loads the entire workbook into memory as a DOM.

```python
# openpyxl: opt-in streaming
wb = load_workbook("large.xlsx", read_only=True)  # Streaming read
wb = Workbook(write_only=True)                     # Streaming write
```

In OpenSheet Core, streaming is the default and only mode. The Rust core uses a SAX-style parser for reads and streams data directly to disk for writes. There is no DOM mode and no way to load an entire workbook into memory as a mutable object tree.

### Row-by-Row vs Random Cell Access

openpyxl supports random access to any cell by reference:

```python
ws["B2"] = 42
ws.cell(row=3, column=4, value="hello")
value = ws["A1"].value
```

OpenSheet Core writes data sequentially, one row at a time:

```python
writer.write_row(["col1", "col2", "col3"])  # Row 1
writer.write_row([None, 42, None])           # Row 2
```

You cannot go back and modify a previously written row. If you need to set a value in a specific cell, you must plan your row layout in advance. For reads, data is returned as a complete list of rows -- there is no cell-by-cell random access.

### Context Manager Pattern

OpenSheet Core requires (or strongly encourages) use of the `with` statement for writing:

```python
with XlsxWriter("output.xlsx") as writer:
    writer.add_sheet("Sheet1")
    writer.write_row(["data"])
# File is finalized and closed here
```

Without the context manager, you must call `writer.close()` explicitly. Failing to close the writer will produce an incomplete or corrupt file.

### Zero Dependencies

openpyxl depends on `et-xmlfile` (and optionally `lxml` for performance). OpenSheet Core has zero Python dependencies. The entire implementation is a single compiled native extension built with Rust and PyO3.

### Type Handling Differences

| Data type | openpyxl read | OpenSheet Core read |
|-----------|--------------|---------------------|
| String | `str` | `str` |
| Integer | `int` | `int` |
| Float | `float` | `float` |
| Boolean | `bool` | `bool` |
| Date | `datetime.datetime` | `datetime.date` |
| DateTime | `datetime.datetime` | `datetime.datetime` |
| Formula | `str` (prefixed with `=`) | `Formula` object |
| Formatted number | `float` or `int` | `FormattedCell` object |
| Styled cell | value only (style on cell object) | `StyledCell` object |
| Empty | `None` | `None` |

Notable differences:
- **Dates**: openpyxl returns all date-like values as `datetime.datetime`. OpenSheet Core distinguishes between `datetime.date` (date-only) and `datetime.datetime` (date with time).
- **Formulas**: openpyxl returns formulas as strings with a leading `=`. OpenSheet Core returns `Formula` objects with `.formula` (no `=` prefix) and `.cached_value` attributes.
- **Formatting**: openpyxl keeps number format and style as properties on the cell object. OpenSheet Core bundles them with the value as `FormattedCell` or `StyledCell` wrapper objects.

---

## Performance Comparison

Benchmarked against openpyxl 3.1.5 on a 100,000-row x 10-column dataset (1 million cells), 5 interleaved runs, current RSS measurement:

| Operation | OpenSheet Core | openpyxl | Speedup | Memory (RSS delta) |
|-----------|---------------|----------|---------|---------------------|
| **Write** | 2.3s | 3.7s | **1.6x faster** | **1.7x less** (1.2 MB vs 2.1 MB) |
| **Read** | 253ms | 3.5s | **13.8x faster** | **2.5x less** (13.5 MB vs 33.3 MB) |

### Why is it faster?

- **Rust streaming core**: The parser uses SAX-style XML processing. No DOM tree is ever constructed.
- **Deferred shared-string resolution**: Shared strings are stored as integer indices during parsing and only converted to Python objects at the boundary, avoiding duplicate string allocations.
- **Direct-to-disk writes**: The writer streams compressed XML directly to disk. Memory usage stays constant regardless of row count.
- **Single FFI boundary crossing**: `write_rows()` sends all row data to Rust in a single call, minimizing Python-to-Rust overhead.

### Reproducing the benchmarks

```bash
python benchmarks/benchmark.py
```

See [docs/benchmarking.md](benchmarking.md) for the full methodology, including how measurement avoids common pitfalls like high-water-mark RSS reporting.

---

## What Is Not Supported Yet

The following openpyxl features are not yet available in OpenSheet Core. Check the [roadmap](https://github.com/0xNadr/opensheet-core#roadmap) for planned additions.

| Feature | Status |
|---------|--------|
| Charts (bar, line, pie, scatter, etc.) | Planned |
| Image embedding (PNG, JPEG) | Planned |
| Conditional formatting (color scales, data bars, icon sets) | Planned |
| Pattern fills (gradient, hatching) | Not planned -- solid fills only |
| Named styles (reusable style presets) | Planned |
| In-place editing (open, modify, save) | Not supported -- write-from-scratch only |
| .xltx / .xltm template files | Not planned |
| .xlsm write support | Not planned -- read-only |
| Rich text within a single cell | Planned |
| Row/column insert/delete | Not supported |
| Row/column grouping (outline) | Not supported |
| Pivot tables | Not supported |
| Print settings (page setup, headers/footers) | Planned |
| Workbook-level protection | Not supported |

### Workaround for in-place editing

OpenSheet Core does not support opening an existing file, modifying cells, and saving it back (openpyxl's primary workflow). Instead, read the data, transform it in Python, and write a new file:

```python
from opensheet_core import read_xlsx, XlsxWriter

# Read existing data
sheets = read_xlsx("input.xlsx")

# Write modified data to a new file
with XlsxWriter("output.xlsx") as writer:
    for sheet in sheets:
        writer.add_sheet(sheet["name"])
        for row in sheet["rows"]:
            # Apply transformations as needed
            writer.write_row(row)
```

This read-transform-write pattern works well for ETL pipelines and report generation, which are the primary use cases OpenSheet Core is optimized for.
