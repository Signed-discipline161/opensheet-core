"""
FastAPI example: generate and download an XLSX file.

Install dependencies:
    pip install fastapi uvicorn opensheet-core

Run the server:
    uvicorn fastapi_download:app

Then open http://127.0.0.1:8000/download in your browser to download the file.
You can append ?rows=5000 to control how many data rows are generated.
"""

import os
import tempfile
from datetime import date

from fastapi import FastAPI, Query
from fastapi.responses import FileResponse

from opensheet_core import (
    CellStyle,
    FormattedCell,
    Formula,
    StyledCell,
    XlsxWriter,
)

app = FastAPI(title="opensheet-core XLSX Download Example")

# ---------------------------------------------------------------------------
# Styles
# ---------------------------------------------------------------------------

HEADER_STYLE = CellStyle(
    bold=True,
    font_size=11.0,
    font_color="FFFFFF",
    fill_color="4472C4",
    border="thin",
    border_color="2F5496",
    horizontal_alignment="center",
)

CURRENCY_FMT = "$#,##0.00"
PERCENT_FMT = "0.0%"
DATE_FMT = "yyyy-mm-dd"

# ---------------------------------------------------------------------------
# Helpers
# ---------------------------------------------------------------------------

PRODUCTS = ["Widget A", "Widget B", "Gadget X", "Gadget Y", "Gizmo Pro"]
REGIONS = ["North", "South", "East", "West"]


def _sample_rows(count: int) -> list[list]:
    """Return *count* rows of sample sales data."""
    rows = []
    for i in range(count):
        product = PRODUCTS[i % len(PRODUCTS)]
        region = REGIONS[i % len(REGIONS)]
        units = 50 + (i * 7) % 200
        price = 9.99 + (i % 10) * 5
        revenue = units * price
        tax_rate = 0.06 + (i % 5) * 0.01
        order_date = date(2025, 1, 1 + (i % 28))

        rows.append([
            i + 1,
            product,
            region,
            FormattedCell(order_date, DATE_FMT),
            units,
            FormattedCell(price, CURRENCY_FMT),
            FormattedCell(revenue, CURRENCY_FMT),
            FormattedCell(tax_rate, PERCENT_FMT),
        ])
    return rows


def _build_xlsx(path: str, row_count: int) -> None:
    """Write a styled XLSX workbook to *path*."""
    headers = [
        "Order #",
        "Product",
        "Region",
        "Date",
        "Units",
        "Unit Price",
        "Revenue",
        "Tax Rate",
    ]

    with XlsxWriter(path) as w:
        w.add_sheet("Sales Report")

        # Document properties
        w.set_document_property("title", "Sales Report")
        w.set_document_property("creator", "opensheet-core FastAPI example")

        # Column widths (set before writing rows)
        widths = [10, 14, 10, 12, 8, 12, 14, 10]
        for col_idx, width in enumerate(widths):
            w.set_column_width(col_idx, width)

        # Freeze the header row
        w.freeze_panes(row=1, col=0)

        # Header row with styling
        styled_headers = [
            StyledCell(h, HEADER_STYLE) for h in headers
        ]
        w.write_row(styled_headers)

        # Data rows (write in batches for efficiency)
        data = _sample_rows(row_count)
        w.write_rows(data)

        # Summary row with formulas
        last_row = row_count + 1  # 1-based, accounting for header
        w.write_row([
            None,
            None,
            None,
            None,
            Formula(f"SUM(E2:E{last_row})", None),
            None,
            FormattedCell(
                Formula(f"SUM(G2:G{last_row})", None),
                CURRENCY_FMT,
            ),
            None,
        ])

        # Auto-filter on the header range
        last_col_letter = "H"
        w.auto_filter(f"A1:{last_col_letter}{last_row}")

        # Structured table covering the data
        w.add_table(
            reference=f"A1:{last_col_letter}{last_row}",
            columns=headers,
            name="SalesTable",
            style="TableStyleMedium9",
        )


# ---------------------------------------------------------------------------
# Endpoint
# ---------------------------------------------------------------------------

@app.get("/download")
async def download_xlsx(
    rows: int = Query(default=100, ge=1, le=100_000, description="Number of data rows"),
):
    """Generate an XLSX file and return it as a download."""
    tmp = tempfile.NamedTemporaryFile(suffix=".xlsx", delete=False)
    tmp.close()

    try:
        _build_xlsx(tmp.name, rows)
        return FileResponse(
            path=tmp.name,
            media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            filename="sales_report.xlsx",
            # FileResponse deletes the file after sending when background=...
            # is not set; we clean up manually below via a custom approach.
        )
    except Exception:
        os.unlink(tmp.name)
        raise
