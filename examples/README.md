# Examples

Runnable examples that show how to use `opensheet-core` in common scenarios.

## Files

| File | Description |
|---|---|
| `fastapi_download.py` | Generate a styled XLSX file and serve it as a download with FastAPI. |
| `flask_download.py` | Same concept using Flask and `send_file`. |

## Quick start

### FastAPI

```bash
pip install fastapi uvicorn opensheet-core
uvicorn fastapi_download:app
# Open http://127.0.0.1:8000/download
# Append ?rows=5000 to control the number of data rows
```

### Flask

```bash
pip install flask opensheet-core
python flask_download.py
# Open http://127.0.0.1:5000/download
# Append ?rows=5000 to control the number of data rows
```

## What the examples demonstrate

- Writing styled header rows with `StyledCell` and `CellStyle` (bold, colors, borders, alignment).
- Number formatting with `FormattedCell` (currency, percentages, dates).
- Summary formulas with `Formula`.
- Column widths, frozen panes, auto-filters, and structured tables.
- Document properties via `set_document_property`.
- Efficient bulk writes with `write_rows`.
- Proper temp-file handling for web server downloads.
