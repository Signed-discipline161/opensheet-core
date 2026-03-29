use pyo3::prelude::*;
use pyo3::types::{PyList, PyDict, PyNone};
use std::fs::File;
use std::io::BufReader;

mod reader;
mod types;

use types::CellValue;

/// Convert a CellValue to a Python object.
fn cell_to_py(py: Python<'_>, cell: &CellValue) -> PyResult<Py<PyAny>> {
    match cell {
        CellValue::String(s) => Ok(s.into_pyobject(py)?.into_any().unbind()),
        CellValue::Number(n) => {
            if n.fract() == 0.0 && n.abs() < i64::MAX as f64 {
                Ok((*n as i64).into_pyobject(py)?.into_any().unbind())
            } else {
                Ok(n.into_pyobject(py)?.into_any().unbind())
            }
        }
        CellValue::Bool(b) => Ok(b.into_pyobject(py)?.to_owned().into_any().unbind()),
        CellValue::Empty => Ok(PyNone::get(py).to_owned().into_any().unbind()),
    }
}

/// Convert rows to a Python list of lists.
fn rows_to_py(py: Python<'_>, rows: &[Vec<CellValue>]) -> PyResult<Py<PyAny>> {
    let py_rows = PyList::empty(py);
    for row in rows {
        let py_row = PyList::empty(py);
        for cell in row {
            py_row.append(cell_to_py(py, cell)?)?;
        }
        py_rows.append(py_row)?;
    }
    Ok(py_rows.into_any().unbind())
}

/// Returns version information about the native core.
#[pyfunction]
fn version() -> &'static str {
    env!("CARGO_PKG_VERSION")
}

/// Read an XLSX file and return a list of sheets.
///
/// Each sheet is a dict with:
///   - "name": sheet name (str)
///   - "rows": list of lists of cell values
#[pyfunction]
fn read_xlsx(py: Python<'_>, path: &str) -> PyResult<Py<PyAny>> {
    let file = File::open(path)
        .map_err(|e| pyo3::exceptions::PyFileNotFoundError::new_err(format!("{path}: {e}")))?;
    let reader = BufReader::new(file);
    let sheets = reader::xlsx::read_xlsx(reader)?;

    let result = PyList::empty(py);
    for sheet in sheets {
        let dict = PyDict::new(py);
        dict.set_item("name", &sheet.name)?;
        dict.set_item("rows", rows_to_py(py, &sheet.rows)?)?;
        result.append(dict)?;
    }

    Ok(result.into_any().unbind())
}

/// Read a specific sheet by name or index from an XLSX file.
///
/// Returns a list of rows (list of lists of cell values).
#[pyfunction]
#[pyo3(signature = (path, sheet_name=None, sheet_index=None))]
fn read_sheet(
    py: Python<'_>,
    path: &str,
    sheet_name: Option<&str>,
    sheet_index: Option<usize>,
) -> PyResult<Py<PyAny>> {
    let file = File::open(path)
        .map_err(|e| pyo3::exceptions::PyFileNotFoundError::new_err(format!("{path}: {e}")))?;
    let reader = BufReader::new(file);
    let sheets = reader::xlsx::read_xlsx(reader)?;

    let sheet = if let Some(name) = sheet_name {
        sheets
            .iter()
            .find(|s| s.name == name)
            .ok_or_else(|| {
                pyo3::exceptions::PyValueError::new_err(format!("Sheet '{name}' not found"))
            })?
    } else if let Some(idx) = sheet_index {
        sheets.get(idx).ok_or_else(|| {
            pyo3::exceptions::PyValueError::new_err(format!(
                "Sheet index {idx} out of range (file has {} sheets)",
                sheets.len()
            ))
        })?
    } else {
        sheets.first().ok_or_else(|| {
            pyo3::exceptions::PyValueError::new_err("No sheets found in file")
        })?
    };

    rows_to_py(py, &sheet.rows)
}

/// List sheet names in an XLSX file.
#[pyfunction]
fn sheet_names(path: &str) -> PyResult<Vec<String>> {
    let file = File::open(path)
        .map_err(|e| pyo3::exceptions::PyFileNotFoundError::new_err(format!("{path}: {e}")))?;
    let reader = BufReader::new(file);
    let sheets = reader::xlsx::read_xlsx(reader)?;
    Ok(sheets.iter().map(|s| s.name.clone()).collect())
}

/// A Python module implemented in Rust.
#[pymodule]
fn _native(m: &Bound<'_, PyModule>) -> PyResult<()> {
    m.add_function(wrap_pyfunction!(version, m)?)?;
    m.add_function(wrap_pyfunction!(read_xlsx, m)?)?;
    m.add_function(wrap_pyfunction!(read_sheet, m)?)?;
    m.add_function(wrap_pyfunction!(sheet_names, m)?)?;
    Ok(())
}
