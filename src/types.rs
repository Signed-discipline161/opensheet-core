/// Represents a single cell value from a spreadsheet.
#[derive(Debug, Clone)]
pub enum CellValue {
    String(std::string::String),
    Number(f64),
    Bool(bool),
    Empty,
}

/// A worksheet with a name and rows of cell values.
#[derive(Debug, Clone)]
pub struct Sheet {
    pub name: std::string::String,
    pub rows: Vec<Vec<CellValue>>,
}
