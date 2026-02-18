//! Shared protocol types for communication between the native Linux client
//! and the Windows COM bridge process running under WINE.
//!
//! The protocol is JSON-over-stdio: one JSON object per line in each direction.

use serde::{Deserialize, Serialize};

/// A command sent from the Linux client to the WINE bridge process.
#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct Request {
    /// Monotonically increasing request ID for correlating responses.
    pub id: u64,
    /// The command to execute.
    #[serde(flatten)]
    pub command: Command,
}

/// Commands the client can send to the bridge.
#[derive(Debug, Clone, Serialize, Deserialize)]
#[serde(tag = "cmd", content = "params")]
pub enum Command {
    /// Initialize COM and create the Excel.Application instance.
    Init,

    /// Create a new empty workbook. Returns a workbook handle.
    CreateWorkbook,

    /// Open an existing workbook from a file path (Windows path).
    OpenWorkbook { path: String },

    /// Set a cell's value (number, string, or bool).
    SetCellValue {
        workbook: u64,
        sheet: SheetRef,
        cell: String,
        value: CellValue,
    },

    /// Set a cell's formula (e.g., "=SUM(A1:A10)").
    SetCellFormula {
        workbook: u64,
        sheet: SheetRef,
        cell: String,
        formula: String,
    },

    /// Get a cell's computed value after recalculation.
    GetCellValue {
        workbook: u64,
        sheet: SheetRef,
        cell: String,
    },

    /// Get a cell's formula string (empty string if no formula).
    GetCellFormula {
        workbook: u64,
        sheet: SheetRef,
        cell: String,
    },

    /// Force a full recalculation of all open workbooks.
    Recalculate,

    /// Save the workbook to a file path (Windows path).
    /// Format is inferred from extension (.xlsx, .xls, .csv).
    SaveWorkbook { workbook: u64, path: String },

    /// Close a workbook without saving.
    CloseWorkbook { workbook: u64 },

    /// Shut down the bridge: close all workbooks, quit Excel, uninitialize COM.
    Shutdown,
}

/// Reference to a worksheet — by 0-based index or by name.
#[derive(Debug, Clone, Serialize, Deserialize)]
#[serde(untagged)]
pub enum SheetRef {
    Index(u32),
    Name(String),
}

/// A cell value that can be sent to/from Excel.
#[derive(Debug, Clone, Serialize, Deserialize)]
#[serde(untagged)]
pub enum CellValue {
    Null,
    Bool(bool),
    Number(f64),
    String(String),
    Error(CellError),
}

/// Excel error values.
#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct CellError {
    pub code: String,
}

/// A response sent from the WINE bridge back to the Linux client.
#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct Response {
    /// The request ID this response corresponds to.
    pub id: u64,
    /// The result of the command.
    #[serde(flatten)]
    pub result: ResponseResult,
}

#[derive(Debug, Clone, Serialize, Deserialize)]
#[serde(tag = "status")]
pub enum ResponseResult {
    #[serde(rename = "ok")]
    Ok {
        #[serde(skip_serializing_if = "Option::is_none")]
        data: Option<ResponseData>,
    },
    #[serde(rename = "error")]
    Error { message: String },
}

/// Data returned in successful responses.
#[derive(Debug, Clone, Serialize, Deserialize)]
#[serde(untagged)]
pub enum ResponseData {
    /// Handle to a newly created/opened workbook.
    WorkbookHandle { workbook: u64 },
    /// A cell value.
    Value { value: CellValue },
    /// A formula string.
    Formula { formula: String },
}

impl CellValue {
    pub fn is_null(&self) -> bool {
        matches!(self, CellValue::Null)
    }

    pub fn as_f64(&self) -> Option<f64> {
        match self {
            CellValue::Number(n) => Some(*n),
            _ => None,
        }
    }

    pub fn as_str(&self) -> Option<&str> {
        match self {
            CellValue::String(s) => Some(s),
            _ => None,
        }
    }

    pub fn as_bool(&self) -> Option<bool> {
        match self {
            CellValue::Bool(b) => Some(*b),
            _ => None,
        }
    }
}

impl From<&str> for CellValue {
    fn from(s: &str) -> Self {
        CellValue::String(s.to_string())
    }
}

impl From<String> for CellValue {
    fn from(s: String) -> Self {
        CellValue::String(s)
    }
}

impl From<f64> for CellValue {
    fn from(n: f64) -> Self {
        CellValue::Number(n)
    }
}

impl From<f32> for CellValue {
    fn from(n: f32) -> Self {
        CellValue::Number(n as f64)
    }
}

impl From<i32> for CellValue {
    fn from(n: i32) -> Self {
        CellValue::Number(n as f64)
    }
}

impl From<i64> for CellValue {
    fn from(n: i64) -> Self {
        CellValue::Number(n as f64)
    }
}

impl From<bool> for CellValue {
    fn from(b: bool) -> Self {
        CellValue::Bool(b)
    }
}

impl std::fmt::Display for CellValue {
    fn fmt(&self, f: &mut std::fmt::Formatter<'_>) -> std::fmt::Result {
        match self {
            CellValue::Null => write!(f, "<empty>"),
            CellValue::Bool(b) => write!(f, "{}", if *b { "TRUE" } else { "FALSE" }),
            CellValue::Number(n) => write!(f, "{n}"),
            CellValue::String(s) => write!(f, "{s}"),
            CellValue::Error(e) => write!(f, "{}", e.code),
        }
    }
}
