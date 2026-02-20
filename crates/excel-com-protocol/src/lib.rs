//! Generic COM proxy protocol for remote Excel automation.
//!
//! Instead of having a specific command for every Excel operation, this protocol
//! provides three generic primitives — `Get`, `Set`, and `Invoke` — that navigate
//! a chain of COM object properties from a stored handle. This means the bridge
//! server is a thin, stable COM proxy that never needs modification when new
//! Excel features are added. All Excel-specific knowledge lives in the client.
//!
//! ## Wire format
//!
//! Newline-delimited JSON (NDJSON), one object per line in each direction over TCP.
//!
//! ## Handles
//!
//! The server maintains a handle table of COM object references:
//! - Handle 0 = `Excel.Application` (available after `Init`)
//! - Handles 1+ = workbooks, worksheets, ranges, etc. (allocated on demand)
//!
//! When an `Invoke` returns a COM object, it is automatically stored and a handle
//! is returned. Clients should `Release` handles when done.
//!
//! ## Chain navigation
//!
//! A chain is a series of steps to navigate from a stored handle to a target object.
//! Each step is either:
//! - A property access: `"PropertyName"` (string)
//! - An indexed access: `["PropertyName", index]` (array with name + index value)
//!
//! Examples:
//! ```json
//! // app.Workbooks
//! {"handle": 0, "chain": ["Workbooks"]}
//!
//! // workbook.Worksheets[1].Range["A1"]
//! {"handle": 1, "chain": [["Worksheets", 1], ["Range", "A1"]]}
//!
//! // workbook.Worksheets[1].Range["A1"].Font
//! {"handle": 1, "chain": [["Worksheets", 1], ["Range", "A1"], "Font"]}
//! ```

use serde::{Deserialize, Serialize};

// ---------------------------------------------------------------------------
// Request
// ---------------------------------------------------------------------------

/// A request sent from the client to the bridge server.
#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct Request {
    /// Monotonically increasing request ID for correlating responses.
    pub id: u64,
    /// The command to execute.
    #[serde(flatten)]
    pub command: Command,
}

/// Commands the client can send.
///
/// Uses `#[serde(tag = "cmd", content = "params")]` so the wire format is:
/// `{"id": 1, "cmd": "Get", "params": {...}}`
#[derive(Debug, Clone, Serialize, Deserialize)]
#[serde(tag = "cmd", content = "params")]
pub enum Command {
    /// Initialize COM and create the Excel.Application instance.
    /// The Application is stored as handle 0.
    Init,

    /// Get a property value from a COM object.
    ///
    /// Navigate `chain` from the object at `handle`, then read `property`.
    /// If the property value is a COM object, it is stored and a handle is returned.
    Get {
        handle: u64,
        #[serde(default)]
        chain: Vec<ChainStep>,
        property: String,
    },

    /// Set a property value on a COM object.
    ///
    /// Navigate `chain` from the object at `handle`, then set `property` to `value`.
    Set {
        handle: u64,
        #[serde(default)]
        chain: Vec<ChainStep>,
        property: String,
        value: serde_json::Value,
    },

    /// Invoke a method on a COM object.
    ///
    /// Navigate `chain` from the object at `handle`, then call `method` with `args`.
    /// If the return value is a COM object, it is stored and a handle is returned.
    Invoke {
        handle: u64,
        #[serde(default)]
        chain: Vec<ChainStep>,
        method: String,
        #[serde(default)]
        args: Vec<serde_json::Value>,
    },

    /// Release a stored COM object handle.
    ///
    /// Frees the server-side reference. The handle becomes invalid.
    /// Handle 0 (Excel.Application) cannot be released — use `Shutdown` instead.
    Release { handle: u64 },

    /// Shut down: release all handles, quit Excel, uninitialize COM.
    Shutdown,
}

/// A step in navigating a COM object chain.
///
/// Wire format (untagged):
/// - Simple property: `"PropertyName"` (a JSON string)
/// - Indexed property: `["PropertyName", index]` (a JSON array)
///
/// The index can be a number (for 1-based Excel collection indexing) or
/// a string (for named access like `Range["A1"]`).
#[derive(Debug, Clone, Serialize, Deserialize)]
#[serde(untagged)]
pub enum ChainStep {
    /// Simple property access: `obj.PropertyName`
    Property(String),
    /// Indexed property access: `obj.PropertyName(index)`
    /// The tuple is `(property_name, index_value)`.
    Indexed(String, serde_json::Value),
}

// ---------------------------------------------------------------------------
// Response
// ---------------------------------------------------------------------------

/// A response sent from the bridge server to the client.
#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct Response {
    /// The request ID this response corresponds to.
    pub id: u64,
    /// The result.
    #[serde(flatten)]
    pub result: ResponseResult,
}

#[derive(Debug, Clone, Serialize, Deserialize)]
#[serde(tag = "status")]
pub enum ResponseResult {
    /// Command succeeded.
    #[serde(rename = "ok")]
    Ok {
        #[serde(skip_serializing_if = "Option::is_none")]
        data: Option<ResponseData>,
    },
    /// Command failed.
    #[serde(rename = "error")]
    Error { message: String },
}

/// Data returned in successful responses.
#[derive(Debug, Clone, Serialize, Deserialize)]
#[serde(untagged)]
pub enum ResponseData {
    /// A stored COM object handle (returned when Get/Invoke yields an IDispatch).
    Handle { handle: u64 },
    /// A primitive value (number, string, bool, null, or error object).
    Value { value: serde_json::Value },
}

// ---------------------------------------------------------------------------
// Convenience types for the client
// ---------------------------------------------------------------------------

/// Reference to a worksheet — by 0-based index or by name.
///
/// This is a client-side convenience; the wire protocol uses raw JSON values.
/// Use `SheetRef::to_chain_step()` to convert to a `ChainStep` for navigation.
#[derive(Debug, Clone)]
pub enum SheetRef {
    /// 0-based sheet index (converted to 1-based for Excel).
    Index(u32),
    /// Sheet name.
    Name(String),
}

impl SheetRef {
    /// Convert to a `ChainStep` that indexes into the `Worksheets` collection.
    ///
    /// Excel's Worksheets collection is 1-based, so `Index(0)` becomes
    /// `Worksheets(1)` on the wire.
    pub fn to_chain_step(&self) -> ChainStep {
        match self {
            SheetRef::Index(i) => {
                ChainStep::Indexed("Worksheets".into(), serde_json::Value::from(*i + 1))
            }
            SheetRef::Name(name) => {
                ChainStep::Indexed("Worksheets".into(), serde_json::Value::from(name.as_str()))
            }
        }
    }
}

/// A cell value that can be sent to/from Excel.
///
/// This is a client-side convenience type; the wire protocol uses raw
/// `serde_json::Value`. Use the `From`/`Into` conversions.
#[derive(Debug, Clone, PartialEq)]
pub enum CellValue {
    Null,
    Bool(bool),
    Number(f64),
    String(String),
    Error(String),
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

    /// Convert to a JSON value for the wire protocol.
    pub fn to_json(&self) -> serde_json::Value {
        match self {
            CellValue::Null => serde_json::Value::Null,
            CellValue::Bool(b) => serde_json::Value::Bool(*b),
            CellValue::Number(n) => serde_json::json!(*n),
            CellValue::String(s) => serde_json::Value::String(s.clone()),
            CellValue::Error(code) => serde_json::json!({"code": code}),
        }
    }

    /// Parse from a JSON value received over the wire.
    pub fn from_json(v: &serde_json::Value) -> Self {
        match v {
            serde_json::Value::Null => CellValue::Null,
            serde_json::Value::Bool(b) => CellValue::Bool(*b),
            serde_json::Value::Number(n) => CellValue::Number(n.as_f64().unwrap_or(0.0)),
            serde_json::Value::String(s) => CellValue::String(s.clone()),
            serde_json::Value::Object(map) => {
                if let Some(code) = map.get("code").and_then(|v| v.as_str()) {
                    CellValue::Error(code.to_string())
                } else {
                    CellValue::Null
                }
            }
            _ => CellValue::Null,
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

impl From<i32> for CellValue {
    fn from(n: i32) -> Self {
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
            CellValue::Error(code) => write!(f, "{code}"),
        }
    }
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn test_get_serialization() {
        let req = Request {
            id: 1,
            command: Command::Get {
                handle: 1,
                chain: vec![
                    ChainStep::Indexed("Worksheets".into(), serde_json::json!(1)),
                    ChainStep::Indexed("Range".into(), serde_json::json!("A1")),
                ],
                property: "Value".into(),
            },
        };
        let json = serde_json::to_string(&req).unwrap();
        assert!(json.contains("\"cmd\":\"Get\""));
        assert!(json.contains("\"property\":\"Value\""));
        assert!(json.contains("[\"Worksheets\",1]"));
        assert!(json.contains("[\"Range\",\"A1\"]"));

        // Round-trip
        let parsed: Request = serde_json::from_str(&json).unwrap();
        assert_eq!(parsed.id, 1);
    }

    #[test]
    fn test_set_serialization() {
        let req = Request {
            id: 2,
            command: Command::Set {
                handle: 1,
                chain: vec![
                    ChainStep::Indexed("Worksheets".into(), serde_json::json!(1)),
                    ChainStep::Indexed("Range".into(), serde_json::json!("A1")),
                ],
                property: "Value".into(),
                value: serde_json::json!(42.0),
            },
        };
        let json = serde_json::to_string(&req).unwrap();
        assert!(json.contains("\"cmd\":\"Set\""));
        assert!(json.contains("\"value\":42.0"));

        let parsed: Request = serde_json::from_str(&json).unwrap();
        assert_eq!(parsed.id, 2);
    }

    #[test]
    fn test_invoke_serialization() {
        let req = Request {
            id: 3,
            command: Command::Invoke {
                handle: 0,
                chain: vec![ChainStep::Property("Workbooks".into())],
                method: "Add".into(),
                args: vec![],
            },
        };
        let json = serde_json::to_string(&req).unwrap();
        assert!(json.contains("\"cmd\":\"Invoke\""));
        assert!(json.contains("\"method\":\"Add\""));
        assert!(json.contains("\"chain\":[\"Workbooks\"]"));

        let parsed: Request = serde_json::from_str(&json).unwrap();
        assert_eq!(parsed.id, 3);
    }

    #[test]
    fn test_response_ok_with_handle() {
        let resp = Response {
            id: 1,
            result: ResponseResult::Ok {
                data: Some(ResponseData::Handle { handle: 5 }),
            },
        };
        let json = serde_json::to_string(&resp).unwrap();
        assert!(json.contains("\"status\":\"ok\""));
        assert!(json.contains("\"handle\":5"));

        let parsed: Response = serde_json::from_str(&json).unwrap();
        match parsed.result {
            ResponseResult::Ok {
                data: Some(ResponseData::Handle { handle }),
            } => {
                assert_eq!(handle, 5);
            }
            _ => panic!("Expected Ok with Handle"),
        }
    }

    #[test]
    fn test_response_ok_with_value() {
        let resp = Response {
            id: 2,
            result: ResponseResult::Ok {
                data: Some(ResponseData::Value {
                    value: serde_json::json!(42.0),
                }),
            },
        };
        let json = serde_json::to_string(&resp).unwrap();
        assert!(json.contains("\"value\":42.0"));

        let parsed: Response = serde_json::from_str(&json).unwrap();
        match parsed.result {
            ResponseResult::Ok {
                data: Some(ResponseData::Value { value }),
            } => {
                assert_eq!(value, serde_json::json!(42.0));
            }
            _ => panic!("Expected Ok with Value"),
        }
    }

    #[test]
    fn test_response_error() {
        let resp = Response {
            id: 3,
            result: ResponseResult::Error {
                message: "something broke".into(),
            },
        };
        let json = serde_json::to_string(&resp).unwrap();
        assert!(json.contains("\"status\":\"error\""));
        assert!(json.contains("\"message\":\"something broke\""));
    }

    #[test]
    fn test_chain_step_property() {
        let step = ChainStep::Property("Font".into());
        let json = serde_json::to_string(&step).unwrap();
        assert_eq!(json, "\"Font\"");

        let parsed: ChainStep = serde_json::from_str(&json).unwrap();
        match parsed {
            ChainStep::Property(s) => assert_eq!(s, "Font"),
            _ => panic!("Expected Property"),
        }
    }

    #[test]
    fn test_chain_step_indexed() {
        let step = ChainStep::Indexed("Range".into(), serde_json::json!("A1"));
        let json = serde_json::to_string(&step).unwrap();
        assert_eq!(json, "[\"Range\",\"A1\"]");

        let parsed: ChainStep = serde_json::from_str(&json).unwrap();
        match parsed {
            ChainStep::Indexed(name, val) => {
                assert_eq!(name, "Range");
                assert_eq!(val, serde_json::json!("A1"));
            }
            _ => panic!("Expected Indexed"),
        }
    }

    #[test]
    fn test_sheet_ref_to_chain_step() {
        let by_idx = SheetRef::Index(0);
        let step = by_idx.to_chain_step();
        let json = serde_json::to_string(&step).unwrap();
        // Index 0 -> 1-based -> Worksheets[1]
        assert_eq!(json, "[\"Worksheets\",1]");

        let by_name = SheetRef::Name("Sheet2".into());
        let step = by_name.to_chain_step();
        let json = serde_json::to_string(&step).unwrap();
        assert_eq!(json, "[\"Worksheets\",\"Sheet2\"]");
    }

    #[test]
    fn test_cell_value_round_trip() {
        let cases = vec![
            CellValue::Null,
            CellValue::Bool(true),
            CellValue::Number(3.14),
            CellValue::String("hello".into()),
            CellValue::Error("#DIV/0!".into()),
        ];
        for cv in cases {
            let json = cv.to_json();
            let back = CellValue::from_json(&json);
            assert_eq!(cv, back, "Round-trip failed for {cv:?}");
        }
    }
}
