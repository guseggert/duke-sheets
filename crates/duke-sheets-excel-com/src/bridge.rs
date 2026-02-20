//! TCP client for the Excel COM bridge server.
//!
//! Connects to a C# bridge server running inside a Windows VM (QEMU/KVM)
//! that provides generic COM proxy operations over NDJSON-over-TCP.

use std::io::{BufRead, BufReader, BufWriter, Write};
use std::net::{SocketAddr, TcpStream, ToSocketAddrs};
use std::sync::atomic::{AtomicU64, Ordering};
use std::sync::Mutex;
use std::time::Duration;

use excel_com_protocol::{
    CellValue, ChainStep, Command, Request, Response, ResponseData, ResponseResult, SheetRef,
};

use crate::workbook::Workbook;

/// Errors from the Excel COM bridge.
#[derive(Debug, thiserror::Error)]
pub enum BridgeError {
    #[error("Failed to connect to bridge server at {0}: {1}")]
    ConnectFailed(SocketAddr, std::io::Error),

    #[error("Connection lost")]
    ConnectionLost,

    #[error("Failed to send command: {0}")]
    SendFailed(String),

    #[error("Failed to read response: {0}")]
    ReadFailed(String),

    #[error("JSON error: {0}")]
    JsonError(#[from] serde_json::Error),

    #[error("Bridge returned error: {0}")]
    BridgeError(String),

    #[error("Expected a handle in response")]
    ExpectedHandle,

    #[error("Expected a value in response")]
    ExpectedValue,
}

/// Configuration for connecting to the Excel COM bridge server.
pub struct ExcelBridgeConfig {
    /// Address of the bridge server. Default: `127.0.0.1:9876`.
    pub addr: SocketAddr,

    /// Connection timeout.
    pub connect_timeout: Duration,

    /// Read timeout for waiting for responses.
    pub read_timeout: Duration,
}

impl Default for ExcelBridgeConfig {
    fn default() -> Self {
        Self {
            addr: "127.0.0.1:9876".parse().unwrap(),
            connect_timeout: Duration::from_secs(10),
            read_timeout: Duration::from_secs(30),
        }
    }
}

/// TCP connection to the Excel COM bridge server.
///
/// Provides both low-level generic COM proxy methods (`get`, `set`, `invoke`)
/// and high-level Excel-specific convenience methods (`create_workbook`,
/// `set_cell_value`, etc.).
pub struct ExcelBridge {
    reader: Mutex<BufReader<TcpStream>>,
    writer: Mutex<BufWriter<TcpStream>>,
    next_id: AtomicU64,
}

impl ExcelBridge {
    /// Connect to a running bridge server and initialize Excel.
    pub fn connect(config: ExcelBridgeConfig) -> Result<Self, BridgeError> {
        let stream = TcpStream::connect_timeout(&config.addr, config.connect_timeout)
            .map_err(|e| BridgeError::ConnectFailed(config.addr, e))?;

        stream
            .set_read_timeout(Some(config.read_timeout))
            .map_err(|e| BridgeError::ReadFailed(e.to_string()))?;

        stream
            .set_nodelay(true)
            .map_err(|e| BridgeError::SendFailed(e.to_string()))?;

        let reader = stream
            .try_clone()
            .map_err(|e| BridgeError::ReadFailed(e.to_string()))?;

        let bridge = Self {
            reader: Mutex::new(BufReader::new(reader)),
            writer: Mutex::new(BufWriter::new(stream)),
            next_id: AtomicU64::new(1),
        };

        // Initialize COM and Excel
        bridge.send_command(Command::Init)?;

        Ok(bridge)
    }

    /// Connect to localhost on a specific port.
    pub fn connect_local(port: u16) -> Result<Self, BridgeError> {
        Self::connect(ExcelBridgeConfig {
            addr: format!("127.0.0.1:{port}").parse().unwrap(),
            ..Default::default()
        })
    }

    /// Connect using a hostname:port string.
    pub fn connect_addr(addr: impl ToSocketAddrs) -> Result<Self, BridgeError> {
        let socket_addr = addr
            .to_socket_addrs()
            .map_err(|e| BridgeError::SendFailed(format!("Invalid address: {e}")))?
            .next()
            .ok_or_else(|| BridgeError::SendFailed("No address resolved".into()))?;

        Self::connect(ExcelBridgeConfig {
            addr: socket_addr,
            ..Default::default()
        })
    }

    // -----------------------------------------------------------------------
    // Low-level generic COM proxy operations
    // -----------------------------------------------------------------------

    /// Send a command and wait for the response.
    fn send_command(&self, command: Command) -> Result<Option<ResponseData>, BridgeError> {
        let id = self.next_id.fetch_add(1, Ordering::Relaxed);
        let request = Request { id, command };
        let json = serde_json::to_string(&request)?;

        // Send
        {
            let mut writer = self.writer.lock().unwrap();
            writeln!(writer, "{json}").map_err(|e| BridgeError::SendFailed(e.to_string()))?;
            writer
                .flush()
                .map_err(|e| BridgeError::SendFailed(e.to_string()))?;
        }

        // Read
        let response: Response = {
            let mut reader = self.reader.lock().unwrap();
            let mut line = String::new();
            reader
                .read_line(&mut line)
                .map_err(|e| BridgeError::ReadFailed(e.to_string()))?;
            if line.is_empty() {
                return Err(BridgeError::ConnectionLost);
            }
            serde_json::from_str(&line)?
        };

        match response.result {
            ResponseResult::Ok { data } => Ok(data),
            ResponseResult::Error { message } => Err(BridgeError::BridgeError(message)),
        }
    }

    /// Get a property from a COM object, navigating a chain from a handle.
    ///
    /// Returns `Ok(Value)` for primitives, `Ok(Handle)` for COM objects.
    pub fn get(
        &self,
        handle: u64,
        chain: Vec<ChainStep>,
        property: &str,
    ) -> Result<Option<ResponseData>, BridgeError> {
        self.send_command(Command::Get {
            handle,
            chain,
            property: property.to_string(),
        })
    }

    /// Set a property on a COM object, navigating a chain from a handle.
    pub fn set(
        &self,
        handle: u64,
        chain: Vec<ChainStep>,
        property: &str,
        value: serde_json::Value,
    ) -> Result<(), BridgeError> {
        self.send_command(Command::Set {
            handle,
            chain,
            property: property.to_string(),
            value,
        })?;
        Ok(())
    }

    /// Invoke a method on a COM object, navigating a chain from a handle.
    ///
    /// Returns `Ok(Value)` for primitives, `Ok(Handle)` for COM objects.
    pub fn invoke(
        &self,
        handle: u64,
        chain: Vec<ChainStep>,
        method: &str,
        args: Vec<serde_json::Value>,
    ) -> Result<Option<ResponseData>, BridgeError> {
        self.send_command(Command::Invoke {
            handle,
            chain,
            method: method.to_string(),
            args,
        })
    }

    /// Release a stored COM object handle on the server.
    pub fn release(&self, handle: u64) -> Result<(), BridgeError> {
        self.send_command(Command::Release { handle })?;
        Ok(())
    }

    // -----------------------------------------------------------------------
    // High-level Excel convenience methods
    // -----------------------------------------------------------------------

    /// Create a new empty workbook. Returns a `Workbook` handle.
    ///
    /// Equivalent to `Excel.Application.Workbooks.Add()`.
    pub fn create_workbook(&self) -> Result<Workbook<'_>, BridgeError> {
        let data = self.invoke(0, vec![cs_prop("Workbooks")], "Add", vec![])?;
        let handle = extract_handle(data)?;
        Ok(Workbook::new(self, handle))
    }

    /// Open a workbook from a Windows file path. Returns a `Workbook` handle.
    ///
    /// The path must be a Windows path as seen inside the VM
    /// (e.g., `C:\Users\...` or a UNC path like `\\10.0.2.4\qemu\file.xlsx`).
    pub fn open_workbook(&self, windows_path: &str) -> Result<Workbook<'_>, BridgeError> {
        let data = self.invoke(
            0,
            vec![cs_prop("Workbooks")],
            "Open",
            vec![serde_json::Value::from(windows_path)],
        )?;
        let handle = extract_handle(data)?;
        Ok(Workbook::new(self, handle))
    }

    /// Force a full recalculation of all open workbooks.
    ///
    /// Equivalent to `Excel.Application.Calculate()`.
    pub fn recalculate(&self) -> Result<(), BridgeError> {
        self.invoke(0, vec![], "Calculate", vec![])?;
        Ok(())
    }

    /// Shut down: close all workbooks, quit Excel, end the session.
    pub fn shutdown(self) -> Result<(), BridgeError> {
        let _ = self.send_command(Command::Shutdown);
        Ok(())
    }

    // -- Cell operations (used by Workbook) --

    pub(crate) fn set_cell_value(
        &self,
        workbook: u64,
        sheet: SheetRef,
        cell: &str,
        value: CellValue,
    ) -> Result<(), BridgeError> {
        let chain = vec![sheet.to_chain_step(), cs_idx("Range", cell)];
        self.set(workbook, chain, "Value", value.to_json())
    }

    pub(crate) fn set_cell_formula(
        &self,
        workbook: u64,
        sheet: SheetRef,
        cell: &str,
        formula: &str,
    ) -> Result<(), BridgeError> {
        let chain = vec![sheet.to_chain_step(), cs_idx("Range", cell)];
        self.set(workbook, chain, "Formula", serde_json::Value::from(formula))
    }

    pub(crate) fn get_cell_value(
        &self,
        workbook: u64,
        sheet: SheetRef,
        cell: &str,
    ) -> Result<CellValue, BridgeError> {
        let chain = vec![sheet.to_chain_step(), cs_idx("Range", cell)];
        let data = self.get(workbook, chain, "Value")?;
        match data {
            Some(ResponseData::Value { value }) => Ok(CellValue::from_json(&value)),
            None => Ok(CellValue::Null),
            _ => Ok(CellValue::Null),
        }
    }

    pub(crate) fn get_cell_formula(
        &self,
        workbook: u64,
        sheet: SheetRef,
        cell: &str,
    ) -> Result<String, BridgeError> {
        let chain = vec![sheet.to_chain_step(), cs_idx("Range", cell)];
        let data = self.get(workbook, chain, "Formula")?;
        match data {
            Some(ResponseData::Value { value }) => {
                Ok(value.as_str().unwrap_or_default().to_string())
            }
            _ => Ok(String::new()),
        }
    }

    pub(crate) fn save_workbook(
        &self,
        workbook: u64,
        path: &str,
        format: i32,
    ) -> Result<(), BridgeError> {
        self.invoke(
            workbook,
            vec![],
            "SaveAs",
            vec![
                serde_json::Value::from(path),
                serde_json::Value::from(format),
            ],
        )?;
        Ok(())
    }

    pub(crate) fn close_workbook(&self, workbook: u64) -> Result<(), BridgeError> {
        self.invoke(
            workbook,
            vec![],
            "Close",
            vec![serde_json::Value::from(false)],
        )?;
        self.release(workbook)?;
        Ok(())
    }
}

// ---------------------------------------------------------------------------
// Helpers
// ---------------------------------------------------------------------------

/// Create a `ChainStep::Property`.
fn cs_prop(name: &str) -> ChainStep {
    ChainStep::Property(name.to_string())
}

/// Create a `ChainStep::Indexed` with a string index.
fn cs_idx(name: &str, index: &str) -> ChainStep {
    ChainStep::Indexed(name.to_string(), serde_json::Value::from(index))
}

/// Extract a handle from a response.
fn extract_handle(data: Option<ResponseData>) -> Result<u64, BridgeError> {
    match data {
        Some(ResponseData::Handle { handle }) => Ok(handle),
        _ => Err(BridgeError::ExpectedHandle),
    }
}
