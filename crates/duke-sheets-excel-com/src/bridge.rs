//! Subprocess management and JSON IPC for the WINE bridge process.

use std::io::{BufRead, BufReader, Write};
use std::path::{Path, PathBuf};
use std::process::{Child, Stdio};
use std::sync::atomic::{AtomicU64, Ordering};
use std::sync::Mutex;
use std::time::Duration;

use excel_com_protocol::{
    self, CellValue, Command as BridgeCommand, Request, Response, ResponseData, ResponseResult,
    SheetRef,
};

use crate::workbook::Workbook;

/// Errors from the Excel COM bridge.
#[derive(Debug, thiserror::Error)]
pub enum BridgeError {
    #[error("Failed to spawn WINE bridge process: {0}")]
    SpawnFailed(#[from] std::io::Error),

    #[error("Bridge process not running")]
    NotRunning,

    #[error("Failed to send command to bridge: {0}")]
    SendFailed(String),

    #[error("Failed to read response from bridge: {0}")]
    ReadFailed(String),

    #[error("JSON serialization error: {0}")]
    JsonError(#[from] serde_json::Error),

    #[error("Bridge returned error: {0}")]
    BridgeError(String),

    #[error("Unexpected response data")]
    UnexpectedResponse,

    #[error("WINE not found. Install WINE and ensure 'wine' is in PATH.")]
    WineNotFound,

    #[error("Bridge executable not found at: {0}")]
    BridgeExeNotFound(String),
}

/// Configuration for the Excel COM bridge.
pub struct ExcelBridgeConfig {
    /// Path to the `excel-com-bridge.exe` Windows executable.
    /// If None, will search in common locations relative to the current binary.
    pub bridge_exe_path: Option<PathBuf>,

    /// Path to the WINE executable. Defaults to "wine".
    pub wine_path: PathBuf,

    /// Optional WINEPREFIX to use (for isolating the WINE environment).
    pub wine_prefix: Option<PathBuf>,

    /// Timeout for waiting for bridge responses.
    pub timeout: Duration,
}

impl Default for ExcelBridgeConfig {
    fn default() -> Self {
        Self {
            bridge_exe_path: None,
            wine_path: PathBuf::from("wine"),
            wine_prefix: None,
            timeout: Duration::from_secs(30),
        }
    }
}

/// The main handle for communicating with the Excel COM bridge.
///
/// This manages the WINE subprocess lifecycle and provides methods
/// for Excel automation operations.
pub struct ExcelBridge {
    child: Mutex<Child>,
    stdin: Mutex<std::process::ChildStdin>,
    stdout: Mutex<BufReader<std::process::ChildStdout>>,
    next_id: AtomicU64,
}

impl ExcelBridge {
    /// Start the bridge process and initialize Excel.
    pub fn start(config: ExcelBridgeConfig) -> Result<Self, BridgeError> {
        let exe_path = config.bridge_exe_path.unwrap_or_else(|| find_bridge_exe());

        if !exe_path.exists() {
            return Err(BridgeError::BridgeExeNotFound(
                exe_path.display().to_string(),
            ));
        }

        let mut cmd = std::process::Command::new(&config.wine_path);

        if let Some(prefix) = &config.wine_prefix {
            cmd.env("WINEPREFIX", prefix);
        }

        cmd.arg(&exe_path);
        cmd.stdin(Stdio::piped());
        cmd.stdout(Stdio::piped());
        cmd.stderr(Stdio::inherit()); // Bridge diagnostics go to our stderr

        let mut child = cmd.spawn().map_err(|e| {
            if e.kind() == std::io::ErrorKind::NotFound {
                BridgeError::WineNotFound
            } else {
                BridgeError::SpawnFailed(e)
            }
        })?;

        let stdin = child.stdin.take().expect("stdin was piped");
        let stdout = child.stdout.take().expect("stdout was piped");

        let bridge = Self {
            child: Mutex::new(child),
            stdin: Mutex::new(stdin),
            stdout: Mutex::new(BufReader::new(stdout)),
            next_id: AtomicU64::new(1),
        };

        // Initialize COM and Excel
        bridge.send_command(BridgeCommand::Init)?;

        Ok(bridge)
    }

    /// Send a command to the bridge and wait for the response.
    fn send_command(&self, command: BridgeCommand) -> Result<Option<ResponseData>, BridgeError> {
        let id = self.next_id.fetch_add(1, Ordering::Relaxed);

        let request = Request { id, command };
        let json = serde_json::to_string(&request)?;

        // Send the request
        {
            let mut stdin = self.stdin.lock().unwrap();
            writeln!(stdin, "{json}").map_err(|e| BridgeError::SendFailed(e.to_string()))?;
            stdin
                .flush()
                .map_err(|e| BridgeError::SendFailed(e.to_string()))?;
        }

        // Read the response
        let response: Response = {
            let mut stdout = self.stdout.lock().unwrap();
            let mut line = String::new();
            stdout
                .read_line(&mut line)
                .map_err(|e| BridgeError::ReadFailed(e.to_string()))?;

            if line.is_empty() {
                return Err(BridgeError::NotRunning);
            }

            serde_json::from_str(&line)?
        };

        match response.result {
            ResponseResult::Ok { data } => Ok(data),
            ResponseResult::Error { message } => Err(BridgeError::BridgeError(message)),
        }
    }

    /// Create a new empty workbook.
    pub fn create_workbook(&self) -> Result<Workbook<'_>, BridgeError> {
        let data = self.send_command(BridgeCommand::CreateWorkbook)?;
        match data {
            Some(ResponseData::WorkbookHandle { workbook }) => Ok(Workbook::new(self, workbook)),
            _ => Err(BridgeError::UnexpectedResponse),
        }
    }

    /// Open an existing workbook from a file path.
    ///
    /// The path should be a Windows-style path as seen by WINE.
    /// Use `linux_to_wine_path` to convert if needed.
    pub fn open_workbook(&self, path: &str) -> Result<Workbook<'_>, BridgeError> {
        let data = self.send_command(BridgeCommand::OpenWorkbook {
            path: path.to_string(),
        })?;
        match data {
            Some(ResponseData::WorkbookHandle { workbook }) => Ok(Workbook::new(self, workbook)),
            _ => Err(BridgeError::UnexpectedResponse),
        }
    }

    /// Force Excel to recalculate all open workbooks.
    pub fn recalculate(&self) -> Result<(), BridgeError> {
        self.send_command(BridgeCommand::Recalculate)?;
        Ok(())
    }

    /// Shut down the bridge: close all workbooks, quit Excel, and terminate the process.
    pub fn shutdown(self) -> Result<(), BridgeError> {
        let _ = self.send_command(BridgeCommand::Shutdown);

        // Wait for the child process to exit
        let mut child = self.child.lock().unwrap();
        let _ = child.wait();

        Ok(())
    }

    // -- Internal methods used by Workbook --

    pub(crate) fn set_cell_value(
        &self,
        workbook: u64,
        sheet: SheetRef,
        cell: &str,
        value: CellValue,
    ) -> Result<(), BridgeError> {
        self.send_command(BridgeCommand::SetCellValue {
            workbook,
            sheet,
            cell: cell.to_string(),
            value,
        })?;
        Ok(())
    }

    pub(crate) fn set_cell_formula(
        &self,
        workbook: u64,
        sheet: SheetRef,
        cell: &str,
        formula: &str,
    ) -> Result<(), BridgeError> {
        self.send_command(BridgeCommand::SetCellFormula {
            workbook,
            sheet,
            cell: cell.to_string(),
            formula: formula.to_string(),
        })?;
        Ok(())
    }

    pub(crate) fn get_cell_value(
        &self,
        workbook: u64,
        sheet: SheetRef,
        cell: &str,
    ) -> Result<CellValue, BridgeError> {
        let data = self.send_command(BridgeCommand::GetCellValue {
            workbook,
            sheet,
            cell: cell.to_string(),
        })?;
        match data {
            Some(ResponseData::Value { value }) => Ok(value),
            _ => Err(BridgeError::UnexpectedResponse),
        }
    }

    pub(crate) fn get_cell_formula(
        &self,
        workbook: u64,
        sheet: SheetRef,
        cell: &str,
    ) -> Result<String, BridgeError> {
        let data = self.send_command(BridgeCommand::GetCellFormula {
            workbook,
            sheet,
            cell: cell.to_string(),
        })?;
        match data {
            Some(ResponseData::Formula { formula }) => Ok(formula),
            _ => Err(BridgeError::UnexpectedResponse),
        }
    }

    pub(crate) fn save_workbook(&self, workbook: u64, path: &str) -> Result<(), BridgeError> {
        self.send_command(BridgeCommand::SaveWorkbook {
            workbook,
            path: path.to_string(),
        })?;
        Ok(())
    }

    pub(crate) fn close_workbook(&self, workbook: u64) -> Result<(), BridgeError> {
        self.send_command(BridgeCommand::CloseWorkbook { workbook })?;
        Ok(())
    }
}

/// Convert a Linux filesystem path to a WINE (Windows) path.
///
/// WINE maps `/` to `Z:\`, so `/home/user/file.xlsx` becomes `Z:\home\user\file.xlsx`.
/// The WINE prefix's `drive_c` maps to `C:\`.
pub fn linux_to_wine_path(linux_path: &Path) -> String {
    let abs = if linux_path.is_absolute() {
        linux_path.to_path_buf()
    } else {
        std::env::current_dir().unwrap_or_default().join(linux_path)
    };

    // WINE maps the root filesystem to Z:
    format!("Z:{}", abs.display()).replace('/', "\\")
}

/// Attempt to locate the bridge exe relative to the current executable or in common paths.
fn find_bridge_exe() -> PathBuf {
    // Check next to the current executable
    if let Ok(mut exe) = std::env::current_exe() {
        exe.pop();
        let candidate = exe.join("excel-com-bridge.exe");
        if candidate.exists() {
            return candidate;
        }
    }

    // Check in the target directory (for development)
    let target_path = PathBuf::from("target/x86_64-pc-windows-gnu/release/excel-com-bridge.exe");
    if target_path.exists() {
        return target_path;
    }

    let target_path = PathBuf::from("target/x86_64-pc-windows-gnu/debug/excel-com-bridge.exe");
    if target_path.exists() {
        return target_path;
    }

    // Default: assume it's in the current directory
    PathBuf::from("excel-com-bridge.exe")
}
