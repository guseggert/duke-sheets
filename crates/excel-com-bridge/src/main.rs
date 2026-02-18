//! Excel COM Bridge — a Windows process that automates Excel via COM,
//! controlled by JSON commands over stdin/stdout.
//!
//! Designed to be cross-compiled from Linux and run under WINE.
//!
//! Protocol: one JSON object per line (newline-delimited JSON).
//! - Reads `Request` objects from stdin
//! - Writes `Response` objects to stdout
//! - Diagnostic/log messages go to stderr (never stdout)

#[cfg(windows)]
mod dispatch;
#[cfg(windows)]
mod excel;

#[cfg(not(windows))]
fn main() {
    eprintln!("excel-com-bridge must be compiled for Windows (--target x86_64-pc-windows-gnu)");
    eprintln!("and run under WINE on Linux.");
    std::process::exit(1);
}

#[cfg(windows)]
fn main() {
    use std::io::{self, BufRead, Write};

    use excel_com_protocol::*;

    // Use stderr for all diagnostic output so stdout stays clean for protocol
    eprintln!("[excel-com-bridge] Starting up...");

    let stdin = io::stdin();
    let stdout = io::stdout();
    let mut out = stdout.lock();

    let mut excel: Option<excel::ExcelApp> = None;

    for line in stdin.lock().lines() {
        let line = match line {
            Ok(l) => l,
            Err(e) => {
                eprintln!("[excel-com-bridge] stdin read error: {e}");
                break;
            }
        };

        let line = line.trim();
        if line.is_empty() {
            continue;
        }

        let request: Request = match serde_json::from_str(line) {
            Ok(r) => r,
            Err(e) => {
                eprintln!("[excel-com-bridge] JSON parse error: {e}");
                eprintln!("[excel-com-bridge] Line was: {line}");
                // Send an error response with id=0 since we couldn't parse the request
                let resp = Response {
                    id: 0,
                    result: ResponseResult::Error {
                        message: format!("JSON parse error: {e}"),
                    },
                };
                let _ = writeln!(out, "{}", serde_json::to_string(&resp).unwrap());
                let _ = out.flush();
                continue;
            }
        };

        let response = handle_command(&mut excel, &request);
        let json = serde_json::to_string(&response).unwrap();
        let _ = writeln!(out, "{json}");
        let _ = out.flush();

        // If it was a shutdown command and it succeeded, exit
        if matches!(request.command, Command::Shutdown) {
            if matches!(response.result, ResponseResult::Ok { .. }) {
                eprintln!("[excel-com-bridge] Shutdown complete, exiting.");
                break;
            }
        }
    }

    // If Excel is still running when stdin closes, try to clean up
    if let Some(app) = excel {
        eprintln!("[excel-com-bridge] stdin closed, shutting down Excel...");
        let _ = app.shutdown();
    }

    eprintln!("[excel-com-bridge] Process exiting.");
}

#[cfg(windows)]
fn handle_command(
    excel: &mut Option<excel::ExcelApp>,
    request: &excel_com_protocol::Request,
) -> excel_com_protocol::Response {
    use excel_com_protocol::*;

    let id = request.id;

    let result = match &request.command {
        Command::Init => init_com_and_excel(excel),
        Command::CreateWorkbook => with_excel(excel, |app| {
            let handle = app.create_workbook()?;
            Ok(ResponseResult::Ok {
                data: Some(ResponseData::WorkbookHandle { workbook: handle }),
            })
        }),
        Command::OpenWorkbook { path } => with_excel(excel, |app| {
            let handle = app.open_workbook(path)?;
            Ok(ResponseResult::Ok {
                data: Some(ResponseData::WorkbookHandle { workbook: handle }),
            })
        }),
        Command::SetCellValue {
            workbook,
            sheet,
            cell,
            value,
        } => with_excel(excel, |app| {
            app.set_cell_value(*workbook, sheet, cell, value)?;
            Ok(ResponseResult::Ok { data: None })
        }),
        Command::SetCellFormula {
            workbook,
            sheet,
            cell,
            formula,
        } => with_excel(excel, |app| {
            app.set_cell_formula(*workbook, sheet, cell, formula)?;
            Ok(ResponseResult::Ok { data: None })
        }),
        Command::GetCellValue {
            workbook,
            sheet,
            cell,
        } => with_excel(excel, |app| {
            let value = app.get_cell_value(*workbook, sheet, cell)?;
            Ok(ResponseResult::Ok {
                data: Some(ResponseData::Value { value }),
            })
        }),
        Command::GetCellFormula {
            workbook,
            sheet,
            cell,
        } => with_excel(excel, |app| {
            let formula = app.get_cell_formula(*workbook, sheet, cell)?;
            Ok(ResponseResult::Ok {
                data: Some(ResponseData::Formula { formula }),
            })
        }),
        Command::Recalculate => with_excel(excel, |app| {
            app.recalculate()?;
            Ok(ResponseResult::Ok { data: None })
        }),
        Command::SaveWorkbook { workbook, path } => with_excel(excel, |app| {
            app.save_workbook(*workbook, path)?;
            Ok(ResponseResult::Ok { data: None })
        }),
        Command::CloseWorkbook { workbook } => with_excel(excel, |app| {
            app.close_workbook(*workbook)?;
            Ok(ResponseResult::Ok { data: None })
        }),
        Command::Shutdown => match excel.take() {
            Some(app) => match app.shutdown() {
                Ok(()) => {
                    uninit_com();
                    ResponseResult::Ok { data: None }
                }
                Err(e) => ResponseResult::Error {
                    message: format!("Shutdown failed: {e}"),
                },
            },
            None => ResponseResult::Ok { data: None },
        },
    };

    Response { id, result }
}

#[cfg(windows)]
fn init_com_and_excel(excel: &mut Option<excel::ExcelApp>) -> excel_com_protocol::ResponseResult {
    use excel_com_protocol::ResponseResult;
    use windows::Win32::System::Com::{CoInitializeEx, COINIT_APARTMENTTHREADED};

    if excel.is_some() {
        return ResponseResult::Ok { data: None }; // Already initialized
    }

    // Initialize COM in Single-Threaded Apartment mode (required by Excel)
    unsafe {
        let hr = CoInitializeEx(None, COINIT_APARTMENTTHREADED);
        if let Err(e) = hr.ok() {
            return ResponseResult::Error {
                message: format!("CoInitializeEx failed: {e}"),
            };
        }
    }

    eprintln!("[excel-com-bridge] COM initialized (STA)");

    match excel::ExcelApp::new() {
        Ok(app) => {
            eprintln!("[excel-com-bridge] Excel.Application created successfully");
            *excel = Some(app);
            ResponseResult::Ok { data: None }
        }
        Err(e) => ResponseResult::Error {
            message: format!("Failed to create Excel.Application: {e}"),
        },
    }
}

#[cfg(windows)]
fn uninit_com() {
    unsafe {
        windows::Win32::System::Com::CoUninitialize();
    }
    eprintln!("[excel-com-bridge] COM uninitialized");
}

#[cfg(windows)]
fn with_excel(
    excel: &mut Option<excel::ExcelApp>,
    f: impl FnOnce(&mut excel::ExcelApp) -> Result<excel_com_protocol::ResponseResult, String>,
) -> excel_com_protocol::ResponseResult {
    match excel.as_mut() {
        Some(app) => match f(app) {
            Ok(r) => r,
            Err(e) => excel_com_protocol::ResponseResult::Error { message: e },
        },
        None => excel_com_protocol::ResponseResult::Error {
            message: "Excel not initialized. Send 'Init' command first.".to_string(),
        },
    }
}
