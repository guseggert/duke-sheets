//! Common utilities for Excel COM E2E tests.
//!
//! Provides a global bridge connection singleton and helper functions
//! for temp file management. The bridge is shared across all tests
//! because Excel COM startup is slow (~2s).
//!
//! File transfer: Excel saves to `C:\temp\` inside the VM. The test then
//! pulls the file to the Linux host via WinRM (base64-encoded PowerShell
//! response), reads it with `XlsxReader`, and cleans up both sides.

use std::io::Write;
use std::path::PathBuf;
use std::sync::atomic::{AtomicU64, Ordering};
use std::sync::{Mutex, OnceLock};

use base64::prelude::*;
use duke_sheets_excel_com::{ExcelBridge, ExcelBridgeConfig};

/// Host-side directory for downloaded test fixtures.
const HOST_DIR: &str = "/tmp/duke-sheets-excel";

/// VM-side directory where Excel saves files.
const VM_DIR: &str = r"C:\temp";

/// Counter for generating unique temp file names / UUIDs.
static COUNTER: AtomicU64 = AtomicU64::new(0);

/// Global bridge connection, initialized once and shared across tests.
static BRIDGE: OnceLock<Mutex<ExcelBridge>> = OnceLock::new();

/// Get the global Excel bridge, connecting on first call.
///
/// Panics if the bridge server is not available — tests require a running
/// Windows VM with the bridge server on localhost:9876.
pub fn excel_bridge() -> &'static Mutex<ExcelBridge> {
    BRIDGE.get_or_init(|| {
        let bridge = ExcelBridge::connect(ExcelBridgeConfig::default()).expect(
            "Failed to connect to Excel COM bridge on localhost:9876. \
             Start the VM with: bash tools/vm/qemu-start.sh",
        );
        Mutex::new(bridge)
    })
}

/// A temp file path pair: the host-side path (for reading back) and the
/// Windows path (for telling Excel where to save).
pub struct TempFixture {
    pub host_path: PathBuf,
    pub vm_path: String,
    name: String,
}

/// Generate a unique temp fixture with both host and VM paths.
pub fn temp_fixture() -> TempFixture {
    let n = COUNTER.fetch_add(1, Ordering::SeqCst);
    let pid = std::process::id();
    let name = format!("test_{pid}_{n}.xlsx");
    let _ = std::fs::create_dir_all(HOST_DIR);
    TempFixture {
        host_path: PathBuf::from(format!("{HOST_DIR}/{name}")),
        vm_path: format!("{VM_DIR}\\{name}"),
        name: name.clone(),
    }
}

/// Generate a unique temp fixture with .xls extension for XLS format tests.
pub fn temp_fixture_xls() -> TempFixture {
    let n = COUNTER.fetch_add(1, Ordering::SeqCst);
    let pid = std::process::id();
    let name = format!("test_{pid}_{n}.xls");
    let _ = std::fs::create_dir_all(HOST_DIR);
    TempFixture {
        host_path: PathBuf::from(format!("{HOST_DIR}/{name}")),
        vm_path: format!("{VM_DIR}\\{name}"),
        name: name.clone(),
    }
}

/// Ensure C:\temp exists inside the VM (called once).
pub fn ensure_vm_temp_dir() {
    static DONE: std::sync::Once = std::sync::Once::new();
    DONE.call_once(|| {
        let _ = run_winrm_ps(&format!(
            "if (-not (Test-Path '{VM_DIR}')) {{ New-Item -ItemType Directory -Path '{VM_DIR}' -Force | Out-Null }}"
        ));
    });
}

/// Pull a file from the VM to the host via WinRM.
///
/// Runs a PowerShell command that base64-encodes the file, then decodes
/// and writes it on the host side.
pub fn pull_file_from_vm(fixture: &TempFixture) {
    let ps_script = format!(
        "[Convert]::ToBase64String([IO.File]::ReadAllBytes('{path}'))",
        path = fixture.vm_path
    );
    let output = run_winrm_ps(&ps_script).expect("WinRM pull_file_from_vm failed");

    let b64: String = output.chars().filter(|c| !c.is_whitespace()).collect();
    let bytes = BASE64_STANDARD.decode(&b64).unwrap_or_else(|e| {
        panic!(
            "base64 decode failed for {}: {e}\nfirst 200 chars: {}",
            fixture.name,
            &b64[..b64.len().min(200)]
        )
    });

    let mut f = std::fs::File::create(&fixture.host_path)
        .unwrap_or_else(|e| panic!("create {}: {e}", fixture.host_path.display()));
    f.write_all(&bytes)
        .unwrap_or_else(|e| panic!("write {}: {e}", fixture.host_path.display()));
}

/// Push a file from the host to the VM via WinRM.
///
/// Reads the host file, base64-encodes it, and writes it inside the VM
/// in chunks (WinRM has a command-line length limit).  Each chunk is
/// appended to a temp .b64 file, then the whole thing is decoded to
/// the final binary path.
pub fn push_file_to_vm(fixture: &TempFixture) {
    let bytes = std::fs::read(&fixture.host_path)
        .unwrap_or_else(|e| panic!("read {}: {e}", fixture.host_path.display()));
    let b64 = BASE64_STANDARD.encode(&bytes);

    let b64_path = format!("{}.b64", fixture.vm_path);

    // Clear any previous temp file
    let _ = run_winrm_ps(&format!(
        "Remove-Item -Force -ErrorAction SilentlyContinue '{b64_path}'"
    ));

    // Write base64 in chunks (~2000 chars each to stay under cmd line limits
    // after UTF-16LE encoding for -EncodedCommand)
    const CHUNK: usize = 2000;
    for chunk in b64.as_bytes().chunks(CHUNK) {
        let chunk_str = std::str::from_utf8(chunk).unwrap();
        let ps = format!("Add-Content -NoNewline -Path '{b64_path}' -Value '{chunk_str}'");
        run_winrm_ps(&ps).unwrap_or_else(|e| panic!("push chunk failed: {e}"));
    }

    // Decode base64 file to final binary
    let ps = format!(
        "[IO.File]::WriteAllBytes('{path}', [Convert]::FromBase64String([IO.File]::ReadAllText('{b64_path}'))); Remove-Item '{b64_path}'",
        path = fixture.vm_path,
        b64_path = b64_path,
    );
    run_winrm_ps(&ps).expect("WinRM base64 decode failed");
}

/// Clean up fixture files on both host and VM. Ignores errors.
pub fn cleanup_fixture(fixture: &TempFixture) {
    let _ = std::fs::remove_file(&fixture.host_path);
    let ps = format!(
        "Remove-Item -Force -ErrorAction SilentlyContinue '{}'",
        fixture.vm_path
    );
    let _ = run_winrm_ps(&ps);
}

// ---------------------------------------------------------------------------
// Writer E2E helper: duke-sheets write → Excel open+re-save → duke-sheets read
// ---------------------------------------------------------------------------

/// Write a workbook with duke-sheets, push to the VM, open in real Excel
/// (asserting no repair), re-save to a second file, pull back, and read
/// with `XlsxReader`.  Returns the re-read `Workbook`.
///
/// This is the "double roundtrip" pattern: our writer produces the XLSX,
/// Excel normalises it by opening and re-saving, and we read the result
/// back.  If Excel rejects or misinterprets anything in our output the
/// assertions in the calling test will catch the difference.
pub fn roundtrip_through_excel(wb: &duke_sheets_core::Workbook) -> duke_sheets_core::Workbook {
    use duke_sheets_xlsx::{XlsxReader, XlsxWriter};
    use std::io::Cursor;

    let input = temp_fixture();
    let output = temp_fixture();

    // Write XLSX bytes with duke-sheets
    let mut buf = Vec::new();
    XlsxWriter::write(wb, Cursor::new(&mut buf)).expect("XlsxWriter::write");
    std::fs::write(&input.host_path, &buf)
        .unwrap_or_else(|e| panic!("write {}: {e}", input.host_path.display()));

    // Push to VM and open in Excel
    ensure_vm_temp_dir();
    push_file_to_vm(&input);

    let bridge = excel_bridge();
    let excel = bridge.lock().unwrap();
    let opened = excel
        .open_workbook(&input.vm_path)
        .expect("Excel should open our file without error");

    // Assert no repair
    let wb_name = opened.name().expect("get workbook name");
    assert!(
        !wb_name.contains("Repaired"),
        "Excel repaired the file! Workbook name: {wb_name}"
    );
    let read_only = opened.is_read_only().expect("get ReadOnly");
    assert!(
        !read_only,
        "Excel opened the file as read-only (possible repair)"
    );

    // Re-save to a second path (Excel normalises the XML)
    opened.save(&output.vm_path).expect("Excel save");
    opened.close().expect("close workbook");

    // Pull the re-saved file back and read with duke-sheets
    pull_file_from_vm(&output);
    let result = XlsxReader::read_file(&output.host_path).expect("XlsxReader::read_file");

    cleanup_fixture(&input);
    cleanup_fixture(&output);

    result
}

// ---------------------------------------------------------------------------
// WinRM helper (raw SOAP/WS-Man over HTTP, Basic auth)
// ---------------------------------------------------------------------------

/// Run a PowerShell command on the VM via WinRM and return stdout.
fn run_winrm_ps(script: &str) -> Result<String, String> {
    let utf16: Vec<u8> = script
        .encode_utf16()
        .flat_map(|c| c.to_le_bytes())
        .collect();
    let encoded = BASE64_STANDARD.encode(&utf16);
    winrm_exec(
        "powershell.exe",
        &format!("-NoProfile -EncodedCommand {encoded}"),
    )
}

/// Execute a command via WinRM SOAP/WS-Man (Basic auth, unencrypted).
fn winrm_exec(program: &str, arguments: &str) -> Result<String, String> {
    let shell_id = winrm_create_shell()?;
    let command_id = winrm_run_command(&shell_id, program, arguments)?;
    let output = winrm_receive(&shell_id, &command_id)?;
    let _ = winrm_delete_shell(&shell_id);
    Ok(output)
}

fn winrm_post(soap: &str) -> Result<String, String> {
    use std::io::Read;
    let auth = BASE64_STANDARD.encode(b"user:test");
    let http_req = format!(
        "POST /wsman HTTP/1.1\r\n\
         Host: localhost:5985\r\n\
         Authorization: Basic {auth}\r\n\
         Content-Type: application/soap+xml;charset=UTF-8\r\n\
         Content-Length: {len}\r\n\
         Connection: close\r\n\
         \r\n\
         {soap}",
        auth = auth,
        len = soap.len(),
        soap = soap,
    );

    let mut stream = std::net::TcpStream::connect("127.0.0.1:5985")
        .map_err(|e| format!("WinRM connect: {e}"))?;
    stream
        .set_read_timeout(Some(std::time::Duration::from_secs(120)))
        .ok();
    std::io::Write::write_all(&mut stream, http_req.as_bytes())
        .map_err(|e| format!("WinRM write: {e}"))?;

    let mut response = String::new();
    stream
        .read_to_string(&mut response)
        .map_err(|e| format!("WinRM read: {e}"))?;

    // Strip HTTP headers
    if let Some(idx) = response.find("\r\n\r\n") {
        Ok(response[idx + 4..].to_string())
    } else {
        Ok(response)
    }
}

fn winrm_create_shell() -> Result<String, String> {
    let soap = soap_envelope(
        "http://schemas.xmlsoap.org/ws/2004/09/transfer/Create",
        None,
        r#"<s:Shell xmlns:s="http://schemas.microsoft.com/wbem/wsman/1/windows/shell">
             <s:InputStreams>stdin</s:InputStreams>
             <s:OutputStreams>stdout stderr</s:OutputStreams>
           </s:Shell>"#,
    );
    let resp = winrm_post(&soap)?;
    extract_xml_value(&resp, "ShellId")
        .ok_or_else(|| format!("No ShellId in response: {}", &resp[..resp.len().min(500)]))
}

fn winrm_run_command(shell_id: &str, program: &str, arguments: &str) -> Result<String, String> {
    let body = format!(
        r#"<rsp:CommandLine xmlns:rsp="http://schemas.microsoft.com/wbem/wsman/1/windows/shell">
             <rsp:Command>{program}</rsp:Command>
             <rsp:Arguments>{arguments}</rsp:Arguments>
           </rsp:CommandLine>"#
    );
    let soap = soap_envelope(
        "http://schemas.microsoft.com/wbem/wsman/1/windows/shell/Command",
        Some(shell_id),
        &body,
    );
    let resp = winrm_post(&soap)?;
    extract_xml_value(&resp, "CommandId")
        .ok_or_else(|| format!("No CommandId in response: {}", &resp[..resp.len().min(500)]))
}

fn winrm_receive(shell_id: &str, command_id: &str) -> Result<String, String> {
    let body = format!(
        r#"<rsp:Receive xmlns:rsp="http://schemas.microsoft.com/wbem/wsman/1/windows/shell" SequenceId="0">
             <rsp:DesiredStream CommandId="{command_id}">stdout</rsp:DesiredStream>
           </rsp:Receive>"#
    );
    let soap = soap_envelope(
        "http://schemas.microsoft.com/wbem/wsman/1/windows/shell/Receive",
        Some(shell_id),
        &body,
    );
    let resp = winrm_post(&soap)?;

    // Extract stdout Stream content (base64-encoded chunks)
    let mut output = String::new();
    for chunk in extract_all_stream_stdout(&resp) {
        if let Ok(bytes) = BASE64_STANDARD.decode(&chunk) {
            if let Ok(s) = String::from_utf8(bytes) {
                output.push_str(&s);
            }
        }
    }
    Ok(output)
}

fn winrm_delete_shell(shell_id: &str) -> Result<(), String> {
    let soap = soap_envelope(
        "http://schemas.xmlsoap.org/ws/2004/09/transfer/Delete",
        Some(shell_id),
        "",
    );
    let _ = winrm_post(&soap);
    Ok(())
}

/// Build a WS-Man SOAP envelope. If `shell_id` is Some, adds a SelectorSet
/// header and the WINRS_SKIP_CMD_SHELL option.
fn soap_envelope(action: &str, shell_id: Option<&str>, body: &str) -> String {
    let uuid = simple_uuid();
    let selector = match shell_id {
        Some(id) => format!(
            r#"<wsman:SelectorSet><wsman:Selector Name="ShellId">{id}</wsman:Selector></wsman:SelectorSet>
               <wsman:OptionSet xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance">
                 <wsman:Option Name="WINRS_CONSOLEMODE_STDIN">TRUE</wsman:Option>
                 <wsman:Option Name="WINRS_SKIP_CMD_SHELL">TRUE</wsman:Option>
               </wsman:OptionSet>"#
        ),
        None => String::new(),
    };
    format!(
        r#"<?xml version="1.0" encoding="UTF-8"?>
<s:Envelope xmlns:s="http://www.w3.org/2003/05/soap-envelope"
            xmlns:wsa="http://schemas.xmlsoap.org/ws/2004/08/addressing"
            xmlns:wsman="http://schemas.dmtf.org/wbem/wsman/1/wsman.xsd">
  <s:Header>
    <wsa:To>http://localhost:5985/wsman</wsa:To>
    <wsman:ResourceURI s:mustUnderstand="true">http://schemas.microsoft.com/wbem/wsman/1/windows/shell/cmd</wsman:ResourceURI>
    <wsa:Action s:mustUnderstand="true">{action}</wsa:Action>
    <wsman:MaxEnvelopeSize s:mustUnderstand="true">153600</wsman:MaxEnvelopeSize>
    <wsa:MessageID>uuid:{uuid}</wsa:MessageID>
    <wsman:Locale xml:lang="en-US" s:mustUnderstand="false"/>
    <wsman:OperationTimeout>PT120S</wsman:OperationTimeout>
    <wsa:ReplyTo><wsa:Address s:mustUnderstand="true">http://schemas.xmlsoap.org/ws/2004/08/addressing/role/anonymous</wsa:Address></wsa:ReplyTo>
    {selector}
  </s:Header>
  <s:Body>{body}</s:Body>
</s:Envelope>"#
    )
}

// ---------------------------------------------------------------------------
// Minimal XML / UUID helpers
// ---------------------------------------------------------------------------

/// Extract a simple XML element value by tag name (handles namespace prefixes).
fn extract_xml_value(xml: &str, tag: &str) -> Option<String> {
    for pat in [format!("<{tag}>"), format!(":{tag}>")] {
        if let Some(start) = xml.find(&pat) {
            let val_start = start + pat.len();
            if let Some(end) = xml[val_start..].find('<') {
                return Some(xml[val_start..val_start + end].to_string());
            }
        }
    }
    None
}

/// Extract all base64-encoded stdout Stream chunks from a WinRM Receive response.
fn extract_all_stream_stdout(xml: &str) -> Vec<String> {
    let mut chunks = Vec::new();
    let mut pos = 0;
    while let Some(idx) = xml[pos..].find("Name=\"stdout\"") {
        let abs = pos + idx;
        if let Some(gt) = xml[abs..].find('>') {
            let content_start = abs + gt + 1;
            if let Some(end) = xml[content_start..].find("</") {
                let b64 = xml[content_start..content_start + end].trim();
                if !b64.is_empty() {
                    chunks.push(b64.to_string());
                }
            }
        }
        pos = abs + 1;
    }
    chunks
}

fn simple_uuid() -> String {
    use std::time::{SystemTime, UNIX_EPOCH};
    let t = SystemTime::now()
        .duration_since(UNIX_EPOCH)
        .unwrap()
        .as_nanos();
    let n = COUNTER.fetch_add(1, Ordering::Relaxed);
    format!(
        "{:08x}-{:04x}-{:04x}-{:04x}-{:012x}",
        (t >> 96) as u32,
        (t >> 80) as u16,
        (t >> 64) as u16,
        n as u16,
        t as u64 & 0xffffffffffff
    )
}

// ---------------------------------------------------------------------------
// Assertion helpers for tests
// ---------------------------------------------------------------------------

/// Assert a cell contains a number (or formula with cached number).
pub fn assert_number(
    sheet: &duke_sheets_core::Worksheet,
    row: u32,
    col: u16,
    expected: f64,
    label: &str,
) {
    let cell = sheet
        .cell_at(row, col)
        .unwrap_or_else(|| panic!("{label} should exist"));
    match &cell.value {
        duke_sheets_core::CellValue::Number(n) => {
            assert!(
                (*n - expected).abs() < 0.001,
                "{label}: expected {expected}, got {n}"
            );
        }
        duke_sheets_core::CellValue::Formula { cached_value, .. } => {
            if let Some(cached) = cached_value {
                match cached.as_ref() {
                    duke_sheets_core::CellValue::Number(n) => {
                        assert!(
                            (*n - expected).abs() < 0.001,
                            "{label}: expected {expected}, got {n} (cached)"
                        );
                    }
                    other => panic!("{label}: expected Number in formula cache, got {other:?}"),
                }
            } else {
                panic!("{label}: formula has no cached value");
            }
        }
        other => panic!("{label}: expected Number, got {other:?}"),
    }
}

/// Assert a cell contains a string.
pub fn assert_string(
    sheet: &duke_sheets_core::Worksheet,
    row: u32,
    col: u16,
    expected: &str,
    label: &str,
) {
    let cell = sheet
        .cell_at(row, col)
        .unwrap_or_else(|| panic!("{label} should exist"));
    match &cell.value {
        duke_sheets_core::CellValue::String(s) => {
            assert_eq!(s.as_ref(), expected, "{label}");
        }
        other => panic!("{label}: expected String, got {other:?}"),
    }
}

/// Assert a cell contains a string (case-insensitive substring match).
pub fn assert_string_contains(
    sheet: &duke_sheets_core::Worksheet,
    row: u32,
    col: u16,
    substring: &str,
    label: &str,
) {
    let cell = sheet
        .cell_at(row, col)
        .unwrap_or_else(|| panic!("{label} should exist"));
    match &cell.value {
        duke_sheets_core::CellValue::String(s) => {
            assert!(
                s.as_ref().contains(substring),
                "{label}: expected to contain '{substring}', got '{}'",
                s.as_ref()
            );
        }
        other => panic!("{label}: expected String, got {other:?}"),
    }
}

/// Assert a cell contains a boolean.
pub fn assert_bool(
    sheet: &duke_sheets_core::Worksheet,
    row: u32,
    col: u16,
    expected: bool,
    label: &str,
) {
    let cell = sheet
        .cell_at(row, col)
        .unwrap_or_else(|| panic!("{label} should exist"));
    match &cell.value {
        duke_sheets_core::CellValue::Boolean(b) => {
            assert_eq!(*b, expected, "{label}");
        }
        other => panic!("{label}: expected Boolean, got {other:?}"),
    }
}

/// Assert a cell has a formula (text preserved).
pub fn assert_has_formula(sheet: &duke_sheets_core::Worksheet, row: u32, col: u16, label: &str) {
    let formula = sheet.get_formula_at(row, col);
    assert!(formula.is_some(), "{label} should have a formula");
}

/// Assert a cell is an error value (via effective_value).
pub fn assert_is_error(sheet: &duke_sheets_core::Worksheet, row: u32, col: u16, label: &str) {
    let cell = sheet
        .cell_at(row, col)
        .unwrap_or_else(|| panic!("{label} should exist"));
    match cell.value.effective_value() {
        duke_sheets_core::CellValue::Error(_) => {}
        other => panic!("{label}: expected Error, got {other:?}"),
    }
}

/// Assert a formula cell has a cached string value.
pub fn assert_formula_string(
    sheet: &duke_sheets_core::Worksheet,
    row: u32,
    col: u16,
    expected: &str,
    label: &str,
) {
    let cell = sheet
        .cell_at(row, col)
        .unwrap_or_else(|| panic!("{label} should exist"));
    match &cell.value {
        duke_sheets_core::CellValue::Formula { cached_value, .. } => {
            if let Some(cached) = cached_value {
                match cached.as_ref() {
                    duke_sheets_core::CellValue::String(s) => {
                        assert_eq!(s.as_ref(), expected, "{label}");
                    }
                    other => panic!("{label}: expected String in formula cache, got {other:?}"),
                }
            } else {
                panic!("{label}: formula has no cached value");
            }
        }
        duke_sheets_core::CellValue::String(s) => {
            assert_eq!(s.as_ref(), expected, "{label}");
        }
        other => panic!("{label}: expected Formula or String, got {other:?}"),
    }
}
