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
use duke_sheets_test_harness::excel::{ensure_excel_bridge, SHARED_DIR as HOST_DIR};

/// VM-side directory where Excel saves files.
const VM_DIR: &str = r"C:\temp";

/// Counter for generating unique temp file names / UUIDs.
static COUNTER: AtomicU64 = AtomicU64::new(0);

/// Global bridge connection, initialized once and shared across tests.
static BRIDGE: OnceLock<Mutex<ExcelBridge>> = OnceLock::new();

/// Get the global Excel bridge, auto-starting the Windows VM on first
/// call if the bridge isn't already responsive. Panics on timeout.
static VM_READY: OnceLock<()> = OnceLock::new();

/// Ensure the Windows VM is fully usable before any interaction with it.
///
/// `ensure_excel_bridge` only waits for the COM bridge (port 9876) to answer
/// an Init ping. On a cold VM boot the bridge comes up minutes before the
/// Windows WinRM service (port 5985), and the test helpers push/pull each
/// fixture over WinRM *before* connecting the COM bridge — so gating only
/// `excel_bridge()` lets the first WinRM call race a still-booting VM and fail
/// with "Connection refused/reset". This must therefore run at the top of
/// every VM-touching entry point (file push/pull, temp-dir setup, and the
/// bridge getter). Mirrors the LibreOffice harness's "real read probe".
fn ensure_vm_ready() {
    VM_READY.get_or_init(|| {
        ensure_excel_bridge(); // spawn qemu + wait for COM bridge (9876)
        wait_for_winrm_ready(); // wait for a real WinRM round-trip (5985)
    });
}

pub fn excel_bridge() -> &'static Mutex<ExcelBridge> {
    BRIDGE.get_or_init(|| {
        ensure_vm_ready();
        let bridge = ExcelBridge::connect(ExcelBridgeConfig::default())
            .expect("Failed to connect to Excel COM bridge on localhost:9876");
        Mutex::new(bridge)
    })
}

/// Poll WinRM (port 5985) with a trivial, side-effect-free command until it
/// completes a full round-trip, so file push/pull won't race a still-booting
/// VM. Panics (does not silently skip) if WinRM never becomes ready.
fn wait_for_winrm_ready() {
    let start = std::time::Instant::now();
    let timeout = std::time::Duration::from_secs(240);
    let mut attempt = 0u32;
    loop {
        if winrm_probe_once(std::time::Duration::from_secs(8)) {
            eprintln!(
                "[e2e] WinRM ready on port 5985 (took {:.1}s)",
                start.elapsed().as_secs_f64()
            );
            return;
        }
        if start.elapsed() >= timeout {
            panic!(
                "WinRM (localhost:5985) did not become ready within {timeout:?}; \
                 the Windows VM may still be booting or WinRM is misconfigured"
            );
        }
        attempt += 1;
        if attempt % 10 == 0 {
            eprintln!(
                "[e2e] still waiting for WinRM ({:.0}s elapsed)...",
                start.elapsed().as_secs_f64()
            );
        }
        std::thread::sleep(std::time::Duration::from_secs(3));
    }
}

/// Run one WinRM readiness probe, bounded to `per_attempt` wall-clock time.
/// `run_winrm_ps` uses a 120s socket read timeout, so a connection that is
/// accepted but stalls (WinRM mid-startup) would otherwise block the retry
/// loop for the full 120s. Running it on a throwaway thread and waiting only
/// `per_attempt` keeps retries responsive for every failure mode — reset
/// (fast Err), refused (fast Err), or stall (abandoned after `per_attempt`).
fn winrm_probe_once(per_attempt: std::time::Duration) -> bool {
    let (tx, rx) = std::sync::mpsc::channel();
    std::thread::spawn(move || {
        let ok = matches!(
            run_winrm_ps("Write-Output winrm-ready"),
            Ok(out) if out.contains("winrm-ready")
        );
        let _ = tx.send(ok);
    });
    rx.recv_timeout(per_attempt).unwrap_or(false)
}

/// A temp file path pair: the host-side path (for reading back) and the
/// Windows path (for telling Excel where to save).
pub struct TempFixture {
    pub host_path: PathBuf,
    pub vm_path: String,
    pub name: String,
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

/// Generate a unique temp fixture with .xlsb extension for XLSB format tests.
pub fn temp_fixture_xlsb() -> TempFixture {
    let n = COUNTER.fetch_add(1, Ordering::SeqCst);
    let pid = std::process::id();
    let name = format!("test_{pid}_{n}.xlsb");
    let _ = std::fs::create_dir_all(HOST_DIR);
    TempFixture {
        host_path: PathBuf::from(format!("{HOST_DIR}/{name}")),
        vm_path: format!("{VM_DIR}\\{name}"),
        name: name.clone(),
    }
}

/// Ensure C:\temp exists inside the VM (called once).
pub fn ensure_vm_temp_dir() {
    ensure_vm_ready();
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
    ensure_vm_ready();
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
    ensure_vm_ready();
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

    let verify_ps = format!(
        "(Get-Item '{}' -ErrorAction SilentlyContinue).Length",
        fixture.vm_path
    );
    let verify = run_winrm_ps(&verify_ps).unwrap_or_default();
    eprintln!(
        "[push_file_to_vm] {} -> {} (VM size: {})",
        fixture.host_path.display(),
        fixture.vm_path,
        verify.trim()
    );
    assert!(
        verify.trim().parse::<usize>().unwrap_or(0) > 0,
        "File push failed: {} not found on VM",
        fixture.vm_path
    );
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

// Writer E2E helper: duke-sheets write → Excel open+re-save → duke-sheets read

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

/// Write a workbook as XLS (BIFF8) with duke-sheets, push to the VM, open
/// in real Excel (asserting no repair), re-save as XLS (FileFormat=56),
/// pull back, and read with `XlsReader`.
pub fn roundtrip_through_excel_xls(wb: &duke_sheets_core::Workbook) -> duke_sheets_core::Workbook {
    roundtrip_through_excel_xls_bytes(wb).0
}

/// XLS writer parity helper that also returns the pre-Excel writer bytes and
/// Excel's re-saved bytes. Tests use the bytes to compare BIFF formula token
/// streams against Excel's canonical save output without another VM trip.
pub fn roundtrip_through_excel_xls_bytes(
    wb: &duke_sheets_core::Workbook,
) -> (duke_sheets_core::Workbook, Vec<u8>, Vec<u8>) {
    use duke_sheets_xls::{XlsReader, XlsWriter};

    let input = temp_fixture_xls();
    let output = temp_fixture_xls();

    let buf = XlsWriter::write_to_bytes(wb).expect("XlsWriter::write_to_bytes");
    std::fs::write(&input.host_path, &buf)
        .unwrap_or_else(|e| panic!("write {}: {e}", input.host_path.display()));

    ensure_vm_temp_dir();
    push_file_to_vm(&input);

    let bridge = excel_bridge();
    let excel = bridge.lock().unwrap();
    let opened = excel
        .open_workbook(&input.vm_path)
        .expect("Excel should open our XLS without error");

    let wb_name = opened.name().expect("get workbook name");
    assert!(
        !wb_name.contains("Repaired"),
        "Excel repaired the XLS file! Workbook name: {wb_name}"
    );

    // FileFormat=56 = xlExcel8 (.xls / BIFF8 / Excel 97-2003)
    opened
        .save_as(&output.vm_path, 56)
        .expect("Excel SaveAs xls");
    opened.close().expect("close workbook");

    pull_file_from_vm(&output);
    let output_bytes = std::fs::read(&output.host_path)
        .unwrap_or_else(|e| panic!("read {}: {e}", output.host_path.display()));
    let result = XlsReader::read_file(&output.host_path).expect("XlsReader::read_file");

    cleanup_fixture(&input);
    cleanup_fixture(&output);

    (result, buf, output_bytes)
}

#[derive(Debug, Clone, PartialEq, Eq)]
pub struct XlsFormulaPtgStream {
    pub row: u16,
    pub col: u16,
    pub tokens: Vec<u8>,
    /// The rgcb (extra data) that follows the rgce — array-constant element
    /// data, etc. Empty for formulas without such data.
    pub rgcb: Vec<u8>,
}

/// Extract raw FORMULA-record PTG bytes keyed by cell position.
///
/// This deliberately preserves opcode/class bytes instead of going through the
/// normal formula parser, because token class is exactly what byte-parity tests
/// need to pin.
pub fn xls_formula_ptg_streams(bytes: &[u8]) -> Vec<XlsFormulaPtgStream> {
    use duke_sheets_xls::{biff, cfb::CompoundFile};
    use std::io::Cursor;

    let cfb = CompoundFile::open(Cursor::new(bytes)).expect("open XLS CFB");
    let stream_path = if cfb.exists("/Workbook") {
        "/Workbook"
    } else {
        "/Book"
    };
    let stream = cfb.read_stream(stream_path).expect("read workbook stream");
    let records = biff::read_all_records(&mut Cursor::new(stream)).expect("read BIFF records");
    records
        .iter()
        .filter(|rec| rec.record_type == biff::records::FORMULA)
        .map(|rec| {
            assert!(rec.data.len() >= 22, "FORMULA record too short");
            let row = u16::from_le_bytes([rec.data[0], rec.data[1]]);
            let col = u16::from_le_bytes([rec.data[2], rec.data[3]]);
            let cce = u16::from_le_bytes([rec.data[20], rec.data[21]]) as usize;
            assert!(
                rec.data.len() >= 22 + cce,
                "FORMULA record token stream truncated"
            );
            XlsFormulaPtgStream {
                row,
                col,
                tokens: rec.data[22..22 + cce].to_vec(),
                rgcb: rec.data[22 + cce..].to_vec(),
            }
        })
        .collect()
}

/// Extract FORMULA PTGs in the form used for byte parity checks.
///
/// Identical to [`xls_formula_ptg_streams`] except that the two bytes after
/// a PtgAttrSum's flags byte are zeroed. MS-XLS §2.5.198.41 PtgAttrSum
/// defines those two bytes as `unused (2 bytes): Undefined and MUST be
/// ignored.` so authored files may carry uninitialized stack values there.
/// Likewise MS-XLS §2.5.198.42 PtgAttrVolatile defines two `unused` bytes
/// in the same position; zero them out as well so a future Excel emission
/// that puts garbage there doesn't make tests flake.
pub fn xls_formula_ptg_streams_for_compare(bytes: &[u8]) -> Vec<XlsFormulaPtgStream> {
    xls_formula_ptg_streams(bytes)
        .into_iter()
        .map(|mut stream| {
            normalize_attr_reserved_bytes(&mut stream.tokens);
            stream
        })
        .collect()
}

/// Extract the raw bodies of all EXTERNNAME (0x0023) records, in file order.
///
/// Used to byte-compare Analysis-ToolPak add-in function EXTERNNAME emission
/// against Excel's native output — the FORMULA-stream comparison only pins the
/// PtgNameX `nameindex`, not the external-name record contents themselves.
pub fn xls_externname_record_bodies(bytes: &[u8]) -> Vec<Vec<u8>> {
    use duke_sheets_xls::{biff, cfb::CompoundFile};
    use std::io::Cursor;

    let cfb = CompoundFile::open(Cursor::new(bytes)).expect("open XLS CFB");
    let stream_path = if cfb.exists("/Workbook") {
        "/Workbook"
    } else {
        "/Book"
    };
    let stream = cfb.read_stream(stream_path).expect("read workbook stream");
    let records = biff::read_all_records(&mut Cursor::new(stream)).expect("read BIFF records");
    records
        .iter()
        .filter(|rec| rec.record_type == biff::records::EXTERNNAME)
        .map(|rec| rec.data.clone())
        .collect()
}

/// Every Analysis-ToolPak add-in function (Ftab 384..=476) paired with a
/// minimal valid call: `(cell, formula)` with the function's `min_args`
/// arguments drawn from A1.. so a single workbook exercises the whole range
/// in one Excel round-trip.
///
/// Emitted in Ftab-index order (NOT alphabetical), so the comparison also
/// stress-tests the XLS EXTERNNAME ordering / nameindex assignment — a
/// scrambled first-use order would diverge from our alphabetical sort if the
/// assumption were wrong.
///
pub fn atp_all_formulas() -> Vec<(String, String)> {
    use duke_sheets_xls::biff::formula::function_table::{function_min_args, function_name};

    let mut out = Vec::new();
    let mut row = 1u32;
    for idx in 384u16..=476 {
        let name = function_name(idx);
        if name.is_empty() {
            continue;
        }
        let n = function_min_args(idx).unwrap_or(1).max(1);
        let args: Vec<String> = (1..=n).map(|i| format!("A{i}")).collect();
        out.push((format!("B{row}"), format!("={}({})", name, args.join(","))));
        row += 1;
    }
    out
}

/// Zero out the per-MS-XLS "undefined" bytes inside PtgAttrSum / PtgAttrVolatile
/// so test comparisons aren't sensitive to whatever Excel left in them.
fn normalize_attr_reserved_bytes(tokens: &mut [u8]) {
    use duke_sheets_xls::biff::formula::ptg;

    let mut pos = 0usize;
    while pos < tokens.len() {
        let raw = tokens[pos];
        let base = ptg::base_ptg(raw);
        pos += 1;
        match base {
            ptg::PTG_ATTR => {
                if pos + 3 > tokens.len() {
                    break;
                }
                let flags = tokens[pos];
                let attr_data = u16::from_le_bytes([tokens[pos + 1], tokens[pos + 2]]) as usize;
                // PtgAttrSum (§2.5.198.41) and PtgAttrVolatile (§2.5.198.42)
                // both have 2 unused bytes right after the flags byte.
                if (flags & (ptg::ATTR_SUM | ptg::ATTR_VOLATILE)) != 0 {
                    tokens[pos + 1] = 0;
                    tokens[pos + 2] = 0;
                }
                pos += 3;
                if (flags & ptg::ATTR_CHOOSE) != 0 {
                    pos = pos.saturating_add((attr_data + 1) * 2);
                }
            }
            ptg::PTG_ARRAY => {
                // PtgArray (§2.5.198.32) has 7 unused bytes after the opcode;
                // Excel may leave uninitialized stack values there. Zero them.
                let n = ptg::token_data_size(ptg::PTG_ARRAY).unwrap_or(7);
                let end = (pos + n).min(tokens.len());
                for b in &mut tokens[pos..end] {
                    *b = 0;
                }
                pos = pos.saturating_add(n);
            }
            ptg::PTG_STR => {
                if pos + 2 > tokens.len() {
                    break;
                }
                let len = tokens[pos] as usize;
                let flags = tokens[pos + 1];
                pos += 2 + if (flags & 0x01) != 0 { len * 2 } else { len };
            }
            _ => match ptg::token_data_size(base) {
                Some(size) => pos = pos.saturating_add(size),
                None => break,
            },
        }
    }
}

/// Write a workbook as XLSB with duke-sheets, push to the VM, open in real
/// Excel (asserting no repair), re-save as XLSB (FileFormat=50), pull back,
/// and read with `XlsbReader`.
pub fn roundtrip_through_excel_xlsb(wb: &duke_sheets_core::Workbook) -> duke_sheets_core::Workbook {
    use duke_sheets_xlsb::{XlsbReader, XlsbWriter};
    use std::io::Cursor;

    let input = temp_fixture_xlsb();
    let output = temp_fixture_xlsb();

    let mut buf = Vec::new();
    XlsbWriter::write(wb, Cursor::new(&mut buf)).expect("XlsbWriter::write");
    std::fs::write(&input.host_path, &buf)
        .unwrap_or_else(|e| panic!("write {}: {e}", input.host_path.display()));

    ensure_vm_temp_dir();
    push_file_to_vm(&input);

    let bridge = excel_bridge();
    let excel = bridge.lock().unwrap();
    let opened = excel
        .open_workbook(&input.vm_path)
        .expect("Excel should open our XLSB without error");

    let wb_name = opened.name().expect("get workbook name");
    assert!(
        !wb_name.contains("Repaired"),
        "Excel repaired the XLSB file! Workbook name: {wb_name}"
    );

    // FileFormat=50 = xlExcel12 (.xlsb)
    opened
        .save_as(&output.vm_path, 50)
        .expect("Excel SaveAs xlsb");
    opened.close().expect("close workbook");

    pull_file_from_vm(&output);
    let result = XlsbReader::read_file(&output.host_path).expect("XlsbReader::read_file");

    cleanup_fixture(&input);
    cleanup_fixture(&output);

    result
}

/// Like [`roundtrip_through_excel_xlsb`] but also returns the raw writer
/// bytes and Excel's re-saved bytes, for byte-parity comparison of the
/// formula PTG streams. Mirrors [`roundtrip_through_excel_xls_bytes`].
pub fn roundtrip_through_excel_xlsb_bytes(
    wb: &duke_sheets_core::Workbook,
) -> (duke_sheets_core::Workbook, Vec<u8>, Vec<u8>) {
    use duke_sheets_xlsb::{XlsbReader, XlsbWriter};
    use std::io::Cursor;

    let input = temp_fixture_xlsb();
    let output = temp_fixture_xlsb();

    let mut buf = Vec::new();
    XlsbWriter::write(wb, Cursor::new(&mut buf)).expect("XlsbWriter::write");
    std::fs::write(&input.host_path, &buf)
        .unwrap_or_else(|e| panic!("write {}: {e}", input.host_path.display()));

    ensure_vm_temp_dir();
    push_file_to_vm(&input);

    let bridge = excel_bridge();
    let excel = bridge.lock().unwrap();
    let opened = excel
        .open_workbook(&input.vm_path)
        .expect("Excel should open our XLSB without error");

    let wb_name = opened.name().expect("get workbook name");
    assert!(
        !wb_name.contains("Repaired"),
        "Excel repaired the XLSB file! Workbook name: {wb_name}"
    );

    opened
        .save_as(&output.vm_path, 50)
        .expect("Excel SaveAs xlsb");
    opened.close().expect("close workbook");

    pull_file_from_vm(&output);
    let output_bytes = std::fs::read(&output.host_path)
        .unwrap_or_else(|e| panic!("read {}: {e}", output.host_path.display()));
    let result = XlsbReader::read_file(&output.host_path).expect("XlsbReader::read_file");

    cleanup_fixture(&input);
    cleanup_fixture(&output);

    (result, buf, output_bytes)
}

#[derive(Debug, Clone, PartialEq, Eq)]
pub struct XlsbFormulaPtgStream {
    pub row: u32,
    pub col: u32,
    pub rgce: Vec<u8>,
    /// The rgcb (extra data) that follows the rgce in BrtCellFmla — array
    /// constant element data, etc. Empty for formulas without such data.
    pub rgcb: Vec<u8>,
}

/// Extract raw BIFF12 formula token (`rgce`) bytes keyed by cell, from every
/// worksheet binary part inside an `.xlsb` zip. Preserves opcode/class bytes
/// (like the XLS helper) so byte-parity tests can pin token class.
///
/// Walks `BrtRowHdr` to track the current row and `BrtCellFmla{Num,Bool,
/// Error,String}` records for the formula structure
/// (`grbit(2) + cce(u32) + rgce + cb(u32) + rgcb`). Results are sorted by
/// (row, col).
pub fn xlsb_formula_ptg_streams(bytes: &[u8]) -> Vec<XlsbFormulaPtgStream> {
    use duke_sheets_xlsb::biff12::{records, RecordIter};
    use std::io::Cursor;

    let mut streams = Vec::new();
    let reader = Cursor::new(bytes);
    let mut zip = zip::ZipArchive::new(reader).expect("open xlsb zip");

    // Collect worksheet bin names first (borrow of zip is released before reads).
    let sheet_names: Vec<String> = (0..zip.len())
        .filter_map(|i| {
            let f = zip.by_index(i).ok()?;
            let name = f.name().to_string();
            if name.starts_with("xl/worksheets/") && name.ends_with(".bin") {
                Some(name)
            } else {
                None
            }
        })
        .collect();

    for name in sheet_names {
        let mut data = Vec::new();
        {
            use std::io::Read;
            let mut f = zip.by_name(&name).expect("read sheet bin");
            f.read_to_end(&mut data).expect("read sheet bin bytes");
        }

        let mut iter = RecordIter::new(Cursor::new(&data));
        let mut buf = Vec::new();
        let mut current_row: u32 = 0;
        loop {
            let (rec_type, len) = match iter.next_record(&mut buf) {
                Ok(t) => t,
                Err(_) => break,
            };
            let payload = &buf[..len];
            match rec_type {
                records::BRT_ROW_HDR => {
                    if payload.len() >= 4 {
                        current_row =
                            u32::from_le_bytes([payload[0], payload[1], payload[2], payload[3]]);
                    }
                }
                records::BRT_FMLA_NUM
                | records::BRT_FMLA_BOOL
                | records::BRT_FMLA_ERROR
                | records::BRT_FMLA_STRING => {
                    let Some((col, grbit_off)) = fmla_col_and_grbit_offset(rec_type, payload)
                    else {
                        continue;
                    };
                    let cce_off = grbit_off + 2;
                    if cce_off + 4 > payload.len() {
                        continue;
                    }
                    let cce = u32::from_le_bytes([
                        payload[cce_off],
                        payload[cce_off + 1],
                        payload[cce_off + 2],
                        payload[cce_off + 3],
                    ]) as usize;
                    let rgce_start = cce_off + 4;
                    if rgce_start + cce > payload.len() {
                        continue;
                    }
                    // rgcb follows the rgce: cb(u32) + rgcb(cb bytes).
                    let cb_off = rgce_start + cce;
                    let rgcb = if cb_off + 4 <= payload.len() {
                        let cb = u32::from_le_bytes([
                            payload[cb_off],
                            payload[cb_off + 1],
                            payload[cb_off + 2],
                            payload[cb_off + 3],
                        ]) as usize;
                        let start = cb_off + 4;
                        if start + cb <= payload.len() {
                            payload[start..start + cb].to_vec()
                        } else {
                            Vec::new()
                        }
                    } else {
                        Vec::new()
                    };
                    streams.push(XlsbFormulaPtgStream {
                        row: current_row,
                        col,
                        rgce: payload[rgce_start..rgce_start + cce].to_vec(),
                        rgcb,
                    });
                }
                _ => {}
            }
        }
    }

    streams.sort_by_key(|s| (s.row, s.col));
    streams
}

/// Column index and the byte offset of the formula `grbit` for a
/// `BrtCellFmla*` record. The Cell header is `col(u32) + iStyleRef/flags(u32)`
/// = 8 bytes, followed by the cached value whose width depends on the record
/// type; the formula structure begins right after.
fn fmla_col_and_grbit_offset(rec_type: u16, payload: &[u8]) -> Option<(u32, usize)> {
    use duke_sheets_xlsb::biff12::records;
    if payload.len() < 8 {
        return None;
    }
    let col = u32::from_le_bytes([payload[0], payload[1], payload[2], payload[3]]);
    let grbit_off = match rec_type {
        records::BRT_FMLA_NUM => 16,  // 8 + xnum(8)
        records::BRT_FMLA_BOOL => 9,  // 8 + bool(1)
        records::BRT_FMLA_ERROR => 9, // 8 + err(1)
        records::BRT_FMLA_STRING => {
            // 8 + XLWideString(cch u32 + UTF-16LE chars)
            if payload.len() < 12 {
                return None;
            }
            let cch =
                u32::from_le_bytes([payload[8], payload[9], payload[10], payload[11]]) as usize;
            12 + cch * 2
        }
        _ => return None,
    };
    Some((col, grbit_off))
}

/// XLSB formula PTG streams in the form used for byte parity checks: the two
/// "undefined" bytes after a PtgAttrSum / PtgAttrVolatile flags byte are
/// zeroed (same rationale as the XLS variant — Excel may carry uninitialized
/// stack values there). Requires a BIFF12 token walk because reserved bytes
/// can appear mid-stream.
pub fn xlsb_formula_ptg_streams_for_compare(bytes: &[u8]) -> Vec<XlsbFormulaPtgStream> {
    xlsb_formula_ptg_streams(bytes)
        .into_iter()
        .map(|mut s| {
            normalize_xlsb_attr_reserved_bytes(&mut s.rgce);
            s
        })
        .collect()
}

/// Zero the per-MS-XLS "undefined" bytes inside PtgAttrSum / PtgAttrVolatile
/// in a BIFF12 token stream. Walks tokens using BIFF12 sizes (wider refs than
/// BIFF8).
fn normalize_xlsb_attr_reserved_bytes(tokens: &mut [u8]) {
    use duke_sheets_xlsb::biff12::ptg;

    let mut pos = 0usize;
    while pos < tokens.len() {
        let base = ptg::base_ptg(tokens[pos]);
        pos += 1;
        match base {
            ptg::PTG_ATTR => {
                if pos + 3 > tokens.len() {
                    break;
                }
                let flags = tokens[pos];
                let attr_data = u16::from_le_bytes([tokens[pos + 1], tokens[pos + 2]]) as usize;
                if (flags & (ptg::ATTR_SUM | ptg::ATTR_VOLATILE)) != 0 {
                    tokens[pos + 1] = 0;
                    tokens[pos + 2] = 0;
                }
                pos += 3;
                if (flags & ptg::ATTR_CHOOSE) != 0 {
                    pos = pos.saturating_add((attr_data + 1) * 2);
                }
            }
            ptg::PTG_ARRAY => {
                // BIFF12 PtgArray has 14 unused bytes after the opcode; Excel
                // leaves uninitialized stack values there. Zero them.
                let n = 14usize;
                let end = (pos + n).min(tokens.len());
                for b in &mut tokens[pos..end] {
                    *b = 0;
                }
                pos = pos.saturating_add(n);
            }
            ptg::PTG_STR => {
                if pos + 2 > tokens.len() {
                    break;
                }
                let cch = u16::from_le_bytes([tokens[pos], tokens[pos + 1]]) as usize;
                pos += 2 + cch * 2;
            }
            _ => match biff12_token_data_size(base) {
                Some(size) => pos = pos.saturating_add(size),
                None => break,
            },
        }
    }
}

/// Data byte count (excluding the 1 opcode byte) for fixed-size BIFF12 tokens
/// we emit. `None` for variable/unknown tokens (walk stops). Sizes verified
/// against `duke_sheets_xlsb::biff12::token_parser`.
fn biff12_token_data_size(base: u8) -> Option<usize> {
    use duke_sheets_xlsb::biff12::ptg;
    Some(match base {
        // Operators / no-data tokens
        ptg::PTG_ADD
        | ptg::PTG_SUB
        | ptg::PTG_MUL
        | ptg::PTG_DIV
        | ptg::PTG_POWER
        | ptg::PTG_CONCAT
        | ptg::PTG_LT
        | ptg::PTG_LE
        | ptg::PTG_EQ
        | ptg::PTG_GE
        | ptg::PTG_GT
        | ptg::PTG_NE
        | ptg::PTG_ISECT
        | ptg::PTG_LIST
        | ptg::PTG_RANGE
        | ptg::PTG_UPLUS
        | ptg::PTG_UMINUS
        | ptg::PTG_PERCENT
        | ptg::PTG_PAREN
        | ptg::PTG_MISS_ARG => 0,
        ptg::PTG_ERR | ptg::PTG_BOOL => 1,
        ptg::PTG_INT => 2,
        ptg::PTG_NUM => 8,
        ptg::PTG_FUNC => 2,
        ptg::PTG_FUNC_VAR => 3,
        ptg::PTG_NAME => 4,
        ptg::PTG_MEM_FUNC => 2,
        ptg::PTG_REF | ptg::PTG_REF_ERR => 6,
        ptg::PTG_AREA | ptg::PTG_AREA_ERR => 12,
        ptg::PTG_NAME_X => 6,
        ptg::PTG_REF_3D | ptg::PTG_REF_ERR_3D => 8,
        ptg::PTG_AREA_3D | ptg::PTG_AREA_ERR_3D => 14,
        _ => return None,
    })
}

// WinRM helper (raw SOAP/WS-Man over HTTP, Basic auth)

/// Run a PowerShell command on the VM via WinRM and return stdout.
pub fn run_winrm_ps(script: &str) -> Result<String, String> {
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

// Minimal XML / UUID helpers

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

// Assertion helpers for tests

pub fn assert_number(
    sheet: &duke_sheets_core::Worksheet,
    row: u32,
    col: u16,
    expected: f64,
    label: &str,
) {
    let value = sheet.get_value_at(row, col);
    match value {
        duke_sheets_core::CellValue::Number(n) => {
            assert!(
                (n - expected).abs() < 0.001,
                "{label}: expected {expected}, got {n}"
            );
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
    let value = sheet.get_value_at(row, col);
    match value {
        duke_sheets_core::CellValue::String(s) => {
            assert_eq!(s.as_str(), expected, "{label}");
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
    let value = sheet.get_value_at(row, col);
    match value {
        duke_sheets_core::CellValue::String(s) => {
            assert!(
                s.as_str().contains(substring),
                "{label}: expected to contain '{substring}', got '{}'",
                s.as_str()
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
    let value = sheet.get_value_at(row, col);
    match value {
        duke_sheets_core::CellValue::Boolean(b) => {
            assert_eq!(b, expected, "{label}");
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
    match sheet.get_value_at(row, col).effective_value() {
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
    assert!(
        sheet.get_formula_at(row, col).is_some(),
        "{label}: expected formula"
    );
    match sheet.get_value_at(row, col) {
        duke_sheets_core::CellValue::String(s) => {
            assert_eq!(s.as_str(), expected, "{label}");
        }
        other => panic!("{label}: expected cached String, got {other:?}"),
    }
}
