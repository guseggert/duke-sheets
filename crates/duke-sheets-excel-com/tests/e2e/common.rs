//! Common utilities for Excel COM E2E tests.
//!
//! Provides a global bridge connection singleton and helper functions
//! for temp file management. The bridge is shared across all tests
//! because Excel COM startup is slow (~2s).
//!
//! File transfer: Excel saves to `C:\temp\` inside the VM, and the test
//! moves the file over the bridge connection itself (`PutFile`/`GetFile`/
//! `DeleteFile`), reads it with `XlsxReader`, and cleans up both sides.
//!
//! This used to go over WinRM, which cost ~420ms per SOAP session; pushing a
//! fixture needed ~50 of them because the base64 payload had to be chunked
//! under the command-line limit. That transfer overhead, not Excel, was the
//! bulk of this suite's runtime.

use std::path::PathBuf;
use std::sync::atomic::{AtomicU64, Ordering};
use std::sync::OnceLock;

use duke_sheets_excel_com::{ExcelBridge, ExcelBridgeConfig};
use duke_sheets_test_harness::excel::{ensure_excel_bridge, SHARED_DIR as HOST_DIR};

/// VM-side directory where Excel saves files.
const VM_DIR: &str = r"C:\temp";

/// Counter for generating unique temp file names / UUIDs.
static COUNTER: AtomicU64 = AtomicU64::new(0);

/// Global bridge connection, initialized once and shared across tests.
static BRIDGE: OnceLock<BridgeCell> = OnceLock::new();

/// Holder for the shared bridge.
///
/// `ExcelBridge::send_command` serializes whole exchanges internally, so a
/// `&ExcelBridge` is already safe to share and no outer `Mutex` is needed.
/// The `lock()` shim keeps the several hundred existing
/// `bridge.lock().unwrap()` call sites compiling, and — unlike a real
/// `Mutex` — it cannot deadlock when a helper such as `pull_file_from_vm`
/// reaches for the bridge while a test still holds its "guard".
pub struct BridgeCell(ExcelBridge);

impl BridgeCell {
    pub fn lock(&self) -> Result<&ExcelBridge, std::convert::Infallible> {
        Ok(&self.0)
    }
}

pub fn excel_bridge() -> &'static BridgeCell {
    BRIDGE.get_or_init(|| {
        ensure_excel_bridge();
        let bridge = ExcelBridge::connect(ExcelBridgeConfig::default())
            .expect("Failed to connect to Excel COM bridge on localhost:9876");
        BridgeCell(bridge)
    })
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
    static DONE: std::sync::Once = std::sync::Once::new();
    DONE.call_once(|| {
        let bridge = excel_bridge();
        let excel = bridge.lock().unwrap();
        excel
            .create_dir(VM_DIR)
            .unwrap_or_else(|e| panic!("create {VM_DIR} in VM: {e}"));
    });
}

/// Pull a file from the VM to the host over the bridge connection.
pub fn pull_file_from_vm(fixture: &TempFixture) {
    let bridge = excel_bridge();
    let excel = bridge.lock().unwrap();
    let bytes = excel
        .get_file(&fixture.vm_path)
        .unwrap_or_else(|e| panic!("pull {} from VM: {e}", fixture.vm_path));
    std::fs::write(&fixture.host_path, &bytes)
        .unwrap_or_else(|e| panic!("write {}: {e}", fixture.host_path.display()));
}

/// Push a file from the host to the VM over the bridge connection.
pub fn push_file_to_vm(fixture: &TempFixture) {
    let bytes = std::fs::read(&fixture.host_path)
        .unwrap_or_else(|e| panic!("read {}: {e}", fixture.host_path.display()));
    let bridge = excel_bridge();
    let excel = bridge.lock().unwrap();
    excel
        .put_file(&fixture.vm_path, &bytes)
        .unwrap_or_else(|e| panic!("push {} to VM: {e}", fixture.vm_path));
}

/// Clean up fixture files on both host and VM. Ignores errors.
pub fn cleanup_fixture(fixture: &TempFixture) {
    let _ = std::fs::remove_file(&fixture.host_path);
    let bridge = excel_bridge();
    let excel = bridge.lock().unwrap();
    let _ = excel.delete_file(&fixture.vm_path);
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
    use duke_sheets_xls::{XlsReadOptions, XlsReader, XlsWriter};

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
    let result = XlsReader::read_file_with(&output.host_path, &XlsReadOptions { try_velvet_sweatshop: true, ..Default::default() })
        .expect("XlsReader::read_file_with");

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
                    let Some((col, grbit_off)) = fmla_col_and_grbit_offset(rec_type, payload) else {
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
        records::BRT_FMLA_NUM => 16,   // 8 + xnum(8)
        records::BRT_FMLA_BOOL => 9,   // 8 + bool(1)
        records::BRT_FMLA_ERROR => 9,  // 8 + err(1)
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
                let attr_data =
                    u16::from_le_bytes([tokens[pos + 1], tokens[pos + 2]]) as usize;
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
        ptg::PTG_ADD | ptg::PTG_SUB | ptg::PTG_MUL | ptg::PTG_DIV | ptg::PTG_POWER
        | ptg::PTG_CONCAT | ptg::PTG_LT | ptg::PTG_LE | ptg::PTG_EQ | ptg::PTG_GE
        | ptg::PTG_GT | ptg::PTG_NE | ptg::PTG_ISECT | ptg::PTG_LIST | ptg::PTG_RANGE
        | ptg::PTG_UPLUS | ptg::PTG_UMINUS | ptg::PTG_PERCENT | ptg::PTG_PAREN
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
