//! XLS (BIFF8) writer.
//!
//! Currently emits the structural minimum required for round-tripping
//! a [`Workbook`] through [`XlsReader`](crate::reader::XlsReader): one
//! globals stream containing a `BOF`, a `BoundSheet8` per worksheet,
//! and an `EOF`, followed by one (empty) worksheet stream per
//! `BoundSheet8`. Cell values, formatting, formulas, and comments are
//! deliberately not emitted yet — they land in subsequent slices.

use std::collections::{BTreeMap, BTreeSet, HashMap, HashSet};
use std::io::Write;
use std::path::Path;

use duke_sheets_core::style::{
    Alignment, BorderEdge, BorderLineStyle, BorderStyle, Color, DiagonalDirection, FillStyle,
    FontStyle, HorizontalAlignment, NumberFormat, PatternType, ReadingOrder, Underline,
    VerticalAlignment,
};
use duke_sheets_core::workbook::Workbook;
use duke_sheets_core::worksheet::Worksheet;
use duke_sheets_core::{
    CellValue, PivotAggregate, PivotDateGroupUnit, PivotField, PivotFilter, PivotGrouping,
    PivotManualGroup, PivotValue, PivotValuesAxis,
};
use duke_sheets_pivot::{FormatPivotCache, FormatPivotPlan, FormatPivotSource};
use ssfmt::{
    date_serial::{serial_to_date, serial_to_time},
    DateSystem,
};

use crate::cfb::CompoundFileBuilder;
use crate::error::{XlsError, XlsResult};

const BOF_RECORD: u16 = 0x0809;
const EOF_RECORD: u16 = 0x000A;
const CONTINUE_RECORD: u16 = 0x003C;
const DATEMODE_RECORD: u16 = 0x0022;
const BOUND_SHEET_8: u16 = 0x0085;
const COUNTRY_RECORD: u16 = 0x008C;
const SST_RECORD: u16 = 0x00FC;
const EXTSST_RECORD: u16 = 0x00FF;
const FONT_RECORD: u16 = 0x0031;
const FORMAT_RECORD: u16 = 0x041E;
const XF_RECORD: u16 = 0x00E0;
const DIMENSION_RECORD: u16 = 0x0200;
const WINDOW2_RECORD: u16 = 0x023E;
const CALCCOUNT_RECORD: u16 = 0x000C;
const CALCMODE_RECORD: u16 = 0x000D;
const REFMODE_RECORD: u16 = 0x000F;
const DELTA_RECORD: u16 = 0x0010;
const ITERATION_RECORD: u16 = 0x0011;
const BLANK_RECORD: u16 = 0x0201;
const NUMBER_RECORD: u16 = 0x0203;
const BOOLERR_RECORD: u16 = 0x0205;
const LABELSST_RECORD: u16 = 0x00FD;
const FORMULA_RECORD: u16 = 0x0006;
const STRING_RECORD: u16 = 0x0207;
const MERGECELLS_RECORD: u16 = 0x00E5;
const INDEX_RECORD: u16 = 0x020B;
const ROW_RECORD: u16 = 0x0208;
const DBCELL_RECORD: u16 = 0x00D7;
const COLINFO_RECORD: u16 = 0x007D;
const PANE_RECORD: u16 = 0x0041;
const WINDOW1_RECORD: u16 = 0x003D;
const DEFCOLWIDTH_RECORD: u16 = 0x0055;
const SAVERECALC_RECORD: u16 = 0x005F;
const GUTS_RECORD: u16 = 0x0080;
const WSBOOL_RECORD: u16 = 0x0081;
const GRIDSET_RECORD: u16 = 0x0082;
const HCENTER_RECORD: u16 = 0x0083;
const VCENTER_RECORD: u16 = 0x0084;
const PROTECT_RECORD: u16 = 0x0012;
const PASSWORD_RECORD: u16 = 0x0013;
const DEFAULTROWHEIGHT_RECORD: u16 = 0x0225;
const SETUP_RECORD: u16 = 0x00A1;
const HEADER_RECORD: u16 = 0x0014;
const FOOTER_RECORD: u16 = 0x0015;
const LEFT_MARGIN_RECORD: u16 = 0x0026;
const RIGHT_MARGIN_RECORD: u16 = 0x0027;
const TOP_MARGIN_RECORD: u16 = 0x0028;
const BOTTOM_MARGIN_RECORD: u16 = 0x0029;
const PRINTHEADERS_RECORD: u16 = 0x002A;
const PRINTGRIDLINES_RECORD: u16 = 0x002B;
const HPAGEBREAKS_RECORD: u16 = 0x001B;
const VPAGEBREAKS_RECORD: u16 = 0x001A;
const SELECTION_RECORD: u16 = 0x001D;
const SCL_RECORD: u16 = 0x00A0;
const HLINK_RECORD: u16 = 0x01B8;
const NAME_RECORD: u16 = 0x0018;
const SUPBOOK_RECORD: u16 = 0x01AE;
const EXTERNSHEET_RECORD: u16 = 0x0017;
const EXTERNNAME_RECORD: u16 = 0x0023;
const AUTOFILTER_RECORD: u16 = 0x009E;
const FILTERMODE_RECORD: u16 = 0x009B;
const DVAL_RECORD: u16 = 0x01B2;
const DV_RECORD: u16 = 0x01BE;
const CONDFMT_RECORD: u16 = 0x01B0;
const CF_RECORD: u16 = 0x01B1;
const MSODRAWINGGROUP_RECORD: u16 = 0x00EB;
const MSODRAWING_RECORD: u16 = 0x00EC;
const NOTE_RECORD: u16 = 0x001C;
const OBJ_RECORD: u16 = 0x005D;
const TXO_RECORD: u16 = 0x01B6;
const PIVOT_SXFMLA_RECORD: u16 = 0x00F8;
const PIVOT_SXNAME_RECORD: u16 = 0x00F5;
const PIVOT_SXFDB_RECORD: u16 = 0x00C7;
const PIVOT_SXDBB_RECORD: u16 = 0x00C8;
const PIVOT_SXNUM_RECORD: u16 = 0x00C9;
const PIVOT_SXINT_RECORD: u16 = 0x00CC;
const PIVOT_SXSTRING_RECORD: u16 = 0x00CD;
const PIVOT_SXDTR_RECORD: u16 = 0x00CE;
const PIVOT_SXRNG_RECORD: u16 = 0x00D8;
const PIVOT_SXIDSTM_RECORD: u16 = 0x00D9;
const BOOKEXT_RECORD: u16 = 0x0863;
const COMPAT12_RECORD: u16 = 0x088C;
const TABLESTYLES_RECORD: u16 = 0x088E;
const RECALCID_RECORD: u16 = 0x01C1;
const COMPRESSPICTURES_RECORD: u16 = 0x089B;

struct XlsPivotGroupingInfo<'a> {
    grouping: &'a PivotGrouping,
    source_numbers: Vec<f64>,
    source_items: Vec<PivotValue>,
    source_item_ids: Vec<u32>,
    group_items: Vec<PivotValue>,
    base_item_group_ids: Vec<u32>,
    group_item_ids: Vec<u32>,
}

#[derive(Debug, Clone)]
struct XlsPivotCacheLayout {
    cache_num: usize,
    row_count: usize,
    base_field_count: usize,
    fields: Vec<XlsPivotFieldLayout>,
}

impl XlsPivotCacheLayout {
    fn field_index(&self, name: &str) -> Option<usize> {
        self.fields
            .iter()
            .position(|field| field.name.eq_ignore_ascii_case(name))
    }

    fn axis_field_index(&self, name: &str) -> Option<usize> {
        self.fields.iter().position(|field| {
            let axis_name = match &field.kind {
                XlsPivotFieldKind::DateSource { .. } => return false,
                XlsPivotFieldKind::DateGroup {
                    source_field_index, ..
                } => self
                    .fields
                    .get(*source_field_index)
                    .map(|source| source.name.as_str())
                    .unwrap_or(field.name.as_str()),
                _ => field.name.as_str(),
            };
            axis_name.eq_ignore_ascii_case(name)
        })
    }
}

#[derive(Debug, Clone)]
struct XlsPivotCacheLayouts {
    date_system: DateSystem,
    by_cache_num: HashMap<usize, XlsPivotCacheLayout>,
}

impl XlsPivotCacheLayouts {
    fn get(&self, cache_num: usize) -> XlsResult<&XlsPivotCacheLayout> {
        self.by_cache_num.get(&cache_num).ok_or_else(|| {
            XlsError::InvalidFormat(format!("pivot cache layout {cache_num} not found"))
        })
    }
}

#[derive(Debug, Clone)]
struct XlsPivotFieldLayout {
    name: String,
    formula: Option<String>,
    shared_items: Vec<PivotValue>,
    item_ids: Vec<u32>,
    kind: XlsPivotFieldKind,
}

#[derive(Debug, Clone)]
enum XlsPivotFieldKind {
    Regular,
    NumberGroup {
        start: Option<f64>,
        end: Option<f64>,
        interval: f64,
        source_numbers: Vec<f64>,
    },
    DateSource {
        derived_field_index: usize,
        source_numbers: Vec<f64>,
    },
    DateGroup {
        source_field_index: usize,
        unit: PivotDateGroupUnit,
        source_numbers: Vec<f64>,
    },
    ManualSource {
        derived_field_index: usize,
    },
    ManualGroup {
        source_field_index: usize,
        source_item_group_ids: Vec<u32>,
    },
}

#[derive(Debug)]
struct XlsDateSourceData {
    shared_items: Vec<PivotValue>,
    item_ids: Vec<u32>,
    row_numbers: Vec<f64>,
    source_numbers: Vec<f64>,
}

/// Built-in NAME index for `Print_Area` (MS-XLS §2.5.4).
const BUILTIN_NAME_PRINT_AREA: u8 = 0x06;
/// Built-in NAME index for `Print_Titles`.
const BUILTIN_NAME_PRINT_TITLES: u8 = 0x07;
/// Built-in NAME index for `_FilterDatabase` (the AutoFilter range).
const BUILTIN_NAME_FILTER_DATABASE: u8 = 0x0D;

/// Hyperlink CLSID (MS-XLS §2.4.144) - StdLink class id
/// 79EAC9D0-BAF9-11CE-8C82-00AA004BA90B, on-disk LE-mixed format
/// (Data1/2/3 little-endian, Data4 byte-ordered).
const HLINK_CLSID: [u8; 16] = [
    0xD0, 0xC9, 0xEA, 0x79, 0xF9, 0xBA, 0xCE, 0x11, 0x8C, 0x82, 0x00, 0xAA, 0x00, 0x4B, 0xA9, 0x0B,
];

/// URL moniker CLSID 79EAC9E0-BAF9-11CE-8C82-00AA004BA90B in on-disk
/// LE-mixed format (Data1/Data2/Data3 stored little-endian, Data4
/// stored byte-ordered). Identifies the moniker block as a URL rather
/// than a file path. Must match the reader's URL_MONIKER constant
/// byte-for-byte and the bytes Excel/LibreOffice write — earlier
/// drafts had Data1 in the wrong direction, which produced files
/// Excel rejected with "Open method of Workbooks class failed".
const URL_MONIKER_CLSID: [u8; 16] = [
    0xE0, 0xC9, 0xEA, 0x79, 0xF9, 0xBA, 0xCE, 0x11, 0x8C, 0x82, 0x00, 0xAA, 0x00, 0x4B, 0xA9, 0x0B,
];

/// MS-XLS §2.4.169: a single MERGECELLS record can hold at most 1027
/// merged ranges (8 bytes each, plus the 2-byte cmcs count, fits in
/// the 8224-byte BIFF8 record body cap with margin).
const MERGECELLS_MAX_PER_RECORD: usize = 1027;

/// MS-XLS §2.4.126 user-defined number-format index base. Built-in
/// formats use ifmt 0..=49; user-defined custom format strings start
/// at index 164.
const FORMAT_USER_INDEX_BASE: u16 = 164;

/// BIFF8 reserves the first 16 XF records for built-in cell-format
/// slots; user-defined cell XFs start at 16. Also see MS-XLS §2.4.353.
const XF_USER_BASE: u16 = 16;

/// BIFF8 emits at least 5 FONT records before any can be referenced
/// from XFs, with a "skip 4" quirk in cell-XF font_index decoding
/// (MS-XLS §2.4.122). The writer always emits exactly 5 default fonts
/// up front; user-defined fonts append after.
const FONT_BUILTIN_COUNT: u16 = 5;

/// MS-XLS §2.4.122 Auto color sentinel.
const COLOR_AUTO: u16 = 0x7FFF;

/// Maximum BIFF record body size. Records larger than this must be
/// split with `CONTINUE` records.
const BIFF_MAX_RECORD_BODY: usize = 8224;

const BIFF8_VERSION: u16 = 0x0600;
const DT_WORKBOOK_GLOBALS: u16 = 0x0005;
const DT_WORKSHEET: u16 = 0x0010;

/// MS-XLS §2.4.21 BOF body fields beyond `vers`/`dt`. These don't affect
/// our reader but matter for real Excel: a 4-byte BOF body is rejected
/// outright. Use Office 97 reference values for parity with what
/// historical writers produced.
const BOF_RUP_BUILD: u16 = 0x4F5A;
const BOF_RUP_YEAR: u16 = 0x07CD;
const BOF_BFH: u32 = 0x0002_00C1;
const BOF_SFH: u32 = 0x0806;

/// Writes [`Workbook`] instances to the BIFF8 (`.xls`) format.
pub struct XlsWriter;

impl XlsWriter {
    /// Serialize a workbook to BIFF8 bytes (CFB envelope with a single
    /// `/Workbook` stream).
    pub fn write_to_bytes(workbook: &Workbook) -> XlsResult<Vec<u8>> {
        let package = build_workbook_package(workbook)?;
        wrap_workbook_package_in_cfb(package)
    }

    /// Write a workbook to a filesystem path.
    pub fn write_file<P: AsRef<Path>>(workbook: &Workbook, path: P) -> XlsResult<()> {
        let bytes = Self::write_to_bytes(workbook)?;
        let mut f = std::fs::File::create(path.as_ref())?;
        f.write_all(&bytes)?;
        Ok(())
    }

    /// Serialize a workbook to encrypted BIFF8 bytes. Builds the
    /// plaintext `/Workbook` stream, runs it through
    /// [`duke_sheets_crypto::xls::encrypt_workbook_stream`] for the
    /// requested variant, then re-wraps the encrypted stream in a
    /// fresh CFB envelope.
    pub fn write_to_bytes_encrypted(
        workbook: &Workbook,
        password: &str,
        variant: duke_sheets_crypto::xls::XlsEncryptionVariant,
    ) -> XlsResult<Vec<u8>> {
        let mut package = build_workbook_package(workbook)?;
        let plain = std::mem::take(&mut package.workbook_stream);
        let encrypted = duke_sheets_crypto::xls::encrypt_workbook_stream(&plain, password, variant)
            .map_err(XlsError::from)?;
        package.workbook_stream = encrypted;
        wrap_workbook_package_in_cfb(package)
    }

    /// Write a workbook to a filesystem path with FilePass encryption.
    pub fn write_file_encrypted<P: AsRef<Path>>(
        workbook: &Workbook,
        path: P,
        password: &str,
        variant: duke_sheets_crypto::xls::XlsEncryptionVariant,
    ) -> XlsResult<()> {
        let bytes = Self::write_to_bytes_encrypted(workbook, password, variant)?;
        let mut f = std::fs::File::create(path.as_ref())?;
        f.write_all(&bytes)?;
        Ok(())
    }
}

fn wrap_workbook_package_in_cfb(package: XlsWritePackage) -> XlsResult<Vec<u8>> {
    let mut builder = CompoundFileBuilder::new();
    builder.set_root_clsid(EXCEL_WORKBOOK_ROOT_CLSID);
    builder
        .add_stream("/Workbook", package.workbook_stream)
        .map_err(cfb_to_xls)?;
    for storage in package.extra_storages {
        builder.add_storage(&storage).map_err(cfb_to_xls)?;
    }
    for (path, data) in package.extra_streams {
        builder.add_stream(&path, data).map_err(cfb_to_xls)?;
    }
    builder.build().map_err(cfb_to_xls)
}

struct XlsWritePackage {
    workbook_stream: Vec<u8>,
    extra_storages: Vec<String>,
    extra_streams: Vec<(String, Vec<u8>)>,
}

const EXCEL_WORKBOOK_ROOT_CLSID: [u8; 16] = [
    0x20, 0x08, 0x02, 0x00, 0x00, 0x00, 0x00, 0x00, 0xC0, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x46,
];

fn cfb_to_xls(err: crate::cfb::CfbError) -> XlsError {
    XlsError::Io(std::io::Error::from(err))
}

/// Assemble the full BIFF8 Workbook stream for `workbook`. Layout:
///
/// ```text
/// [globals]
///   BOF  (dt=0x0005)
///   BoundSheet8 × N            ← lbPlyPos backfilled after worksheets emit
///   EOF
/// [worksheet 1]
///   BOF  (dt=0x0010)
///   EOF
/// [worksheet 2]
///   BOF
///   EOF
/// ...
/// ```
fn build_workbook_package(workbook: &Workbook) -> XlsResult<XlsWritePackage> {
    let pivot_plan = duke_sheets_pivot::plan_format_pivots(workbook)?;
    let pivot_layouts = build_xls_pivot_cache_layouts(workbook, &pivot_plan)?;
    let workbook_stream = build_workbook_stream_with_pivots(workbook, &pivot_plan, &pivot_layouts)?;
    let (extra_storages, extra_streams) = build_pivot_cache_streams(&pivot_plan, &pivot_layouts)?;
    Ok(XlsWritePackage {
        workbook_stream,
        extra_storages,
        extra_streams,
    })
}

#[cfg(test)]
fn build_workbook_stream(workbook: &Workbook) -> XlsResult<Vec<u8>> {
    build_workbook_package(workbook).map(|package| package.workbook_stream)
}

fn build_workbook_stream_with_pivots(
    workbook: &Workbook,
    pivot_plan: &FormatPivotPlan,
    pivot_layouts: &XlsPivotCacheLayouts,
) -> XlsResult<Vec<u8>> {
    if workbook.sheet_count() == 0 {
        return Err(XlsError::InvalidFormat(
            "workbook must have at least one worksheet".into(),
        ));
    }

    let styles = StyleTables::collect(workbook);
    let sst = SstTable::collect(workbook, &styles);

    let mut stream = Vec::new();
    write_bof(&mut stream, DT_WORKBOOK_GLOBALS);
    write_window1(&mut stream, workbook);
    write_date_mode(&mut stream, workbook);
    styles.write_font_records(&mut stream)?;
    styles.write_format_records(&mut stream)?;
    styles.write_xf_records(&mut stream);

    write_table_styles_record(&mut stream, pivot_plan);

    let mut lbplypos_field_offsets = Vec::with_capacity(workbook.sheet_count());
    write_pivot_pre_boundsheet_records(&mut stream, pivot_plan)?;
    for sheet in workbook.worksheets() {
        let body_start = stream.len() + 4;
        write_boundsheet8_with_placeholder_offset(&mut stream, sheet)?;
        lbplypos_field_offsets.push(body_start);
    }

    let addin_table = build_addin_table(workbook);
    let externsheet_table = build_externsheet_table(workbook, !addin_table.is_empty());
    let name_table = build_name_table(workbook);
    write_supbook_and_externsheet(&mut stream, &externsheet_table, &addin_table);
    write_user_name_records(
        &mut stream,
        workbook,
        &externsheet_table,
        &name_table,
        &addin_table,
    );
    write_print_name_records(&mut stream, workbook);
    write_pivot_global_records(&mut stream, pivot_plan);
    sst.write_records(&mut stream)?;
    let drawing_state = compute_drawing_state(workbook);
    write_msodrawinggroup(&mut stream, &drawing_state);
    write_pivot_workbook_extension_records(&mut stream, pivot_plan);
    write_eof(&mut stream);

    let mut sheet_bof_offsets = Vec::with_capacity(workbook.sheet_count());
    for (sheet_idx, sheet) in workbook.worksheets().enumerate() {
        let has_pivot_tables = sheet_has_pivot_tables(pivot_plan, sheet_idx);
        let bof_pos = stream.len() as u32;
        write_bof(&mut stream, DT_WORKSHEET);
        let emits_cell_records = sheet_has_biff_cell_records(sheet, sheet_idx, &styles);
        let index_record_pos = write_index_placeholder(&mut stream, emits_cell_records);
        write_sheet_calculation_records(&mut stream);
        write_protect_records(&mut stream, sheet);
        write_colinfo_records(&mut stream, sheet);
        write_page_break_records(&mut stream, sheet);
        write_sheet_display_default_records(&mut stream, sheet);
        write_header_footer_records(&mut stream, sheet);
        write_margin_records(&mut stream, sheet);
        write_print_flags(&mut stream, sheet);
        write_setup_record(&mut stream, sheet);
        write_default_column_width_record(&mut stream);
        write_dimension(&mut stream, sheet);
        let row_record_positions = write_row_records(&mut stream, sheet);
        let mut first_cell_positions = BTreeMap::new();
        write_cell_records(
            &mut stream,
            sheet,
            sheet_idx,
            &sst,
            &styles,
            &externsheet_table,
            &name_table,
            &addin_table,
            &mut first_cell_positions,
        );
        let dbcell_pos =
            write_dbcell_record(&mut stream, &row_record_positions, &first_cell_positions);
        patch_index_record(
            &mut stream,
            index_record_pos,
            sheet,
            sheet_idx,
            &styles,
            row_record_positions.values().next().copied(),
            dbcell_pos,
        );
        write_pivot_sheet_records(&mut stream, workbook, pivot_plan, pivot_layouts, sheet_idx)?;
        write_window2(&mut stream, sheet, has_pivot_tables);
        write_scl(&mut stream, sheet);
        write_pane(&mut stream, sheet);
        if !has_pivot_tables {
            write_selection_records(&mut stream, sheet);
        }
        write_mergecells(&mut stream, sheet);
        write_hlink_records(&mut stream, sheet);
        write_autofilter_records(&mut stream, sheet);
        write_data_validations(
            &mut stream,
            sheet,
            &externsheet_table,
            &name_table,
            &addin_table,
        );
        write_conditional_formats(
            &mut stream,
            sheet,
            &externsheet_table,
            &name_table,
            &addin_table,
        );
        if let Some(sheet_drawing) = drawing_state.sheets.get(&sheet_idx) {
            write_sheet_drawing_records(&mut stream, sheet_drawing);
        }
        write_pivot_sheet_tail_records(&mut stream, sheet, pivot_plan, sheet_idx);
        write_eof(&mut stream);
        sheet_bof_offsets.push(bof_pos);
    }

    for (offset, sheet_bof) in lbplypos_field_offsets.iter().zip(sheet_bof_offsets.iter()) {
        stream[*offset..*offset + 4].copy_from_slice(&sheet_bof.to_le_bytes());
    }

    Ok(stream)
}

fn sheet_has_pivot_tables(pivot_plan: &FormatPivotPlan, sheet_idx: usize) -> bool {
    pivot_plan
        .tables
        .iter()
        .any(|part| part.sheet_index == sheet_idx)
}

fn build_pivot_cache_streams(
    pivot_plan: &FormatPivotPlan,
    layouts: &XlsPivotCacheLayouts,
) -> XlsResult<(Vec<String>, Vec<(String, Vec<u8>)>)> {
    if pivot_plan.caches.is_empty() {
        return Ok((Vec::new(), Vec::new()));
    }

    let storages = vec!["/_SX_DB_CUR".to_string()];
    let mut streams = Vec::with_capacity(pivot_plan.caches.len());
    for cache in &pivot_plan.caches {
        let layout = layouts.get(cache.cache_num)?;
        streams.push((
            format!("/_SX_DB_CUR/{:04}", cache.cache_num),
            build_pivot_cache_stream(cache, layout, layouts.date_system)?,
        ));
    }
    Ok((storages, streams))
}

fn build_xls_pivot_cache_layouts(
    workbook: &Workbook,
    pivot_plan: &FormatPivotPlan,
) -> XlsResult<XlsPivotCacheLayouts> {
    let date_system = workbook_date_system(workbook.settings().date_1904);
    let mut by_cache_num = HashMap::with_capacity(pivot_plan.caches.len());
    for cache in &pivot_plan.caches {
        let groupings = groupings_for_cache(workbook, pivot_plan, cache)?;
        validate_xls_pivot_groupings(cache, groupings)?;
        let grouping_infos = xls_pivot_grouping_infos(workbook, cache, groupings, date_system)?;
        let layout = build_xls_pivot_cache_layout(workbook, cache, &grouping_infos, date_system)?;
        by_cache_num.insert(cache.cache_num, layout);
    }
    Ok(XlsPivotCacheLayouts {
        date_system,
        by_cache_num,
    })
}

fn build_pivot_cache_stream(
    cache: &FormatPivotCache,
    layout: &XlsPivotCacheLayout,
    date_system: DateSystem,
) -> XlsResult<Vec<u8>> {
    if !matches!(cache.source, FormatPivotSource::Worksheet { .. }) {
        return Err(XlsError::InvalidFormat(
            "XLS pivot cache writing requires worksheet or table sources".into(),
        ));
    }

    let mut stream = Vec::new();
    let mut sxdb = Vec::new();
    sxdb.extend_from_slice(&(layout.row_count as u32).to_le_bytes());
    sxdb.extend_from_slice(&(layout.cache_num as u16).to_le_bytes());
    sxdb.extend_from_slice(&0x0003u16.to_le_bytes());
    let has_grouped_cache_field = layout.fields.iter().any(|field| {
        matches!(
            field.kind,
            XlsPivotFieldKind::DateGroup { .. } | XlsPivotFieldKind::ManualGroup { .. }
        )
    });
    sxdb.extend_from_slice(&if has_grouped_cache_field {
        0x0FFFu16.to_le_bytes()
    } else {
        0x1999u16.to_le_bytes()
    });
    sxdb.extend_from_slice(
        &checked_u16(layout.base_field_count, "pivot cache base field count")?.to_le_bytes(),
    );
    sxdb.extend_from_slice(
        &checked_u16(layout.fields.len(), "pivot cache total field count")?.to_le_bytes(),
    );
    sxdb.extend_from_slice(&if has_grouped_cache_field {
        checked_u16(layout.row_count, "pivot cache used row count")?.to_le_bytes()
    } else {
        0u16.to_le_bytes()
    });
    sxdb.extend_from_slice(&1u16.to_le_bytes());
    if has_grouped_cache_field {
        sxdb.extend_from_slice(&4u16.to_le_bytes());
        sxdb.push(0);
        sxdb.extend_from_slice(b"user");
    } else {
        sxdb.extend_from_slice(&0xFFFFu16.to_le_bytes());
    }
    write_biff_record(&mut stream, 0x00C6, &sxdb);
    write_biff_record(&mut stream, 0x0122, &[0; 12]);

    for field in &layout.fields {
        write_sxfdb_record(&mut stream, field)?;
        if let Some(formula) = field.formula.as_deref() {
            write_pivot_calculated_field_formula_records(&mut stream, cache, formula)?;
        }
        write_biff_record(&mut stream, 0x01BB, &0u16.to_le_bytes());
        write_pivot_cache_field_items(&mut stream, field, date_system)?;
    }
    write_numeric_cache_records(&mut stream, layout)?;

    write_eof(&mut stream);
    Ok(stream)
}

fn write_numeric_cache_records(
    stream: &mut Vec<u8>,
    layout: &XlsPivotCacheLayout,
) -> XlsResult<()> {
    let row_marker_field = xls_cache_row_marker_field(layout);
    let numeric_fields = layout
        .fields
        .iter()
        .filter(|field| field_is_numeric(field) && matches!(field.kind, XlsPivotFieldKind::Regular))
        .collect::<Vec<_>>();
    if numeric_fields.is_empty() {
        return Ok(());
    }

    for row in 0..layout.row_count {
        let row_marker = row_marker_field
            .and_then(|field| field.item_ids.get(row).copied())
            .map(|item_id| checked_u8(item_id as usize, "pivot cache row item index"))
            .unwrap_or_else(|| checked_u8(row, "pivot cache row index"))?;
        write_biff_record(stream, PIVOT_SXDBB_RECORD, &[row_marker]);
        for field in &numeric_fields {
            if let Some(number) = numeric_cache_value(field, row) {
                write_biff_record(stream, PIVOT_SXNUM_RECORD, &number.to_le_bytes());
            }
        }
    }
    Ok(())
}

fn xls_cache_row_marker_field(layout: &XlsPivotCacheLayout) -> Option<&XlsPivotFieldLayout> {
    layout
        .fields
        .iter()
        .take(layout.base_field_count)
        .find(|field| !matches!(field.kind, XlsPivotFieldKind::Regular) || !field_is_numeric(field))
}

fn numeric_cache_value(field: &XlsPivotFieldLayout, row: usize) -> Option<f64> {
    let item_id = *field.item_ids.get(row)? as usize;
    let PivotValue::Number(number) = field.shared_items.get(item_id)? else {
        return None;
    };
    Some(*number)
}

fn groupings_for_cache<'a>(
    workbook: &'a Workbook,
    plan: &'a FormatPivotPlan,
    cache: &FormatPivotCache,
) -> XlsResult<&'a [PivotGrouping]> {
    let Some(part) = plan
        .tables
        .iter()
        .find(|part| part.cache_num == cache.cache_num)
    else {
        return Ok(&[]);
    };
    let worksheet = workbook
        .worksheet(part.sheet_index)
        .ok_or_else(|| XlsError::InvalidFormat("pivot table sheet not found".into()))?;
    let pivot = worksheet
        .pivot_tables()
        .get(part.pivot_index)
        .ok_or_else(|| XlsError::InvalidFormat("pivot table not found".into()))?;
    Ok(&pivot.groupings)
}

fn validate_xls_pivot_groupings(
    cache: &FormatPivotCache,
    groupings: &[PivotGrouping],
) -> XlsResult<()> {
    let mut grouped_fields = HashSet::new();
    for grouping in groupings {
        let field_name = grouping_field_name(grouping);
        if cache.field_index(field_name).is_none() {
            return Err(XlsError::InvalidFormat(format!(
                "XLS pivot grouping references unknown cache field: {field_name}"
            )));
        }
        if !grouped_fields.insert(field_name.to_lowercase()) {
            return Err(XlsError::InvalidFormat(format!(
                "XLS pivot cache has more than one grouping for field {field_name}"
            )));
        }

        match grouping {
            PivotGrouping::Number {
                start,
                end,
                interval,
                ..
            } => {
                if !interval.is_finite() || *interval <= 0.0 {
                    return Err(XlsError::InvalidFormat(format!(
                        "XLS pivot grouping for field {field_name} has an invalid interval"
                    )));
                }
                if start.is_some_and(|value| !value.is_finite())
                    || end.is_some_and(|value| !value.is_finite())
                {
                    return Err(XlsError::InvalidFormat(format!(
                        "XLS pivot grouping for field {field_name} has a non-finite bound"
                    )));
                }
            }
            PivotGrouping::Date { units, .. } => {
                if units.len() != 1 {
                    return Err(XlsError::InvalidFormat(format!(
                        "XLS pivot date grouping currently supports exactly one date unit: {field_name}"
                    )));
                }
            }
            PivotGrouping::Manual { groups, .. } => {
                validate_xls_manual_grouping(field_name, groups)?;
            }
        }
    }
    Ok(())
}

fn validate_xls_pivot_grouping_axes(pivot: &duke_sheets_core::PivotTable) -> XlsResult<()> {
    for grouping in &pivot.groupings {
        let PivotGrouping::Manual { field, .. } = grouping else {
            continue;
        };
        let field_name = field.name.as_str();
        if !pivot_axis_contains_field(&pivot.rows, field_name) {
            return Err(XlsError::InvalidFormat(format!(
                "XLS pivot manual grouping currently supports only row-axis fields: {field_name}"
            )));
        }
    }
    Ok(())
}

fn pivot_axis_contains_field(fields: &[PivotField], field_name: &str) -> bool {
    fields
        .iter()
        .any(|field| field.field.name.eq_ignore_ascii_case(field_name))
}

fn validate_xls_manual_grouping(field_name: &str, groups: &[PivotManualGroup]) -> XlsResult<()> {
    if groups.is_empty() {
        return Err(XlsError::InvalidFormat(format!(
            "XLS pivot manual grouping for field {field_name} has no groups"
        )));
    }

    let mut names = HashSet::new();
    let mut members = HashSet::new();
    for group in groups {
        if group.name.trim().is_empty() {
            return Err(XlsError::InvalidFormat(format!(
                "XLS pivot manual grouping for field {field_name} has a blank group name"
            )));
        }
        if group.members.is_empty() {
            return Err(XlsError::InvalidFormat(format!(
                "XLS pivot manual group {} has no members",
                group.name
            )));
        }
        if !names.insert(group.name.to_lowercase()) {
            return Err(XlsError::InvalidFormat(format!(
                "XLS pivot manual grouping for field {field_name} has duplicate group name {}",
                group.name
            )));
        }
        for member in &group.members {
            if !members.insert(member.clone()) {
                return Err(XlsError::InvalidFormat(format!(
                    "XLS pivot manual grouping for field {field_name} assigns item {member} to more than one group"
                )));
            }
        }
    }
    Ok(())
}

fn xls_pivot_grouping_infos<'a>(
    workbook: &Workbook,
    cache: &FormatPivotCache,
    groupings: &'a [PivotGrouping],
    date_system: DateSystem,
) -> XlsResult<Vec<XlsPivotGroupingInfo<'a>>> {
    groupings
        .iter()
        .map(|grouping| {
            if let PivotGrouping::Manual { groups, .. } = grouping {
                let field_name = grouping_field_name(grouping);
                let field_index = cache.field_index(field_name).ok_or_else(|| {
                    XlsError::InvalidFormat(format!(
                        "XLS pivot grouping references unknown cache field: {field_name}"
                    ))
                })?;
                let (source_items, source_item_ids) =
                    xls_manual_group_source_items(workbook, cache, field_index, field_name)?;
                let (group_items, base_item_group_ids) =
                    manual_group_items_and_ids(field_name, &source_items, groups)?;
                let group_item_ids = source_item_ids
                    .iter()
                    .map(|item_id| {
                        base_item_group_ids
                            .get(*item_id as usize)
                            .copied()
                            .ok_or_else(|| {
                                XlsError::InvalidFormat(format!(
                                    "XLS pivot manual grouping for field {field_name} has an out-of-range source item index"
                                ))
                            })
                    })
                    .collect::<XlsResult<Vec<_>>>()?;
                return Ok(XlsPivotGroupingInfo {
                    grouping,
                    source_numbers: Vec::new(),
                    source_items,
                    source_item_ids,
                    group_items,
                    base_item_group_ids,
                    group_item_ids,
                });
            }

            Ok(XlsPivotGroupingInfo {
                grouping,
                source_numbers: xls_grouping_source_numbers(
                    workbook,
                    cache,
                    grouping,
                    date_system,
                )?,
                source_items: Vec::new(),
                source_item_ids: Vec::new(),
                group_items: Vec::new(),
                base_item_group_ids: Vec::new(),
                group_item_ids: Vec::new(),
            })
        })
        .collect()
}

fn xls_grouping_source_numbers(
    workbook: &Workbook,
    cache: &FormatPivotCache,
    grouping: &PivotGrouping,
    date_system: DateSystem,
) -> XlsResult<Vec<f64>> {
    let field_name = grouping_field_name(grouping);
    let field_index = cache.field_index(field_name).ok_or_else(|| {
        XlsError::InvalidFormat(format!(
            "XLS pivot grouping references unknown cache field: {field_name}"
        ))
    })?;
    let FormatPivotSource::Worksheet {
        sheet_index, range, ..
    } = &cache.source
    else {
        return Err(XlsError::InvalidFormat(
            "XLS pivot grouping requires worksheet-range source data".into(),
        ));
    };
    let source_col = u32::from(range.start.col)
        .checked_add(field_index as u32)
        .ok_or_else(|| XlsError::InvalidFormat("XLS pivot grouping field index overflow".into()))?;
    if source_col > u16::MAX as u32 || source_col > u32::from(range.end.col) {
        return Err(XlsError::InvalidFormat(format!(
            "XLS pivot grouping for field {field_name} must reference a source field"
        )));
    }
    let worksheet = workbook
        .worksheet(*sheet_index)
        .ok_or_else(|| XlsError::InvalidFormat("pivot source worksheet not found".into()))?;
    let is_date_grouping = matches!(grouping, PivotGrouping::Date { .. });
    let mut seen = HashSet::new();
    let mut numbers = Vec::new();
    for row in range.start.row.saturating_add(1)..=range.end.row {
        match worksheet.get_value_at(row, source_col as u16) {
            CellValue::Number(value) if value.is_finite() => {
                if is_date_grouping && !valid_xls_date_serial(value, date_system) {
                    return Err(XlsError::InvalidFormat(format!(
                        "XLS pivot date grouping for field {field_name} has invalid date serial: {value}"
                    )));
                }
                if seen.insert(value.to_bits()) {
                    numbers.push(value);
                }
            }
            CellValue::Number(value) if is_date_grouping => {
                return Err(XlsError::InvalidFormat(format!(
                    "XLS pivot date grouping for field {field_name} has non-finite source value: {value}"
                )));
            }
            CellValue::Empty if !is_date_grouping => {}
            _ if is_date_grouping => {
                return Err(XlsError::InvalidFormat(format!(
                    "XLS pivot date grouping for field {field_name} requires numeric date source values"
                )));
            }
            _ => {}
        }
    }
    if is_date_grouping && numbers.is_empty() {
        return Err(XlsError::InvalidFormat(format!(
            "XLS pivot date grouping for field {field_name} has no source dates"
        )));
    }
    Ok(numbers)
}

fn grouping_field_name(grouping: &PivotGrouping) -> &str {
    match grouping {
        PivotGrouping::Number { field, .. }
        | PivotGrouping::Date { field, .. }
        | PivotGrouping::Manual { field, .. } => &field.name,
    }
}

fn manual_group_items_and_ids(
    field_name: &str,
    source_items: &[PivotValue],
    groups: &[PivotManualGroup],
) -> XlsResult<(Vec<PivotValue>, Vec<u32>)> {
    if source_items.is_empty() {
        return Err(XlsError::InvalidFormat(format!(
            "XLS pivot manual grouping for field {field_name} has no source items"
        )));
    }
    if source_items
        .iter()
        .any(|item| !xls_manual_group_item_is_string_like(item))
    {
        return Err(XlsError::InvalidFormat(format!(
            "XLS pivot manual grouping for field {field_name} currently supports only text or blank source items"
        )));
    }

    let mut member_to_group = HashMap::new();
    for group in groups {
        for member in &group.members {
            if !xls_manual_group_item_is_string_like(member) {
                return Err(XlsError::InvalidFormat(format!(
                    "XLS pivot manual group {} references a non-text item in field {field_name}: {member}",
                    group.name
                )));
            }
            if !source_items.iter().any(|item| item == member) {
                return Err(XlsError::InvalidFormat(format!(
                    "XLS pivot manual group {} references item not found in field {field_name}: {member}",
                    group.name
                )));
            }
            member_to_group.insert(member.clone(), group.name.clone());
        }
    }

    let mut group_items = Vec::new();
    let mut ungrouped_item_indexes = HashMap::new();
    for item in source_items {
        if member_to_group.contains_key(item) {
            continue;
        }
        let index = checked_u32(group_items.len(), "pivot manual ungrouped item index")?;
        ungrouped_item_indexes.insert(item.clone(), index);
        group_items.push(item.clone());
    }

    let mut group_name_indexes = HashMap::new();
    for group in groups {
        let index = checked_u32(group_items.len(), "pivot manual group item index")?;
        group_name_indexes.insert(group.name.clone(), index);
        group_items.push(PivotValue::String(group.name.clone()));
    }

    let base_item_group_ids = source_items
        .iter()
        .map(|item| {
            if let Some(group_name) = member_to_group.get(item) {
                group_name_indexes.get(group_name).copied().ok_or_else(|| {
                    XlsError::InvalidFormat(format!(
                        "XLS pivot manual grouping for field {field_name} could not map group {group_name}"
                    ))
                })
            } else {
                ungrouped_item_indexes.get(item).copied().ok_or_else(|| {
                    XlsError::InvalidFormat(format!(
                        "XLS pivot manual grouping for field {field_name} could not map ungrouped item {item}"
                    ))
                })
            }
        })
        .collect::<XlsResult<Vec<_>>>()?;

    Ok((group_items, base_item_group_ids))
}

fn xls_manual_group_item_is_string_like(item: &PivotValue) -> bool {
    matches!(item, PivotValue::Blank | PivotValue::String(_))
}

fn xls_manual_group_source_items(
    workbook: &Workbook,
    cache: &FormatPivotCache,
    field_index: usize,
    field_name: &str,
) -> XlsResult<(Vec<PivotValue>, Vec<u32>)> {
    let FormatPivotSource::Worksheet {
        sheet_index, range, ..
    } = &cache.source
    else {
        return Err(XlsError::InvalidFormat(
            "XLS pivot manual grouping requires worksheet-range source data".into(),
        ));
    };
    let source_col = u32::from(range.start.col)
        .checked_add(field_index as u32)
        .ok_or_else(|| XlsError::InvalidFormat("XLS pivot grouping field index overflow".into()))?;
    if source_col > u16::MAX as u32 || source_col > u32::from(range.end.col) {
        return Err(XlsError::InvalidFormat(format!(
            "XLS pivot manual grouping for field {field_name} must reference a source field"
        )));
    }
    let worksheet = workbook
        .worksheet(*sheet_index)
        .ok_or_else(|| XlsError::InvalidFormat("pivot source worksheet not found".into()))?;

    let mut source_items = Vec::new();
    let mut source_item_ids = Vec::new();
    let mut lookup = HashMap::new();
    for row in range.start.row.saturating_add(1)..=range.end.row {
        let cell_value = worksheet.get_value_at(row, source_col as u16);
        let value = PivotValue::from_cell_value(&cell_value);
        let item_id = if let Some(item_id) = lookup.get(&value) {
            *item_id
        } else {
            let item_id = checked_u32(source_items.len(), "pivot manual source item index")?;
            lookup.insert(value.clone(), item_id);
            source_items.push(value);
            item_id
        };
        source_item_ids.push(item_id);
    }
    if source_items.is_empty() {
        return Err(XlsError::InvalidFormat(format!(
            "XLS pivot manual grouping for field {field_name} has no source items"
        )));
    }
    Ok((source_items, source_item_ids))
}

fn build_xls_pivot_cache_layout(
    workbook: &Workbook,
    cache: &FormatPivotCache,
    grouping_infos: &[XlsPivotGroupingInfo<'_>],
    date_system: DateSystem,
) -> XlsResult<XlsPivotCacheLayout> {
    let mut layout = XlsPivotCacheLayout {
        cache_num: cache.cache_num,
        row_count: cache.row_count,
        base_field_count: cache.fields.len(),
        fields: cache
            .fields
            .iter()
            .map(|field| XlsPivotFieldLayout {
                name: field.name.clone(),
                formula: field.formula.clone(),
                shared_items: field.shared_items.clone(),
                item_ids: field.item_ids.clone(),
                kind: XlsPivotFieldKind::Regular,
            })
            .collect(),
    };

    for info in grouping_infos {
        let field_name = grouping_field_name(info.grouping);
        let field_index = cache.field_index(field_name).ok_or_else(|| {
            XlsError::InvalidFormat(format!(
                "XLS pivot grouping references unknown cache field: {field_name}"
            ))
        })?;
        match info.grouping {
            PivotGrouping::Number {
                start,
                end,
                interval,
                ..
            } => {
                layout.fields[field_index].kind = XlsPivotFieldKind::NumberGroup {
                    start: *start,
                    end: *end,
                    interval: *interval,
                    source_numbers: info.source_numbers.clone(),
                };
            }
            PivotGrouping::Date { units, .. } => {
                let unit = units.first().copied().ok_or_else(|| {
                    XlsError::InvalidFormat(format!(
                        "XLS pivot date grouping has no unit: {field_name}"
                    ))
                })?;
                let source = xls_date_source_data(workbook, cache, field_index, date_system)?;
                let (start, end) =
                    source_number_min_max(&source.source_numbers).ok_or_else(|| {
                        XlsError::InvalidFormat(format!(
                            "XLS pivot date grouping for field {field_name} has no source dates"
                        ))
                    })?;
                let derived_items = xls_date_group_shared_items(unit, start, end, date_system)?;
                let derived_item_ids = source
                    .row_numbers
                    .iter()
                    .map(|serial| xls_date_group_item_id(unit, *serial, start, end, date_system))
                    .collect::<XlsResult<Vec<_>>>()?;
                let derived_field_index = layout.fields.len();
                layout.fields[field_index].shared_items = source.shared_items;
                layout.fields[field_index].item_ids = source.item_ids;
                layout.fields[field_index].kind = XlsPivotFieldKind::DateSource {
                    derived_field_index,
                    source_numbers: source.source_numbers.clone(),
                };
                let derived_name = unique_xls_date_group_field_name(
                    &layout.fields,
                    &layout.fields[field_index].name,
                    unit,
                );
                layout.fields.push(XlsPivotFieldLayout {
                    name: derived_name,
                    formula: None,
                    shared_items: derived_items,
                    item_ids: derived_item_ids,
                    kind: XlsPivotFieldKind::DateGroup {
                        source_field_index: field_index,
                        unit,
                        source_numbers: source.source_numbers,
                    },
                });
            }
            PivotGrouping::Manual { .. } => {
                let derived_field_index = layout.fields.len();
                layout.fields[field_index].shared_items = info.source_items.clone();
                layout.fields[field_index].item_ids = info.source_item_ids.clone();
                layout.fields[field_index].kind = XlsPivotFieldKind::ManualSource {
                    derived_field_index,
                };
                let derived_name = unique_xls_manual_grouped_field_name(
                    &layout.fields,
                    &layout.fields[field_index].name,
                );
                layout.fields.push(XlsPivotFieldLayout {
                    name: derived_name,
                    formula: None,
                    shared_items: info.group_items.clone(),
                    item_ids: info.group_item_ids.clone(),
                    kind: XlsPivotFieldKind::ManualGroup {
                        source_field_index: field_index,
                        source_item_group_ids: info.base_item_group_ids.clone(),
                    },
                });
            }
        }
    }

    Ok(layout)
}

fn unique_xls_manual_grouped_field_name(
    fields: &[XlsPivotFieldLayout],
    source_name: &str,
) -> String {
    for suffix in 2usize.. {
        let candidate = format!("{source_name}{suffix}");
        if fields
            .iter()
            .all(|field| !field.name.eq_ignore_ascii_case(&candidate))
        {
            return candidate;
        }
    }
    unreachable!("unbounded manual grouped field name suffix search should return")
}

fn xls_date_source_data(
    workbook: &Workbook,
    cache: &FormatPivotCache,
    field_index: usize,
    date_system: DateSystem,
) -> XlsResult<XlsDateSourceData> {
    let FormatPivotSource::Worksheet {
        sheet_index, range, ..
    } = &cache.source
    else {
        return Err(XlsError::InvalidFormat(
            "XLS pivot date grouping requires worksheet-range source data".into(),
        ));
    };
    let source_col = u32::from(range.start.col)
        .checked_add(field_index as u32)
        .ok_or_else(|| XlsError::InvalidFormat("XLS pivot grouping field index overflow".into()))?;
    if source_col > u16::MAX as u32 || source_col > u32::from(range.end.col) {
        return Err(XlsError::InvalidFormat(
            "XLS pivot date grouping must reference a source field".into(),
        ));
    }
    let worksheet = workbook
        .worksheet(*sheet_index)
        .ok_or_else(|| XlsError::InvalidFormat("pivot source worksheet not found".into()))?;
    let mut shared_items = Vec::new();
    let mut item_ids = Vec::new();
    let mut row_numbers = Vec::new();
    let mut index_by_bits = HashMap::new();
    for row in range.start.row.saturating_add(1)..=range.end.row {
        let value = match worksheet.get_value_at(row, source_col as u16) {
            CellValue::Number(value) if value.is_finite() => value,
            CellValue::Number(value) => {
                return Err(XlsError::InvalidFormat(format!(
                    "XLS pivot date grouping has non-finite source value: {value}"
                )));
            }
            _ => {
                return Err(XlsError::InvalidFormat(
                    "XLS pivot date grouping requires numeric date source values".into(),
                ));
            }
        };
        if !valid_xls_date_serial(value, date_system) {
            return Err(XlsError::InvalidFormat(format!(
                "XLS pivot date grouping has invalid date serial: {value}"
            )));
        }
        let id = if let Some(id) = index_by_bits.get(&value.to_bits()) {
            *id
        } else {
            let id = checked_u32(shared_items.len(), "pivot date shared item index")?;
            shared_items.push(PivotValue::Number(value));
            index_by_bits.insert(value.to_bits(), id);
            id
        };
        item_ids.push(id);
        row_numbers.push(value);
    }
    if row_numbers.len() != cache.row_count {
        return Err(XlsError::InvalidFormat(
            "XLS pivot date source row count does not match the cache".into(),
        ));
    }
    let source_numbers = shared_items
        .iter()
        .filter_map(|value| match value {
            PivotValue::Number(number) => Some(*number),
            _ => None,
        })
        .collect();
    Ok(XlsDateSourceData {
        shared_items,
        item_ids,
        row_numbers,
        source_numbers,
    })
}

fn unique_xls_date_group_field_name(
    fields: &[XlsPivotFieldLayout],
    source_name: &str,
    unit: PivotDateGroupUnit,
) -> String {
    let base = format!("{} ({source_name})", xls_date_group_unit_name(unit));
    if !fields
        .iter()
        .any(|field| field.name.eq_ignore_ascii_case(&base))
    {
        return base;
    }
    for suffix in 2.. {
        let candidate = format!("{base} {suffix}");
        if !fields
            .iter()
            .any(|field| field.name.eq_ignore_ascii_case(&candidate))
        {
            return candidate;
        }
    }
    unreachable!("unbounded pivot date group name suffix search should return")
}

fn write_sxfdb_record(stream: &mut Vec<u8>, field: &XlsPivotFieldLayout) -> XlsResult<()> {
    let mut body = Vec::new();
    let mut flags = match field.kind {
        XlsPivotFieldKind::NumberGroup { .. } => 0x0571u16,
        XlsPivotFieldKind::DateSource { .. } => 0x0909u16,
        XlsPivotFieldKind::DateGroup { .. } => 0x0011u16,
        XlsPivotFieldKind::ManualSource { .. } => 0x0489u16,
        XlsPivotFieldKind::ManualGroup { .. } => 0x0001u16,
        XlsPivotFieldKind::Regular if field_is_numeric(field) => 0x0560u16,
        XlsPivotFieldKind::Regular => 0x0481u16,
    };
    if field.formula.is_some() {
        flags |= 0x8000;
    }
    body.extend_from_slice(&flags.to_le_bytes());
    match &field.kind {
        XlsPivotFieldKind::NumberGroup { source_numbers, .. } => {
            body.extend_from_slice(&(-1i16).to_le_bytes());
            body.extend_from_slice(&(-1i16).to_le_bytes());
            let item_count = checked_u16(field.shared_items.len(), "pivot group item count")?;
            body.extend_from_slice(&item_count.to_le_bytes());
            body.extend_from_slice(&item_count.to_le_bytes());
            body.extend_from_slice(&0u16.to_le_bytes());
            body.extend_from_slice(
                &checked_u16(source_numbers.len(), "pivot grouped source atom count")?
                    .to_le_bytes(),
            );
        }
        XlsPivotFieldKind::DateSource {
            derived_field_index,
            source_numbers,
        } => {
            body.extend_from_slice(
                &checked_i16(*derived_field_index, "pivot date group field index")?.to_le_bytes(),
            );
            body.extend_from_slice(&(-1i16).to_le_bytes());
            let item_count = checked_u16(field.shared_items.len(), "pivot date item count")?;
            body.extend_from_slice(&item_count.to_le_bytes());
            body.extend_from_slice(&0u16.to_le_bytes());
            body.extend_from_slice(&0u16.to_le_bytes());
            body.extend_from_slice(
                &checked_u16(source_numbers.len(), "pivot date source atom count")?.to_le_bytes(),
            );
        }
        XlsPivotFieldKind::DateGroup {
            source_field_index, ..
        } => {
            body.extend_from_slice(&(-1i16).to_le_bytes());
            body.extend_from_slice(
                &checked_i16(*source_field_index, "pivot date source field index")?.to_le_bytes(),
            );
            let item_count = checked_u16(field.shared_items.len(), "pivot date group item count")?;
            body.extend_from_slice(&item_count.to_le_bytes());
            body.extend_from_slice(&item_count.to_le_bytes());
            body.extend_from_slice(&0u16.to_le_bytes());
            body.extend_from_slice(&0u16.to_le_bytes());
        }
        XlsPivotFieldKind::ManualSource {
            derived_field_index,
        } => {
            body.extend_from_slice(
                &checked_i16(*derived_field_index, "pivot manual group field index")?.to_le_bytes(),
            );
            body.extend_from_slice(&(-1i16).to_le_bytes());
            let item_count =
                checked_u16(field.shared_items.len(), "pivot manual source item count")?;
            body.extend_from_slice(&item_count.to_le_bytes());
            body.extend_from_slice(&0u16.to_le_bytes());
            body.extend_from_slice(&0u16.to_le_bytes());
            body.extend_from_slice(&item_count.to_le_bytes());
        }
        XlsPivotFieldKind::ManualGroup {
            source_field_index,
            source_item_group_ids,
        } => {
            body.extend_from_slice(&(-1i16).to_le_bytes());
            body.extend_from_slice(
                &checked_i16(*source_field_index, "pivot manual source field index")?.to_le_bytes(),
            );
            let item_count =
                checked_u16(field.shared_items.len(), "pivot manual group item count")?;
            body.extend_from_slice(&item_count.to_le_bytes());
            body.extend_from_slice(&item_count.to_le_bytes());
            body.extend_from_slice(
                &checked_u16(
                    source_item_group_ids.len(),
                    "pivot manual grouped source atom count",
                )?
                .to_le_bytes(),
            );
            body.extend_from_slice(&0u16.to_le_bytes());
        }
        XlsPivotFieldKind::Regular => {
            body.extend_from_slice(&0xFFFF_FFFFu32.to_le_bytes());
            body.extend_from_slice(&(field.shared_items.len() as u32).to_le_bytes());
            body.extend_from_slice(&0u16.to_le_bytes());
            body.extend_from_slice(&if field_is_numeric(field) {
                0u16.to_le_bytes()
            } else {
                2u16.to_le_bytes()
            });
        }
    }
    push_xlunicode_string(&mut body, &field.name)?;
    write_biff_record(stream, PIVOT_SXFDB_RECORD, &body);
    Ok(())
}

fn write_pivot_cache_field_items(
    stream: &mut Vec<u8>,
    field: &XlsPivotFieldLayout,
    date_system: DateSystem,
) -> XlsResult<()> {
    match &field.kind {
        XlsPivotFieldKind::NumberGroup {
            start,
            end,
            interval,
            source_numbers,
        } => {
            for value in &field.shared_items {
                write_biff_record(
                    stream,
                    PIVOT_SXSTRING_RECORD,
                    &pivot_value_string_payload(value)?,
                );
            }

            let mut flags = 0x0040u16;
            if start.is_none() {
                flags |= 0x0001;
            }
            if end.is_none() {
                flags |= 0x0002;
            }
            write_biff_record(stream, PIVOT_SXRNG_RECORD, &flags.to_le_bytes());
            write_biff_record(
                stream,
                PIVOT_SXNUM_RECORD,
                &start.unwrap_or(0.0).to_le_bytes(),
            );
            write_biff_record(
                stream,
                PIVOT_SXNUM_RECORD,
                &end.unwrap_or(0.0).to_le_bytes(),
            );
            write_biff_record(stream, PIVOT_SXNUM_RECORD, &interval.to_le_bytes());
            for number in source_numbers {
                write_biff_record(stream, PIVOT_SXNUM_RECORD, &number.to_le_bytes());
            }
        }
        XlsPivotFieldKind::DateSource { source_numbers, .. } => {
            for number in source_numbers {
                write_pivot_sxdtr_record(stream, *number, date_system)?;
            }
        }
        XlsPivotFieldKind::DateGroup {
            unit,
            source_numbers,
            ..
        } => {
            for value in &field.shared_items {
                write_biff_record(
                    stream,
                    PIVOT_SXSTRING_RECORD,
                    &pivot_value_string_payload(value)?,
                );
            }
            let flags = 0x0003u16 | (u16::from(xls_date_group_by(*unit)) << 2);
            write_biff_record(stream, PIVOT_SXRNG_RECORD, &flags.to_le_bytes());
            let (start, end) = source_number_min_max(source_numbers).ok_or_else(|| {
                XlsError::InvalidFormat(format!(
                    "XLS pivot date grouping for field {} has no source dates",
                    field.name
                ))
            })?;
            write_pivot_sxdtr_record(stream, start, date_system)?;
            write_pivot_sxdtr_record(stream, end, date_system)?;
            write_biff_record(stream, PIVOT_SXINT_RECORD, &1u16.to_le_bytes());
        }
        XlsPivotFieldKind::ManualSource { .. } => {
            for value in &field.shared_items {
                write_biff_record(
                    stream,
                    PIVOT_SXSTRING_RECORD,
                    &pivot_value_string_payload(value)?,
                );
            }
        }
        XlsPivotFieldKind::ManualGroup {
            source_item_group_ids,
            ..
        } => {
            for value in &field.shared_items {
                write_biff_record(
                    stream,
                    PIVOT_SXSTRING_RECORD,
                    &pivot_value_string_payload(value)?,
                );
            }
            let mut body = Vec::new();
            for item_id in source_item_group_ids {
                body.extend_from_slice(
                    &checked_u16(*item_id as usize, "pivot manual source item group index")?
                        .to_le_bytes(),
                );
            }
            write_biff_record(stream, PIVOT_SXIDSTM_RECORD, &body);
        }
        XlsPivotFieldKind::Regular if !field_is_numeric(field) => {
            for value in &field.shared_items {
                write_biff_record(
                    stream,
                    PIVOT_SXSTRING_RECORD,
                    &pivot_value_string_payload(value)?,
                );
            }
        }
        XlsPivotFieldKind::Regular => {}
    }
    Ok(())
}

fn xls_date_group_shared_items(
    unit: PivotDateGroupUnit,
    start: f64,
    end: f64,
    date_system: DateSystem,
) -> XlsResult<Vec<PivotValue>> {
    let mut items = Vec::new();
    items.push(PivotValue::String(format!(
        "<{}",
        format_xls_pivot_date_bound(start, date_system)?
    )));
    match unit {
        PivotDateGroupUnit::Seconds | PivotDateGroupUnit::Minutes => {
            for value in 0..=59 {
                items.push(PivotValue::String(value.to_string()));
            }
        }
        PivotDateGroupUnit::Hours => {
            for value in 0..=23 {
                items.push(PivotValue::String(value.to_string()));
            }
        }
        PivotDateGroupUnit::Days => {
            for value in 1..=31 {
                items.push(PivotValue::String(value.to_string()));
            }
        }
        PivotDateGroupUnit::Months => {
            for label in [
                "Jan", "Feb", "Mar", "Apr", "May", "Jun", "Jul", "Aug", "Sep", "Oct", "Nov", "Dec",
            ] {
                items.push(PivotValue::String(label.to_string()));
            }
        }
        PivotDateGroupUnit::Quarters => {
            for value in 1..=4 {
                items.push(PivotValue::String(format!("Qtr{value}")));
            }
        }
        PivotDateGroupUnit::Years => {
            let (start_year, _, _) = xls_serial_date_tuple(start, date_system)?;
            let (end_year, _, _) = xls_serial_date_tuple(end, date_system)?;
            for year in start_year..=end_year {
                items.push(PivotValue::String(year.to_string()));
            }
        }
    }
    items.push(PivotValue::String(format!(
        ">{}",
        format_xls_pivot_date_bound(end, date_system)?
    )));
    Ok(items)
}

fn xls_date_group_item_id(
    unit: PivotDateGroupUnit,
    serial: f64,
    start: f64,
    end: f64,
    date_system: DateSystem,
) -> XlsResult<u32> {
    if serial < start {
        return Ok(0);
    }
    let (year, month, day) = xls_serial_date_tuple(serial, date_system)?;
    let (hour, minute, second) = serial_to_time(serial);
    let index = match unit {
        PivotDateGroupUnit::Seconds => 1 + u32::from(second),
        PivotDateGroupUnit::Minutes => 1 + u32::from(minute),
        PivotDateGroupUnit::Hours => 1 + u32::from(hour),
        PivotDateGroupUnit::Days => u32::from(day),
        PivotDateGroupUnit::Months => u32::from(month),
        PivotDateGroupUnit::Quarters => u32::from((month - 1) / 3 + 1),
        PivotDateGroupUnit::Years => {
            let (start_year, _, _) = xls_serial_date_tuple(start, date_system)?;
            1 + (year - start_year) as u32
        }
    };
    if serial > end {
        Ok(index.saturating_add(1))
    } else {
        Ok(index)
    }
}

fn xls_date_group_unit_name(unit: PivotDateGroupUnit) -> &'static str {
    match unit {
        PivotDateGroupUnit::Seconds => "Seconds",
        PivotDateGroupUnit::Minutes => "Minutes",
        PivotDateGroupUnit::Hours => "Hours",
        PivotDateGroupUnit::Days => "Days",
        PivotDateGroupUnit::Months => "Months",
        PivotDateGroupUnit::Quarters => "Quarters",
        PivotDateGroupUnit::Years => "Years",
    }
}

fn format_xls_pivot_date_bound(serial: f64, date_system: DateSystem) -> XlsResult<String> {
    let (year, month, day) = xls_serial_date_tuple(serial, date_system)?;
    Ok(format!("{month}/{day}/{year}"))
}

fn xls_serial_date_tuple(serial: f64, date_system: DateSystem) -> XlsResult<(i32, u32, u32)> {
    serial_to_date(serial, date_system).ok_or_else(|| {
        XlsError::InvalidFormat(format!(
            "XLS pivot date grouping has invalid date serial: {serial}"
        ))
    })
}

fn xls_date_group_by(unit: PivotDateGroupUnit) -> u8 {
    match unit {
        PivotDateGroupUnit::Seconds => 0x01,
        PivotDateGroupUnit::Minutes => 0x02,
        PivotDateGroupUnit::Hours => 0x03,
        PivotDateGroupUnit::Days => 0x04,
        PivotDateGroupUnit::Months => 0x05,
        PivotDateGroupUnit::Quarters => 0x06,
        PivotDateGroupUnit::Years => 0x07,
    }
}

fn write_pivot_sxdtr_record(
    stream: &mut Vec<u8>,
    serial: f64,
    date_system: DateSystem,
) -> XlsResult<()> {
    let Some((year, month, day)) = serial_to_date(serial, date_system) else {
        return Err(XlsError::InvalidFormat(format!(
            "XLS pivot date grouping has invalid date serial: {serial}"
        )));
    };
    if !(1900..=9999).contains(&year) {
        return Err(XlsError::InvalidFormat(format!(
            "XLS pivot date grouping date year is out of range: {year}"
        )));
    }
    let (hour, minute, second) = serial_to_time(serial);
    let mut payload = Vec::with_capacity(8);
    payload.extend_from_slice(&(year as u16).to_le_bytes());
    payload.extend_from_slice(&(month as u16).to_le_bytes());
    payload.push(day as u8);
    payload.push(hour as u8);
    payload.push(minute as u8);
    payload.push(second as u8);
    write_biff_record(stream, PIVOT_SXDTR_RECORD, &payload);
    Ok(())
}

fn source_number_min_max(numbers: &[f64]) -> Option<(f64, f64)> {
    let mut min = None::<f64>;
    let mut max = None::<f64>;
    for number in numbers {
        min = Some(min.map_or(*number, |current| current.min(*number)));
        max = Some(max.map_or(*number, |current| current.max(*number)));
    }
    min.zip(max)
}

fn valid_xls_date_serial(serial: f64, date_system: DateSystem) -> bool {
    serial.is_finite()
        && serial_to_date(serial, date_system).is_some_and(|(year, month, day)| {
            (1900..=9999).contains(&year) && (1..=12).contains(&month) && day <= 31
        })
}

fn workbook_date_system(date_1904: bool) -> DateSystem {
    if date_1904 {
        DateSystem::Date1904
    } else {
        DateSystem::Date1900
    }
}

fn write_pivot_calculated_field_formula_records(
    stream: &mut Vec<u8>,
    cache: &FormatPivotCache,
    formula: &str,
) -> XlsResult<()> {
    let trimmed = formula.trim();
    let parse_input = if trimmed.starts_with('=') {
        trimmed.to_string()
    } else {
        format!("={trimmed}")
    };
    let expr = duke_sheets_formula::parse_formula(&parse_input).map_err(|err| {
        XlsError::InvalidFormat(format!(
            "XLS pivot calculated field formula could not be parsed: {err}"
        ))
    })?;

    let mut ptgs = Vec::new();
    let mut sx_names = Vec::new();
    compile_pivot_calculated_formula_expr(&expr, &mut ptgs, &mut sx_names, cache).map_err(
        |_| {
            XlsError::InvalidFormat(format!(
                "XLS pivot calculated field formula uses unsupported syntax: {formula}"
            ))
        },
    )?;

    let mut body = Vec::new();
    body.extend_from_slice(
        &checked_u16(ptgs.len(), "pivot calculated formula token length")?.to_le_bytes(),
    );
    body.extend_from_slice(
        &checked_u16(
            sx_names.len(),
            "pivot calculated formula field reference count",
        )?
        .to_le_bytes(),
    );
    body.extend_from_slice(&ptgs);
    write_biff_record(stream, PIVOT_SXFMLA_RECORD, &body);

    for field_index in sx_names {
        write_pivot_sxname_record(stream, field_index)?;
    }

    Ok(())
}

fn write_pivot_sxname_record(stream: &mut Vec<u8>, field_index: usize) -> XlsResult<()> {
    let field_index = checked_i16(field_index, "pivot calculated formula source field index")?;
    let mut body = Vec::with_capacity(8);
    body.extend_from_slice(&0u16.to_le_bytes());
    body.extend_from_slice(&field_index.to_le_bytes());
    body.extend_from_slice(&0xFFFFu16.to_le_bytes());
    body.extend_from_slice(&0u16.to_le_bytes());
    write_biff_record(stream, PIVOT_SXNAME_RECORD, &body);
    Ok(())
}

fn compile_pivot_calculated_formula_expr(
    expr: &duke_sheets_formula::FormulaExpr,
    out: &mut Vec<u8>,
    sx_names: &mut Vec<usize>,
    cache: &FormatPivotCache,
) -> Result<(), UnsupportedToken> {
    use duke_sheets_formula::ast::{BinaryOperator, UnaryOperator};
    use duke_sheets_formula::FormulaExpr;

    match expr {
        FormulaExpr::Number(n) => {
            if let Some(i) = number_as_ptg_int(*n) {
                out.push(0x1E);
                out.extend_from_slice(&i.to_le_bytes());
            } else {
                out.push(0x1F);
                out.extend_from_slice(&n.to_le_bytes());
            }
        }
        FormulaExpr::String(s) => {
            out.push(0x17);
            push_short_xlunicode_string(out, s).map_err(|_| UnsupportedToken)?;
        }
        FormulaExpr::Boolean(b) => {
            out.push(0x1D);
            out.push(if *b { 1 } else { 0 });
        }
        FormulaExpr::Error(e) => {
            out.push(0x1C);
            out.push(e.code());
        }
        FormulaExpr::NameRef(name) => {
            push_pivot_sxname_ptg(out, sx_names, cache, name)?;
        }
        FormulaExpr::StructuredRef(reference) => {
            let Some(column) = reference.column.as_deref() else {
                return Err(UnsupportedToken);
            };
            push_pivot_sxname_ptg(out, sx_names, cache, column)?;
        }
        FormulaExpr::BinaryOp { op, left, right } => {
            compile_pivot_calculated_formula_expr(left, out, sx_names, cache)?;
            compile_pivot_calculated_formula_expr(right, out, sx_names, cache)?;
            out.push(match op {
                BinaryOperator::Add => 0x03,
                BinaryOperator::Subtract => 0x04,
                BinaryOperator::Multiply => 0x05,
                BinaryOperator::Divide => 0x06,
                BinaryOperator::Power => 0x07,
                BinaryOperator::Concat => 0x08,
                BinaryOperator::LessThan => 0x09,
                BinaryOperator::LessEqual => 0x0A,
                BinaryOperator::Equal => 0x0B,
                BinaryOperator::GreaterEqual => 0x0C,
                BinaryOperator::GreaterThan => 0x0D,
                BinaryOperator::NotEqual => 0x0E,
                BinaryOperator::Intersect | BinaryOperator::Union | BinaryOperator::Range => {
                    return Err(UnsupportedToken);
                }
            });
        }
        FormulaExpr::UnaryOp { op, operand } => {
            compile_pivot_calculated_formula_expr(operand, out, sx_names, cache)?;
            out.push(match op {
                UnaryOperator::Plus => 0x12,
                UnaryOperator::Negate => 0x13,
                UnaryOperator::Percent => 0x14,
                UnaryOperator::Paren => 0x15,
                UnaryOperator::ImplicitIntersection | UnaryOperator::SpillRange => {
                    return Err(UnsupportedToken);
                }
            });
        }
        FormulaExpr::CellRef(_)
        | FormulaExpr::RangeRef(_)
        | FormulaExpr::Function { .. }
        | FormulaExpr::ExternalFunction { .. }
        | FormulaExpr::Array(_)
        | FormulaExpr::ExternalRef(_)
        | FormulaExpr::Empty => return Err(UnsupportedToken),
    }

    Ok(())
}

fn push_pivot_sxname_ptg(
    out: &mut Vec<u8>,
    sx_names: &mut Vec<usize>,
    cache: &FormatPivotCache,
    field_name: &str,
) -> Result<(), UnsupportedToken> {
    let field_index = cache.field_index(field_name).ok_or(UnsupportedToken)?;
    let sx_name_index = checked_u32(sx_names.len(), "pivot calculated formula SXNAME index")
        .map_err(|_| UnsupportedToken)?;
    sx_names.push(field_index);
    out.extend_from_slice(&[0x18, 0x1D]);
    out.extend_from_slice(&sx_name_index.to_le_bytes());
    Ok(())
}

fn write_pivot_global_records(stream: &mut Vec<u8>, pivot_plan: &FormatPivotPlan) {
    if pivot_plan.caches.is_empty() {
        return;
    }

    write_biff_record(
        stream,
        0x089A,
        &[
            0x9A, 0x08, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x01, 0x00,
            0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x02, 0x00, 0x00, 0x00,
        ],
    );
    write_biff_record(
        stream,
        0x08A3,
        &[
            0xA3, 0x08, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00,
            0x00, 0x00,
        ],
    );
}

fn write_table_styles_record(stream: &mut Vec<u8>, pivot_plan: &FormatPivotPlan) {
    if pivot_plan.caches.is_empty() {
        return;
    }

    let default_table_style = "TableStyleMedium2";
    let default_pivot_style = "PivotStyleMedium9";
    let mut body = Vec::with_capacity(88);
    body.extend_from_slice(&TABLESTYLES_RECORD.to_le_bytes());
    body.extend_from_slice(&0u16.to_le_bytes());
    body.extend_from_slice(&[0; 8]);
    body.extend_from_slice(&0u32.to_le_bytes());
    body.extend_from_slice(&(default_table_style.encode_utf16().count() as u16).to_le_bytes());
    body.extend_from_slice(&(default_pivot_style.encode_utf16().count() as u16).to_le_bytes());
    for unit in default_table_style.encode_utf16() {
        body.extend_from_slice(&unit.to_le_bytes());
    }
    for unit in default_pivot_style.encode_utf16() {
        body.extend_from_slice(&unit.to_le_bytes());
    }
    write_biff_record(stream, TABLESTYLES_RECORD, &body);
}

fn write_pivot_workbook_extension_records(stream: &mut Vec<u8>, pivot_plan: &FormatPivotPlan) {
    if pivot_plan.caches.is_empty() {
        return;
    }

    write_biff_record(stream, COUNTRY_RECORD, &[0x01, 0x00, 0x01, 0x00]);
    write_biff_record(
        stream,
        RECALCID_RECORD,
        &[0xC1, 0x01, 0x00, 0x00, 0x35, 0xEA, 0x02, 0x00],
    );
    write_biff_record(
        stream,
        BOOKEXT_RECORD,
        &[
            0x63, 0x08, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x16, 0x00,
            0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x02, 0x00,
        ],
    );
    write_biff_record(
        stream,
        COMPRESSPICTURES_RECORD,
        &[
            0x9B, 0x08, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x01, 0x00,
            0x00, 0x00,
        ],
    );
    write_biff_record(
        stream,
        COMPAT12_RECORD,
        &[
            0x8C, 0x08, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00,
            0x00, 0x00,
        ],
    );
}

fn write_pivot_pre_boundsheet_records(
    stream: &mut Vec<u8>,
    pivot_plan: &FormatPivotPlan,
) -> XlsResult<()> {
    if pivot_plan.caches.is_empty() {
        return Ok(());
    }

    for cache in &pivot_plan.caches {
        write_biff_record(stream, 0x00D5, &(cache.cache_num as u16).to_le_bytes());
        write_biff_record(stream, 0x00E3, &1u16.to_le_bytes());
        write_dconref_record(stream, cache)?;
        for payload in [
            &[
                0x64, 0x08, 0x00, 0x00, 0x03, 0x00, 0x01, 0x00, 0x00, 0x00, 0x00, 0x00,
            ][..],
            &[
                0x64, 0x08, 0x00, 0x00, 0x03, 0x02, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0xFF, 0xFF,
                0xFF, 0xFF, 0x04, 0x03, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00,
            ][..],
            &[
                0x64, 0x08, 0x00, 0x00, 0x03, 0x18, 0x04, 0x00, 0x00, 0x00, 0x00, 0x00,
            ][..],
            &[
                0x64, 0x08, 0x00, 0x00, 0x03, 0x01, 0x02, 0x00, 0x00, 0x00, 0x00, 0x00,
            ][..],
            &[
                0x64, 0x08, 0x00, 0x00, 0x03, 0x41, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00,
            ][..],
            &[
                0x64, 0x08, 0x00, 0x00, 0x03, 0x34, 0x01, 0x00, 0x00, 0x00, 0x00, 0x00,
            ][..],
            &[
                0x64, 0x08, 0x00, 0x00, 0x03, 0x01, 0xFF, 0x00, 0x00, 0x00, 0x00, 0x00,
            ][..],
            &[
                0x64, 0x08, 0x00, 0x00, 0x03, 0xFF, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00,
            ][..],
        ] {
            write_biff_record(stream, 0x0864, payload);
        }
    }
    write_biff_record(stream, 0x0160, &0u16.to_le_bytes());
    Ok(())
}

fn write_dconref_record(stream: &mut Vec<u8>, cache: &FormatPivotCache) -> XlsResult<()> {
    let FormatPivotSource::Worksheet {
        sheet_name, range, ..
    } = &cache.source
    else {
        return Err(XlsError::InvalidFormat(
            "XLS pivot cache sources must be worksheet ranges".into(),
        ));
    };
    if range.start.row > u16::MAX as u32 || range.end.row > u16::MAX as u32 {
        return Err(XlsError::InvalidFormat(
            "XLS pivot source range exceeds BIFF8 row limits".into(),
        ));
    }
    if range.start.col > u8::MAX as u16 || range.end.col > u8::MAX as u16 {
        return Err(XlsError::InvalidFormat(
            "XLS pivot source range exceeds BIFF8 column limits".into(),
        ));
    }

    let encoded_len = 1usize + sheet_name.encode_utf16().count();
    if encoded_len > u16::MAX as usize {
        return Err(XlsError::InvalidFormat(
            "XLS pivot source sheet name is too long".into(),
        ));
    }

    let mut body = Vec::new();
    body.extend_from_slice(&(range.start.row as u16).to_le_bytes());
    body.extend_from_slice(&(range.end.row as u16).to_le_bytes());
    body.push(range.start.col as u8);
    body.push(range.end.col as u8);
    body.extend_from_slice(&(encoded_len as u16).to_le_bytes());
    if sheet_name.is_ascii() {
        body.push(0);
        body.push(0x02);
        body.extend_from_slice(sheet_name.as_bytes());
    } else {
        body.push(1);
        body.extend_from_slice(&0x0002u16.to_le_bytes());
        for unit in sheet_name.encode_utf16() {
            body.extend_from_slice(&unit.to_le_bytes());
        }
    }
    write_biff_record(stream, 0x0051, &body);
    Ok(())
}

fn write_pivot_sheet_records(
    stream: &mut Vec<u8>,
    workbook: &Workbook,
    pivot_plan: &FormatPivotPlan,
    pivot_layouts: &XlsPivotCacheLayouts,
    sheet_idx: usize,
) -> XlsResult<()> {
    for part in pivot_plan
        .tables
        .iter()
        .filter(|part| part.sheet_index == sheet_idx)
    {
        let cache = pivot_plan
            .caches
            .iter()
            .find(|cache| cache.cache_num == part.cache_num)
            .ok_or_else(|| XlsError::InvalidFormat("pivot cache part not found".into()))?;
        let sheet = workbook
            .worksheet(sheet_idx)
            .ok_or_else(|| XlsError::InvalidFormat("pivot sheet not found".into()))?;
        let pivot = sheet
            .pivot_tables()
            .get(part.pivot_index)
            .ok_or_else(|| XlsError::InvalidFormat("pivot table not found".into()))?;
        validate_xls_pivot_groupings(cache, &pivot.groupings)?;
        validate_xls_pivot_grouping_axes(pivot)?;
        let layout = pivot_layouts.get(cache.cache_num)?;

        let multi_measure = pivot.measures.len() > 1;
        if pivot.rows.len() != 1
            || pivot.columns.len() > 1
            || pivot.page_fields.len() > 1
            || pivot.measures.is_empty()
            || (multi_measure
                && (!pivot.page_fields.is_empty() || !xls_values_field_on_columns(pivot)))
        {
            return Err(XlsError::InvalidFormat(format!(
                "XLS pivot table {} uses a layout this BIFF8 writer slice does not encode yet",
                pivot.name
            )));
        }

        write_classic_pivot_view_records(stream, pivot, cache, &layout)?;
        let has_expanded_row_axis = expanded_axis_field_count(layout, &pivot.rows)?
            > checked_u16(pivot.rows.len(), "pivot visible row field count")?;
        write_biff_record(
            stream,
            0x00F1,
            &[
                0x00, 0x00, 0xFF, 0xFF, 0xFF, 0xFF, 0xFF, 0xFF, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00,
                0x00, 0x02, 0x4F, 0x00, 0xFF, 0xFF, 0xFF, 0xFF, 0xFF, 0xFF,
            ],
        );
        write_sxview_record(stream, pivot, cache, has_expanded_row_axis)?;
        let sxview_ex_flags = if has_expanded_row_axis { 0x24 } else { 0x04 };
        write_biff_record(
            stream,
            0x0810,
            &[
                0x10,
                0x08,
                0x02,
                0x00,
                0x00,
                0x00,
                0x00,
                0x00,
                sxview_ex_flags,
                0x00,
                0x00,
                0x00,
                0x01,
                0x00,
                0x00,
                0x00,
                0x00,
            ],
        );
        write_pivot_frt_records(stream, pivot, &layout, has_expanded_row_axis)?;
    }
    Ok(())
}

fn write_classic_pivot_view_records(
    stream: &mut Vec<u8>,
    pivot: &duke_sheets_core::PivotTable,
    cache: &FormatPivotCache,
    layout: &XlsPivotCacheLayout,
) -> XlsResult<()> {
    write_sxview_record_biff8(stream, pivot, cache, layout)?;
    for (field_index, field) in layout.fields.iter().enumerate() {
        let axis = xls_pivot_field_axis(pivot, layout, field_index);
        write_sxvd_record(stream, field_index, field, axis)?;
    }
    let values_on_columns = xls_values_field_on_columns(pivot);
    write_sxivd_record(stream, layout, &pivot.rows)?;
    if values_on_columns {
        write_values_sxivd_record(stream);
    } else if !pivot.columns.is_empty() {
        write_sxivd_record(stream, layout, &pivot.columns)?;
    }
    write_sxpi_records(stream, pivot, layout)?;
    write_sxdi_records(stream, pivot, layout)?;
    write_sxli_collection(stream, pivot, layout, &pivot.rows)?;
    if values_on_columns {
        write_values_axis_sxli_collection(stream, pivot)?;
    } else {
        write_sxli_collection(stream, pivot, layout, &pivot.columns)?;
    }
    Ok(())
}

fn write_sxview_record_biff8(
    stream: &mut Vec<u8>,
    pivot: &duke_sheets_core::PivotTable,
    cache: &FormatPivotCache,
    layout: &XlsPivotCacheLayout,
) -> XlsResult<()> {
    let data_caption = if pivot.layout.data_caption.trim().is_empty() {
        "Values"
    } else {
        pivot.layout.data_caption.as_str()
    };
    let row_line_count = axis_line_count(pivot, layout, &pivot.rows)?;
    let values_on_columns = xls_values_field_on_columns(pivot);
    let column_line_count = if values_on_columns {
        checked_u16(pivot.measures.len(), "pivot values axis line count")?
    } else {
        axis_line_count(pivot, layout, &pivot.columns)?
    };
    let row_axis_count = expanded_axis_field_count(layout, &pivot.rows)?;
    let visible_row_axis_count = checked_u16(pivot.rows.len(), "pivot visible row field count")?;
    let column_axis_count = expanded_axis_field_count(layout, &pivot.columns)?
        .saturating_add(if values_on_columns { 1 } else { 0 });
    let page_axis_count = checked_u16(pivot.page_fields.len(), "pivot page field count")?;
    let data_field_count = checked_u16(pivot.measures.len(), "pivot data field count")?;
    let (page_rows, _) = page_field_area_size(pivot);
    let first_row_offset = if page_rows == 0 {
        0
    } else {
        page_rows.saturating_add(1)
    };
    let first_row = checked_biff8_row(
        pivot.target.row.saturating_add(first_row_offset),
        "pivot target row",
    )?;
    let first_col = checked_biff8_col(pivot.target.col, "pivot target column")?;
    let first_header_row = if values_on_columns && pivot.columns.is_empty() {
        first_row
    } else {
        first_row.saturating_add(1)
    };
    let first_data_row = if values_on_columns && pivot.columns.is_empty() {
        first_row.saturating_add(1)
    } else {
        first_header_row.saturating_add(column_axis_count)
    };
    let first_data_col = first_col.saturating_add(visible_row_axis_count);
    let last_row = if values_on_columns && pivot.columns.is_empty() {
        first_row.saturating_add(row_line_count)
    } else {
        first_row
            .saturating_add(column_axis_count)
            .saturating_add(row_line_count)
    };
    let last_col = first_col
        .saturating_add(visible_row_axis_count)
        .saturating_add(column_line_count)
        .saturating_sub(1);
    let mut body = Vec::new();
    body.extend_from_slice(&first_row.to_le_bytes());
    body.extend_from_slice(&last_row.to_le_bytes());
    body.extend_from_slice(&first_col.to_le_bytes());
    body.extend_from_slice(&last_col.to_le_bytes());
    body.extend_from_slice(&first_header_row.to_le_bytes());
    body.extend_from_slice(&first_data_row.to_le_bytes());
    body.extend_from_slice(&first_data_col.to_le_bytes());
    body.extend_from_slice(&((cache.cache_num.saturating_sub(1)) as u16).to_le_bytes());
    body.extend_from_slice(&0u16.to_le_bytes());
    body.extend_from_slice(&0x0002u16.to_le_bytes());
    body.extend_from_slice(&(-1i16).to_le_bytes());
    body.extend_from_slice(
        &checked_u16(layout.fields.len(), "pivot view field count")?.to_le_bytes(),
    );
    body.extend_from_slice(&row_axis_count.to_le_bytes());
    body.extend_from_slice(&column_axis_count.to_le_bytes());
    body.extend_from_slice(&page_axis_count.to_le_bytes());
    body.extend_from_slice(&data_field_count.to_le_bytes());
    body.extend_from_slice(&row_line_count.to_le_bytes());
    body.extend_from_slice(&column_line_count.to_le_bytes());
    let view_flags: u16 = if column_axis_count > 0
        || page_axis_count > 0
        || row_axis_count > visible_row_axis_count
    {
        0x020B
    } else {
        0x0003
    };
    body.extend_from_slice(&view_flags.to_le_bytes());
    body.extend_from_slice(&0x0001u16.to_le_bytes());
    body.extend_from_slice(&xlunicode_len_u16(&pivot.name)?.to_le_bytes());
    body.extend_from_slice(&xlunicode_len_u16(data_caption)?.to_le_bytes());
    push_xlunicode_string_no_cch(&mut body, &pivot.name)?;
    push_xlunicode_string_no_cch(&mut body, data_caption)?;
    write_biff_record(stream, 0x00B0, &body);
    Ok(())
}

fn write_sxvd_record(
    stream: &mut Vec<u8>,
    field_index: usize,
    field: &XlsPivotFieldLayout,
    axis: u16,
) -> XlsResult<()> {
    if field_index > u16::MAX as usize {
        return Err(XlsError::InvalidFormat(
            "pivot field index exceeds BIFF8 limits".into(),
        ));
    }
    let item_count = if matches!(axis, 0x0001 | 0x0002 | 0x0004)
        || matches!(field.kind, XlsPivotFieldKind::DateSource { .. })
    {
        checked_u16(
            field.shared_items.len().saturating_add(1),
            "pivot item count",
        )?
    } else {
        0
    };

    let mut body = Vec::new();
    body.extend_from_slice(&axis.to_le_bytes());
    body.extend_from_slice(&1u16.to_le_bytes());
    body.extend_from_slice(&1u16.to_le_bytes());
    body.extend_from_slice(&item_count.to_le_bytes());
    body.extend_from_slice(&(-1i16).to_le_bytes());
    write_biff_record(stream, 0x00B1, &body);

    if item_count > 0 {
        for item_index in 0..field.shared_items.len() {
            let item_index = checked_u16(item_index, "pivot item index")?;
            let mut item = Vec::new();
            item.extend_from_slice(&0u16.to_le_bytes());
            item.extend_from_slice(&0u16.to_le_bytes());
            item.extend_from_slice(&item_index.to_le_bytes());
            item.extend_from_slice(&(-1i16).to_le_bytes());
            write_biff_record(stream, 0x00B2, &item);
        }
        let mut default_item = Vec::new();
        default_item.extend_from_slice(&1u16.to_le_bytes());
        default_item.extend_from_slice(&0u16.to_le_bytes());
        default_item.extend_from_slice(&(-1i16).to_le_bytes());
        default_item.extend_from_slice(&(-1i16).to_le_bytes());
        write_biff_record(stream, 0x00B2, &default_item);
    }

    write_sxvdex_record(stream, axis);
    Ok(())
}

fn write_sxvdex_record(stream: &mut Vec<u8>, _axis: u16) {
    let mut body = Vec::new();
    body.extend_from_slice(&[0x1E, 0x14, 0xA0, 0x0A]);
    body.extend_from_slice(&(-1i32).to_le_bytes());
    body.extend_from_slice(&0u16.to_le_bytes());
    body.extend_from_slice(&(-1i16).to_le_bytes());
    body.extend_from_slice(&0u64.to_le_bytes());
    write_biff_record(stream, 0x0100, &body);
}

fn write_sxivd_record(
    stream: &mut Vec<u8>,
    layout: &XlsPivotCacheLayout,
    fields: &[duke_sheets_core::PivotField],
) -> XlsResult<()> {
    if fields.is_empty() {
        return Ok(());
    }
    let mut body = Vec::new();
    for field_index in expanded_axis_field_indexes(layout, fields)? {
        body.extend_from_slice(&checked_u16(field_index, "pivot axis field index")?.to_le_bytes());
    }
    write_biff_record(stream, 0x00B4, &body);
    Ok(())
}

fn write_values_sxivd_record(stream: &mut Vec<u8>) {
    write_biff_record(stream, 0x00B4, &(-2i16).to_le_bytes());
}

fn write_sxpi_records(
    stream: &mut Vec<u8>,
    pivot: &duke_sheets_core::PivotTable,
    layout: &XlsPivotCacheLayout,
) -> XlsResult<()> {
    for field in &pivot.page_fields {
        let field_index = layout.axis_field_index(&field.field.name).ok_or_else(|| {
            XlsError::InvalidFormat(format!(
                "pivot references unknown page field {}",
                field.field.name
            ))
        })?;
        let selected_item =
            selected_page_item_index(pivot, &field.field.name, &layout.fields[field_index])
                .unwrap_or(0xFFFF);
        let mut body = Vec::new();
        body.extend_from_slice(&checked_u16(field_index, "pivot page field index")?.to_le_bytes());
        body.extend_from_slice(&selected_item.to_le_bytes());
        body.extend_from_slice(&1u16.to_le_bytes());
        write_biff_record(stream, 0x00B6, &body);
    }
    Ok(())
}

fn write_sxdi_records(
    stream: &mut Vec<u8>,
    pivot: &duke_sheets_core::PivotTable,
    layout: &XlsPivotCacheLayout,
) -> XlsResult<()> {
    if pivot.measures.is_empty() {
        return Err(XlsError::InvalidFormat(
            "pivot table has no data field".into(),
        ));
    }

    for measure in &pivot.measures {
        let field_index = layout.field_index(&measure.field.name).ok_or_else(|| {
            XlsError::InvalidFormat(format!(
                "pivot references unknown measure field {}",
                measure.field.name
            ))
        })?;
        if field_index > u16::MAX as usize {
            return Err(XlsError::InvalidFormat(
                "pivot data field index exceeds BIFF8 limits".into(),
            ));
        }

        let mut body = Vec::new();
        body.extend_from_slice(&(field_index as u16).to_le_bytes());
        body.extend_from_slice(&xls_pivot_aggregate_code(measure.aggregate).to_le_bytes());
        body.extend_from_slice(&0u16.to_le_bytes());
        body.extend_from_slice(&0u16.to_le_bytes());
        body.extend_from_slice(&0u16.to_le_bytes());
        body.extend_from_slice(&0u16.to_le_bytes());
        push_xlunicode_string(&mut body, &measure.caption())?;
        write_biff_record(stream, 0x00C5, &body);
    }
    Ok(())
}

fn write_sxli_collection(
    stream: &mut Vec<u8>,
    pivot: &duke_sheets_core::PivotTable,
    layout: &XlsPivotCacheLayout,
    fields: &[PivotField],
) -> XlsResult<()> {
    let mut body = Vec::new();
    if fields.is_empty() {
        write_sxli_item(&mut body, &[], false);
    } else {
        let tuples = axis_item_tuples(pivot, layout, fields)?;
        for tuple in &tuples {
            write_sxli_item(&mut body, tuple, false);
        }
        write_sxli_item(
            &mut body,
            &vec![0; expanded_axis_field_count(layout, fields)? as usize],
            true,
        );
    }
    write_biff_record(stream, 0x00B5, &body);
    Ok(())
}

fn write_sxli_item(body: &mut Vec<u8>, item_indexes: &[u16], grand_total: bool) {
    let item_type: u16 = if grand_total { 0x000D } else { 0x0000 };
    let item_flags: u16 = if grand_total { 0x0A00 } else { 0x0000 };
    write_sxli_structure(body, 0, item_type, item_flags, item_indexes);
}

fn write_sxli_structure(
    body: &mut Vec<u8>,
    c_sic: u16,
    item_type: u16,
    item_flags: u16,
    item_indexes: &[u16],
) {
    body.extend_from_slice(&c_sic.to_le_bytes());
    body.extend_from_slice(&item_type.to_le_bytes());
    let isxvi_mac = item_indexes.len().saturating_sub(1) as u16;
    body.extend_from_slice(&isxvi_mac.to_le_bytes());
    body.extend_from_slice(&item_flags.to_le_bytes());
    for item_index in item_indexes {
        body.extend_from_slice(&item_index.to_le_bytes());
    }
}

fn write_values_axis_sxli_collection(
    stream: &mut Vec<u8>,
    pivot: &duke_sheets_core::PivotTable,
) -> XlsResult<()> {
    let mut body = Vec::new();
    for data_item in 0..pivot.measures.len() {
        let data_item = checked_u16(data_item, "pivot data item index")?;
        body.extend_from_slice(&0u16.to_le_bytes());
        body.extend_from_slice(&0u16.to_le_bytes());
        body.extend_from_slice(&0u16.to_le_bytes());
        body.extend_from_slice(&(0x1000u16 | data_item.saturating_mul(2)).to_le_bytes());
        body.extend_from_slice(&data_item.to_le_bytes());
    }
    write_biff_record(stream, 0x00B5, &body);
    Ok(())
}

fn axis_line_count(
    pivot: &duke_sheets_core::PivotTable,
    layout: &XlsPivotCacheLayout,
    fields: &[PivotField],
) -> XlsResult<u16> {
    if fields.is_empty() {
        return Ok(1);
    }
    checked_u16(
        axis_item_tuples(pivot, layout, fields)?
            .len()
            .saturating_add(1),
        "pivot line count",
    )
}

fn axis_item_tuples(
    pivot: &duke_sheets_core::PivotTable,
    layout: &XlsPivotCacheLayout,
    fields: &[PivotField],
) -> XlsResult<Vec<Vec<u16>>> {
    let indexes = expanded_axis_field_indexes(layout, fields)?;

    let mut seen = HashSet::new();
    let mut tuples = Vec::new();
    for row in 0..layout.row_count {
        if !row_matches_page_filters(pivot, layout, row)? {
            continue;
        }
        let tuple = indexes
            .iter()
            .map(|index| {
                let item_id = layout.fields[*index]
                    .item_ids
                    .get(row)
                    .copied()
                    .unwrap_or(0);
                checked_u16(item_id as usize, "pivot item index")
            })
            .collect::<XlsResult<Vec<_>>>()?;
        if seen.insert(tuple.clone()) {
            tuples.push(tuple);
        }
    }
    Ok(tuples)
}

fn expanded_axis_field_count(
    layout: &XlsPivotCacheLayout,
    fields: &[PivotField],
) -> XlsResult<u16> {
    checked_u16(
        expanded_axis_field_indexes(layout, fields)?.len(),
        "pivot axis field count",
    )
}

fn expanded_axis_field_indexes(
    layout: &XlsPivotCacheLayout,
    fields: &[PivotField],
) -> XlsResult<Vec<usize>> {
    let mut indexes = Vec::new();
    for field in fields {
        let field_index = layout.axis_field_index(&field.field.name).ok_or_else(|| {
            XlsError::InvalidFormat(format!(
                "pivot references unknown axis field {}",
                field.field.name
            ))
        })?;
        match layout.fields.get(field_index).map(|field| &field.kind) {
            Some(XlsPivotFieldKind::DateSource {
                derived_field_index,
                ..
            }) => indexes.push(*derived_field_index),
            Some(XlsPivotFieldKind::ManualSource {
                derived_field_index,
            }) => {
                indexes.push(*derived_field_index);
                indexes.push(field_index);
            }
            _ => indexes.push(field_index),
        }
    }
    Ok(indexes)
}

fn row_matches_page_filters(
    pivot: &duke_sheets_core::PivotTable,
    layout: &XlsPivotCacheLayout,
    row: usize,
) -> XlsResult<bool> {
    for field in &pivot.page_fields {
        let field_index = layout.axis_field_index(&field.field.name).ok_or_else(|| {
            XlsError::InvalidFormat(format!(
                "pivot references unknown page field {}",
                field.field.name
            ))
        })?;
        let Some(selected_item) =
            selected_page_item_index(pivot, &field.field.name, &layout.fields[field_index])
        else {
            continue;
        };
        let row_item = layout.fields[field_index]
            .item_ids
            .get(row)
            .copied()
            .unwrap_or(0);
        if row_item != u32::from(selected_item) {
            return Ok(false);
        }
    }
    Ok(true)
}

fn selected_page_item_index(
    pivot: &duke_sheets_core::PivotTable,
    field_name: &str,
    field: &XlsPivotFieldLayout,
) -> Option<u16> {
    let PivotFilter::FieldItems { allowed_items, .. } = pivot.filters.iter().find(|filter| {
        matches!(
            filter,
            PivotFilter::FieldItems {
                field: filter_field,
                ..
            } if filter_field.name.eq_ignore_ascii_case(field_name)
        )
    })?
    else {
        return None;
    };

    let [item] = allowed_items.as_slice() else {
        return None;
    };
    field
        .shared_items
        .iter()
        .position(|candidate| candidate == item)
        .and_then(|index| u16::try_from(index).ok())
}

fn page_field_area_size(pivot: &duke_sheets_core::PivotTable) -> (u32, u32) {
    let count = pivot.page_fields.len();
    if count == 0 {
        return (0, 0);
    }

    let wrap = pivot.layout.page_wrap as usize;
    let row_count = if wrap == 0 {
        count
    } else if pivot.layout.page_over_then_down {
        (count + wrap - 1) / wrap
    } else {
        wrap.min(count)
    };
    let col_count = if wrap == 0 {
        1
    } else if pivot.layout.page_over_then_down {
        wrap.min(count)
    } else {
        (count + row_count - 1) / row_count
    };
    (row_count as u32, col_count as u32)
}

fn checked_u16(value: usize, label: &str) -> XlsResult<u16> {
    if value > u16::MAX as usize {
        return Err(XlsError::InvalidFormat(format!(
            "{label} exceeds BIFF8 limits"
        )));
    }
    Ok(value as u16)
}

fn checked_u8(value: usize, label: &str) -> XlsResult<u8> {
    if value > u8::MAX as usize {
        return Err(XlsError::InvalidFormat(format!(
            "{label} exceeds BIFF8 limits"
        )));
    }
    Ok(value as u8)
}

fn checked_i16(value: usize, label: &str) -> XlsResult<i16> {
    if value > i16::MAX as usize {
        return Err(XlsError::InvalidFormat(format!(
            "{label} exceeds BIFF8 limits"
        )));
    }
    Ok(value as i16)
}

fn checked_u32(value: usize, label: &str) -> XlsResult<u32> {
    if value > u32::MAX as usize {
        return Err(XlsError::InvalidFormat(format!(
            "{label} exceeds BIFF8 limits"
        )));
    }
    Ok(value as u32)
}

fn checked_biff8_row(value: u32, label: &str) -> XlsResult<u16> {
    if value > u16::MAX as u32 {
        return Err(XlsError::InvalidFormat(format!(
            "{label} exceeds BIFF8 row limits"
        )));
    }
    Ok(value as u16)
}

fn checked_biff8_col(value: u16, label: &str) -> XlsResult<u16> {
    if value > u8::MAX as u16 {
        return Err(XlsError::InvalidFormat(format!(
            "{label} exceeds BIFF8 column limits"
        )));
    }
    Ok(value)
}

fn xls_pivot_field_axis(
    pivot: &duke_sheets_core::PivotTable,
    layout: &XlsPivotCacheLayout,
    field_index: usize,
) -> u16 {
    let Some(field) = layout.fields.get(field_index) else {
        return 0x0000;
    };
    let field_name = match field.kind {
        XlsPivotFieldKind::DateSource { .. } => return 0x0000,
        XlsPivotFieldKind::DateGroup {
            source_field_index, ..
        }
        | XlsPivotFieldKind::ManualGroup {
            source_field_index, ..
        } => layout
            .fields
            .get(source_field_index)
            .map(|source| source.name.as_str())
            .unwrap_or(field.name.as_str()),
        _ => field.name.as_str(),
    };
    if pivot
        .rows
        .iter()
        .any(|field| field.field.name.eq_ignore_ascii_case(field_name))
    {
        0x0001
    } else if pivot
        .columns
        .iter()
        .any(|field| field.field.name.eq_ignore_ascii_case(field_name))
    {
        0x0002
    } else if pivot
        .page_fields
        .iter()
        .any(|field| field.field.name.eq_ignore_ascii_case(field_name))
    {
        0x0004
    } else if pivot
        .measures
        .iter()
        .any(|measure| measure.field.name.eq_ignore_ascii_case(field_name))
    {
        0x0008
    } else {
        0x0000
    }
}

fn xls_values_field_on_columns(pivot: &duke_sheets_core::PivotTable) -> bool {
    pivot.measures.len() > 1
        && pivot.columns.is_empty()
        && matches!(pivot.layout.values_axis, PivotValuesAxis::Columns)
}

fn xls_pivot_aggregate_code(aggregate: PivotAggregate) -> u16 {
    match aggregate {
        PivotAggregate::Sum => 0,
        PivotAggregate::Count => 1,
        PivotAggregate::Average => 2,
        PivotAggregate::Max => 3,
        PivotAggregate::Min => 4,
        PivotAggregate::Product => 5,
        PivotAggregate::CountNumbers => 6,
        PivotAggregate::StdDev => 7,
        PivotAggregate::StdDevP => 8,
        PivotAggregate::Var => 9,
        PivotAggregate::VarP => 10,
    }
}

fn write_pivot_sheet_tail_records(
    stream: &mut Vec<u8>,
    sheet: &Worksheet,
    pivot_plan: &FormatPivotPlan,
    sheet_idx: usize,
) {
    for _part in pivot_plan
        .tables
        .iter()
        .filter(|part| part.sheet_index == sheet_idx)
    {
        write_biff_record(
            stream,
            0x088B,
            &[
                0x8B, 0x08, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00,
                0x12, 0x00,
            ],
        );
        if sheet.selections().is_empty() {
            write_default_selection_record(stream);
        } else {
            write_selection_records(stream, sheet);
        }
        write_biff_record(
            stream,
            0x0867,
            &[
                0x67, 0x08, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x02, 0x00,
                0x01, 0xFF, 0xFF, 0xFF, 0xFF, 0x03, 0x44, 0x00, 0x00,
            ],
        );
    }
}

fn write_default_selection_record(stream: &mut Vec<u8>) {
    write_biff_record(
        stream,
        SELECTION_RECORD,
        &[
            0x03, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x01, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00,
            0x00,
        ],
    );
}

fn write_sxview_record(
    stream: &mut Vec<u8>,
    pivot: &duke_sheets_core::PivotTable,
    _cache: &FormatPivotCache,
    has_expanded_row_axis: bool,
) -> XlsResult<()> {
    let mut body = Vec::new();
    body.extend_from_slice(&[0x02, 0x08, 0x00, 0x00]);
    body.extend_from_slice(&1u16.to_le_bytes());
    body.extend_from_slice(&2u16.to_le_bytes());
    body.extend_from_slice(&0u32.to_le_bytes());
    body.extend_from_slice(&[0x08, 0x03, 0x10, 0x00]);
    push_xlunicode_string(&mut body, &pivot.name)?;
    body.push(0);
    body.push(if has_expanded_row_axis { 2 } else { 1 });
    write_biff_record(stream, 0x0802, &body);
    Ok(())
}

fn write_pivot_frt_records(
    stream: &mut Vec<u8>,
    pivot: &duke_sheets_core::PivotTable,
    layout: &XlsPivotCacheLayout,
    has_expanded_row_axis: bool,
) -> XlsResult<()> {
    write_frt0864_name(stream, 0x0000, &pivot.name)?;
    if has_expanded_row_axis {
        write_frt0864_raw(stream, &[0x00, 0x02, 0x08, 0x41, 0x40, 0x00, 0x00, 0x00]);
        write_frt0864_raw(stream, &[0x00, 0x19, 0x9F, 0x00, 0x40, 0x00, 0x00, 0x00]);
    } else {
        write_frt0864_raw(stream, &[0x00, 0x02, 0x00, 0x41, 0x40, 0x01, 0x00, 0x00]);
        write_frt0864_raw(stream, &[0x00, 0x19, 0x19, 0x00, 0x40, 0x01, 0x00, 0x00]);
    }
    for field in &layout.fields {
        write_frt0864_name(stream, 0x0017, &field.name)?;
        write_frt0864_raw(stream, &[0x17, 0x19, 0x28, 0x00, 0x00, 0x00, 0x00, 0x00]);
        write_frt0864_raw(stream, &[0x17, 0x01, 0x02, 0x00, 0x00, 0x00, 0x00, 0x00]);
        write_frt0864_raw(stream, &[0x17, 0x01, 0xFF, 0x00, 0x00, 0x00, 0x00, 0x00]);
        write_frt0864_raw(stream, &[0x17, 0xFF, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00]);
    }
    write_frt0864_style(stream, &pivot.style)?;
    write_frt0864_raw(stream, &[0x00, 0x01, 0x02, 0x00, 0x00, 0x00, 0x00, 0x00]);
    write_frt0864_raw(stream, &[0x00, 0x01, 0xFF, 0x00, 0x00, 0x00, 0x00, 0x00]);
    write_frt0864_raw(stream, &[0x00, 0xFF, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00]);
    Ok(())
}

fn write_frt0864_name(stream: &mut Vec<u8>, subtype: u16, name: &str) -> XlsResult<()> {
    let mut tail = Vec::new();
    tail.extend_from_slice(&subtype.to_le_bytes());
    tail.extend_from_slice(&(name.encode_utf16().count() as u32).to_le_bytes());
    tail.extend_from_slice(&0u16.to_le_bytes());
    push_xlunicode_string(&mut tail, name)?;
    write_frt0864_raw(stream, &tail);
    Ok(())
}

fn write_frt0864_style(
    stream: &mut Vec<u8>,
    style: &duke_sheets_core::PivotStyle,
) -> XlsResult<()> {
    let style_name = style
        .name
        .as_deref()
        .filter(|name| !name.trim().is_empty())
        .unwrap_or("PivotStyleMedium9");
    let mut tail = Vec::new();
    tail.extend_from_slice(&0x1E00u16.to_le_bytes());
    tail.extend_from_slice(&0u32.to_le_bytes());
    tail.extend_from_slice(&0u16.to_le_bytes());
    tail.extend_from_slice(&pivot_style_flags(style).to_le_bytes());
    tail.extend_from_slice(
        &checked_u16(style_name.encode_utf16().count(), "pivot style name length")?.to_le_bytes(),
    );
    for unit in style_name.encode_utf16() {
        tail.extend_from_slice(&unit.to_le_bytes());
    }
    write_frt0864_raw(stream, &tail);
    Ok(())
}

fn pivot_style_flags(style: &duke_sheets_core::PivotStyle) -> u16 {
    let mut flags = 0u16;
    if style.show_last_column {
        flags |= 0x02;
    }
    if style.show_row_stripes {
        flags |= 0x04;
    }
    if style.show_column_stripes {
        flags |= 0x08;
    }
    if style.show_row_headers {
        flags |= 0x10;
    }
    if style.show_column_headers {
        flags |= 0x20;
    }
    flags
}

fn write_frt0864_raw(stream: &mut Vec<u8>, tail: &[u8]) {
    let mut body = Vec::with_capacity(4 + tail.len());
    body.extend_from_slice(&[0x64, 0x08, 0x00, 0x00]);
    body.extend_from_slice(tail);
    write_biff_record(stream, 0x0864, &body);
}

fn pivot_value_string_payload(value: &PivotValue) -> XlsResult<Vec<u8>> {
    let text = match value {
        PivotValue::Blank => String::new(),
        PivotValue::Boolean(value) => {
            if *value {
                "TRUE".to_string()
            } else {
                "FALSE".to_string()
            }
        }
        PivotValue::Number(value) => value.to_string(),
        PivotValue::String(value) => value.clone(),
        PivotValue::Error(value) => value.to_string(),
    };
    let mut body = Vec::new();
    push_xlunicode_string(&mut body, &text)?;
    Ok(body)
}

fn field_is_numeric(field: &XlsPivotFieldLayout) -> bool {
    !field.shared_items.is_empty()
        && field
            .shared_items
            .iter()
            .all(|value| matches!(value, PivotValue::Number(_)))
}

/// Workbook-global font + format + XF tables. Slice 4a/4b customize
/// the font_index and format_index axes of an XF record; alignment,
/// fill, and border stay at defaults.
struct StyleTables {
    /// FONT records to emit, in disk order. The first
    /// `FONT_BUILTIN_COUNT` entries are the BIFF8-required built-ins
    /// (all defaulted Calibri 11). User-defined fonts append after.
    fonts_in_order: Vec<FontStyle>,
    /// `FontStyle -> on-disk-font-index` (with the BIFF8 "skip 4"
    /// quirk applied: positions 0..3 are the index, position 4 is
    /// reserved, position 5+ become index +1). Exposed so the SST
    /// rich-text emitter can look up font indices for run fonts that
    /// were also interned in this table.
    font_xf_index: HashMap<FontStyle, u16>,
    /// User-defined number-format strings to emit as FORMAT records.
    /// Their on-disk ifmt values start at `FORMAT_USER_INDEX_BASE` and
    /// increment by one per entry.
    user_formats: Vec<String>,
    /// User-defined XF records to emit after the 16 built-ins.
    user_xfs: Vec<UserXf>,
    /// `(sheet_idx, style_index_in_pool) -> ixfe`. Cells consult this
    /// to pick their `XF` reference; absent entries fall back to 0.
    cell_ixfe: HashMap<(usize, u32), u16>,
}

#[derive(Debug, Clone)]
struct UserXf {
    font_index: u16,
    format_index: u16,
    alignment: Alignment,
    border: BorderStyle,
    fill: FillStyle,
    protection: duke_sheets_core::style::Protection,
}

impl StyleTables {
    fn collect(workbook: &Workbook) -> Self {
        let default_font = FontStyle::default();
        let mut fonts_in_order = vec![default_font.clone(); FONT_BUILTIN_COUNT as usize];
        let mut font_xf_index: HashMap<FontStyle, u16> = HashMap::new();
        font_xf_index.insert(default_font, 0u16);

        let mut user_formats: Vec<String> = Vec::new();
        let mut format_index_for_custom: HashMap<String, u16> = HashMap::new();

        let mut user_xfs: Vec<UserXf> = Vec::new();
        type XfKey = (
            u16,
            u16,
            Alignment,
            BorderStyle,
            FillStyle,
            duke_sheets_core::style::Protection,
        );
        let mut xf_key_to_ixfe: HashMap<XfKey, u16> = HashMap::new();
        let mut cell_ixfe = HashMap::new();

        for (sheet_idx, sheet) in workbook.worksheets().enumerate() {
            for (_row, _col, cell) in sheet.iter_cells() {
                if cell.style_index == 0 {
                    continue;
                }
                let Some(style) = sheet.style_by_index(cell.style_index) else {
                    continue;
                };

                let font = style.font.clone();
                let font_idx = match font_xf_index.get(&font) {
                    Some(&idx) => idx,
                    None => {
                        let on_disk = fonts_in_order.len() as u16;
                        let xf_idx = if on_disk < 4 { on_disk } else { on_disk + 1 };
                        fonts_in_order.push(font.clone());
                        font_xf_index.insert(font.clone(), xf_idx);
                        xf_idx
                    }
                };

                let format_idx = match &style.number_format {
                    NumberFormat::General => 0u16,
                    NumberFormat::BuiltIn(id) => *id as u16,
                    NumberFormat::Custom(s) => match format_index_for_custom.get(s) {
                        Some(&idx) => idx,
                        None => {
                            let idx = FORMAT_USER_INDEX_BASE + user_formats.len() as u16;
                            user_formats.push(s.clone());
                            format_index_for_custom.insert(s.clone(), idx);
                            idx
                        }
                    },
                };

                let xf_key: XfKey = (
                    font_idx,
                    format_idx,
                    style.alignment.clone(),
                    style.border.clone(),
                    style.fill.clone(),
                    style.protection,
                );
                let ixfe = match xf_key_to_ixfe.get(&xf_key) {
                    Some(&i) => i,
                    None => {
                        let new_ixfe = XF_USER_BASE + user_xfs.len() as u16;
                        user_xfs.push(UserXf {
                            font_index: font_idx,
                            format_index: format_idx,
                            alignment: style.alignment.clone(),
                            border: style.border.clone(),
                            fill: style.fill.clone(),
                            protection: style.protection,
                        });
                        xf_key_to_ixfe.insert(xf_key, new_ixfe);
                        new_ixfe
                    }
                };

                cell_ixfe.insert((sheet_idx, cell.style_index), ixfe);
            }
        }

        // Second pass: intern fonts for any rich-text run that has a
        // RunFont attached. Each run gets a complete FontStyle (with
        // unspecified RunFont fields filled from FontStyle::default)
        // so the SST emitter can look up its on-disk font index.
        for sheet in workbook.worksheets() {
            for (_row, _col, cell) in sheet.iter_cells() {
                if let CellValue::RichText(runs) = &cell.value {
                    for run in runs.iter() {
                        if let Some(rf) = &run.font {
                            let font = run_font_to_font_style(rf);
                            if !font_xf_index.contains_key(&font) {
                                let on_disk = fonts_in_order.len() as u16;
                                let xf_idx = if on_disk < 4 { on_disk } else { on_disk + 1 };
                                fonts_in_order.push(font.clone());
                                font_xf_index.insert(font, xf_idx);
                            }
                        }
                    }
                }
            }
        }

        StyleTables {
            fonts_in_order,
            font_xf_index,
            user_formats,
            user_xfs,
            cell_ixfe,
        }
    }

    fn ixfe_for_cell(&self, sheet_idx: usize, style_index: u32) -> u16 {
        if style_index == 0 {
            return 0;
        }
        self.cell_ixfe
            .get(&(sheet_idx, style_index))
            .copied()
            .unwrap_or(0)
    }

    fn write_font_records(&self, stream: &mut Vec<u8>) -> XlsResult<()> {
        for font in &self.fonts_in_order {
            write_font_record(stream, font)?;
        }
        Ok(())
    }

    fn write_format_records(&self, stream: &mut Vec<u8>) -> XlsResult<()> {
        for (i, fmt) in self.user_formats.iter().enumerate() {
            let ifmt = FORMAT_USER_INDEX_BASE + i as u16;
            write_format_record(stream, ifmt, fmt)?;
        }
        Ok(())
    }

    fn write_xf_records(&self, stream: &mut Vec<u8>) {
        for _ in 0..XF_USER_BASE {
            write_xf_record(stream, /* is_style_xf */ false, &XF_DEFAULTS);
        }
        for xf in &self.user_xfs {
            write_xf_record(stream, false, xf);
        }
    }
}

/// Materialise a [`RunFont`] (where each property is optional) into a
/// complete [`FontStyle`]. Unspecified `RunFont` properties fall back
/// to [`FontStyle::default()`] so the resulting struct can be emitted
/// as a complete BIFF8 FONT record.
fn run_font_to_font_style(rf: &duke_sheets_core::rich_text::RunFont) -> FontStyle {
    let mut font = FontStyle::default();
    if let Some(b) = rf.bold {
        font.bold = b;
    }
    if let Some(i) = rf.italic {
        font.italic = i;
    }
    if let Some(s) = rf.size {
        font.size = s;
    }
    if let Some(c) = rf.color.as_ref() {
        font.color = c.clone();
    }
    if let Some(name) = rf.name.as_ref() {
        font.name = name.clone();
    }
    if let Some(u) = rf.underline {
        font.underline = u;
    }
    if let Some(s) = rf.strikethrough {
        font.strikethrough = s;
    }
    if let Some(va) = rf.vertical_align {
        font.vertical_align = va;
    }
    if let Some(family) = rf.family {
        font.family = Some(family);
    }
    if let Some(charset) = rf.charset {
        font.charset = Some(charset);
    }
    if let Some(scheme) = rf.scheme.as_ref() {
        font.scheme = Some(scheme.clone());
    }
    font
}

const XF_DEFAULTS: UserXf = UserXf {
    font_index: 0,
    format_index: 0,
    alignment: Alignment {
        horizontal: HorizontalAlignment::General,
        vertical: VerticalAlignment::Bottom,
        wrap_text: false,
        shrink_to_fit: false,
        indent: 0,
        rotation: 0,
        reading_order: ReadingOrder::ContextDependent,
    },
    border: BorderStyle {
        left: None,
        right: None,
        top: None,
        bottom: None,
        diagonal: None,
        diagonal_direction: DiagonalDirection::None,
    },
    fill: FillStyle::None,
    protection: duke_sheets_core::style::Protection {
        locked: true,
        hidden: false,
    },
};

/// Emit a FORMAT record (MS-XLS §2.4.126) for a user-defined custom
/// number-format string. Built-in format indices (0..=49) are implicit
/// and don't need a FORMAT record.
fn write_format_record(stream: &mut Vec<u8>, ifmt: u16, format_string: &str) -> XlsResult<()> {
    let mut body = Vec::with_capacity(2 + 3 + format_string.len() * 2);
    body.extend_from_slice(&ifmt.to_le_bytes());
    push_xlunicode_string(&mut body, format_string)?;
    stream.extend_from_slice(&FORMAT_RECORD.to_le_bytes());
    stream.extend_from_slice(&(body.len() as u16).to_le_bytes());
    stream.extend_from_slice(&body);
    Ok(())
}

/// Emit a FONT record (MS-XLS §2.4.122). Only the fields needed by the
/// reader's `parse_font` are populated; family/charset are left at 0
/// (font-style-default). Color resolution: `Auto` → 0x7FFF, `Indexed`
/// passes through; RGB/Theme fall back to `Auto` because BIFF8 needs
/// a PALETTE record to express arbitrary RGB and that's deferred.
fn write_font_record(stream: &mut Vec<u8>, font: &FontStyle) -> XlsResult<()> {
    let mut body = Vec::with_capacity(16 + font.name.len() * 2);
    let height_twips = (font.size * 20.0).round() as u16;
    body.extend_from_slice(&height_twips.to_le_bytes());

    let mut grbit: u16 = 0;
    if font.italic {
        grbit |= 0x0002;
    }
    if font.strikethrough {
        grbit |= 0x0008;
    }
    body.extend_from_slice(&grbit.to_le_bytes());

    let icv = match font.color {
        Color::Auto => COLOR_AUTO,
        Color::Indexed(i) => i as u16,
        // BIFF8 has no first-class RGB / theme support without a
        // PALETTE record. Fall back to auto rather than emit an
        // out-of-range icv that the reader would reject.
        _ => COLOR_AUTO,
    };
    body.extend_from_slice(&icv.to_le_bytes());

    let bls: u16 = if font.bold { 700 } else { 400 };
    body.extend_from_slice(&bls.to_le_bytes());

    let sss: u16 = match font.vertical_align {
        duke_sheets_core::style::FontVerticalAlign::Superscript => 1,
        duke_sheets_core::style::FontVerticalAlign::Subscript => 2,
        _ => 0,
    };
    body.extend_from_slice(&sss.to_le_bytes());

    let uls: u8 = match font.underline {
        Underline::Single => 0x01,
        Underline::Double => 0x02,
        Underline::SingleAccounting => 0x21,
        Underline::DoubleAccounting => 0x22,
        Underline::None => 0x00,
    };
    body.push(uls);
    body.push(0); // bFamily
    body.push(0); // bCharSet
    body.push(0); // reserved

    push_short_xlunicode_string(&mut body, &font.name)?;

    stream.extend_from_slice(&FONT_RECORD.to_le_bytes());
    stream.extend_from_slice(&(body.len() as u16).to_le_bytes());
    stream.extend_from_slice(&body);
    Ok(())
}

/// Emit a 20-byte XF record (MS-XLS §2.4.353) with the supplied
/// font/format/alignment/border/fill axes encoded into the bit-packed
/// fields. `is_style_xf` flips the type/protect bit that tells the
/// reader whether this XF is a cell XF or a named-style XF.
fn write_xf_record(stream: &mut Vec<u8>, is_style_xf: bool, xf: &UserXf) {
    stream.extend_from_slice(&XF_RECORD.to_le_bytes());
    stream.extend_from_slice(&20u16.to_le_bytes());
    stream.extend_from_slice(&xf.font_index.to_le_bytes());
    stream.extend_from_slice(&xf.format_index.to_le_bytes());
    // type_prot bit layout (MS-XLS §2.4.353):
    //   bit 0 (0x0001): fLocked
    //   bit 1 (0x0002): fHidden (formula hidden when sheet is protected)
    //   bit 2 (0x0004): fStyle (1 = style XF, 0 = cell XF)
    //   bits 4-15: ifmt-XF parent (cell XF parent index, 0xFFF for style XFs)
    let type_prot: u16 = if is_style_xf {
        // 0xFFF5 = parent=0xFFF, fStyle=1, fHidden=0, fLocked=1 (default).
        0xFFF5
    } else {
        let mut bits: u16 = 0;
        if xf.protection.locked {
            bits |= 0x0001;
        }
        if xf.protection.hidden {
            bits |= 0x0002;
        }
        bits
    };
    stream.extend_from_slice(&type_prot.to_le_bytes());

    let halign = encode_horizontal_alignment(xf.alignment.horizontal);
    let valign = encode_vertical_alignment(xf.alignment.vertical);
    let align1: u8 =
        (halign & 0x07) | (if xf.alignment.wrap_text { 0x08 } else { 0 }) | ((valign & 0x07) << 4);
    stream.push(align1);
    stream.push(encode_rotation(xf.alignment.rotation));
    let reading_order = encode_reading_order(xf.alignment.reading_order);
    let align2: u8 = (xf.alignment.indent.min(15))
        | (if xf.alignment.shrink_to_fit { 0x10 } else { 0 })
        | ((reading_order & 0x03) << 6);
    stream.push(align2);
    stream.push(0); // used_attribs

    let (border_left, border_right, border_top, border_bottom) = (
        encode_border_line(xf.border.left.as_ref()),
        encode_border_line(xf.border.right.as_ref()),
        encode_border_line(xf.border.top.as_ref()),
        encode_border_line(xf.border.bottom.as_ref()),
    );
    let (icv_left, icv_right) = (
        encode_border_color(xf.border.left.as_ref()),
        encode_border_color(xf.border.right.as_ref()),
    );
    let diagonal_dir = encode_diagonal_direction(xf.border.diagonal_direction);
    let border1: u32 = (border_left as u32 & 0x0F)
        | ((border_right as u32 & 0x0F) << 4)
        | ((border_top as u32 & 0x0F) << 8)
        | ((border_bottom as u32 & 0x0F) << 12)
        | ((icv_left as u32 & 0x7F) << 16)
        | ((icv_right as u32 & 0x7F) << 23)
        | ((diagonal_dir as u32 & 0x03) << 30);
    stream.extend_from_slice(&border1.to_le_bytes());

    let icv_top = encode_border_color(xf.border.top.as_ref());
    let icv_bottom = encode_border_color(xf.border.bottom.as_ref());
    let icv_diag = encode_border_color(xf.border.diagonal.as_ref());
    let border_diag = encode_border_line(xf.border.diagonal.as_ref());
    let fill_pattern = encode_fill_pattern(&xf.fill);
    let border2: u32 = (icv_top as u32 & 0x7F)
        | ((icv_bottom as u32 & 0x7F) << 7)
        | ((icv_diag as u32 & 0x7F) << 14)
        | ((border_diag as u32 & 0x0F) << 21)
        | ((fill_pattern as u32 & 0x3F) << 26);
    stream.extend_from_slice(&border2.to_le_bytes());

    let (fill_fg, fill_bg) = encode_fill_colors(&xf.fill);
    let fill_colors: u16 = (fill_fg as u16 & 0x7F) | ((fill_bg as u16 & 0x7F) << 7);
    stream.extend_from_slice(&fill_colors.to_le_bytes());
}

fn encode_horizontal_alignment(h: HorizontalAlignment) -> u8 {
    match h {
        HorizontalAlignment::General => 0,
        HorizontalAlignment::Left => 1,
        HorizontalAlignment::Center => 2,
        HorizontalAlignment::Right => 3,
        HorizontalAlignment::Fill => 4,
        HorizontalAlignment::Justify => 5,
        HorizontalAlignment::CenterContinuous => 6,
        HorizontalAlignment::Distributed => 7,
    }
}

fn encode_vertical_alignment(v: VerticalAlignment) -> u8 {
    match v {
        VerticalAlignment::Top => 0,
        VerticalAlignment::Center => 1,
        VerticalAlignment::Bottom => 2,
        VerticalAlignment::Justify => 3,
        VerticalAlignment::Distributed => 4,
    }
}

fn encode_reading_order(r: ReadingOrder) -> u8 {
    match r {
        ReadingOrder::ContextDependent => 0,
        ReadingOrder::LeftToRight => 1,
        ReadingOrder::RightToLeft => 2,
    }
}

fn encode_rotation(rotation: i16) -> u8 {
    if rotation == 255 {
        return 255;
    }
    match rotation {
        0 => 0,
        1..=90 => rotation as u8,
        // The reader maps BIFF values 91..=180 back to negative
        // (anti-clockwise) angles. Inverse the relation here.
        -90..=-1 => (90i16 - rotation) as u8,
        _ => 0,
    }
}

fn encode_border_line(edge: Option<&BorderEdge>) -> u8 {
    let style = edge.map(|e| e.style).unwrap_or(BorderLineStyle::None);
    match style {
        BorderLineStyle::None => 0,
        BorderLineStyle::Thin => 1,
        BorderLineStyle::Medium => 2,
        BorderLineStyle::Dashed => 3,
        BorderLineStyle::Dotted => 4,
        BorderLineStyle::Thick => 5,
        BorderLineStyle::Double => 6,
        BorderLineStyle::Hair => 7,
        BorderLineStyle::MediumDashed => 8,
        BorderLineStyle::DashDot => 9,
        BorderLineStyle::MediumDashDot => 10,
        BorderLineStyle::DashDotDot => 11,
        BorderLineStyle::MediumDashDotDot => 12,
        BorderLineStyle::SlantDashDot => 13,
    }
}

fn encode_border_color(edge: Option<&BorderEdge>) -> u8 {
    edge.map(|e| color_to_icv7(&e.color)).unwrap_or(0x40)
}

fn encode_diagonal_direction(d: DiagonalDirection) -> u8 {
    match d {
        DiagonalDirection::None => 0,
        DiagonalDirection::Down => 1,
        DiagonalDirection::Up => 2,
        DiagonalDirection::Both => 3,
    }
}

fn encode_fill_pattern(fill: &FillStyle) -> u8 {
    match fill {
        FillStyle::None => 0,
        FillStyle::Solid { .. } => 1,
        FillStyle::Pattern { pattern, .. } => match pattern {
            PatternType::None => 0,
            PatternType::Solid => 1,
            PatternType::MediumGray => 2,
            PatternType::DarkGray => 3,
            PatternType::LightGray => 4,
            PatternType::DarkHorizontal => 5,
            PatternType::DarkVertical => 6,
            PatternType::DarkDown => 7,
            PatternType::DarkUp => 8,
            PatternType::DarkGrid => 9,
            PatternType::DarkTrellis => 10,
            PatternType::LightHorizontal => 11,
            PatternType::LightVertical => 12,
            PatternType::LightDown => 13,
            PatternType::LightUp => 14,
            PatternType::LightGrid => 15,
            PatternType::LightTrellis => 16,
            PatternType::Gray125 => 17,
            PatternType::Gray0625 => 18,
        },
        FillStyle::Gradient { .. } => 1, // fall back to solid for now
    }
}

fn encode_fill_colors(fill: &FillStyle) -> (u8, u8) {
    match fill {
        FillStyle::None => (0x40, 0x41),
        FillStyle::Solid { color } => (color_to_icv7(color), 0x41),
        FillStyle::Pattern {
            foreground,
            background,
            ..
        } => (color_to_icv7(foreground), color_to_icv7(background)),
        FillStyle::Gradient { stops, .. } => {
            let fg = stops
                .first()
                .map(|s| color_to_icv7(&s.color))
                .unwrap_or(0x40);
            (fg, 0x41)
        }
    }
}

/// Encode a [`Color`] as a 7-bit `icv` value (the on-disk encoding
/// used in border/fill XF fields). The font-side encoding uses 16
/// bits and a different sentinel — see `write_font_record`.
fn color_to_icv7(color: &Color) -> u8 {
    match color {
        Color::Auto => 0x40,
        Color::Indexed(i) => 0x08u8.saturating_add(i.min(&55).clone()),
        // BIFF8 has no first-class RGB without a PALETTE record. Pick
        // 0x40 (system foreground) as the closest no-op default; a
        // future PALETTE-emission slice can express arbitrary RGB.
        _ => 0x40,
    }
}

/// Emit a `ShortXLUnicodeString` (1-byte cch, 1-byte fHighByte, chars).
fn push_short_xlunicode_string(buf: &mut Vec<u8>, s: &str) -> XlsResult<()> {
    let units: Vec<u16> = s.encode_utf16().collect();
    if units.len() > u8::MAX as usize {
        return Err(XlsError::InvalidFormat(format!(
            "short string '{s}' exceeds 255-char ShortXLUnicodeString limit"
        )));
    }
    let high_byte = units.iter().any(|&u| u > 0xFF);
    buf.push(units.len() as u8);
    if high_byte {
        buf.push(0x01);
        for u in &units {
            buf.extend_from_slice(&u.to_le_bytes());
        }
    } else {
        buf.push(0x00);
        for u in &units {
            buf.push(*u as u8);
        }
    }
    Ok(())
}

/// Builder + emitter for the workbook-level Shared String Table (SST).
///
/// String cells in BIFF8 don't store their text inline; instead they
/// reference an entry in this workbook-global table via `LABELSST.isst`.
/// We dedupe identical strings on insert (matching what Excel writes)
/// and emit a single SST record + CONTINUE chain in the globals stream.
/// A single entry in the BIFF8 Shared String Table.
///
/// `Plain` is a flat string with no per-character formatting. `Rich`
/// stores the same text plus a list of formatting runs (each a
/// `(char_pos, font_idx)` pair) that the writer encodes as a BIFF8
/// rich-text SST entry with `fRichSt` set.
#[derive(Debug, Clone, PartialEq, Eq, Hash)]
enum SstEntryKind {
    Plain,
    Rich(Vec<(u16, u16)>),
}

#[derive(Debug, Clone)]
struct SstEntry {
    text: String,
    kind: SstEntryKind,
}

struct SstTable {
    entries: Vec<SstEntry>,
    /// Index keyed by `(text, kind)` so plain and rich variants of the
    /// same text are distinct entries (a plain "hello" and a bolded
    /// "hello" share neither LABELSST index nor SST slot).
    index: HashMap<(String, SstEntryKind), u32>,
    total_refs: u32,
}

impl SstTable {
    fn collect(workbook: &Workbook, styles: &StyleTables) -> Self {
        let mut t = SstTable {
            entries: Vec::new(),
            index: HashMap::new(),
            total_refs: 0,
        };
        for sheet in workbook.worksheets() {
            for (_row, _col, cell) in sheet.iter_cells() {
                match &cell.value {
                    CellValue::String(s) => {
                        t.add_plain(s.as_ref());
                    }
                    CellValue::RichText(runs) => {
                        t.add_rich(runs, styles);
                    }
                    _ => {}
                }
            }
        }
        t
    }

    fn add_plain(&mut self, s: &str) {
        self.total_refs += 1;
        let key = (s.to_string(), SstEntryKind::Plain);
        if !self.index.contains_key(&key) {
            let idx = self.entries.len() as u32;
            self.index.insert(key, idx);
            self.entries.push(SstEntry {
                text: s.to_string(),
                kind: SstEntryKind::Plain,
            });
        }
    }

    fn add_rich(
        &mut self,
        runs: &[duke_sheets_core::rich_text::RichTextRun],
        styles: &StyleTables,
    ) {
        self.total_refs += 1;
        let mut text = String::new();
        let mut formatting: Vec<(u16, u16)> = Vec::new();
        for run in runs {
            // char_pos is the UTF-16 code unit offset where this run
            // begins, which is what BIFF8 expects.
            let char_pos = text.encode_utf16().count().min(u16::MAX as usize) as u16;
            if let Some(rf) = &run.font {
                let font = run_font_to_font_style(rf);
                if let Some(&font_idx) = styles.font_xf_index.get(&font) {
                    formatting.push((char_pos, font_idx));
                }
            }
            text.push_str(&run.text);
        }

        let kind = if formatting.is_empty() {
            SstEntryKind::Plain
        } else {
            SstEntryKind::Rich(formatting)
        };

        let key = (text.clone(), kind.clone());
        if !self.index.contains_key(&key) {
            let idx = self.entries.len() as u32;
            self.index.insert(key, idx);
            self.entries.push(SstEntry { text, kind });
        }
    }

    /// Look up a plain-string SST index. Used by the LABELSST path
    /// for `CellValue::String` cells.
    fn lookup_plain(&self, s: &str) -> Option<u32> {
        self.index
            .get(&(s.to_string(), SstEntryKind::Plain))
            .copied()
    }

    /// Look up the SST index for a rich-text cell, given the text +
    /// formatting runs that were resolved for it.
    fn lookup_rich(&self, text: &str, formatting: &[(u16, u16)]) -> Option<u32> {
        let kind = if formatting.is_empty() {
            SstEntryKind::Plain
        } else {
            SstEntryKind::Rich(formatting.to_vec())
        };
        self.index.get(&(text.to_string(), kind)).copied()
    }

    /// Serialize the SST (and any required CONTINUE records) into the
    /// workbook stream. No-op if no entries were collected.
    fn write_records(&self, stream: &mut Vec<u8>) -> XlsResult<()> {
        if self.entries.is_empty() {
            return Ok(());
        }

        let mut payload = Vec::new();
        payload.extend_from_slice(&self.total_refs.to_le_bytes());
        payload.extend_from_slice(&(self.entries.len() as u32).to_le_bytes());
        for entry in &self.entries {
            push_sst_entry(&mut payload, entry)?;
        }

        // Split between entries (never mid-entry). For each entry,
        // compute its on-disk length and chunk into BIFF_MAX_RECORD_BODY
        // groups.
        let mut chunks: Vec<(usize, usize)> = Vec::new();
        let mut entry_offsets = Vec::with_capacity(self.entries.len());
        let mut cursor = 8usize; // skip cstTotal + cstUnique header
        let mut chunk_start = 0usize;
        let mut chunk_len = 8usize;
        for entry in &self.entries {
            let entry_len = sst_entry_len(entry);
            if chunk_len + entry_len > BIFF_MAX_RECORD_BODY {
                chunks.push((chunk_start, chunk_len));
                chunk_start = cursor;
                chunk_len = 0;
            }
            entry_offsets.push(cursor);
            chunk_len += entry_len;
            cursor += entry_len;
        }
        if chunk_len > 0 {
            chunks.push((chunk_start, chunk_len));
        }

        let mut chunk_record_offsets = Vec::with_capacity(chunks.len());
        for (i, (start, len)) in chunks.iter().enumerate() {
            let record_type = if i == 0 { SST_RECORD } else { CONTINUE_RECORD };
            chunk_record_offsets.push(stream.len());
            stream.extend_from_slice(&record_type.to_le_bytes());
            stream.extend_from_slice(&(*len as u16).to_le_bytes());
            stream.extend_from_slice(&payload[*start..*start + *len]);
        }
        write_extsst_record(stream, &entry_offsets, &chunks, &chunk_record_offsets);
        Ok(())
    }
}

fn write_extsst_record(
    stream: &mut Vec<u8>,
    entry_offsets: &[usize],
    chunks: &[(usize, usize)],
    chunk_record_offsets: &[usize],
) {
    const STRINGS_PER_BUCKET: usize = 8;

    if entry_offsets.is_empty() {
        return;
    }

    let bucket_count = entry_offsets.len().div_ceil(STRINGS_PER_BUCKET);
    let body_len = 2 + bucket_count * 8;
    if body_len > BIFF_MAX_RECORD_BODY {
        return;
    }

    let mut body = Vec::with_capacity(body_len);
    body.extend_from_slice(&(STRINGS_PER_BUCKET as u16).to_le_bytes());
    for entry_offset in entry_offsets.iter().step_by(STRINGS_PER_BUCKET) {
        if let Some((chunk_idx, (chunk_start, _chunk_len))) = chunks
            .iter()
            .enumerate()
            .find(|(_, (start, len))| *entry_offset >= *start && *entry_offset < *start + *len)
        {
            let record_relative = entry_offset - chunk_start;
            let absolute = chunk_record_offsets[chunk_idx] + 4 + record_relative;
            let relative = 4 + record_relative;
            body.extend_from_slice(&(absolute as u32).to_le_bytes());
            body.extend_from_slice(&(relative as u16).to_le_bytes());
            body.extend_from_slice(&0u16.to_le_bytes());
        } else {
            break;
        }
    }
    write_biff_record(stream, EXTSST_RECORD, &body);
}

/// Length of an `XLUnicodeRichExtendedString`-formatted plaintext entry
/// (cch + flags + chars) without rich/ext fields.
fn xlunicode_string_len(s: &str) -> usize {
    let units: Vec<u16> = s.encode_utf16().collect();
    let high_byte = units.iter().any(|&u| u > 0xFF);
    let chars_len = if high_byte {
        units.len() * 2
    } else {
        units.len()
    };
    2 + 1 + chars_len
}

/// Length of an SST entry on disk: plain entries are
/// cch + flags + chars; rich entries add a 2-byte cRun count plus
/// 4 bytes per formatting run inserted between the flags byte and
/// the character data.
fn sst_entry_len(entry: &SstEntry) -> usize {
    let plain = xlunicode_string_len(&entry.text);
    match &entry.kind {
        SstEntryKind::Plain => plain,
        SstEntryKind::Rich(runs) => plain + 2 + runs.len() * 4,
    }
}

fn push_sst_entry(buf: &mut Vec<u8>, entry: &SstEntry) -> XlsResult<()> {
    match &entry.kind {
        SstEntryKind::Plain => push_xlunicode_string(buf, &entry.text),
        SstEntryKind::Rich(runs) => push_rich_xlunicode_string(buf, &entry.text, runs),
    }
}

/// Emit a rich `XLUnicodeRichExtendedString`: cch + flags (with
/// fRichSt bit set) + cRun + chars + run array. Each run is
/// (char_pos u16, font_idx u16) = 4 bytes.
fn push_rich_xlunicode_string(buf: &mut Vec<u8>, s: &str, runs: &[(u16, u16)]) -> XlsResult<()> {
    let units: Vec<u16> = s.encode_utf16().collect();
    if units.len() > u16::MAX as usize {
        return Err(XlsError::InvalidFormat(format!(
            "string of {} UTF-16 units exceeds BIFF8 cch limit (u16)",
            units.len()
        )));
    }
    if runs.len() > u16::MAX as usize {
        return Err(XlsError::InvalidFormat(format!(
            "rich-text run count {} exceeds BIFF8 limit (u16)",
            runs.len()
        )));
    }
    let high_byte = units.iter().any(|&u| u > 0xFF);
    buf.extend_from_slice(&(units.len() as u16).to_le_bytes());
    let flags: u8 = 0x08 | if high_byte { 0x01 } else { 0x00 };
    buf.push(flags);
    buf.extend_from_slice(&(runs.len() as u16).to_le_bytes());
    if high_byte {
        for u in &units {
            buf.extend_from_slice(&u.to_le_bytes());
        }
    } else {
        for u in &units {
            buf.push(*u as u8);
        }
    }
    for (char_pos, font_idx) in runs {
        buf.extend_from_slice(&char_pos.to_le_bytes());
        buf.extend_from_slice(&font_idx.to_le_bytes());
    }
    Ok(())
}

/// Emit an `XLUnicodeRichExtendedString` (no rich runs, no ExtRst).
/// Picks the compact Latin-1 encoding when every code unit fits in a
/// byte; otherwise emits UTF-16LE.
fn push_xlunicode_string(buf: &mut Vec<u8>, s: &str) -> XlsResult<()> {
    let units: Vec<u16> = s.encode_utf16().collect();
    let len = xlunicode_units_len_u16(&units)?;
    let high_byte = units.iter().any(|&u| u > 0xFF);
    buf.extend_from_slice(&len.to_le_bytes());
    push_xlunicode_units_no_cch(buf, &units, high_byte);
    Ok(())
}

fn push_xlunicode_string_no_cch(buf: &mut Vec<u8>, s: &str) -> XlsResult<()> {
    let units: Vec<u16> = s.encode_utf16().collect();
    xlunicode_units_len_u16(&units)?;
    let high_byte = units.iter().any(|&u| u > 0xFF);
    push_xlunicode_units_no_cch(buf, &units, high_byte);
    Ok(())
}

fn xlunicode_len_u16(s: &str) -> XlsResult<u16> {
    let units = s.encode_utf16().collect::<Vec<_>>();
    xlunicode_units_len_u16(&units)
}

fn xlunicode_units_len_u16(units: &[u16]) -> XlsResult<u16> {
    if units.len() > u16::MAX as usize {
        return Err(XlsError::InvalidFormat(format!(
            "string of {} UTF-16 units exceeds BIFF8 cch limit (u16)",
            units.len()
        )));
    }
    Ok(units.len() as u16)
}

fn push_xlunicode_units_no_cch(buf: &mut Vec<u8>, units: &[u16], high_byte: bool) {
    if high_byte {
        buf.push(0x01); // fHighByte = 1
        for u in units {
            buf.extend_from_slice(&u.to_le_bytes());
        }
    } else {
        buf.push(0x00); // fHighByte = 0 (Latin-1)
        for u in units {
            buf.push(*u as u8);
        }
    }
}

/// Emit a DIMENSION record (MS-XLS §2.4.62) bounding the populated
/// cell rectangle. LibreOffice and real Excel scan the cell records in
/// the half-open `[firstRow..lastRow), [firstCol..lastCol)` range
/// declared here; without DIMENSION they treat the worksheet as empty
/// even when cell records are physically present.
fn write_dimension(stream: &mut Vec<u8>, sheet: &Worksheet) {
    let mut first_row = u32::MAX;
    let mut last_row_excl = 0u32;
    let mut first_col = u16::MAX;
    let mut last_col_excl = 0u16;
    for (row, col, _) in sheet.iter_cells() {
        if row < first_row {
            first_row = row;
        }
        if row + 1 > last_row_excl {
            last_row_excl = row + 1;
        }
        if col < first_col {
            first_col = col;
        }
        if col + 1 > last_col_excl {
            last_col_excl = col + 1;
        }
    }
    if last_row_excl == 0 {
        first_row = 0;
        first_col = 0;
    }

    stream.extend_from_slice(&DIMENSION_RECORD.to_le_bytes());
    stream.extend_from_slice(&14u16.to_le_bytes());
    stream.extend_from_slice(&first_row.to_le_bytes());
    stream.extend_from_slice(&last_row_excl.to_le_bytes());
    stream.extend_from_slice(&(first_col as u16).to_le_bytes());
    stream.extend_from_slice(&(last_col_excl as u16).to_le_bytes());
    stream.extend_from_slice(&0u16.to_le_bytes());
}

/// Emit a WINDOW1 record (MS-XLS §2.4.346). Bytes 10-11 carry
/// `itabCur` (active sheet index); the reader uses that to set the
/// workbook's active sheet. The other fields use sensible defaults.
fn write_window1(stream: &mut Vec<u8>, workbook: &Workbook) {
    let active = workbook.active_sheet().min(u16::MAX as usize) as u16;
    let visible_count = workbook
        .worksheets()
        .filter(|s| {
            !matches!(
                s.visibility(),
                duke_sheets_core::worksheet::SheetVisibility::Hidden
                    | duke_sheets_core::worksheet::SheetVisibility::VeryHidden
            )
        })
        .count()
        .min(u16::MAX as usize) as u16;

    stream.extend_from_slice(&WINDOW1_RECORD.to_le_bytes());
    stream.extend_from_slice(&18u16.to_le_bytes());
    stream.extend_from_slice(&0i16.to_le_bytes()); // xWn
    stream.extend_from_slice(&0i16.to_le_bytes()); // yWn
    stream.extend_from_slice(&15000u16.to_le_bytes()); // dxWn
    stream.extend_from_slice(&9000u16.to_le_bytes()); // dyWn
    stream.extend_from_slice(&0x0038u16.to_le_bytes()); // grbit: tab visible, scroll bars visible
    stream.extend_from_slice(&active.to_le_bytes()); // itabCur
    stream.extend_from_slice(&0u16.to_le_bytes()); // itabFirst
    stream.extend_from_slice(&visible_count.max(1).to_le_bytes()); // ctabSel
    stream.extend_from_slice(&600u16.to_le_bytes()); // wTabRatio (600 = 60%)
}

fn write_date_mode(stream: &mut Vec<u8>, workbook: &Workbook) {
    let mode = if workbook.settings().date_1904 {
        1u16
    } else {
        0u16
    };
    write_biff_record(stream, DATEMODE_RECORD, &mode.to_le_bytes());
}

/// Emit a minimal WINDOW2 record (MS-XLS §2.4.349). The reader extracts
/// only the frozen-pane bit, but real Excel and LibreOffice expect this
/// record as a structural cue that the worksheet stream is well-formed.
/// Sets fFrozen (bit 3) and fFrozenNoSplit (bit 8) when the sheet has
/// freeze panes set; the matching PANE record carries the split
/// position.
fn write_window2(stream: &mut Vec<u8>, sheet: &Worksheet, has_pivot_tables: bool) {
    let mut options: u16 = 0x06B6;
    if sheet.freeze_panes().is_some() {
        options |= 0x0008 | 0x0100;
    }
    let row_pos: u16 = 0;
    let col_pos: u16 = 0;
    let grid_color: u32 = if has_pivot_tables { 0x40 } else { 0 };
    let preview_zoom: u16 = 0;
    let normal_zoom: u16 = 0;
    let reserved: u32 = if has_pivot_tables { 0x11 } else { 0 };

    stream.extend_from_slice(&WINDOW2_RECORD.to_le_bytes());
    stream.extend_from_slice(&18u16.to_le_bytes());
    stream.extend_from_slice(&options.to_le_bytes());
    stream.extend_from_slice(&row_pos.to_le_bytes());
    stream.extend_from_slice(&col_pos.to_le_bytes());
    stream.extend_from_slice(&grid_color.to_le_bytes());
    stream.extend_from_slice(&preview_zoom.to_le_bytes());
    stream.extend_from_slice(&normal_zoom.to_le_bytes());
    stream.extend_from_slice(&reserved.to_le_bytes());
}

/// Emit a PANE record (MS-XLS §2.4.187) when the worksheet has freeze
/// panes set. Body: x (col split, u16), y (row split, u16), top_row
/// (u16), left_col (u16), active_pane (u16, 0=bottomRight). The
/// matching WINDOW2 sets fFrozen + fFrozenNoSplit so the reader knows
/// the split values are row/col indices rather than twip offsets.
fn write_pane(stream: &mut Vec<u8>, sheet: &Worksheet) {
    let Some(freeze) = sheet.freeze_panes() else {
        return;
    };
    if freeze.row == 0 && freeze.col == 0 {
        return;
    }
    let x = freeze.col;
    let y = freeze.row.min(u16::MAX as u32) as u16;
    let top_row: u16 = y;
    let left_col: u16 = x;
    // Active pane is the bottom-right when both axes are frozen,
    // else the corresponding half. Excel uses 0=bottomRight,
    // 1=topRight, 2=bottomLeft, 3=topLeft.
    let active_pane: u16 = match (x > 0, y > 0) {
        (true, true) => 0,
        (true, false) => 0,
        (false, true) => 2,
        (false, false) => 3,
    };

    stream.extend_from_slice(&PANE_RECORD.to_le_bytes());
    stream.extend_from_slice(&10u16.to_le_bytes());
    stream.extend_from_slice(&x.to_le_bytes());
    stream.extend_from_slice(&y.to_le_bytes());
    stream.extend_from_slice(&top_row.to_le_bytes());
    stream.extend_from_slice(&left_col.to_le_bytes());
    stream.extend_from_slice(&active_pane.to_le_bytes());
}

/// Emit an SCL record (MS-XLS §2.4.249) carrying the worksheet zoom
/// level as a numerator/denominator ratio. The model stores zoom as a
/// percentage (10..=400); reduce against 100 for compactness.
fn write_scl(stream: &mut Vec<u8>, sheet: &Worksheet) {
    let Some(zoom) = sheet.zoom_scale() else {
        return;
    };
    let zoom = zoom.clamp(10, 400);
    let (num, den) = if zoom % 25 == 0 && zoom <= 200 {
        (zoom / 25, 4u16)
    } else {
        (zoom, 100u16)
    };
    stream.extend_from_slice(&SCL_RECORD.to_le_bytes());
    stream.extend_from_slice(&4u16.to_le_bytes());
    stream.extend_from_slice(&num.to_le_bytes());
    stream.extend_from_slice(&den.to_le_bytes());
}

/// Emit one SELECTION record (MS-XLS §2.4.247) per Selection in the
/// worksheet's selection list. Body: pane (u8), active_row (u16),
/// active_col (u16), active_ref (u16), ref_count (u16), then
/// `ref_count` × Ref8U (r1 u16, r2 u16, c1 u8, c2 u8 = 6 bytes).
/// Selections targeting cells beyond the BIFF8 sheet extent (col >
/// 255 or row > 65535) are skipped because the on-disk Ref8U format
/// can't represent them.
fn write_selection_records(stream: &mut Vec<u8>, sheet: &Worksheet) {
    use duke_sheets_core::{CellAddress, CellRange};

    for selection in sheet.selections() {
        let pane_byte: u8 = match selection.pane.as_deref() {
            Some("topRight") => 1,
            Some("bottomLeft") => 2,
            Some("topLeft") => 3,
            _ => 0, // bottomRight or unspecified
        };

        let (active_row, active_col) = selection
            .active_cell
            .as_deref()
            .and_then(|s| CellAddress::parse(s).ok())
            .map(|a| (a.row, a.col))
            .unwrap_or((0u32, 0u16));
        if active_row > u16::MAX as u32 || active_col > u8::MAX as u16 {
            continue;
        }

        let ranges: Vec<CellRange> = selection
            .sqref
            .as_deref()
            .map(|s| {
                s.split_whitespace()
                    .filter_map(|piece| {
                        CellRange::parse(piece).ok().or_else(|| {
                            CellAddress::parse(piece).ok().map(|a| CellRange::new(a, a))
                        })
                    })
                    .collect()
            })
            .unwrap_or_default();

        let usable: Vec<&CellRange> = ranges
            .iter()
            .filter(|r| {
                r.start.row <= u16::MAX as u32
                    && r.end.row <= u16::MAX as u32
                    && r.start.col <= u8::MAX as u16
                    && r.end.col <= u8::MAX as u16
            })
            .collect();

        let ref_count = usable.len() as u16;
        let body_len = 9u16 + ref_count * 6;
        stream.extend_from_slice(&SELECTION_RECORD.to_le_bytes());
        stream.extend_from_slice(&body_len.to_le_bytes());
        stream.push(pane_byte);
        stream.extend_from_slice(&(active_row as u16).to_le_bytes());
        stream.extend_from_slice(&active_col.to_le_bytes());
        stream.extend_from_slice(&0u16.to_le_bytes()); // active_ref index
        stream.extend_from_slice(&ref_count.to_le_bytes());
        for r in usable {
            stream.extend_from_slice(&(r.start.row as u16).to_le_bytes());
            stream.extend_from_slice(&(r.end.row as u16).to_le_bytes());
            stream.push(r.start.col as u8);
            stream.push(r.end.col as u8);
        }
    }
}

/// Emit PROTECT (MS-XLS §2.4.196) and optional PASSWORD (MS-XLS
/// §2.4.190) records when the worksheet has protection set. PROTECT
/// is a single u16 = 1; PASSWORD is the precomputed 16-bit verifier
/// (zero means "no password required, but sheet is protected").
fn write_sheet_calculation_records(stream: &mut Vec<u8>) {
    write_biff_record(stream, CALCMODE_RECORD, &1u16.to_le_bytes());
    write_biff_record(stream, CALCCOUNT_RECORD, &100u16.to_le_bytes());
    write_biff_record(stream, REFMODE_RECORD, &1u16.to_le_bytes());
    write_biff_record(stream, ITERATION_RECORD, &0u16.to_le_bytes());
    write_biff_record(stream, DELTA_RECORD, &0.001f64.to_le_bytes());
    write_biff_record(stream, SAVERECALC_RECORD, &1u16.to_le_bytes());
}

fn write_sheet_display_default_records(stream: &mut Vec<u8>, sheet: &Worksheet) {
    let ps = sheet.page_setup();
    if !ps.print_headings {
        write_biff_record(stream, PRINTHEADERS_RECORD, &0u16.to_le_bytes());
    }
    if !ps.print_gridlines {
        write_biff_record(stream, PRINTGRIDLINES_RECORD, &0u16.to_le_bytes());
    }
    write_biff_record(stream, GRIDSET_RECORD, &1u16.to_le_bytes());
    write_biff_record(stream, GUTS_RECORD, &[0; 8]);

    let height_twips = ((sheet.default_row_height() * 20.0).round() as u32).min(0x7FFF) as u16;
    let mut row_height = Vec::with_capacity(4);
    row_height.extend_from_slice(&0u16.to_le_bytes());
    row_height.extend_from_slice(&height_twips.to_le_bytes());
    write_biff_record(stream, DEFAULTROWHEIGHT_RECORD, &row_height);

    write_biff_record(stream, WSBOOL_RECORD, &[0xC1, 0x04]);
}

fn write_default_column_width_record(stream: &mut Vec<u8>) {
    write_biff_record(stream, DEFCOLWIDTH_RECORD, &8u16.to_le_bytes());
}

fn write_protect_records(stream: &mut Vec<u8>, sheet: &Worksheet) {
    let Some(protection) = sheet.protection() else {
        return;
    };
    if !protection.protected {
        return;
    }
    stream.extend_from_slice(&PROTECT_RECORD.to_le_bytes());
    stream.extend_from_slice(&2u16.to_le_bytes());
    stream.extend_from_slice(&1u16.to_le_bytes());

    let password_hash = protection.password_hash.unwrap_or(0);
    stream.extend_from_slice(&PASSWORD_RECORD.to_le_bytes());
    stream.extend_from_slice(&2u16.to_le_bytes());
    stream.extend_from_slice(&password_hash.to_le_bytes());
}

fn write_index_placeholder(stream: &mut Vec<u8>, has_dbcell: bool) -> usize {
    let record_pos = stream.len();
    let payload_len = if has_dbcell { 20u16 } else { 16u16 };
    stream.extend_from_slice(&INDEX_RECORD.to_le_bytes());
    stream.extend_from_slice(&payload_len.to_le_bytes());
    stream.resize(stream.len() + payload_len as usize, 0);
    record_pos
}

fn patch_index_record(
    stream: &mut [u8],
    record_pos: usize,
    sheet: &Worksheet,
    sheet_idx: usize,
    styles: &StyleTables,
    first_row_record_pos: Option<usize>,
    dbcell_pos: Option<usize>,
) {
    let payload_pos = record_pos + 4;
    if payload_pos + 16 > stream.len() {
        return;
    }

    let (first_row, last_row_exclusive) =
        biff_cell_row_bounds(sheet, sheet_idx, styles).unwrap_or((0, 0));
    let first_row_record_pos = first_row_record_pos.unwrap_or(0);

    stream[payload_pos..payload_pos + 4].copy_from_slice(&0u32.to_le_bytes());
    stream[payload_pos + 4..payload_pos + 8].copy_from_slice(&first_row.to_le_bytes());
    stream[payload_pos + 8..payload_pos + 12].copy_from_slice(&last_row_exclusive.to_le_bytes());
    stream[payload_pos + 12..payload_pos + 16]
        .copy_from_slice(&(first_row_record_pos as u32).to_le_bytes());

    if let Some(dbcell_pos) = dbcell_pos {
        if payload_pos + 20 <= stream.len() {
            stream[payload_pos + 16..payload_pos + 20]
                .copy_from_slice(&(dbcell_pos as u32).to_le_bytes());
        }
    }
}

fn biff_cell_row_bounds(
    sheet: &Worksheet,
    sheet_idx: usize,
    styles: &StyleTables,
) -> Option<(u32, u32)> {
    let mut first_row = u32::MAX;
    let mut last_row_exclusive = 0u32;
    for (row, _col, data) in sheet.iter_cells() {
        if row > u16::MAX as u32 || !cell_will_emit_biff_record(sheet_idx, data, styles) {
            continue;
        }
        first_row = first_row.min(row);
        last_row_exclusive = last_row_exclusive.max(row.saturating_add(1));
    }
    (last_row_exclusive != 0).then_some((first_row, last_row_exclusive))
}

fn sheet_has_biff_cell_records(sheet: &Worksheet, sheet_idx: usize, styles: &StyleTables) -> bool {
    biff_cell_row_bounds(sheet, sheet_idx, styles).is_some()
}

fn cell_will_emit_biff_record(
    sheet_idx: usize,
    data: &duke_sheets_core::cell::CellData,
    styles: &StyleTables,
) -> bool {
    !matches!(&data.value, CellValue::Empty)
        || styles.ixfe_for_cell(sheet_idx, data.style_index) != 0
}

/// Emit a ROW record (MS-XLS §2.4.220) per row that has cells or any
/// non-default row property. Bit layout in the 32-bit options field at
/// body offset 12 matches the reader's parse_row: 0x10 collapsed, 0x20
/// hidden, 0x40 fUnsynced (custom height), bits 8-10 outline level.
/// Rows beyond u16::MAX are silently skipped because BIFF8 can't
/// address them.
fn write_row_records(stream: &mut Vec<u8>, sheet: &Worksheet) -> BTreeMap<u32, usize> {
    let heights = sheet.custom_row_heights();
    let hidden = sheet.hidden_rows();
    let outlines = sheet.row_outline_levels();
    let collapsed = sheet.collapsed_rows();
    let mut row_record_positions = BTreeMap::new();

    let mut cell_spans: BTreeMap<u32, (u16, u16)> = BTreeMap::new();
    for (row, col, _cell) in sheet.iter_cells() {
        if row > u16::MAX as u32 {
            continue;
        }
        cell_spans
            .entry(row)
            .and_modify(|(first_col, last_col)| {
                *first_col = (*first_col).min(col);
                *last_col = (*last_col).max(col);
            })
            .or_insert((col, col));
    }

    let mut rows: BTreeSet<u32> = BTreeSet::new();
    rows.extend(cell_spans.keys().copied());
    rows.extend(heights.keys());
    rows.extend(hidden.keys());
    rows.extend(outlines.keys());
    rows.extend(collapsed.keys());

    for row in rows {
        if row > u16::MAX as u32 {
            continue;
        }
        let height_pt = heights.get(&row).copied();
        let is_hidden = hidden.get(&row).copied().unwrap_or(false);
        let outline_level = outlines.get(&row).copied().unwrap_or(0);
        let is_collapsed = collapsed.get(&row).copied().unwrap_or(false);
        let (first_col, last_col) = cell_spans.get(&row).copied().unwrap_or((0, 0));
        let last_col_plus_one = if cell_spans.contains_key(&row) {
            last_col.saturating_add(1)
        } else {
            0
        };

        let mut body = Vec::with_capacity(16);
        body.extend_from_slice(&(row as u16).to_le_bytes());
        body.extend_from_slice(&first_col.to_le_bytes());
        body.extend_from_slice(&last_col_plus_one.to_le_bytes());
        let height_twips = match height_pt {
            Some(h) if h > 0.0 => ((h * 20.0).round() as u32).min(0x7FFF) as u16,
            _ => ((sheet.default_row_height() * 20.0).round() as u32).min(0x7FFF) as u16,
        };
        body.extend_from_slice(&height_twips.to_le_bytes());
        body.extend_from_slice(&0u16.to_le_bytes()); // reserved1
        body.extend_from_slice(&0u16.to_le_bytes()); // unused1
        let mut options: u32 = 0;
        if is_collapsed {
            options |= 0x10;
        }
        if is_hidden {
            options |= 0x20;
        }
        if height_pt.is_some() {
            options |= 0x40; // fUnsynced (custom height)
        }
        if outline_level > 0 {
            options |= ((outline_level as u32) & 0x07) << 8;
        }
        body.extend_from_slice(&options.to_le_bytes());

        row_record_positions.insert(row, stream.len());
        stream.extend_from_slice(&ROW_RECORD.to_le_bytes());
        stream.extend_from_slice(&(body.len() as u16).to_le_bytes());
        stream.extend_from_slice(&body);
    }
    row_record_positions
}

fn write_dbcell_record(
    stream: &mut Vec<u8>,
    row_record_positions: &BTreeMap<u32, usize>,
    first_cell_positions: &BTreeMap<u32, usize>,
) -> Option<usize> {
    if first_cell_positions.is_empty() {
        return None;
    }

    let Some((&first_row, &first_cell_pos)) = first_cell_positions.iter().next() else {
        return None;
    };
    let Some(&first_row_record_pos) = row_record_positions.get(&first_row) else {
        return None;
    };

    let dbcell_pos = stream.len();
    let Ok(db_rtrw) = u32::try_from(dbcell_pos.saturating_sub(first_row_record_pos)) else {
        return None;
    };
    let first_cell_offset = first_cell_pos.saturating_sub(first_row_record_pos + 20);
    let Ok(first_cell_offset) = u16::try_from(first_cell_offset) else {
        return None;
    };

    let mut payload = Vec::new();
    payload.extend_from_slice(&db_rtrw.to_le_bytes());
    payload.extend_from_slice(&first_cell_offset.to_le_bytes());

    let mut previous_first_cell_pos = first_cell_pos;
    for (_row, cell_pos) in first_cell_positions.iter().skip(1) {
        let delta = cell_pos.saturating_sub(previous_first_cell_pos);
        let Ok(delta) = u16::try_from(delta) else {
            return None;
        };
        payload.extend_from_slice(&delta.to_le_bytes());
        previous_first_cell_pos = *cell_pos;
    }

    write_biff_record(stream, DBCELL_RECORD, &payload);
    Some(dbcell_pos)
}

/// Emit a COLINFO record (MS-XLS §2.4.49) per column with any non-
/// default property: explicit width, hidden flag, outline level, or
/// collapsed state. Each emitted record covers a single column for
/// simplicity; the reader merges adjacent COLINFO ranges fine.
/// Width is converted from the model's "characters" unit to BIFF8's
/// 1/256-of-default-char-width unit.
fn write_colinfo_records(stream: &mut Vec<u8>, sheet: &Worksheet) {
    let widths = sheet.custom_column_widths();
    let hidden = sheet.hidden_columns();
    let outlines = sheet.column_outline_levels();
    let collapsed = sheet.collapsed_columns();

    let mut cols: std::collections::BTreeSet<u16> = std::collections::BTreeSet::new();
    cols.extend(widths.keys());
    cols.extend(hidden.keys());
    cols.extend(outlines.keys());
    cols.extend(collapsed.keys());

    for col in cols {
        let width_chars = widths.get(&col).copied();
        let is_hidden = hidden.get(&col).copied().unwrap_or(false);
        let outline_level = outlines.get(&col).copied().unwrap_or(0);
        let is_collapsed = collapsed.get(&col).copied().unwrap_or(false);

        let mut body = Vec::with_capacity(12);
        body.extend_from_slice(&col.to_le_bytes());
        body.extend_from_slice(&col.to_le_bytes()); // last == first for one-col record
        let coldx = match width_chars {
            Some(w) if w > 0.0 => ((w * 256.0).round() as u32).min(u16::MAX as u32) as u16,
            _ => 0,
        };
        body.extend_from_slice(&coldx.to_le_bytes());
        body.extend_from_slice(&15u16.to_le_bytes()); // ixfe (default cell XF)
        let mut options: u16 = 0;
        if is_hidden {
            options |= 0x0001;
        }
        if outline_level > 0 {
            options |= ((outline_level as u16) & 0x0007) << 8;
        }
        if is_collapsed {
            options |= 0x1000;
        }
        body.extend_from_slice(&options.to_le_bytes());
        body.extend_from_slice(&0u16.to_le_bytes()); // reserved

        stream.extend_from_slice(&COLINFO_RECORD.to_le_bytes());
        stream.extend_from_slice(&(body.len() as u16).to_le_bytes());
        stream.extend_from_slice(&body);
    }
}

/// Emit a SETUP record (MS-XLS §2.4.252) carrying paper size, scale,
/// fit-to-width/height, orientation, and header/footer margins.
/// Body layout: iPaperSize, iScale, iPageStart, iFitWidth, iFitHeight,
/// grbit, iRes, iVRes, numHdr (f64), numFtr (f64), iCopies. The
/// orientation lives in grbit bit 1 (set = landscape); bit 6 (fNoOrient)
/// is cleared so the orientation is honoured.
fn write_setup_record(stream: &mut Vec<u8>, sheet: &Worksheet) {
    let ps = sheet.page_setup();
    let mut body = Vec::with_capacity(34);
    body.extend_from_slice(&(ps.paper_size as u16).to_le_bytes());
    body.extend_from_slice(&ps.scale.clamp(10, 400).to_le_bytes());
    body.extend_from_slice(&1u16.to_le_bytes()); // iPageStart
    body.extend_from_slice(&ps.fit_to_width.unwrap_or(0).to_le_bytes());
    body.extend_from_slice(&ps.fit_to_height.unwrap_or(0).to_le_bytes());

    let mut grbit: u16 = 0;
    if matches!(ps.orientation, duke_sheets_core::PageOrientation::Landscape) {
        grbit |= 0x0002;
    }
    body.extend_from_slice(&grbit.to_le_bytes());
    body.extend_from_slice(&600u16.to_le_bytes()); // iRes (DPI)
    body.extend_from_slice(&600u16.to_le_bytes()); // iVRes (DPI)
    body.extend_from_slice(&ps.header_margin.to_le_bytes());
    body.extend_from_slice(&ps.footer_margin.to_le_bytes());
    body.extend_from_slice(&1u16.to_le_bytes()); // iCopies

    stream.extend_from_slice(&SETUP_RECORD.to_le_bytes());
    stream.extend_from_slice(&(body.len() as u16).to_le_bytes());
    stream.extend_from_slice(&body);
}

/// Emit HEADER (MS-XLS §2.4.137) and FOOTER (MS-XLS §2.4.111) records
/// when the page setup has odd-page header/footer text. Both wrap the
/// text in a XLUnicodeString (cch + flags + chars). Even/first page
/// headers and footers are not yet emitted; BIFF8's HEADERFOOTER ext
/// record carries those.
fn write_header_footer_records(stream: &mut Vec<u8>, sheet: &Worksheet) {
    let ps = sheet.page_setup();
    match ps.odd_header.as_deref() {
        Some(text) => emit_unicode_string_record(stream, HEADER_RECORD, text),
        None => write_biff_record(stream, HEADER_RECORD, &[]),
    }
    match ps.odd_footer.as_deref() {
        Some(text) => emit_unicode_string_record(stream, FOOTER_RECORD, text),
        None => write_biff_record(stream, FOOTER_RECORD, &[]),
    }
    write_biff_record(stream, HCENTER_RECORD, &0u16.to_le_bytes());
    write_biff_record(stream, VCENTER_RECORD, &0u16.to_le_bytes());
}

fn emit_unicode_string_record(stream: &mut Vec<u8>, record_type: u16, text: &str) {
    let units: Vec<u16> = text.encode_utf16().collect();
    let high_byte = units.iter().any(|&u| u > 0xFF);
    let mut body = Vec::with_capacity(3 + units.len() * 2);
    body.extend_from_slice(&(units.len() as u16).to_le_bytes());
    if high_byte {
        body.push(0x01);
        for u in &units {
            body.extend_from_slice(&u.to_le_bytes());
        }
    } else {
        body.push(0x00);
        for u in &units {
            body.push(*u as u8);
        }
    }
    stream.extend_from_slice(&record_type.to_le_bytes());
    stream.extend_from_slice(&(body.len() as u16).to_le_bytes());
    stream.extend_from_slice(&body);
}

/// Emit LEFT_MARGIN, RIGHT_MARGIN, TOP_MARGIN, BOTTOM_MARGIN records.
/// Each is an 8-byte f64 carrying the margin in inches. The SETUP
/// record covers header/footer margins separately.
fn write_margin_records(stream: &mut Vec<u8>, sheet: &Worksheet) {
    let ps = sheet.page_setup();
    for (record_type, value) in [
        (LEFT_MARGIN_RECORD, ps.left_margin),
        (RIGHT_MARGIN_RECORD, ps.right_margin),
        (TOP_MARGIN_RECORD, ps.top_margin),
        (BOTTOM_MARGIN_RECORD, ps.bottom_margin),
    ] {
        stream.extend_from_slice(&record_type.to_le_bytes());
        stream.extend_from_slice(&8u16.to_le_bytes());
        stream.extend_from_slice(&value.to_le_bytes());
    }
}

/// Emit PRINTHEADERS and PRINTGRIDLINES boolean records when their
/// flags are set on the page setup. Each body is a single u16 = 1
/// (omitted when the corresponding flag is false to keep the stream
/// minimal).
fn write_print_flags(stream: &mut Vec<u8>, sheet: &Worksheet) {
    let ps = sheet.page_setup();
    if ps.print_headings {
        stream.extend_from_slice(&PRINTHEADERS_RECORD.to_le_bytes());
        stream.extend_from_slice(&2u16.to_le_bytes());
        stream.extend_from_slice(&1u16.to_le_bytes());
    }
    if ps.print_gridlines {
        stream.extend_from_slice(&PRINTGRIDLINES_RECORD.to_le_bytes());
        stream.extend_from_slice(&2u16.to_le_bytes());
        stream.extend_from_slice(&1u16.to_le_bytes());
    }
}

/// Emit HPAGEBREAKS (row breaks, MS-XLS §2.4.139) and VPAGEBREAKS
/// (column breaks, MS-XLS §2.4.342). Body for each is count(u16)
/// followed by N × (id u16, min u16, max u16). Breaks beyond u16::MAX
/// are skipped because BIFF8 can't address them.
fn write_page_break_records(stream: &mut Vec<u8>, sheet: &Worksheet) {
    fn emit(
        stream: &mut Vec<u8>,
        record_type: u16,
        breaks: &[duke_sheets_core::worksheet::PageBreak],
    ) {
        let usable: Vec<_> = breaks
            .iter()
            .filter(|b| {
                b.id <= u16::MAX as u32 && b.min <= u16::MAX as u32 && b.max <= u16::MAX as u32
            })
            .collect();
        if usable.is_empty() {
            return;
        }
        let mut body = Vec::with_capacity(2 + usable.len() * 6);
        body.extend_from_slice(&(usable.len() as u16).to_le_bytes());
        for b in usable {
            body.extend_from_slice(&(b.id as u16).to_le_bytes());
            body.extend_from_slice(&(b.min as u16).to_le_bytes());
            body.extend_from_slice(&(b.max as u16).to_le_bytes());
        }
        stream.extend_from_slice(&record_type.to_le_bytes());
        stream.extend_from_slice(&(body.len() as u16).to_le_bytes());
        stream.extend_from_slice(&body);
    }
    emit(stream, HPAGEBREAKS_RECORD, sheet.row_breaks());
    emit(stream, VPAGEBREAKS_RECORD, sheet.col_breaks());
}

/// Emit one or more MERGECELLS records for the worksheet
/// (MS-XLS §2.4.169). Body: cmcs (u16, count) followed by
/// `cmcs` × Ref8U (rwFirst, rwLast, colFirst, colLast - all u16).
/// Multiple records are emitted when the merge count exceeds
/// `MERGECELLS_MAX_PER_RECORD`. Ranges with rows beyond u16::MAX or
/// cols beyond u16::MAX are silently skipped because BIFF8 can't
/// address them.
fn write_mergecells(stream: &mut Vec<u8>, sheet: &Worksheet) {
    let regions = sheet.merged_regions();
    if regions.is_empty() {
        return;
    }
    for chunk in regions.chunks(MERGECELLS_MAX_PER_RECORD) {
        let mut body = Vec::with_capacity(2 + chunk.len() * 8);
        let mut emitted: u16 = 0;
        let count_offset = body.len();
        body.extend_from_slice(&0u16.to_le_bytes()); // placeholder for cmcs
        for region in chunk {
            if region.start.row > u16::MAX as u32 || region.end.row > u16::MAX as u32 {
                continue;
            }
            body.extend_from_slice(&(region.start.row as u16).to_le_bytes());
            body.extend_from_slice(&(region.end.row as u16).to_le_bytes());
            body.extend_from_slice(&region.start.col.to_le_bytes());
            body.extend_from_slice(&region.end.col.to_le_bytes());
            emitted += 1;
        }
        if emitted == 0 {
            continue;
        }
        body[count_offset..count_offset + 2].copy_from_slice(&emitted.to_le_bytes());
        stream.extend_from_slice(&MERGECELLS_RECORD.to_le_bytes());
        stream.extend_from_slice(&(body.len() as u16).to_le_bytes());
        stream.extend_from_slice(&body);
    }
}

/// Emit a NAME record (MS-XLS §2.4.176, "Lbl") for each worksheet
/// that has a Print_Area or Print_Titles set on its page setup. The
/// record is sheet-scoped (itab = sheet_idx + 1) and uses the
/// built-in name index byte: 0x06 for Print_Area, 0x07 for
/// Print_Titles.
///
/// Print_Area body: a single tArea3D ptg covering the print area
/// range. Print_Titles body holds row titles, column titles, or
/// both:
///
///   - rows only: tArea3D(first_row..last_row, col 0..0xFF)
///   - cols only: tArea3D(row 0..0xFFFF, first_col..last_col)
///   - both:      tMemFunc + cce + tArea3D(rows) + tArea3D(cols)
///                + tList
///
/// ixti is hardcoded to 0; our XlsReader's extract_filter_db_range
/// and extract_print_titles ignore the EXTERNSHEET index when
/// extracting the row/col fields, so the round-trip works without
/// emitting EXTERNSHEET / SUPBOOK records too.
/// Maps user-defined name strings to their 1-based NAME table index
/// (the value embedded in tName ptgs). Names are matched case-
/// insensitively to mirror Excel's name resolution.
#[derive(Debug, Default)]
struct NameTable {
    by_name: HashMap<String, NameInfo>,
}

#[derive(Debug, Clone, Copy)]
struct NameInfo {
    idx: u16,
    body_class: OperandClass,
}

impl NameTable {
    fn idx_for_name(&self, name: &str) -> Option<u16> {
        self.by_name
            .get(&name.to_ascii_lowercase())
            .map(|info| info.idx)
    }

    fn body_class_for_name(&self, name: &str) -> Option<OperandClass> {
        self.by_name
            .get(&name.to_ascii_lowercase())
            .map(|info| info.body_class)
    }
}

fn build_name_table(workbook: &Workbook) -> NameTable {
    let mut by_name = HashMap::new();
    for (i, nr) in user_names_in_xls_emit_order(workbook)
        .into_iter()
        .enumerate()
    {
        if i >= u16::MAX as usize {
            break;
        }
        let body_class = parse_name_body(&nr.refers_to)
            .as_ref()
            .map(name_body_operand_class)
            .unwrap_or(OperandClass::V);
        // Names are 1-based in tName ptg encoding.
        by_name.insert(
            nr.name.to_ascii_lowercase(),
            NameInfo {
                idx: (i as u16) + 1,
                body_class,
            },
        );
    }
    NameTable { by_name }
}

/// Emit one NAME record (MS-XLS §2.4.176, Lbl) per user-defined named
/// range in the workbook. Layout:
///
///   flags (u16)    - 0 for visible, 0x0001 fHidden if NamedRange.hidden
///   chKey (u8)     - 0
///   cch (u8)       - name string length (UTF-16 code units)
///   cce (u16)      - formula body length
///   reserved (u16) - 0
///   itab (u16)     - 0 for workbook scope, sheet_idx + 1 for sheet
///                    scope; the reader maps itab=0 -> 0xFFFFFFFF
///                    sheet_idx (workbook), else sheet_idx = itab - 1
///   reserved (u32) - 0
///   name string    - flags byte (0=Latin1, 1=UTF-16) + chars
///   formula body   - parsed via duke-sheets-formula and recompiled
///                    via compile_ptgs_with_context. parse failures
///                    emit no formula body (cce = 0) so the reader
///                    sees the name without a body.
fn write_user_name_records(
    stream: &mut Vec<u8>,
    workbook: &Workbook,
    externsheet: &ExternSheetTable,
    name_table: &NameTable,
    addins: &AddinTable,
) {
    use duke_sheets_core::named_range::NameScope;

    for nr in user_names_in_xls_emit_order(workbook) {
        let name_units: Vec<u16> = nr.name.encode_utf16().collect();
        if name_units.is_empty() || name_units.len() > u8::MAX as usize {
            continue;
        }

        let formula_body: Vec<u8> = {
            if let Some(expr) = parse_name_body(&nr.refers_to) {
                let mut bytes = Vec::new();
                let mut extra = Vec::new();
                let operand_class = name_body_operand_class(&expr);
                // Array constants in a defined-name body would need the
                // NAME record's cce/rgcb split (cce counts rgce only); that
                // path is unsupported, so reject a non-empty rgcb here.
                if compile_ptgs_with_context(
                    &expr,
                    &mut bytes,
                    &mut extra,
                    externsheet,
                    name_table,
                    addins,
                    operand_class,
                )
                .is_ok()
                    && extra.is_empty()
                {
                    bytes
                } else {
                    Vec::new()
                }
            } else {
                Vec::new()
            }
        };

        let mut flags: u16 = 0;
        if nr.hidden {
            flags |= 0x0001;
        }
        let itab: u16 = match nr.scope {
            NameScope::Workbook => 0,
            NameScope::Sheet(idx) => (idx.min(u16::MAX as usize) as u16) + 1,
        };

        let high_byte = name_units.iter().any(|&u| u > 0xFF);
        let name_bytes_len = if high_byte {
            1 + name_units.len() * 2
        } else {
            1 + name_units.len()
        };

        let mut body = Vec::with_capacity(15 + name_bytes_len + formula_body.len());
        body.extend_from_slice(&flags.to_le_bytes());
        body.push(0); // chKey
        body.push(name_units.len() as u8);
        body.extend_from_slice(&(formula_body.len() as u16).to_le_bytes());
        body.extend_from_slice(&0u16.to_le_bytes()); // reserved
        body.extend_from_slice(&itab.to_le_bytes());
        body.extend_from_slice(&[0u8; 4]); // reserved
        if high_byte {
            body.push(0x01);
            for u in &name_units {
                body.extend_from_slice(&u.to_le_bytes());
            }
        } else {
            body.push(0x00);
            for u in &name_units {
                body.push(*u as u8);
            }
        }
        body.extend_from_slice(&formula_body);

        stream.extend_from_slice(&NAME_RECORD.to_le_bytes());
        stream.extend_from_slice(&(body.len() as u16).to_le_bytes());
        stream.extend_from_slice(&body);
    }
}

fn parse_name_body(refers_to: &str) -> Option<duke_sheets_formula::FormulaExpr> {
    let to_parse = if refers_to.starts_with('=') {
        refers_to.to_string()
    } else {
        format!("={refers_to}")
    };
    duke_sheets_formula::parse_formula(&to_parse).ok()
}

fn user_names_in_xls_emit_order(
    workbook: &Workbook,
) -> Vec<&duke_sheets_core::named_range::NamedRange> {
    // Excel SaveAs emits user-defined NAME/Lbl records ordered by name.
    // Keep our name table in the same order so PtgName indexes remain
    // stable through Excel's open/save cycle.
    let mut names = workbook.named_ranges().iter().collect::<Vec<_>>();
    names.sort_by(|a, b| {
        a.name
            .to_ascii_lowercase()
            .cmp(&b.name.to_ascii_lowercase())
            .then_with(|| name_scope_sort_key(&a.scope).cmp(&name_scope_sort_key(&b.scope)))
    });
    names
}

fn name_scope_sort_key(scope: &duke_sheets_core::named_range::NameScope) -> usize {
    match scope {
        duke_sheets_core::named_range::NameScope::Workbook => 0,
        duke_sheets_core::named_range::NameScope::Sheet(idx) => idx.saturating_add(1),
    }
}

/// Maps sheet names to their EXTERNSHEET ixti index, with case-
/// insensitive lookup to match Excel's name resolution. Built once per
/// workbook write; passed to `compile_ptgs_with_externsheet` so 3D
/// refs (tRef3D / tArea3D) can encode the right ixti.
#[derive(Debug, Default)]
struct ExternSheetTable {
    /// Lower-cased sheet name → ixti.
    by_name: HashMap<String, u16>,
    /// Sheet count, for SUPBOOK self-ref ctab.
    sheet_count: u16,
    /// EXTERNSHEET entries in ixti order. Values are 0-based sheet indexes.
    entries: Vec<u16>,
    /// True when the workbook uses Analysis-ToolPak add-in functions. The
    /// AddIn SUPBOOK then occupies SUPBOOK index 0 and an XTI for it is
    /// prepended to the EXTERNSHEET at ixti 0, so the sheet XTIs that
    /// `by_name` indexes are shifted up by one and reference SUPBOOK 1.
    addin_present: bool,
}

impl ExternSheetTable {
    fn ixti_for_sheet(&self, name: &str) -> Option<u16> {
        self.by_name.get(&name.to_ascii_lowercase()).copied()
    }

    /// SUPBOOK index the sheet XTIs reference: 1 when an AddIn SUPBOOK
    /// precedes the self-ref SUPBOOK, 0 otherwise.
    fn self_ref_supbook_idx(&self) -> u16 {
        if self.addin_present {
            1
        } else {
            0
        }
    }
}

fn build_externsheet_table(workbook: &Workbook, addin_present: bool) -> ExternSheetTable {
    let sheets = workbook
        .worksheets()
        .enumerate()
        .filter_map(|(idx, sheet)| {
            if idx > u16::MAX as usize {
                None
            } else {
                Some((idx as u16, sheet.name().to_string()))
            }
        })
        .collect::<Vec<_>>();
    let sheet_by_name = sheets
        .iter()
        .map(|(idx, name)| (name.to_ascii_lowercase(), *idx))
        .collect::<HashMap<_, _>>();

    let mut entries = Vec::new();
    for sheet_name in formula_referenced_sheet_names(workbook) {
        if let Some(idx) = sheet_by_name.get(&sheet_name.to_ascii_lowercase()) {
            if !entries.contains(idx) {
                entries.push(*idx);
            }
        }
    }
    for (idx, _) in &sheets {
        if !entries.contains(idx) {
            entries.push(*idx);
        }
    }

    // When an AddIn SUPBOOK is present its XTI is prepended at ixti 0, so
    // every sheet XTI shifts up by one.
    let ixti_base: u16 = if addin_present { 1 } else { 0 };
    let mut by_name = HashMap::new();
    for (ixti, sheet_idx) in entries.iter().enumerate() {
        if let Some((_, name)) = sheets.iter().find(|(idx, _)| idx == sheet_idx) {
            by_name.insert(name.to_ascii_lowercase(), ixti as u16 + ixti_base);
        }
    }
    ExternSheetTable {
        by_name,
        sheet_count: workbook.sheet_count().min(u16::MAX as usize) as u16,
        entries,
        addin_present,
    }
}

/// Table of add-in functions used anywhere in the workbook,
/// in the order Excel writes their EXTERNNAME records: distinct canonical
/// names sorted alphabetically (ASCII), each assigned a 1-based `nameindex`.
///
/// BIFF8 serializes ATP functions (Ftab 384..=476) as an add-in UDF call:
/// `PtgNameX` references an EXTERNNAME record in the AddIn SUPBOOK by this
/// 1-based index. See [`write_supbook_and_externsheet`].
#[derive(Debug, Default)]
struct AddinTable {
    /// Canonical (uppercase) names in EXTERNNAME emission order.
    names: Vec<String>,
    /// Uppercase name → 1-based nameindex.
    by_name: HashMap<String, u16>,
}

impl AddinTable {
    fn is_empty(&self) -> bool {
        self.names.is_empty()
    }

    /// 1-based EXTERNNAME index for an add-in function name, or None when the
    /// name was not collected (caller should fall back to native emission).
    fn nameindex_for(&self, name: &str) -> Option<u16> {
        self.by_name.get(&name.to_ascii_uppercase()).copied()
    }
}

fn build_addin_table(workbook: &Workbook) -> AddinTable {
    use std::collections::BTreeMap;
    let mut set: BTreeMap<String, String> = BTreeMap::new();
    for sheet in workbook.worksheets() {
        for (_, _, formula) in sheet.formula_cells() {
            collect_addin_function_names(formula, &mut set);
        }
    }
    for named_range in user_names_in_xls_emit_order(workbook) {
        collect_addin_function_names(&named_range.refers_to, &mut set);
    }
    let names: Vec<String> = set.into_values().collect();
    let by_name = names
        .iter()
        .enumerate()
        .map(|(i, name)| (name.to_ascii_uppercase(), (i + 1) as u16))
        .collect();
    AddinTable { names, by_name }
}

fn collect_addin_function_names(
    formula: &str,
    out: &mut std::collections::BTreeMap<String, String>,
) {
    let formula = if formula.starts_with('=') {
        formula.to_string()
    } else {
        format!("={formula}")
    };
    if let Ok(expr) = duke_sheets_formula::parse_formula(&formula) {
        collect_addin_names_expr(&expr, out);
    }
}

fn collect_addin_names_expr(
    expr: &duke_sheets_formula::FormulaExpr,
    out: &mut std::collections::BTreeMap<String, String>,
) {
    use duke_sheets_formula::decompile::{function_index, function_is_biff8_addin, function_name};
    use duke_sheets_formula::FormulaExpr;

    match expr {
        FormulaExpr::Function { name, args } => {
            if let Some(idx) = function_index(name) {
                if function_is_biff8_addin(idx) {
                    // Canonical (uppercase) name so EXTERNNAME spelling and
                    // sort order are independent of how the user typed it.
                    let canonical = function_name(idx).to_string();
                    out.entry(canonical.to_ascii_uppercase())
                        .or_insert(canonical);
                }
            }
            for arg in args {
                collect_addin_names_expr(arg, out);
            }
        }
        FormulaExpr::ExternalFunction { name, args, .. } => {
            out.entry(name.to_ascii_uppercase())
                .or_insert_with(|| name.clone());
            for arg in args {
                collect_addin_names_expr(arg, out);
            }
        }
        FormulaExpr::BinaryOp { left, right, .. } => {
            collect_addin_names_expr(left, out);
            collect_addin_names_expr(right, out);
        }
        FormulaExpr::UnaryOp { operand, .. } => collect_addin_names_expr(operand, out),
        FormulaExpr::Array(rows) => {
            for row in rows {
                for item in row {
                    collect_addin_names_expr(item, out);
                }
            }
        }
        FormulaExpr::Number(_)
        | FormulaExpr::String(_)
        | FormulaExpr::Boolean(_)
        | FormulaExpr::Error(_)
        | FormulaExpr::NameRef(_)
        | FormulaExpr::CellRef(_)
        | FormulaExpr::RangeRef(_)
        | FormulaExpr::StructuredRef(_)
        | FormulaExpr::ExternalRef(_)
        | FormulaExpr::Empty => {}
    }
}

fn formula_referenced_sheet_names(workbook: &Workbook) -> Vec<String> {
    let mut names = Vec::new();
    for sheet in workbook.worksheets() {
        for (_, _, formula) in sheet.formula_cells() {
            collect_formula_text_sheet_names(formula, &mut names);
        }
    }
    for named_range in user_names_in_xls_emit_order(workbook) {
        collect_formula_text_sheet_names(&named_range.refers_to, &mut names);
    }
    names
}

fn collect_formula_text_sheet_names(formula: &str, out: &mut Vec<String>) {
    let formula = if formula.starts_with('=') {
        formula.to_string()
    } else {
        format!("={formula}")
    };
    if let Ok(expr) = duke_sheets_formula::parse_formula(&formula) {
        collect_formula_expr_sheet_names(&expr, out);
    }
}

fn collect_formula_expr_sheet_names(
    expr: &duke_sheets_formula::FormulaExpr,
    out: &mut Vec<String>,
) {
    use duke_sheets_formula::FormulaExpr;

    match expr {
        FormulaExpr::CellRef(cell_ref) => {
            if let Some(sheet) = cell_ref.sheet.as_ref() {
                push_unique_sheet_name(out, sheet);
            }
        }
        FormulaExpr::RangeRef(range_ref) => {
            if let Some(sheet) = range_ref.sheet.as_ref() {
                push_unique_sheet_name(out, sheet);
            }
        }
        FormulaExpr::BinaryOp { left, right, .. } => {
            collect_formula_expr_sheet_names(left, out);
            collect_formula_expr_sheet_names(right, out);
        }
        FormulaExpr::UnaryOp { operand, .. } => collect_formula_expr_sheet_names(operand, out),
        FormulaExpr::Function { args, .. } => {
            for arg in args {
                collect_formula_expr_sheet_names(arg, out);
            }
        }
        FormulaExpr::ExternalFunction { args, .. } => {
            for arg in args {
                collect_formula_expr_sheet_names(arg, out);
            }
        }
        FormulaExpr::Array(rows) => {
            for row in rows {
                for item in row {
                    collect_formula_expr_sheet_names(item, out);
                }
            }
        }
        FormulaExpr::Number(_)
        | FormulaExpr::String(_)
        | FormulaExpr::Boolean(_)
        | FormulaExpr::Error(_)
        | FormulaExpr::NameRef(_)
        | FormulaExpr::StructuredRef(_)
        | FormulaExpr::ExternalRef(_)
        | FormulaExpr::Empty => {}
    }
}

fn push_unique_sheet_name(out: &mut Vec<String>, sheet: &str) {
    if !out
        .iter()
        .any(|existing| existing.eq_ignore_ascii_case(sheet))
    {
        out.push(sheet.to_string());
    }
}

/// Emit the SUPBOOK / EXTERNNAME / EXTERNSHEET records (MS-XLS §2.4.273,
/// §2.4.150, §2.4.105) for the global substream.
///
/// Without add-in functions this is a single self-ref SUPBOOK
/// (ctab=sheet_count, cch=0x0401) plus an EXTERNSHEET with one XTI per
/// referenced worksheet (sup_book_idx=0).
///
/// With Analysis-ToolPak add-in functions present, Excel writes:
///   SUPBOOK #0  AddIn sentinel (ctab=0x0001, cch=0x3A01)
///   EXTERNNAME  one per distinct add-in function (nameindex order)
///   SUPBOOK #1  self-ref
///   EXTERNSHEET XTI[0] = AddIn (sup=0, itab=0xFFFE workbook-level), then
///               the sheet XTIs referencing self-ref SUPBOOK #1 at ixti 1+.
///
/// EXTERNSHEET XTI body: sup_book_idx(u16) + itabFirst(i16) + itabLast(i16).
/// XTIs are ordered by first formula use so tRef3D/tArea3D ixti values match
/// Excel's SaveAs ordering for referenced sheets.
fn write_supbook_and_externsheet(
    stream: &mut Vec<u8>,
    table: &ExternSheetTable,
    addins: &AddinTable,
) {
    if table.sheet_count == 0 {
        return;
    }

    if !addins.is_empty() {
        // AddIn SUPBOOK sentinel: ctab=0x0001, cch=0x3A01.
        stream.extend_from_slice(&SUPBOOK_RECORD.to_le_bytes());
        stream.extend_from_slice(&4u16.to_le_bytes());
        stream.extend_from_slice(&0x0001u16.to_le_bytes());
        stream.extend_from_slice(&0x3A01u16.to_le_bytes());
        // EXTERNNAME records belong to the SUPBOOK immediately preceding
        // them, so emit them right after the AddIn SUPBOOK.
        for name in &addins.names {
            write_addin_externname(stream, name);
        }
    }

    // Self-ref SUPBOOK: ctab=sheet_count, cch=0x0401.
    stream.extend_from_slice(&SUPBOOK_RECORD.to_le_bytes());
    stream.extend_from_slice(&4u16.to_le_bytes());
    stream.extend_from_slice(&table.sheet_count.to_le_bytes());
    stream.extend_from_slice(&0x0401u16.to_le_bytes());

    // EXTERNSHEET
    let self_ref_sup = table.self_ref_supbook_idx();
    let sheet_count = table.entries.len().min(u16::MAX as usize);
    let total = (if table.addin_present {
        sheet_count + 1
    } else {
        sheet_count
    })
    .min(u16::MAX as usize) as u16;
    let body_len = 2 + (total as usize) * 6;
    stream.extend_from_slice(&EXTERNSHEET_RECORD.to_le_bytes());
    stream.extend_from_slice(&(body_len as u16).to_le_bytes());
    stream.extend_from_slice(&total.to_le_bytes());
    if table.addin_present {
        // XTI[0]: the AddIn SUPBOOK, workbook-level scope (itab=0xFFFE).
        stream.extend_from_slice(&0u16.to_le_bytes());
        stream.extend_from_slice(&0xFFFEu16.to_le_bytes());
        stream.extend_from_slice(&0xFFFEu16.to_le_bytes());
    }
    for sheet_idx in table.entries.iter().take(sheet_count) {
        stream.extend_from_slice(&self_ref_sup.to_le_bytes()); // sup_book_idx
        stream.extend_from_slice(&sheet_idx.to_le_bytes()); // first_sheet
        stream.extend_from_slice(&sheet_idx.to_le_bytes()); // last_sheet
    }
}

/// Emit one EXTERNNAME record (MS-XLS §2.4.150) describing an Analysis-ToolPak
/// add-in function.
///
/// Body layout (matches Excel's native output):
/// ```text
/// grbit      (2 bytes) = 0x0000
/// reserved   (2 bytes) = 0x0000   (sheet index / not used for add-ins)
/// reserved   (2 bytes) = 0x0000   (not used)
/// cch        (1 byte)  = name length
/// grbit      (1 byte)  = 0x00     (compressed: one byte per character)
/// name       (cch bytes, ASCII)
/// cce        (2 bytes) = 0x0002   name-definition formula length
/// rgce       = 1C 17              PtgErr(#REF!) — never-evaluated placeholder
/// ```
fn write_addin_externname(stream: &mut Vec<u8>, name: &str) {
    let name_bytes = name.as_bytes(); // ATP function names are ASCII
    let mut body = Vec::with_capacity(6 + 2 + name_bytes.len() + 4);
    body.extend_from_slice(&0u16.to_le_bytes()); // grbit
    body.extend_from_slice(&0u16.to_le_bytes()); // reserved
    body.extend_from_slice(&0u16.to_le_bytes()); // reserved
    body.push(name_bytes.len() as u8); // cch
    body.push(0x00); // grbit: compressed/ASCII
    body.extend_from_slice(name_bytes);
    body.extend_from_slice(&[0x02, 0x00, 0x1C, 0x17]); // cce=2, PtgErr #REF!
    stream.extend_from_slice(&EXTERNNAME_RECORD.to_le_bytes());
    stream.extend_from_slice(&(body.len() as u16).to_le_bytes());
    stream.extend_from_slice(&body);
}

fn write_print_name_records(stream: &mut Vec<u8>, workbook: &Workbook) {
    for (sheet_idx, sheet) in workbook.worksheets().enumerate() {
        if sheet_idx > u16::MAX as usize - 1 {
            continue;
        }
        let itab = (sheet_idx as u16) + 1;
        let ps = sheet.page_setup();

        if let Some(range) = ps.print_area.as_ref() {
            let body = build_print_area_body(range);
            emit_builtin_name(stream, itab, BUILTIN_NAME_PRINT_AREA, &body);
        }

        let print_titles_body = build_print_titles_body(ps.repeat_rows, ps.repeat_cols);
        if !print_titles_body.is_empty() {
            emit_builtin_name(stream, itab, BUILTIN_NAME_PRINT_TITLES, &print_titles_body);
        }

        if let Some(af) = sheet.auto_filter() {
            let body = build_print_area_body(&af.range);
            emit_builtin_name(stream, itab, BUILTIN_NAME_FILTER_DATABASE, &body);
        }
    }
}

fn emit_builtin_name(stream: &mut Vec<u8>, itab: u16, builtin_index: u8, formula_body: &[u8]) {
    let cce = formula_body.len() as u16;
    let mut body = Vec::with_capacity(15 + 2 + formula_body.len());
    body.extend_from_slice(&0x0020u16.to_le_bytes()); // flags: fBuiltin
    body.push(0); // chKey
    body.push(1); // cch (one "character" - the built-in index byte)
    body.extend_from_slice(&cce.to_le_bytes());
    body.extend_from_slice(&0u16.to_le_bytes()); // reserved (ixals - external book index)
    body.extend_from_slice(&itab.to_le_bytes());
    body.extend_from_slice(&[0u8; 4]); // 4 reserved bytes
    body.push(0); // name flags: 0 = compressed/Latin-1
    body.push(builtin_index); // the "character" is actually the built-in index
    body.extend_from_slice(formula_body);

    stream.extend_from_slice(&NAME_RECORD.to_le_bytes());
    stream.extend_from_slice(&(body.len() as u16).to_le_bytes());
    stream.extend_from_slice(&body);
}

/// Build a tArea3D ptg body (11 bytes): token + ixti(u16) + first_row
/// + last_row + first_col + last_col. ixti is 0 (no EXTERNSHEET).
fn build_t_area_3d(first_row: u16, last_row: u16, first_col: u16, last_col: u16) -> [u8; 11] {
    let mut buf = [0u8; 11];
    buf[0] = 0x3B; // tArea3D (R class)
                   // bytes 1..3: ixti (u16) = 0
    buf[3..5].copy_from_slice(&first_row.to_le_bytes());
    buf[5..7].copy_from_slice(&last_row.to_le_bytes());
    buf[7..9].copy_from_slice(&(first_col & 0x3FFF).to_le_bytes());
    buf[9..11].copy_from_slice(&(last_col & 0x3FFF).to_le_bytes());
    buf
}

fn build_print_area_body(range: &duke_sheets_core::CellRange) -> Vec<u8> {
    let first_row = range.start.row.min(u16::MAX as u32) as u16;
    let last_row = range.end.row.min(u16::MAX as u32) as u16;
    let first_col = range.start.col;
    let last_col = range.end.col;
    build_t_area_3d(first_row, last_row, first_col, last_col).to_vec()
}

fn build_print_titles_body(
    repeat_rows: Option<(u32, u32)>,
    repeat_cols: Option<(u16, u16)>,
) -> Vec<u8> {
    let row_area = repeat_rows.map(|(r1, r2)| {
        build_t_area_3d(
            r1.min(u16::MAX as u32) as u16,
            r2.min(u16::MAX as u32) as u16,
            0,
            0xFF,
        )
    });
    let col_area = repeat_cols.map(|(c1, c2)| build_t_area_3d(0, 0xFFFF, c1, c2));

    match (row_area, col_area) {
        (None, None) => Vec::new(),
        (Some(rows), None) => rows.to_vec(),
        (None, Some(cols)) => cols.to_vec(),
        (Some(rows), Some(cols)) => {
            // tMemFunc(0x29) + cce(u16) + rows(11) + cols(11) + tList(0x10)
            let inner_len: u16 = 11 + 11 + 1;
            let mut body = Vec::with_capacity(3 + inner_len as usize);
            body.push(0x29);
            body.extend_from_slice(&inner_len.to_le_bytes());
            body.extend_from_slice(&rows);
            body.extend_from_slice(&cols);
            body.push(0x10); // tList (range union)
            body
        }
    }
}

/// Emit one HLINK record per cell-attached hyperlink (MS-XLS §2.4.144).
///
/// Body layout:
///   - Ref8U (8 bytes): row_first/row_last/col_first/col_last (u16
///     each); single-cell hyperlinks use row_first==row_last and
///     col_first==col_last.
///   - HLINK CLSID (16 bytes).
///   - streamVersion (4 bytes) = 0x00000002.
///   - flags (4 bytes): bit 0 has_moniker, bit 1 is_absolute,
///     bit 3 has_location, bit 4 has_display, bit 7 has_frame.
///   - displayName (only when bit 4 set): char_count(u32) + UTF-16LE
///     chars + 0x0000 terminator.
///   - frameName (when bit 7 set): same encoding.
///   - moniker (when bit 0 set): URL_MONIKER_CLSID + url_byte_len(u32)
///     + UTF-16LE chars + 0x0000 terminator.
///   - location (when bit 3 set): char_count(u32) + UTF-16LE chars +
///     0x0000 terminator.
///
/// Internal `#Sheet!A1` targets are emitted with no moniker but a
/// location string. URL-style targets use the URL moniker. File
/// monikers and other types are not yet emitted; the reader can read
/// them but the writer round-trips file paths via the URL path.
fn write_hlink_records(stream: &mut Vec<u8>, sheet: &Worksheet) {
    use duke_sheets_core::CellAddress;

    let mut entries: Vec<(CellAddress, &duke_sheets_core::Hyperlink)> =
        sheet.hyperlinks().iter().map(|(a, h)| (*a, h)).collect();
    entries.sort_by_key(|(addr, _)| (addr.row, addr.col));

    for (addr, hyperlink) in entries {
        if addr.row > u16::MAX as u32 {
            continue;
        }
        let mut body = Vec::with_capacity(64);
        body.extend_from_slice(&(addr.row as u16).to_le_bytes());
        body.extend_from_slice(&(addr.row as u16).to_le_bytes()); // row_last
        body.extend_from_slice(&addr.col.to_le_bytes());
        body.extend_from_slice(&addr.col.to_le_bytes()); // col_last
        body.extend_from_slice(&HLINK_CLSID);
        body.extend_from_slice(&2u32.to_le_bytes()); // streamVersion

        let target = hyperlink.target.as_str();
        let is_internal = target.starts_with('#') || target.is_empty();
        let location_text: Option<String> = if is_internal {
            if target.starts_with('#') {
                Some(target[1..].to_string())
            } else {
                hyperlink.location.clone()
            }
        } else {
            hyperlink.location.clone()
        };

        let display = hyperlink.display.as_deref();

        let mut flags: u32 = 0;
        if !is_internal {
            flags |= 0x0001 | 0x0002; // hlstmfHasMoniker + hlstmfIsAbsolute
        }
        if location_text.is_some() {
            flags |= 0x0008; // hlstmfHasLocationStr (text mark)
        }
        if display.is_some() {
            // hlstmfHasDisplayName only. The 0x04 bit
            // (hlstmfSiteGaveDisplayName) means "displayName is implicit
            // and not stored" — combining it with HasDisplayName is a
            // contradiction Excel rejects with "Open method of Workbooks
            // class failed".
            flags |= 0x0010;
        }
        body.extend_from_slice(&flags.to_le_bytes());

        if let Some(text) = display {
            push_hlink_string(&mut body, text);
        }

        if !is_internal {
            body.extend_from_slice(&URL_MONIKER_CLSID);
            push_url_moniker_payload(&mut body, target);
        }

        if let Some(loc) = location_text.as_deref() {
            push_hlink_string(&mut body, loc);
        }

        stream.extend_from_slice(&HLINK_RECORD.to_le_bytes());
        stream.extend_from_slice(&(body.len() as u16).to_le_bytes());
        stream.extend_from_slice(&body);
    }
}

/// Emit a FILTERMODE record (1 byte body of zero) plus one AUTOFILTER
/// record per FilterColumn when the worksheet has an auto-filter set.
/// FILTERMODE indicates the worksheet has filtering active; the
/// AUTOFILTER body carries per-column criteria.
///
/// AUTOFILTER body (24 bytes minimum + variable strings):
///   bytes 0..2  - i_entry (column offset within the filter range)
///   bytes 2..4  - flags:
///       bits 0-1 - 0x01 = OR-join the two dopers (else AND)
///       bit 4    - fTopN
///       bit 5    - fTop (top vs bottom for top-N)
///       bit 6    - fPercent
///       bits 7-15 - wTopN value
///   bytes 4..14 - doper1 (vt + op + 8-byte payload)
///   bytes 14..24 - doper2
///   bytes 24..  - inline strings for any string-typed dopers
fn write_autofilter_records(stream: &mut Vec<u8>, sheet: &Worksheet) {
    use duke_sheets_core::auto_filter::{ColumnFilter, FilterOperator};

    let Some(af) = sheet.auto_filter() else {
        return;
    };

    stream.extend_from_slice(&FILTERMODE_RECORD.to_le_bytes());
    stream.extend_from_slice(&0u16.to_le_bytes());

    for column in &af.filter_columns {
        if column.col_id > u16::MAX as u32 {
            continue;
        }
        let mut body = Vec::with_capacity(32);
        body.extend_from_slice(&(column.col_id as u16).to_le_bytes());

        let mut flags: u16 = 0;
        let mut doper1 = [0u8; 10];
        let mut doper2 = [0u8; 10];
        let mut trailing_strings: Vec<String> = Vec::new();

        match &column.filter {
            ColumnFilter::Top10(t) => {
                flags |= 0x0010; // fTopN
                if t.top {
                    flags |= 0x0020;
                }
                if t.percent {
                    flags |= 0x0040;
                }
                let n = (t.val as u16) & 0x01FF;
                flags |= n << 7;
            }
            ColumnFilter::Custom(c) => {
                if !c.and {
                    flags |= 0x0001; // OR-join
                }
                if let Some(cond) = c.conditions.first() {
                    encode_custom_condition(cond, &mut doper1, &mut trailing_strings);
                }
                if let Some(cond) = c.conditions.get(1) {
                    encode_custom_condition(cond, &mut doper2, &mut trailing_strings);
                }
            }
            ColumnFilter::Values(v) => {
                flags |= 0x0001; // OR-join (Values are matched as Equal+OR)
                if let Some(value) = v.values.first() {
                    encode_custom_condition(
                        &duke_sheets_core::auto_filter::CustomFilterCondition {
                            operator: FilterOperator::Equal,
                            value: value.clone(),
                        },
                        &mut doper1,
                        &mut trailing_strings,
                    );
                }
                if let Some(value) = v.values.get(1) {
                    encode_custom_condition(
                        &duke_sheets_core::auto_filter::CustomFilterCondition {
                            operator: FilterOperator::Equal,
                            value: value.clone(),
                        },
                        &mut doper2,
                        &mut trailing_strings,
                    );
                }
            }
            ColumnFilter::Dynamic(_) | ColumnFilter::Color(_) => {
                // Reader doesn't decode these from BIFF8 today; skip
                // emit so we don't create records that won't round-
                // trip via XlsReader.
                continue;
            }
        }

        body.extend_from_slice(&flags.to_le_bytes());
        body.extend_from_slice(&doper1);
        body.extend_from_slice(&doper2);
        for s in &trailing_strings {
            push_autofilter_string(&mut body, s);
        }

        stream.extend_from_slice(&AUTOFILTER_RECORD.to_le_bytes());
        stream.extend_from_slice(&(body.len() as u16).to_le_bytes());
        stream.extend_from_slice(&body);
    }
}

/// Encode a `CustomFilterCondition` into a 10-byte doper. Numeric
/// conditions store the value inline (vt=0x04, IEEE 754 f64). String
/// conditions emit vt=0x06 with cch in payload[4]; the actual UTF-8
/// string is collected into `trailing_strings` and appended to the
/// AUTOFILTER record body in order.
fn encode_custom_condition(
    cond: &duke_sheets_core::auto_filter::CustomFilterCondition,
    doper: &mut [u8; 10],
    trailing_strings: &mut Vec<String>,
) {
    use duke_sheets_core::auto_filter::FilterOperator;

    doper[1] = match cond.operator {
        FilterOperator::LessThan => 0x01,
        FilterOperator::Equal => 0x02,
        FilterOperator::LessThanOrEqual => 0x03,
        FilterOperator::GreaterThan => 0x04,
        FilterOperator::NotEqual => 0x05,
        FilterOperator::GreaterThanOrEqual => 0x06,
    };

    if let Ok(n) = cond.value.parse::<f64>() {
        doper[0] = 0x04; // f64
        doper[2..10].copy_from_slice(&n.to_le_bytes());
    } else {
        doper[0] = 0x06; // string
        let units = cond.value.encode_utf16().count() as u8;
        doper[6] = units; // cch lives at offset 4 of payload, i.e. doper[6]
        trailing_strings.push(cond.value.clone());
    }
}

fn push_autofilter_string(buf: &mut Vec<u8>, s: &str) {
    let units: Vec<u16> = s.encode_utf16().collect();
    let high_byte = units.iter().any(|&u| u > 0xFF);
    if high_byte {
        buf.push(0x01);
        for u in &units {
            buf.extend_from_slice(&u.to_le_bytes());
        }
    } else {
        buf.push(0x00);
        for u in &units {
            buf.push(*u as u8);
        }
    }
}

/// Emit one CONDFMT (MS-XLS §2.4.45) + CF (§2.4.43) pair per
/// `ConditionalFormatRule`. Each CONDFMT carries the bounding
/// rectangle and the per-range Ref8U list; each CF carries the rule
/// header + cce1 + cce2 + an empty dxf block (no formatting overrides
/// emitted yet) + formula1 + formula2.
///
/// The reader keeps the most recent CONDFMT range list and applies it
/// to subsequent CF records, so we emit them as alternating
/// CONDFMT → CF pairs to keep the mapping unambiguous.
fn write_conditional_formats(
    stream: &mut Vec<u8>,
    sheet: &Worksheet,
    externsheet: &ExternSheetTable,
    names: &NameTable,
    addins: &AddinTable,
) {
    use duke_sheets_core::conditional_format::{CfOperator, CfRuleType};

    for rule in sheet.conditional_formats() {
        if rule.ranges.is_empty() {
            continue;
        }

        let usable: Vec<&duke_sheets_core::CellRange> = rule
            .ranges
            .iter()
            .filter(|r| r.start.row <= u16::MAX as u32 && r.end.row <= u16::MAX as u32)
            .collect();
        if usable.is_empty() {
            continue;
        }

        // Compute enclosing range = bounding box of all individual
        // ranges. The reader skips this field but Excel/LO use it.
        let mut enc_first_row = u16::MAX;
        let mut enc_last_row = 0u16;
        let mut enc_first_col = u16::MAX;
        let mut enc_last_col = 0u16;
        for r in &usable {
            enc_first_row = enc_first_row.min(r.start.row as u16);
            enc_last_row = enc_last_row.max(r.end.row as u16);
            enc_first_col = enc_first_col.min(r.start.col);
            enc_last_col = enc_last_col.max(r.end.col);
        }

        let mut cfmt_body = Vec::with_capacity(14 + usable.len() * 8);
        cfmt_body.extend_from_slice(&1u16.to_le_bytes()); // cCF: 1 rule follows
        cfmt_body.extend_from_slice(&1u16.to_le_bytes()); // flags (fAlwaysCalc-ish, set to 1)
        cfmt_body.extend_from_slice(&enc_first_row.to_le_bytes());
        cfmt_body.extend_from_slice(&enc_last_row.to_le_bytes());
        cfmt_body.extend_from_slice(&enc_first_col.to_le_bytes());
        cfmt_body.extend_from_slice(&enc_last_col.to_le_bytes());
        cfmt_body.extend_from_slice(&(usable.len() as u16).to_le_bytes());
        for r in &usable {
            cfmt_body.extend_from_slice(&(r.start.row as u16).to_le_bytes());
            cfmt_body.extend_from_slice(&(r.end.row as u16).to_le_bytes());
            cfmt_body.extend_from_slice(&r.start.col.to_le_bytes());
            cfmt_body.extend_from_slice(&r.end.col.to_le_bytes());
        }
        stream.extend_from_slice(&CONDFMT_RECORD.to_le_bytes());
        stream.extend_from_slice(&(cfmt_body.len() as u16).to_le_bytes());
        stream.extend_from_slice(&cfmt_body);

        let (ct, cp, formula1_text, formula2_text) = match &rule.rule_type {
            CfRuleType::CellIs {
                operator,
                formula1,
                formula2,
            } => {
                let cp = match operator {
                    CfOperator::Between => 1u8,
                    CfOperator::NotBetween => 2,
                    CfOperator::Equal => 3,
                    CfOperator::NotEqual => 4,
                    CfOperator::GreaterThan => 5,
                    CfOperator::LessThan => 6,
                    CfOperator::GreaterThanOrEqual => 7,
                    CfOperator::LessThanOrEqual => 8,
                };
                (1u8, cp, Some(formula1.as_str()), formula2.as_deref())
            }
            CfRuleType::Expression { formula } => (2u8, 0u8, Some(formula.as_str()), None),
            _ => continue, // skip rule types we don't know how to emit
        };

        let f1 = formula1_text
            .map(|t| encode_dv_formula(t, externsheet, names, addins))
            .unwrap_or_default();
        let f2 = formula2_text
            .map(|t| encode_dv_formula(t, externsheet, names, addins))
            .unwrap_or_default();

        let mut cf_body = Vec::with_capacity(12 + f1.len() + f2.len());
        cf_body.push(ct);
        cf_body.push(cp);
        cf_body.extend_from_slice(&(f1.len() as u16).to_le_bytes());
        cf_body.extend_from_slice(&(f2.len() as u16).to_le_bytes());
        // dxfn12 header (6 bytes): all-zero flags = "no formatting
        // override". Excel rejects CF records that omit this block
        // ("Open method of Workbooks class failed") even though our
        // permissive reader and LibreOffice both tolerate the
        // shorter shape.
        cf_body.extend_from_slice(&[0u8; 6]);
        cf_body.extend_from_slice(&f1);
        cf_body.extend_from_slice(&f2);

        stream.extend_from_slice(&CF_RECORD.to_le_bytes());
        stream.extend_from_slice(&(cf_body.len() as u16).to_le_bytes());
        stream.extend_from_slice(&cf_body);
    }
}

/// Emit a DVAL header (MS-XLS §2.4.81) plus one DV record (§2.4.79)
/// per `DataValidation` on the worksheet.
///
/// DVAL body (18 bytes):
///   options (u16)        - input-box state flags
///   xLeft, yTop (i32x2)  - input-box position
///   iDvIdInputBox (u32)  - the DV record currently in the input box
///   idv (u32)            - count of DV records that follow
///
/// DV body:
///   flags (u32)          - val_type (bits 0-3), err_style (4-6),
///                          fExplicit list (7), fAllowBlank (8),
///                          fSuppressDropdown (9), fShowInput (18),
///                          fShowError (19), operator (20-23)
///   input_title          - XLUnicodeString
///   error_title          - XLUnicodeString
///   input_msg            - XLUnicodeString
///   error_msg            - XLUnicodeString
///   cce1 (u16) + unused (u16) + formula1 ptgs
///   cce2 (u16) + unused (u16) + formula2 ptgs
///   range_count (u16) + N × Ref8U (8 bytes: r1/r2/c1/c2 u16)
fn write_data_validations(
    stream: &mut Vec<u8>,
    sheet: &Worksheet,
    externsheet: &ExternSheetTable,
    names: &NameTable,
    addins: &AddinTable,
) {
    use duke_sheets_core::validation::{ValidationErrorStyle, ValidationOperator, ValidationType};

    let validations = sheet.data_validations();
    if validations.is_empty() {
        return;
    }

    let mut dval_body = Vec::with_capacity(18);
    dval_body.extend_from_slice(&0u16.to_le_bytes()); // options
    dval_body.extend_from_slice(&0i32.to_le_bytes()); // xLeft
    dval_body.extend_from_slice(&0i32.to_le_bytes()); // yTop
    dval_body.extend_from_slice(&0u32.to_le_bytes()); // iDvIdInputBox
    dval_body.extend_from_slice(&(validations.len() as u32).to_le_bytes());
    stream.extend_from_slice(&DVAL_RECORD.to_le_bytes());
    stream.extend_from_slice(&(dval_body.len() as u16).to_le_bytes());
    stream.extend_from_slice(&dval_body);

    for v in validations {
        let val_type_bits: u32 = match &v.validation_type {
            ValidationType::None => 0,
            ValidationType::Whole { .. } => 1,
            ValidationType::Decimal { .. } => 2,
            ValidationType::List { .. } => 3,
            ValidationType::Date { .. } => 4,
            ValidationType::Time { .. } => 5,
            ValidationType::TextLength { .. } => 6,
            ValidationType::Custom { .. } => 7,
        };
        let err_style_bits: u32 = match v.error_style {
            ValidationErrorStyle::Stop => 0,
            ValidationErrorStyle::Warning => 1,
            ValidationErrorStyle::Information => 2,
        };
        let op_bits: u32 = match &v.validation_type {
            ValidationType::Whole { operator, .. }
            | ValidationType::Decimal { operator, .. }
            | ValidationType::Date { operator, .. }
            | ValidationType::Time { operator, .. }
            | ValidationType::TextLength { operator, .. } => match operator {
                ValidationOperator::Between => 0,
                ValidationOperator::NotBetween => 1,
                ValidationOperator::Equal => 2,
                ValidationOperator::NotEqual => 3,
                ValidationOperator::GreaterThan => 4,
                ValidationOperator::LessThan => 5,
                ValidationOperator::GreaterThanOrEqual => 6,
                ValidationOperator::LessThanOrEqual => 7,
            },
            _ => 0,
        };
        let is_explicit_list = matches!(
            &v.validation_type,
            ValidationType::List { source } if !source.starts_with('=')
        );

        let mut flags: u32 = val_type_bits | (err_style_bits << 4) | (op_bits << 20);
        if is_explicit_list {
            flags |= 0x0080;
        }
        if v.allow_blank {
            flags |= 0x0100;
        }
        if !v.show_dropdown {
            flags |= 0x0200;
        }
        if v.show_input_message {
            flags |= 0x0004_0000;
        }
        if v.show_error_alert {
            flags |= 0x0008_0000;
        }

        let mut body = Vec::new();
        body.extend_from_slice(&flags.to_le_bytes());

        // String headers (always emit, empty when None).
        for s in [
            v.input_title.as_deref(),
            v.error_title.as_deref(),
            v.input_message.as_deref(),
            v.error_message.as_deref(),
        ] {
            push_dv_unicode_string(&mut body, s.unwrap_or(""));
        }

        let (value1, value2) = match &v.validation_type {
            ValidationType::Whole { value1, value2, .. }
            | ValidationType::Decimal { value1, value2, .. }
            | ValidationType::Date { value1, value2, .. }
            | ValidationType::Time { value1, value2, .. }
            | ValidationType::TextLength { value1, value2, .. } => {
                (Some(value1.as_str()), value2.as_deref())
            }
            ValidationType::List { source } => (Some(source.as_str()), None),
            ValidationType::Custom { formula } => (Some(formula.as_str()), None),
            ValidationType::None => (None, None),
        };

        let formula1 = value1
            .map(|t| encode_dv_formula(t, externsheet, names, addins))
            .unwrap_or_default();
        body.extend_from_slice(&(formula1.len() as u16).to_le_bytes());
        body.extend_from_slice(&0u16.to_le_bytes()); // unused
        body.extend_from_slice(&formula1);

        let formula2 = value2
            .map(|t| encode_dv_formula(t, externsheet, names, addins))
            .unwrap_or_default();
        body.extend_from_slice(&(formula2.len() as u16).to_le_bytes());
        body.extend_from_slice(&0u16.to_le_bytes()); // unused
        body.extend_from_slice(&formula2);

        let usable: Vec<&duke_sheets_core::CellRange> = v
            .ranges
            .iter()
            .filter(|r| r.start.row <= u16::MAX as u32 && r.end.row <= u16::MAX as u32)
            .collect();
        body.extend_from_slice(&(usable.len() as u16).to_le_bytes());
        for r in usable {
            body.extend_from_slice(&(r.start.row as u16).to_le_bytes());
            body.extend_from_slice(&(r.end.row as u16).to_le_bytes());
            body.extend_from_slice(&r.start.col.to_le_bytes());
            body.extend_from_slice(&r.end.col.to_le_bytes());
        }

        stream.extend_from_slice(&DV_RECORD.to_le_bytes());
        stream.extend_from_slice(&(body.len() as u16).to_le_bytes());
        stream.extend_from_slice(&body);
    }
}

/// Emit an XLUnicodeString (cch u16 + flags u8 + chars) for DV input
/// titles, error titles, and message text. Empty strings emit cch=0
/// and a flags byte of 0; the reader treats them as None.
fn push_dv_unicode_string(buf: &mut Vec<u8>, s: &str) {
    let units: Vec<u16> = s.encode_utf16().collect();
    let cch = units.len().min(u16::MAX as usize) as u16;
    buf.extend_from_slice(&cch.to_le_bytes());
    let high_byte = units.iter().any(|&u| u > 0xFF);
    if high_byte {
        buf.push(0x01);
        for u in &units {
            buf.extend_from_slice(&u.to_le_bytes());
        }
    } else {
        buf.push(0x00);
        for u in &units {
            buf.push(*u as u8);
        }
    }
}

/// Encode a DataValidation value-string into a ptg formula body.
///
/// The model carries values as decompiled formula strings (e.g.
/// "100", "10.5", "Red,Green,Blue", "=A1+1"). Round-tripping requires
/// re-parsing as a formula and recompiling to ptgs:
///
///   - Numeric/string/cell-ref values reach the parser via a synthetic
///     `=` prefix. parse_formula handles `=100`, `=10.5`, `=A1`, etc.
///   - Custom validation formulas already start with `=`.
///   - Inline list sources like "Red,Green,Blue" don't parse as a
///     formula expression. Fall back to a single tStr ptg that
///     decompiles back to a quoted string; the reader strips the
///     surrounding quotes when the fExplicit-list flag is set.
fn encode_dv_formula(
    value: &str,
    externsheet: &ExternSheetTable,
    names: &NameTable,
    addins: &AddinTable,
) -> Vec<u8> {
    if value.is_empty() {
        return Vec::new();
    }
    let to_parse = if value.starts_with('=') {
        value.to_string()
    } else {
        format!("={value}")
    };
    if let Ok(expr) = duke_sheets_formula::parse_formula(&to_parse) {
        let mut bytes = Vec::new();
        let mut extra = Vec::new();
        if compile_ptgs_with_context(
            &expr,
            &mut bytes,
            &mut extra,
            externsheet,
            names,
            addins,
            OperandClass::V,
        )
        .is_ok()
        {
            bytes.extend_from_slice(&extra);
            return bytes;
        }
    }
    // Fallback: emit the raw value as a tStr ptg literal.
    let mut bytes = Vec::new();
    bytes.push(0x17); // PTG_STR
    let _ = push_short_xlunicode_string(&mut bytes, value);
    bytes
}

/// Encode a length-prefixed UTF-16LE string for HLINK record fields
/// (displayName, frameName, location). Format: char_count(u32,
/// includes null terminator) + UTF-16LE chars + 0x0000.
fn push_hlink_string(buf: &mut Vec<u8>, s: &str) {
    let units: Vec<u16> = s.encode_utf16().collect();
    let total_chars = (units.len() + 1) as u32; // include null
    buf.extend_from_slice(&total_chars.to_le_bytes());
    for u in &units {
        buf.extend_from_slice(&u.to_le_bytes());
    }
    buf.extend_from_slice(&0u16.to_le_bytes()); // null terminator
}

/// Encode the URL moniker payload following the URL_MONIKER_CLSID:
/// byte_len(u32, includes the null terminator's 2 bytes) + UTF-16LE
/// chars + 0x0000. Note: the URL moniker uses BYTE length, not char
/// count, unlike the regular HLINK strings.
fn push_url_moniker_payload(buf: &mut Vec<u8>, url: &str) {
    let units: Vec<u16> = url.encode_utf16().collect();
    let byte_len = ((units.len() + 1) * 2) as u32; // include null
    buf.extend_from_slice(&byte_len.to_le_bytes());
    for u in &units {
        buf.extend_from_slice(&u.to_le_bytes());
    }
    buf.extend_from_slice(&0u16.to_le_bytes()); // null terminator
}

/// Emit cell records (BLANK, NUMBER, BOOLERR, LABELSST, FORMULA) for
/// every non-empty cell in `sheet`, sorted in row-major order.
/// Spill-target cells are silently skipped (dynamic-array machinery;
/// deferred). Formula cells with an unsupported AST shape (named
/// ranges, structured refs, function calls, etc.) fall back to
/// emitting their cached value as a static cell.
///
/// `ixfe` is resolved via `styles.ixfe_for_cell` so cells with a
/// non-default style point at the appropriate user-defined XF.
fn write_cell_records(
    stream: &mut Vec<u8>,
    sheet: &Worksheet,
    sheet_idx: usize,
    sst: &SstTable,
    styles: &StyleTables,
    externsheet: &ExternSheetTable,
    names: &NameTable,
    addins: &AddinTable,
    first_cell_positions: &mut BTreeMap<u32, usize>,
) {
    let mut cells: Vec<_> = sheet.iter_cells().collect();
    cells.sort_by_key(|(row, col, _)| (*row, *col));

    for (row, col, data) in cells {
        if row > u16::MAX as u32 {
            continue;
        }
        let row16 = row as u16;
        let ixfe = styles.ixfe_for_cell(sheet_idx, data.style_index);

        if let Some(formula_text) = sheet.get_formula_at(row, col) {
            let record_pos = stream.len();
            if try_write_formula_record(
                stream,
                row16,
                col,
                ixfe,
                formula_text,
                &data.value,
                sst,
                externsheet,
                names,
                addins,
            ) {
                first_cell_positions.entry(row).or_insert(record_pos);
                continue;
            }
        }

        match &data.value {
            CellValue::Empty => {
                if ixfe != 0 {
                    first_cell_positions.entry(row).or_insert(stream.len());
                    write_blank(stream, row16, col, ixfe);
                }
            }
            CellValue::Number(v) => {
                first_cell_positions.entry(row).or_insert(stream.len());
                write_number(stream, row16, col, ixfe, *v);
            }
            CellValue::Boolean(b) => {
                first_cell_positions.entry(row).or_insert(stream.len());
                write_boolerr(stream, row16, col, ixfe, u8::from(*b), false);
            }
            CellValue::Error(err) => {
                first_cell_positions.entry(row).or_insert(stream.len());
                write_boolerr(stream, row16, col, ixfe, err.code(), true);
            }
            CellValue::String(s) => {
                if let Some(idx) = sst.lookup_plain(s.as_ref()) {
                    first_cell_positions.entry(row).or_insert(stream.len());
                    write_labelsst(stream, row16, col, ixfe, idx);
                }
            }
            CellValue::RichText(runs) => {
                let mut text = String::new();
                let mut formatting: Vec<(u16, u16)> = Vec::new();
                for run in runs.iter() {
                    let char_pos = text.encode_utf16().count().min(u16::MAX as usize) as u16;
                    if let Some(rf) = &run.font {
                        let font = run_font_to_font_style(rf);
                        if let Some(&font_idx) = styles.font_xf_index.get(&font) {
                            formatting.push((char_pos, font_idx));
                        }
                    }
                    text.push_str(&run.text);
                }
                if let Some(idx) = sst.lookup_rich(&text, &formatting) {
                    first_cell_positions.entry(row).or_insert(stream.len());
                    write_labelsst(stream, row16, col, ixfe, idx);
                }
            }
            CellValue::SpillTarget { .. } => {
                // Dynamic-array formula spill targets need formula
                // emission infrastructure; deferred.
            }
        }
    }
}

/// Try to compile and emit a FORMULA record for a cell. Returns true
/// on success, false if the formula AST contains constructs we don't
/// yet emit (named ranges, function calls, structured refs, external
/// refs, arrays); on false the caller falls back to emitting the
/// cell's cached value as a static record.
///
/// Slice 5a supports: numeric/string/bool/error literals, single cell
/// refs (relative + absolute), range refs, binary operators (+, -, *,
/// /, ^, &, comparisons), unary minus / unary plus / percent.
fn try_write_formula_record(
    stream: &mut Vec<u8>,
    row: u16,
    col: u16,
    ixfe: u16,
    formula_text: &str,
    cached: &CellValue,
    sst: &SstTable,
    externsheet: &ExternSheetTable,
    names: &NameTable,
    addins: &AddinTable,
) -> bool {
    // duke-sheets-formula's parse_formula requires the leading '=';
    // ensure it's present without double-prefixing.
    let with_eq_owned: String;
    let parse_input: &str = if formula_text.starts_with('=') {
        formula_text
    } else {
        with_eq_owned = format!("={formula_text}");
        &with_eq_owned
    };
    let Ok(expr) = duke_sheets_formula::parse_formula(parse_input) else {
        return false;
    };
    let mut tokens = Vec::with_capacity(32);
    // MS-XLS Rgce: a formula containing any volatile function (NOW, RAND,
    // TODAY, OFFSET, INDIRECT, etc.) is prefixed with PtgAttrVolatile so the
    // recalculation engine knows to re-evaluate this cell on every change.
    // Excel always emits this prefix; matching it preserves byte parity
    // through an Excel open/save round-trip.
    if expr_calls_volatile_function(&expr) {
        tokens.extend_from_slice(&[0x19, 0x01, 0x00, 0x00]); // PtgAttrVolatile
    }
    // `extra` accumulates the rgcb (array-constant element data etc.) that
    // follows the rgce in the FORMULA record. The cce field counts only the
    // rgce (tokens); rgcb bytes come after.
    let mut extra = Vec::new();
    if compile_ptgs_with_context(
        &expr,
        &mut tokens,
        &mut extra,
        externsheet,
        names,
        addins,
        OperandClass::V,
    )
    .is_err()
    {
        return false;
    }
    // cce is a u16; the whole formula body (22 + rgce + rgcb) must also fit
    // in the u16 record length.
    if tokens.len() > u16::MAX as usize || 22 + tokens.len() + extra.len() > u16::MAX as usize {
        return false;
    }

    let cached_bytes = encode_cached_result(cached);
    stream.extend_from_slice(&FORMULA_RECORD.to_le_bytes());
    let body_len: u16 = 22 + tokens.len() as u16 + extra.len() as u16;
    stream.extend_from_slice(&body_len.to_le_bytes());
    stream.extend_from_slice(&row.to_le_bytes());
    stream.extend_from_slice(&col.to_le_bytes());
    stream.extend_from_slice(&ixfe.to_le_bytes());
    stream.extend_from_slice(&cached_bytes);
    let grbit: u16 = 0x0002; // fAlwaysCalc cleared, fCalcOnLoad set: cause Excel to recompute on open
    stream.extend_from_slice(&grbit.to_le_bytes());
    stream.extend_from_slice(&0u32.to_le_bytes()); // chn (cache key)
    stream.extend_from_slice(&(tokens.len() as u16).to_le_bytes());
    stream.extend_from_slice(&tokens);
    stream.extend_from_slice(&extra);

    if let CellValue::String(s) = cached {
        write_string_followup(stream, s.as_ref());
    } else if let CellValue::RichText(runs) = cached {
        let plain: String = runs
            .iter()
            .map(|r| r.text.as_str())
            .collect::<Vec<_>>()
            .join("");
        write_string_followup(stream, &plain);
    }

    let _ = sst;
    true
}

fn write_string_followup(stream: &mut Vec<u8>, s: &str) {
    let mut body = Vec::with_capacity(3 + s.len() * 2);
    let units: Vec<u16> = s.encode_utf16().collect();
    let high_byte = units.iter().any(|&u| u > 0xFF);
    body.extend_from_slice(&(units.len() as u16).to_le_bytes());
    if high_byte {
        body.push(0x01);
        for u in &units {
            body.extend_from_slice(&u.to_le_bytes());
        }
    } else {
        body.push(0x00);
        for u in &units {
            body.push(*u as u8);
        }
    }
    stream.extend_from_slice(&STRING_RECORD.to_le_bytes());
    stream.extend_from_slice(&(body.len() as u16).to_le_bytes());
    stream.extend_from_slice(&body);
}

/// Encode an 8-byte FORMULA cached_result field (MS-XLS §2.5.133
/// FormulaValue). Numeric results: 8-byte f64 little-endian. Other
/// types use a sentinel encoding where bytes[6..8] = 0xFFFF and
/// bytes[0] selects the variant: 0=string (real value in STRING
/// follow-up), 1=bool, 2=error, 3=empty.
fn encode_cached_result(value: &CellValue) -> [u8; 8] {
    let mut out = [0u8; 8];
    match value {
        CellValue::Number(n) => {
            out.copy_from_slice(&n.to_le_bytes());
        }
        CellValue::String(_) | CellValue::RichText(_) => {
            out[0] = 0x00;
            out[6] = 0xFF;
            out[7] = 0xFF;
        }
        CellValue::Boolean(b) => {
            out[0] = 0x01;
            out[2] = if *b { 1 } else { 0 };
            out[6] = 0xFF;
            out[7] = 0xFF;
        }
        CellValue::Error(e) => {
            out[0] = 0x02;
            out[2] = e.code();
            out[6] = 0xFF;
            out[7] = 0xFF;
        }
        CellValue::Empty | CellValue::SpillTarget { .. } => {
            out[0] = 0x03;
            out[6] = 0xFF;
            out[7] = 0xFF;
        }
    }
    out
}

#[derive(Debug)]
struct UnsupportedToken;

// PTG operand class and per-function metadata live in duke-sheets-formula
// because they're shared between the XLS (BIFF8) and XLSB (BIFF12) writers.
// MS-XLS §2.5.198 defines the V/R/A class distinction; this writer only uses
// V and R. See `OperandClass` in the formula crate for the full rationale.
use duke_sheets_formula::decompile::function_table::{
    expr_calls_volatile_function, function_arg_class, function_index, function_is_biff8_addin,
    function_is_fixed_arity, function_returns_reference, name_body_operand_class, OperandClass,
};

fn compile_ptgs_with_context(
    expr: &duke_sheets_formula::FormulaExpr,
    out: &mut Vec<u8>,
    extra: &mut Vec<u8>,
    externsheet: &ExternSheetTable,
    names: &NameTable,
    addins: &AddinTable,
    operand_class: OperandClass,
) -> Result<(), UnsupportedToken> {
    use duke_sheets_formula::ast::{BinaryOperator, UnaryOperator};
    use duke_sheets_formula::FormulaExpr;

    match expr {
        FormulaExpr::Number(n) => {
            if let Some(i) = number_as_ptg_int(*n) {
                out.push(0x1E); // PTG_INT
                out.extend_from_slice(&i.to_le_bytes());
            } else {
                out.push(0x1F); // PTG_NUM
                out.extend_from_slice(&n.to_le_bytes());
            }
        }
        FormulaExpr::String(s) => {
            out.push(0x17); // PTG_STR
            push_short_xlunicode_string(out, s).map_err(|_| UnsupportedToken)?;
        }
        FormulaExpr::Boolean(b) => {
            out.push(0x1D); // PTG_BOOL
            out.push(if *b { 1 } else { 0 });
        }
        FormulaExpr::Error(e) => {
            out.push(0x1C); // PTG_ERR
            out.push(e.code());
        }
        FormulaExpr::CellRef(cref) => {
            if let Some(sheet_name) = cref.sheet.as_deref() {
                let ixti = externsheet
                    .ixti_for_sheet(sheet_name)
                    .ok_or(UnsupportedToken)?;
                // tRef3D V class = base 0x3A | V-class 0x20 = 0x5A.
                // R-class variant is 0x3A (base) | R-class 0x00 = 0x3A.
                // (0x5C is tRefErr3D - decompiles to "#REF!".)
                let opcode = match operand_class {
                    OperandClass::R => 0x3A,
                    OperandClass::V => 0x5A,
                };
                out.push(opcode);
                out.extend_from_slice(&ixti.to_le_bytes());
                push_ref_payload(out, &cref.address)?;
                return Ok(());
            }
            // PTG_REF base = 0x04; R-class = 0x24, V-class = 0x44.
            let opcode = match operand_class {
                OperandClass::R => 0x24,
                OperandClass::V => 0x44,
            };
            out.push(opcode);
            push_ref_payload(out, &cref.address)?;
        }
        FormulaExpr::RangeRef(rref) => {
            if let Some(sheet_name) = rref.sheet.as_deref() {
                let ixti = externsheet
                    .ixti_for_sheet(sheet_name)
                    .ok_or(UnsupportedToken)?;
                // 3D areas always use R-class — V-class causes
                // Excel to collapse the range to a single cell.
                out.push(0x3B); // PTG_AREA3D (R class)
                out.extend_from_slice(&ixti.to_le_bytes());
                push_area_payload(out, &rref.range)?;
                return Ok(());
            }
            // PTG_AREA base = 0x05; R-class = 0x25, V-class = 0x45.
            let opcode = match operand_class {
                OperandClass::R => 0x25,
                OperandClass::V => 0x45,
            };
            out.push(opcode);
            push_area_payload(out, &rref.range)?;
        }
        FormulaExpr::BinaryOp { op, left, right } => {
            // The `:` operator with NameRef/Number leaves on both sides
            // is how Excel-style full-column (`A:A`) and full-row
            // (`1:1`) refs reach us from the parser - it lacks a
            // first-class FullColumn / FullRow AST node. Detect those
            // shapes and emit a single tArea ptg covering the BIFF8
            // sheet extent.
            if matches!(op, BinaryOperator::Range) {
                if let Some(area) = full_column_or_row_range(left, right) {
                    let opcode = match operand_class {
                        OperandClass::R => 0x25,
                        OperandClass::V => 0x45,
                    };
                    out.push(opcode);
                    push_area_payload(out, &area)?;
                    return Ok(());
                }
            }
            // Intersection / union / range operators take reference
            // operands. All other binary operators (arithmetic, comparison,
            // concat) take values regardless of the surrounding context:
            // multiplying a cell ref always wants the cell's value, even
            // if the parent context was a reference position. Matches
            // MS-XLS §2.5.198 value-vs-reference operator rules.
            let child_class = match op {
                BinaryOperator::Intersect | BinaryOperator::Union | BinaryOperator::Range => {
                    OperandClass::R
                }
                _ => OperandClass::V,
            };
            compile_ptgs_with_context(left, out, extra, externsheet, names, addins, child_class)?;
            compile_ptgs_with_context(right, out, extra, externsheet, names, addins, child_class)?;
            out.push(match op {
                BinaryOperator::Add => 0x03,
                BinaryOperator::Subtract => 0x04,
                BinaryOperator::Multiply => 0x05,
                BinaryOperator::Divide => 0x06,
                BinaryOperator::Power => 0x07,
                BinaryOperator::Concat => 0x08,
                BinaryOperator::LessThan => 0x09,
                BinaryOperator::LessEqual => 0x0A,
                BinaryOperator::Equal => 0x0B,
                BinaryOperator::GreaterEqual => 0x0C,
                BinaryOperator::GreaterThan => 0x0D,
                BinaryOperator::NotEqual => 0x0E,
                // BIFF8 PTGs per MS-XLS §2.5.198.x:
                //   0x0F PtgIsect (intersection, space operator)
                //   0x10 PtgUnion (union, comma inside range parens)
                //   0x11 PtgRange (range, colon operator)
                BinaryOperator::Intersect => 0x0F,
                BinaryOperator::Union => 0x10,
                BinaryOperator::Range => 0x11,
            });
        }
        FormulaExpr::UnaryOp { op, operand } => {
            // Unary value operators force their operand to V-class even
            // inside R-forced argument positions: Excel emits tRefV for
            // =SUM(-A1) / =SUM(A1%) (verified byte-for-byte against
            // Excel-authored output). Paren is class-transparent.
            let inner_class = match op {
                UnaryOperator::Paren => operand_class,
                _ => OperandClass::V,
            };
            compile_ptgs_with_context(
                operand,
                out,
                extra,
                externsheet,
                names,
                addins,
                inner_class,
            )?;
            out.push(match op {
                UnaryOperator::Plus => 0x12,    // PtgUplus
                UnaryOperator::Negate => 0x13,  // PtgUminus
                UnaryOperator::Percent => 0x14, // PtgPercent
                UnaryOperator::Paren => 0x15,   // PtgParen
                UnaryOperator::ImplicitIntersection | UnaryOperator::SpillRange => {
                    return Err(UnsupportedToken);
                }
            });
        }
        FormulaExpr::Function { name, args } => {
            let Some(idx) = function_index(name) else {
                return Err(UnsupportedToken);
            };
            // BIFF8 PtgFuncVar cparams is 7-bit argc + the fPrompt bit
            // ([MS-XLS] §2.5.198.63): 127 arguments max. Beyond that,
            // fall back to the cached value — an argc with bit 7 set
            // reads back as a garbled call.
            if args.len() > 0x7F {
                return Err(UnsupportedToken);
            }
            // Analysis-ToolPak add-in functions (Ftab 384..=476) are not native
            // BIFF8 functions: Excel serializes them as an add-in UDF call —
            // a PtgNameX referencing an EXTERNNAME in the AddIn SUPBOOK,
            // followed by R-class (by-reference) arguments and a PtgFuncVar
            // with the UDF sentinel iftab 0x00FF whose argument count includes
            // the PtgNameX operand. See `write_supbook_and_externsheet`.
            //
            // Falls through to native emission when the name was not collected
            // into the AddinTable (e.g. a function used only in a data-
            // validation formula, which the pre-scan does not visit) or when
            // the argument count would overflow the PtgFuncVar byte.
            // The +1 for the PtgNameX operand must also fit in 7 bits.
            if function_is_biff8_addin(idx) && args.len() < 0x7F {
                if let Some(nameindex) = addins.nameindex_for(name) {
                    out.push(0x39); // PtgNameX
                    out.extend_from_slice(&0u16.to_le_bytes()); // ixti = 0 (AddIn XTI)
                    out.extend_from_slice(&nameindex.to_le_bytes()); // 1-based
                    out.extend_from_slice(&0u16.to_le_bytes()); // reserved
                    for arg in args {
                        if matches!(arg, FormulaExpr::Empty) {
                            out.push(0x16); // PTG_MISS_ARG
                        } else {
                            compile_ptgs_with_context(
                                arg,
                                out,
                                extra,
                                externsheet,
                                names,
                                addins,
                                OperandClass::R,
                            )?;
                        }
                    }
                    // PtgFuncVar argc counts the PtgNameX as the first operand.
                    out.push(ptg_func_var_opcode(func_token_class(idx, operand_class)));
                    out.push((args.len() + 1) as u8);
                    out.extend_from_slice(&0x00FFu16.to_le_bytes()); // iftab = UDF
                    return Ok(());
                }
            }
            if idx == 4 && args.len() == 1 && !matches!(args[0], FormulaExpr::Empty) {
                emit_optimized_sum(&args[0], out, extra, externsheet, names, addins)?;
                return Ok(());
            }
            // IF gets MS-XLS PtgAttrIf / PtgAttrGoto short-circuit optimization
            // ([MS-XLS] §2.5.198.39 / §2.5.198.37). Excel always emits this
            // for IF — matching it is required for byte-for-byte parity.
            if idx == 1 && (args.len() == 2 || args.len() == 3) {
                if emit_optimized_if(args, out, extra, externsheet, names, addins, operand_class)? {
                    return Ok(());
                }
                // emit_optimized_if returns Ok(false) when token offsets
                // would overflow u16; fall through to the plain emission.
            }
            // CHOOSE gets MS-XLS PtgAttrChoose jump-table optimization
            // ([MS-XLS] §2.5.198.40). Like IF, Excel always emits this so
            // matching is required for byte-for-byte parity.
            if idx == 100 && args.len() >= 2 {
                if emit_optimized_choose(
                    args,
                    out,
                    extra,
                    externsheet,
                    names,
                    addins,
                    operand_class,
                )? {
                    return Ok(());
                }
            }
            for (arg_idx, arg) in args.iter().enumerate() {
                if matches!(arg, FormulaExpr::Empty) {
                    out.push(0x16); // PTG_MISS_ARG
                } else {
                    let arg_class = function_arg_class(idx, arg_idx);
                    compile_ptgs_with_context(
                        arg,
                        out,
                        extra,
                        externsheet,
                        names,
                        addins,
                        arg_class,
                    )?;
                }
            }
            // PtgFunc (fixed-arity) or PtgFuncVar (variable-arity). The token's
            // operand-class bits come from `func_token_class`: reference-class
            // functions (IF/CHOOSE/etc.) take the surrounding context class,
            // value functions are always V. See `func_token_class`.
            let class = func_token_class(idx, operand_class);
            if function_is_fixed_arity(idx, args.len()) {
                out.push(ptg_func_opcode(class)); // PtgFunc
                out.extend_from_slice(&idx.to_le_bytes());
            } else {
                out.push(ptg_func_var_opcode(class)); // PtgFuncVar
                out.push(args.len() as u8);
                out.extend_from_slice(&idx.to_le_bytes());
            }
        }
        FormulaExpr::ExternalFunction { name, args, .. } => {
            if args.len() >= 0x7F {
                return Err(UnsupportedToken);
            }
            let Some(nameindex) = addins.nameindex_for(name) else {
                return Err(UnsupportedToken);
            };
            out.push(0x39); // PtgNameX
            out.extend_from_slice(&0u16.to_le_bytes()); // ixti = 0 (AddIn XTI)
            out.extend_from_slice(&nameindex.to_le_bytes()); // 1-based
            out.extend_from_slice(&0u16.to_le_bytes()); // reserved
            for arg in args {
                if matches!(arg, FormulaExpr::Empty) {
                    out.push(0x16); // PTG_MISS_ARG
                } else {
                    compile_ptgs_with_context(
                        arg,
                        out,
                        extra,
                        externsheet,
                        names,
                        addins,
                        OperandClass::R,
                    )?;
                }
            }
            out.push(ptg_func_var_opcode(OperandClass::V));
            out.push((args.len() + 1) as u8);
            out.extend_from_slice(&0x00FFu16.to_le_bytes());
        }
        FormulaExpr::NameRef(name) => {
            let idx = names.idx_for_name(name).ok_or(UnsupportedToken)?;
            let name_class = match operand_class {
                OperandClass::R => OperandClass::R,
                OperandClass::V => names.body_class_for_name(name).unwrap_or(OperandClass::V),
            };
            let opcode = match name_class {
                OperandClass::R => 0x23,
                OperandClass::V => 0x43,
            };
            out.push(opcode);
            out.extend_from_slice(&idx.to_le_bytes());
            out.extend_from_slice(&0u16.to_le_bytes()); // 2 reserved bytes
        }
        FormulaExpr::Array(rows) => {
            emit_array_constant(rows, out, extra)?;
        }
        FormulaExpr::StructuredRef(_) | FormulaExpr::ExternalRef(_) | FormulaExpr::Empty => {
            return Err(UnsupportedToken);
        }
    }
    Ok(())
}

/// Emit a BIFF8 array constant ([MS-XLS] §2.5.198.32 PtgArray + §2.5.8
/// SerAr). The token in `out` is an A-class PtgArray (0x60) followed by 7
/// reserved bytes; the actual element data goes into `extra` (the rgcb that
/// follows the rgce in the FORMULA record):
///
/// ```text
/// rgcb: ccol-1 (1 byte) | crow-1 (2 bytes) | elements (row-major)
/// element: 0x01 + f64                          (number)
///          0x02 + XLUnicodeString (cch + grbit + chars)  (string)
///          0x04 + bool(1) + 7 reserved          (bool)
///          0x10 + errcode(1) + 7 reserved       (error)
/// ```
fn emit_array_constant(
    rows: &[Vec<duke_sheets_formula::FormulaExpr>],
    out: &mut Vec<u8>,
    extra: &mut Vec<u8>,
) -> Result<(), UnsupportedToken> {
    use duke_sheets_formula::FormulaExpr;

    let nrows = rows.len();
    let ncols = rows.first().map_or(0, |r| r.len());
    // Excel arrays are rectangular and non-empty; reject anything else.
    if nrows == 0 || ncols == 0 || rows.iter().any(|r| r.len() != ncols) {
        return Err(UnsupportedToken);
    }
    if ncols > 256 || nrows > 65536 {
        return Err(UnsupportedToken);
    }

    // PtgArray token (A-class) + 7 reserved bytes, in the rgce.
    out.push(0x60);
    out.extend_from_slice(&[0u8; 7]);

    // rgcb: column count - 1 (1 byte), row count - 1 (2 bytes), then elements.
    extra.push((ncols - 1) as u8);
    extra.extend_from_slice(&((nrows - 1) as u16).to_le_bytes());
    for row in rows {
        for cell in row {
            match cell {
                FormulaExpr::Number(n) => {
                    extra.push(0x01);
                    extra.extend_from_slice(&n.to_le_bytes());
                }
                FormulaExpr::Boolean(b) => {
                    extra.push(0x04);
                    extra.push(if *b { 1 } else { 0 });
                    extra.extend_from_slice(&[0u8; 7]);
                }
                FormulaExpr::Error(e) => {
                    extra.push(0x10);
                    extra.push(e.code());
                    extra.extend_from_slice(&[0u8; 7]);
                }
                FormulaExpr::String(s) => {
                    extra.push(0x02);
                    push_short_xlunicode_string(extra, s).map_err(|_| UnsupportedToken)?;
                }
                // Nested arrays / refs / functions aren't valid array
                // constant elements.
                _ => return Err(UnsupportedToken),
            }
        }
    }
    Ok(())
}

fn emit_optimized_sum(
    arg: &duke_sheets_formula::FormulaExpr,
    out: &mut Vec<u8>,
    extra: &mut Vec<u8>,
    externsheet: &ExternSheetTable,
    names: &NameTable,
    addins: &AddinTable,
) -> Result<(), UnsupportedToken> {
    use duke_sheets_formula::ast::{BinaryOperator, UnaryOperator};
    use duke_sheets_formula::FormulaExpr;

    // =SUM((A1,B1)): the parser keeps the semantic parens as a Paren
    // node around the union. Excel's emission for this shape is
    // MemFunc + refs + PtgParen + AttrSum, so unwrap the Paren here
    // and re-emit it after the MemFunc block.
    let mem_arg = match arg {
        FormulaExpr::UnaryOp {
            op: UnaryOperator::Paren,
            operand,
        } if matches!(
            &**operand,
            FormulaExpr::BinaryOp {
                op: BinaryOperator::Union,
                ..
            }
        ) =>
        {
            &**operand
        }
        other => other,
    };

    if let FormulaExpr::BinaryOp { op, .. } = mem_arg {
        if matches!(
            op,
            BinaryOperator::Intersect | BinaryOperator::Union | BinaryOperator::Range
        ) {
            let mut ref_tokens = Vec::new();
            compile_ptgs_with_context(
                mem_arg,
                &mut ref_tokens,
                extra,
                externsheet,
                names,
                addins,
                OperandClass::R,
            )?;
            if ref_tokens.len() > u16::MAX as usize {
                return Err(UnsupportedToken);
            }
            out.push(0x29); // PTG_MEM_FUNC
            out.extend_from_slice(&(ref_tokens.len() as u16).to_le_bytes());
            out.extend_from_slice(&ref_tokens);
            if matches!(op, BinaryOperator::Union) {
                out.push(0x15); // PTG_PAREN, matching Excel's union-in-SUM emit
            }
            push_attr_sum(out);
            return Ok(());
        }
    }

    // Excel emits the PtgAttrSum form with the SUM arg in R-class context
    // when the arg is a reference; for non-ref args (numbers, value
    // expressions) the leaf token has no class bits so passing R is a
    // no-op there. Pull the class from the function metadata so the rule
    // stays in one place.
    let arg_class = function_arg_class(4, 0);
    compile_ptgs_with_context(arg, out, extra, externsheet, names, addins, arg_class)?;
    push_attr_sum(out);
    Ok(())
}

/// Effective operand class for a function's PtgFunc/PtgFuncVar token.
///
/// Reference-class functions (IF, CHOOSE, …) take the class of the position
/// they occupy, so a function used as a reference argument is emitted R-class
/// and one at a value position is V-class. Pure value functions are always
/// V-class regardless of context. MS-XLS [MS-XLS] §2.5.198.103.
fn func_token_class(iftab: u16, context: OperandClass) -> OperandClass {
    if function_returns_reference(iftab) {
        context
    } else {
        OperandClass::V
    }
}

/// PtgFunc opcode for the given class: R=0x21, V=0x41 (A=0x61 unused here).
fn ptg_func_opcode(class: OperandClass) -> u8 {
    match class {
        OperandClass::R => 0x21,
        OperandClass::V => 0x41,
    }
}

/// PtgFuncVar opcode for the given class: R=0x22, V=0x42 (A=0x62 unused here).
fn ptg_func_var_opcode(class: OperandClass) -> u8 {
    match class {
        OperandClass::R => 0x22,
        OperandClass::V => 0x42,
    }
}

/// Emit IF with Excel's MS-XLS short-circuit optimization tokens.
///
/// Excel always emits IF using [MS-XLS] §2.5.198.39 PtgAttrIf and
/// §2.5.198.37 PtgAttrGoto so only one branch is evaluated at runtime.
/// Matching the layout byte-for-byte is required for parity through
/// Excel open/save round-trips.
///
/// `operand_class` is the class of the position the whole IF occupies;
/// IF is a reference-class function so the trailing PtgFuncVar takes that
/// class (R when the IF is itself a reference argument, e.g. `SUM(IF(...))`).
///
/// For `IF(cond, t, f)` (3-arg form):
/// ```text
/// cond
/// PtgAttrIf  [offset = t_size + 4]    skip to f-branch if cond is FALSE
/// t_branch
/// PtgAttrSkip [offset = f_size + 7]   skip past f-branch + trailing skip + 3
/// f_branch
/// PtgAttrSkip [offset = 3]            trailing marker
/// PtgFuncVar(IF, argc=3) V-class
/// ```
///
/// For `IF(cond, t)` (2-arg form):
/// ```text
/// cond
/// PtgAttrIf  [offset = t_size + 4]    skip to PtgFuncVar if cond is FALSE
/// t_branch
/// PtgAttrSkip [offset = 3]            trailing marker
/// PtgFuncVar(IF, argc=2) V-class
/// ```
///
/// The trailing PtgAttrSkip offset of 3 (one less than the 4-byte
/// PtgFuncVar) appears in every Excel-authored IF; the spec is silent
/// on what it points at but the constant matches all observed cases.
///
/// Returns `Ok(false)` when branch byte sizes would overflow the u16
/// offset field; the caller then falls back to the plain
/// PtgFuncVar(IF) emission.
fn emit_optimized_if(
    args: &[duke_sheets_formula::FormulaExpr],
    out: &mut Vec<u8>,
    extra: &mut Vec<u8>,
    externsheet: &ExternSheetTable,
    names: &NameTable,
    addins: &AddinTable,
    operand_class: OperandClass,
) -> Result<bool, UnsupportedToken> {
    use duke_sheets_formula::FormulaExpr;

    // Reject if any branch is Empty (e.g. `=IF(,1,2)`) — Excel handles
    // the omitted form with a different token shape. Fall back to plain
    // PtgFuncVar emission, which the caller will do.
    if args.iter().any(|a| matches!(a, FormulaExpr::Empty)) {
        return Ok(false);
    }

    let cond = &args[0];
    let t_branch = &args[1];
    let f_branch = args.get(2);
    let argc = args.len() as u8;

    // Compile every part into scratch buffers FIRST so we can measure byte
    // counts and bail on overflow WITHOUT having mutated `out`. Writing
    // `cond` into `out` before the overflow check would leave a partial
    // token stream that the caller's fallback would then duplicate.
    // rgcb (array-constant data) goes into a scratch too: leaking it
    // into the real `extra` on an Ok(false) fallback would duplicate
    // it when the caller re-emits the args, shifting every later
    // PtgArray offset.
    let mut scratch_extra = Vec::new();
    let cond_class = function_arg_class(1, 0);
    let mut cond_bytes = Vec::with_capacity(16);
    compile_ptgs_with_context(
        cond,
        &mut cond_bytes,
        &mut scratch_extra,
        externsheet,
        names,
        addins,
        cond_class,
    )?;

    let mut t_bytes = Vec::with_capacity(32);
    let t_class = function_arg_class(1, 1);
    compile_ptgs_with_context(
        t_branch,
        &mut t_bytes,
        &mut scratch_extra,
        externsheet,
        names,
        addins,
        t_class,
    )?;

    let mut f_bytes = Vec::new();
    if let Some(f) = f_branch {
        let f_class = function_arg_class(1, 2);
        compile_ptgs_with_context(
            f,
            &mut f_bytes,
            &mut scratch_extra,
            externsheet,
            names,
            addins,
            f_class,
        )?;
    }

    // PtgAttrIf offset = bytes after PtgAttrIf to skip when cond is FALSE.
    // Lands at start of f-branch (3-arg) or PtgFuncVar (2-arg). The 4 is
    // the size of the PtgAttrSkip that follows t_branch.
    let attr_if_offset = t_bytes.len().checked_add(4).ok_or(UnsupportedToken)?;
    if attr_if_offset > u16::MAX as usize {
        return Ok(false);
    }
    // First PtgAttrSkip offset (3-arg only): jump past f_branch + trailing
    // PtgAttrSkip + 3.
    let skip_after_t = if f_branch.is_some() {
        let s = f_bytes.len().checked_add(7).ok_or(UnsupportedToken)?;
        if s > u16::MAX as usize {
            return Ok(false);
        }
        Some(s as u16)
    } else {
        None
    };

    // All overflow checks passed — commit to `out` and `extra`.
    extra.extend_from_slice(&scratch_extra);
    out.extend_from_slice(&cond_bytes);
    out.push(0x19);
    out.push(0x02); // ATTR_IF
    out.extend_from_slice(&(attr_if_offset as u16).to_le_bytes());
    out.extend_from_slice(&t_bytes);

    if let Some(skip) = skip_after_t {
        // The 3 baked into the trailing PtgAttrSkip is stable across all
        // observed Excel-authored IF formulas (the spec does not describe
        // what byte it points at; we replicate Excel's exact emission).
        out.push(0x19);
        out.push(0x08); // ATTR_SKIP
        out.extend_from_slice(&skip.to_le_bytes());
        out.extend_from_slice(&f_bytes);
    }

    // Trailing PtgAttrSkip with constant offset 3.
    out.push(0x19);
    out.push(0x08);
    out.extend_from_slice(&3u16.to_le_bytes());

    // PtgFuncVar for IF (iftab=1). IF is reference-class, so its token takes
    // the class of the position it fills.
    let class = func_token_class(1, operand_class);
    out.push(ptg_func_var_opcode(class));
    out.push(argc);
    out.extend_from_slice(&1u16.to_le_bytes());

    Ok(true)
}

/// Emit CHOOSE with Excel's MS-XLS short-circuit jump-table optimization.
///
/// [MS-XLS] §2.5.198.40 PtgAttrChoose carries a fixed-shape `nc`
/// (number of choices) and an array of `nc+1` u16 jump offsets. The
/// k-th offset (0..nc) is the byte distance from the start of the
/// offset table to the start of the k-th choice's token bytes. The
/// final entry (index `nc`) points to PtgFuncVar — used when the
/// selector is out of range so no choice is executed.
///
/// Layout for `CHOOSE(selector, c0, c1, ..., c_{nc-1})`:
///
/// ```text
/// selector
/// PtgAttrChoose [nc, off_0, off_1, ..., off_nc]   (4 + (nc+1)*2 bytes)
/// c0
/// PtgAttrSkip [offset = points to last byte]      (4 bytes)
/// c1
/// PtgAttrSkip [offset = points to last byte]      (4 bytes)
/// ...
/// c_{nc-1}
/// PtgAttrSkip [offset = 3]                        (4 bytes, trailing)
/// PtgFuncVar(CHOOSE, argc=nc+1) V-class
/// ```
///
/// Each post-choice PtgAttrSkip jumps to the last byte of the formula
/// (matching IF). Its offset is the sum of remaining `(choice_size + 4)`
/// terms plus the trailing 3 — the trailing PtgAttrSkip itself has
/// offset 3.
///
/// Returns `Ok(false)` when offsets would overflow u16; the caller
/// then falls back to the plain PtgFuncVar(CHOOSE) emission.
fn emit_optimized_choose(
    args: &[duke_sheets_formula::FormulaExpr],
    out: &mut Vec<u8>,
    extra: &mut Vec<u8>,
    externsheet: &ExternSheetTable,
    names: &NameTable,
    addins: &AddinTable,
    operand_class: OperandClass,
) -> Result<bool, UnsupportedToken> {
    use duke_sheets_formula::FormulaExpr;

    if args.len() < 2 || args.iter().any(|a| matches!(a, FormulaExpr::Empty)) {
        return Ok(false);
    }

    let selector = &args[0];
    let choices = &args[1..];
    let nc = choices.len();
    if nc > u16::MAX as usize {
        return Ok(false);
    }
    let argc = args.len() as u8;

    // Compile selector into scratch (V class for CHOOSE arg 0). Everything
    // is compiled into scratch buffers BEFORE any byte reaches `out`, so an
    // overflow bail (Ok(false)) leaves `out` untouched for the caller's
    // fallback to re-emit cleanly.
    let mut scratch_extra = Vec::new();
    let selector_class = function_arg_class(100, 0);
    let mut selector_bytes = Vec::with_capacity(16);
    compile_ptgs_with_context(
        selector,
        &mut selector_bytes,
        &mut scratch_extra,
        externsheet,
        names,
        addins,
        selector_class,
    )?;

    // Compile each choice into scratch.
    let mut choice_bytes: Vec<Vec<u8>> = Vec::with_capacity(nc);
    for (i, c) in choices.iter().enumerate() {
        let mut buf = Vec::with_capacity(16);
        let class = function_arg_class(100, i + 1);
        compile_ptgs_with_context(
            c,
            &mut buf,
            &mut scratch_extra,
            externsheet,
            names,
            addins,
            class,
        )?;
        choice_bytes.push(buf);
    }

    // Jump-table offsets. The k-th offset points to the start of choice k
    // (or PtgFuncVar for k = nc). All are measured from the start of the
    // offset table itself, which sits immediately after `19 04 nc_lo nc_hi`.
    let table_size = (nc + 1).checked_mul(2).ok_or(UnsupportedToken)?;
    let mut offsets: Vec<u16> = Vec::with_capacity(nc + 1);
    let mut running: usize = table_size;
    for choice in &choice_bytes {
        if running > u16::MAX as usize {
            return Ok(false);
        }
        offsets.push(running as u16);
        running = running
            .checked_add(choice.len())
            .and_then(|x| x.checked_add(4)) // PtgAttrSkip after this choice
            .ok_or(UnsupportedToken)?;
    }
    if running > u16::MAX as usize {
        return Ok(false);
    }
    offsets.push(running as u16); // final/exit entry

    // Compute each post-choice PtgAttrSkip offset. The trailing skip uses
    // offset 3 (matching IF); middle skips accumulate the remaining
    // (choice + skip) sizes plus 3.
    //
    // For the k-th PtgAttrSkip (0-indexed, after choice k):
    //   offset = sum_{j>k}(choice_sizes[j] + 4) + 3
    let mut remaining_after: usize = 0;
    let mut skip_offsets: Vec<u16> = vec![0; nc];
    for k in (0..nc).rev() {
        if k + 1 == nc {
            skip_offsets[k] = 3;
        } else {
            // Bytes from end-of-skip-K to last-byte:
            //   = choice_{k+1}_size + 4 (skip_{k+1}) + remaining_after_{k+1}
            let nxt = choice_bytes[k + 1]
                .len()
                .checked_add(4)
                .and_then(|x| x.checked_add(remaining_after))
                .ok_or(UnsupportedToken)?;
            if nxt + 3 > u16::MAX as usize {
                return Ok(false);
            }
            skip_offsets[k] = (nxt + 3) as u16;
            remaining_after = nxt;
        }
    }

    // All overflow checks passed — commit to `out` and `extra`.
    extra.extend_from_slice(&scratch_extra);
    out.extend_from_slice(&selector_bytes);
    out.push(0x19);
    out.push(0x04); // ATTR_CHOOSE
    out.extend_from_slice(&(nc as u16).to_le_bytes());
    for off in &offsets {
        out.extend_from_slice(&off.to_le_bytes());
    }
    for (k, choice) in choice_bytes.iter().enumerate() {
        out.extend_from_slice(choice);
        out.push(0x19);
        out.push(0x08); // ATTR_SKIP
        out.extend_from_slice(&skip_offsets[k].to_le_bytes());
    }

    // PtgFuncVar for CHOOSE (iftab=100). CHOOSE is reference-class, so its
    // token takes the class of the position it fills.
    let class = func_token_class(100, operand_class);
    out.push(ptg_func_var_opcode(class));
    out.push(argc);
    out.extend_from_slice(&100u16.to_le_bytes());

    Ok(true)
}

fn push_attr_sum(out: &mut Vec<u8>) {
    // PtgAttrSum (MS-XLS Rgce: `expression PtgAttrSum`) is Excel's
    // canonical single-argument SUM encoding. The final two bytes are
    // reserved/ignored for this subtype; Excel-authored files may contain
    // non-deterministic values there, so tests normalize them for compare.
    out.extend_from_slice(&[0x19, 0x10, 0x00, 0x00]);
}

fn number_as_ptg_int(n: f64) -> Option<u16> {
    if n.is_finite() && n.fract() == 0.0 && (0.0..=u16::MAX as f64).contains(&n) {
        Some(n as u16)
    } else {
        None
    }
}

fn push_ref_payload(
    out: &mut Vec<u8>,
    addr: &duke_sheets_core::CellAddress,
) -> Result<(), UnsupportedToken> {
    if addr.row > u16::MAX as u32 {
        return Err(UnsupportedToken);
    }
    out.extend_from_slice(&(addr.row as u16).to_le_bytes());
    out.extend_from_slice(
        &encode_col_with_relative_flags(addr.col, addr.row_absolute, addr.col_absolute)
            .to_le_bytes(),
    );
    Ok(())
}

/// Recognise full-column (`A:A`, `B:D`) and full-row (`1:1`, `2:5`)
/// reference shapes that the formula parser leaves as
/// `Range(NameRef, NameRef)` or `Range(Number, Number)` respectively.
/// Returns the equivalent `CellRange` covering the BIFF8 sheet
/// extent, or `None` when the operands aren't a recognised shape.
fn full_column_or_row_range(
    left: &duke_sheets_formula::FormulaExpr,
    right: &duke_sheets_formula::FormulaExpr,
) -> Option<duke_sheets_core::CellRange> {
    use duke_sheets_core::CellAddress;
    use duke_sheets_formula::FormulaExpr;

    if let (FormulaExpr::NameRef(l), FormulaExpr::NameRef(r)) = (left, right) {
        let start_col = CellAddress::letters_to_column(l).ok()?;
        let end_col = CellAddress::letters_to_column(r).ok()?;
        return Some(duke_sheets_core::CellRange {
            start: CellAddress::new(0, start_col.min(end_col)),
            end: CellAddress::new(u16::MAX as u32, start_col.max(end_col)),
        });
    }
    if let (FormulaExpr::Number(l), FormulaExpr::Number(r)) = (left, right) {
        if l.fract() != 0.0 || r.fract() != 0.0 || *l < 1.0 || *r < 1.0 {
            return None;
        }
        let l_idx = (*l as u32).saturating_sub(1);
        let r_idx = (*r as u32).saturating_sub(1);
        if l_idx > u16::MAX as u32 || r_idx > u16::MAX as u32 {
            return None;
        }
        return Some(duke_sheets_core::CellRange {
            start: CellAddress::new(l_idx.min(r_idx), 0),
            end: CellAddress::new(l_idx.max(r_idx), 0xFF),
        });
    }
    None
}

fn push_area_payload(
    out: &mut Vec<u8>,
    range: &duke_sheets_core::CellRange,
) -> Result<(), UnsupportedToken> {
    let start = &range.start;
    let end = &range.end;
    if start.row > u16::MAX as u32 {
        return Err(UnsupportedToken);
    }
    // Clamp end.row to BIFF8's row limit. The XLSX-style parser
    // produces end.row = 1048575 for full-column refs like A:A;
    // BIFF8 can't represent rows beyond 65535, so the closest valid
    // expression is "from start.row through row 65535" - i.e. the
    // entire BIFF8 column (or the entire BIFF8 row, for 1:1 refs).
    let end_row = end.row.min(u16::MAX as u32) as u16;
    out.extend_from_slice(&(start.row as u16).to_le_bytes());
    out.extend_from_slice(&end_row.to_le_bytes());
    out.extend_from_slice(
        &encode_col_with_relative_flags(start.col, start.row_absolute, start.col_absolute)
            .to_le_bytes(),
    );
    out.extend_from_slice(
        &encode_col_with_relative_flags(end.col, end.row_absolute, end.col_absolute).to_le_bytes(),
    );
    Ok(())
}

/// Pack column index + row/col absolute flags into the 16-bit
/// `colIxv` field used by tRef/tArea (MS-XLS §2.5.198.103). Bits 0-13
/// hold the column; bit 14 is fColRel; bit 15 is fRowRel.
fn encode_col_with_relative_flags(col: u16, row_absolute: bool, col_absolute: bool) -> u16 {
    let mut v = col & 0x3FFF;
    if !col_absolute {
        v |= 0x4000;
    }
    if !row_absolute {
        v |= 0x8000;
    }
    v
}

fn write_labelsst(stream: &mut Vec<u8>, row: u16, col: u16, ixfe: u16, isst: u32) {
    stream.extend_from_slice(&LABELSST_RECORD.to_le_bytes());
    stream.extend_from_slice(&10u16.to_le_bytes());
    stream.extend_from_slice(&row.to_le_bytes());
    stream.extend_from_slice(&col.to_le_bytes());
    stream.extend_from_slice(&ixfe.to_le_bytes());
    stream.extend_from_slice(&isst.to_le_bytes());
}

fn write_blank(stream: &mut Vec<u8>, row: u16, col: u16, ixfe: u16) {
    stream.extend_from_slice(&BLANK_RECORD.to_le_bytes());
    stream.extend_from_slice(&6u16.to_le_bytes());
    stream.extend_from_slice(&row.to_le_bytes());
    stream.extend_from_slice(&col.to_le_bytes());
    stream.extend_from_slice(&ixfe.to_le_bytes());
}

fn write_number(stream: &mut Vec<u8>, row: u16, col: u16, ixfe: u16, value: f64) {
    stream.extend_from_slice(&NUMBER_RECORD.to_le_bytes());
    stream.extend_from_slice(&14u16.to_le_bytes());
    stream.extend_from_slice(&row.to_le_bytes());
    stream.extend_from_slice(&col.to_le_bytes());
    stream.extend_from_slice(&ixfe.to_le_bytes());
    stream.extend_from_slice(&value.to_le_bytes());
}

fn write_boolerr(
    stream: &mut Vec<u8>,
    row: u16,
    col: u16,
    ixfe: u16,
    bool_or_err: u8,
    is_error: bool,
) {
    stream.extend_from_slice(&BOOLERR_RECORD.to_le_bytes());
    stream.extend_from_slice(&8u16.to_le_bytes());
    stream.extend_from_slice(&row.to_le_bytes());
    stream.extend_from_slice(&col.to_le_bytes());
    stream.extend_from_slice(&ixfe.to_le_bytes());
    stream.push(bool_or_err);
    stream.push(if is_error { 1 } else { 0 });
}

fn write_bof(stream: &mut Vec<u8>, dt: u16) {
    stream.extend_from_slice(&BOF_RECORD.to_le_bytes());
    stream.extend_from_slice(&16u16.to_le_bytes());
    stream.extend_from_slice(&BIFF8_VERSION.to_le_bytes());
    stream.extend_from_slice(&dt.to_le_bytes());
    stream.extend_from_slice(&BOF_RUP_BUILD.to_le_bytes());
    stream.extend_from_slice(&BOF_RUP_YEAR.to_le_bytes());
    stream.extend_from_slice(&BOF_BFH.to_le_bytes());
    stream.extend_from_slice(&BOF_SFH.to_le_bytes());
}

fn write_eof(stream: &mut Vec<u8>) {
    stream.extend_from_slice(&EOF_RECORD.to_le_bytes());
    stream.extend_from_slice(&0u16.to_le_bytes());
}

/// Emit a BoundSheet8 record with `lbPlyPos` zeroed. The caller backfills
/// the field after the corresponding worksheet stream's BOF position is
/// known.
///
/// Body layout per MS-XLS §2.4.28:
/// ```text
///  u32 lbPlyPos     ← zeroed; caller fixes up
///  u8  hsState      = visibility (0=visible, 1=hidden, 2=very hidden)
///  u8  dt           = 0 (worksheet)
///  ShortXLUnicodeString stName
/// ```
fn write_boundsheet8_with_placeholder_offset(
    stream: &mut Vec<u8>,
    sheet: &Worksheet,
) -> XlsResult<()> {
    let name = sheet.name();
    let utf16_units: Vec<u16> = name.encode_utf16().collect();
    if utf16_units.len() > 31 {
        return Err(XlsError::InvalidFormat(format!(
            "sheet name '{name}' is {} UTF-16 code units; Excel caps sheet names at 31",
            utf16_units.len()
        )));
    }

    let hs_state: u8 = match sheet.visibility() {
        duke_sheets_core::worksheet::SheetVisibility::Visible => 0,
        duke_sheets_core::worksheet::SheetVisibility::Hidden => 1,
        duke_sheets_core::worksheet::SheetVisibility::VeryHidden => 2,
    };

    let mut body = Vec::with_capacity(8 + utf16_units.len() * 2);
    body.extend_from_slice(&[0u8; 4]); // lbPlyPos placeholder
    body.push(hs_state);
    body.push(0); // dt = worksheet
    body.push(utf16_units.len() as u8); // cch
    body.push(1); // fHighByte = 1 (UTF-16LE)
    for unit in utf16_units {
        body.extend_from_slice(&unit.to_le_bytes());
    }

    stream.extend_from_slice(&BOUND_SHEET_8.to_le_bytes());
    stream.extend_from_slice(&(body.len() as u16).to_le_bytes());
    stream.extend_from_slice(&body);
    Ok(())
}

/// First shape ID used by the patriarch of the first drawing. MS-ODRAW
/// reserves shape IDs in clusters of 1024; cluster 0 (IDs 0–1023) is
/// reserved by the spec, so user-allocated shape IDs start at 1024.
const PATRIARCH_SPID_BASE: u32 = 1024;

/// Pre-computed drawing layout for a workbook. Captures which sheets
/// have shapes (comments and/or pictures), what shape IDs have been
/// allocated to their patriarchs, pictures, and comments, plus the
/// global blip store and cluster table that go inside the
/// `MSODRAWINGGROUP`.
#[derive(Debug, Default)]
struct DrawingState {
    /// Per-sheet drawings keyed by sheet index. Sheets without shapes
    /// are absent.
    sheets: HashMap<usize, SheetDrawing>,
    /// Drawing indices in workbook order; lets us emit cluster entries
    /// and sheet drawings deterministically.
    ordered_sheet_indices: Vec<usize>,
    /// Workbook-wide blip store. One entry per embedded image across
    /// all sheets, in deterministic order. The 1-based index into
    /// this vec is the `pib` (picture blip id) referenced by each
    /// picture shape's `FOPT` `0x0104` entry.
    blip_store: Vec<BlipEntry>,
    /// Total shape count across all drawings (patriarch + pictures +
    /// comments per sheet, summed). Goes into `FDGG.csp_saved`.
    csp_total: u32,
    /// Number of drawings (sheets with shapes). Goes into
    /// `FDGG.cdg_saved`.
    cdg_total: u32,
    /// One past the highest shape ID used across the workbook. Goes
    /// into `FDGG.spid_max`.
    spid_max: u32,
}

impl DrawingState {
    fn is_empty(&self) -> bool {
        self.cdg_total == 0
    }
}

/// A single image queued for emission in the workbook-globals blip
/// store. Index in `DrawingState.blip_store` (+1) is the `pib` the
/// picture shape's FOPT references.
#[derive(Debug, Clone)]
struct BlipEntry {
    format: duke_sheets_chart::ImageFormat,
    data: Vec<u8>,
}

#[derive(Debug, Clone)]
struct SheetDrawing {
    /// 1-based drawing ID (`OfficeArtFDG.dgid`).
    dgid: u16,
    /// Shape ID of the patriarch group (the implicit root group).
    patriarch_spid: u32,
    /// One per picture on this sheet, in stable iteration order
    /// (same order as `Worksheet::images()`).
    pictures: Vec<PictureShape>,
    /// One per comment, in stable iteration order (sorted by row,
    /// then column).
    comments: Vec<CommentShape>,
}

#[derive(Debug, Clone)]
struct PictureShape {
    /// Escher shape ID stored in this picture's `OfficeArtFSP.spid`.
    spid: u32,
    /// 1-based per-sheet object ID placed in `OBJ.ftCmo.id`.
    obj_id: u16,
    /// 1-based index into `DrawingState.blip_store`; referenced by
    /// the picture's FOPT `pib` (`0x0104`) property.
    blip_id: u32,
    /// User-visible shape name (e.g. `"Picture 1"`).
    shape_name: String,
    /// Cell-anchor footprint copied from the `EmbeddedImage`.
    anchor: duke_sheets_chart::DrawingAnchor,
    /// Optional rotation in 60,000ths of a degree. Goes into the
    /// picture's FOPT `0x0004` property when set.
    rotation: Option<i32>,
    /// Horizontal flip — sets the FSP `FLIP_H` flag bit.
    flip_h: bool,
    /// Vertical flip — sets the FSP `FLIP_V` flag bit.
    flip_v: bool,
}

#[derive(Debug, Clone)]
struct CommentShape {
    /// Escher shape ID stored in this comment's `OfficeArtFSP.spid`.
    spid: u32,
    /// 1-based per-sheet object ID. Excel's `OBJ.ftCmo.id` and
    /// `NOTE.objId` use this value to link a comment's drawing
    /// object to its cell-anchored note.
    obj_id: u16,
    /// `txid` value placed in the `FOPT` `TEXT_ID` slot. Chosen as a
    /// monotonically-increasing per-comment counter; Excel itself
    /// picks an arbitrary unique value per shape.
    text_id: u32,
    /// Cell row this comment is anchored to.
    row: u32,
    /// Cell column this comment is anchored to.
    col: u16,
    /// Author string from `CellComment.author`.
    author: String,
    /// Comment body text from `CellComment.text`.
    text: String,
    /// Whether the comment box is visible by default (sets `NOTE.flags`
    /// bit 1).
    visible: bool,
}

/// Walk every sheet in `workbook`, allocate drawing IDs and shape
/// IDs, and assemble the [`DrawingState`] used by both
/// [`write_msodrawinggroup`] and [`write_sheet_drawing_records`].
///
/// Within each drawing, shape IDs are allocated in order:
///   patriarch → pictures (in `Worksheet::images()` order) → comments
///   (in sorted row/col order). Each drawing starts in its own
///   1024-aligned cluster.
fn compute_drawing_state(workbook: &Workbook) -> DrawingState {
    let mut state = DrawingState::default();
    let mut next_dgid: u16 = 1;
    let mut next_spid = PATRIARCH_SPID_BASE;
    let mut next_text_id: u32 = 1;
    let mut next_blip_id: u32 = 1;
    let mut highest_spid_used: u32 = 0;

    for (sheet_idx, sheet) in workbook.worksheets().enumerate() {
        let comment_count = sheet.comment_count();
        let picture_count = sheet.image_count();
        if comment_count == 0 && picture_count == 0 {
            continue;
        }

        let patriarch_spid = next_spid;
        next_spid += 1;

        let mut next_obj_id: u16 = 1;
        let mut pictures = Vec::with_capacity(picture_count);
        for image in sheet.images() {
            let spid = next_spid;
            next_spid += 1;
            let blip_id = next_blip_id;
            next_blip_id += 1;
            state.blip_store.push(BlipEntry {
                format: image.format,
                data: image.data.clone(),
            });
            pictures.push(PictureShape {
                spid,
                obj_id: next_obj_id,
                blip_id,
                shape_name: image.name.clone(),
                anchor: image.anchor.clone(),
                rotation: image.rotation,
                flip_h: image.flip_h,
                flip_v: image.flip_v,
            });
            next_obj_id += 1;
        }

        let mut comments_sorted: Vec<_> = sheet.comments().collect();
        comments_sorted.sort_by_key(|((row, col), _)| (*row, *col));

        let mut comments = Vec::with_capacity(comment_count);
        for ((row, col), comment) in comments_sorted.iter() {
            let spid = next_spid;
            next_spid += 1;
            comments.push(CommentShape {
                spid,
                obj_id: next_obj_id,
                text_id: next_text_id,
                row: *row,
                col: *col,
                author: comment.author.clone(),
                text: comment.text.clone(),
                visible: comment.visible,
            });
            next_obj_id += 1;
            next_text_id = next_text_id.wrapping_add(1);
        }
        highest_spid_used = highest_spid_used.max(next_spid - 1);

        let total_shapes = 1 + picture_count + comment_count;
        state.sheets.insert(
            sheet_idx,
            SheetDrawing {
                dgid: next_dgid,
                patriarch_spid,
                pictures,
                comments,
            },
        );
        state.ordered_sheet_indices.push(sheet_idx);
        state.cdg_total += 1;
        state.csp_total += total_shapes as u32;
        next_dgid += 1;

        // Round up to the next 1024-aligned base so the next drawing
        // lives in its own cluster (matching Excel's per-drawing
        // cluster allocation in MS-ODRAW §2.2.46).
        next_spid = next_spid.div_ceil(1024) * 1024;
    }
    state.spid_max = highest_spid_used + 1;
    state
}

/// Emit the workbook-globals `MSODRAWINGGROUP` (BIFF 0xEB) record
/// carrying the global Office Art drawing context: `FDGG` shape-ID
/// allocator state, default-property `FOPT`, and the
/// `SplitMenuColorContainer` palette. No-op if no sheet has comments.
fn write_msodrawinggroup(stream: &mut Vec<u8>, state: &DrawingState) {
    if state.is_empty() {
        return;
    }

    use crate::biff::escher::{
        dgg_default_fopt, rec_type as er, write_bstore_container, write_container, IdCluster,
        OfficeArtBlip, OfficeArtFbse, OfficeArtFdgg, SplitMenuColors,
    };

    let mut dgg_body = Vec::new();
    let clusters: Vec<IdCluster> = state
        .ordered_sheet_indices
        .iter()
        .filter_map(|i| state.sheets.get(i))
        .map(|sd| IdCluster {
            dgid: sd.dgid as u32,
            // patriarch + pictures + comments.
            cspid_cur: (1 + sd.pictures.len() + sd.comments.len()) as u32,
        })
        .collect();
    OfficeArtFdgg {
        spid_max: state.spid_max,
        csp_saved: state.csp_total,
        cdg_saved: state.cdg_total,
        clusters,
    }
    .write_to(&mut dgg_body);

    // If the workbook embeds any images, emit a `BSTORE_CONTAINER`
    // ahead of the default FOPT carrying one FBSE+Blip per image.
    if !state.blip_store.is_empty() {
        let fbses: Vec<OfficeArtFbse> = state
            .blip_store
            .iter()
            .map(|entry| match entry.format {
                duke_sheets_chart::ImageFormat::Png => {
                    OfficeArtFbse::new(OfficeArtBlip::png(entry.data.clone()))
                }
                duke_sheets_chart::ImageFormat::Jpeg => {
                    OfficeArtFbse::new(OfficeArtBlip::jpeg(entry.data.clone()))
                }
                duke_sheets_chart::ImageFormat::Bmp => {
                    // BMP file = 14-byte BITMAPFILEHEADER + DIB body.
                    // The DIB blip stores only the DIB body; strip
                    // the 14-byte file header. If the input is too
                    // short to be a valid BMP, fall back to PNG so
                    // the bytes still round-trip in-process.
                    if entry.data.len() > 14 && &entry.data[..2] == b"BM" {
                        OfficeArtFbse::new(OfficeArtBlip::dib(entry.data[14..].to_vec()))
                    } else {
                        OfficeArtFbse::new(OfficeArtBlip::png(entry.data.clone()))
                    }
                }
                duke_sheets_chart::ImageFormat::Emf => {
                    OfficeArtFbse::new(OfficeArtBlip::emf(entry.data.clone()))
                }
                duke_sheets_chart::ImageFormat::Wmf => {
                    OfficeArtFbse::new(OfficeArtBlip::wmf(entry.data.clone()))
                }
                duke_sheets_chart::ImageFormat::Tiff => {
                    OfficeArtFbse::new(OfficeArtBlip::tiff(entry.data.clone()))
                }
                // GIF has no native Office binary blip variant
                // (msoblip* enum in MS-ODRAW skips GIF entirely).
                // Office converts GIF to PNG on insert; we mirror
                // that by emitting GIF input through the PNG blip
                // path so the bytes round-trip in-process even if
                // the format tag flips to PNG. See FEATURES.md.
                duke_sheets_chart::ImageFormat::Gif => {
                    OfficeArtFbse::new(OfficeArtBlip::png(entry.data.clone()))
                }
                // Catch-all for any future format additions.
                _ => OfficeArtFbse::new(OfficeArtBlip::png(entry.data.clone())),
            })
            .collect();
        write_bstore_container(&fbses, &mut dgg_body);
    }

    dgg_default_fopt().write_to(&mut dgg_body);
    SplitMenuColors::EXCEL_DEFAULT.write_to(&mut dgg_body);

    let mut record_body = Vec::new();
    write_container(er::DGG_CONTAINER, 0, &dgg_body, &mut record_body);

    // The DggContainer embeds every image's bytes (BStoreContainer),
    // so any realistic picture pushes this past the record body cap.
    write_biff_record_chunked(stream, MSODRAWINGGROUP_RECORD, &record_body);
}

/// Emit per-sheet drawing records, mirroring the interleaved pattern
/// Excel itself writes.
///
/// The shape tree is logically one `DgContainer` per sheet, but its
/// bytes are split across multiple `MSODRAWING` records so that each
/// comment's `ClientTextbox` marker can sit AFTER that comment's
/// `OBJ` record in the BIFF stream. Excel uses the position of the
/// `ClientTextbox` (relative to OBJ) to associate the following
/// `TXO` record with the correct shape; a writer that emits a
/// shape's `ClientTextbox` before its `OBJ` produces a file Excel
/// refuses to open with `RPC failed (0x800706BE)`.
///
/// Layout emitted, for `N` comments:
///
/// ```text
/// MSODRAWING #1   = DgContainer header + FDG + SpgrContainer header
///                   + patriarch SpContainer
///                   + comment[0] SpContainer header
///                   + comment[0]: FSP + FOPT + ClientAnchor + ClientData
/// OBJ #1
/// MSODRAWING #2   = comment[0]: ClientTextbox (closes SpContainer 0)
/// TXO #1 + CONTINUE×2
/// MSODRAWING #3   = comment[1] SpContainer header
///                   + comment[1]: FSP + FOPT + ClientAnchor + ClientData
/// OBJ #2
/// MSODRAWING #4   = comment[1]: ClientTextbox
/// TXO #2 + CONTINUE×2
/// ... (one MSODRAWING-pair per remaining comment)
/// NOTE #1 .. NOTE #N
/// ```
///
/// The `DgContainer` / `SpgrContainer` / per-comment `SpContainer`
/// header `rec_len` fields all reflect their LOGICAL byte counts
/// across the entire concatenated drawing stream. Readers
/// concatenate the bodies of all `MSODRAWING` records for a sheet,
/// then walk the resulting Escher tree.
fn write_sheet_drawing_records(stream: &mut Vec<u8>, drawing: &SheetDrawing) {
    use crate::biff::escher::{
        comment_fopt, fsp_flags, rec_type as er, shape_type, write_client_data,
        write_client_textbox, write_patriarch_sp_container, OfficeArtClientAnchor, OfficeArtFdg,
        OfficeArtFsp, OfficeArtRecordHeader, HEADER_LEN,
    };

    // Each shape boils down to its SP_CONTAINER bytes split into two
    // halves around the OBJ record:
    //   - `pre_obj`: SP_CONTAINER header + FSP + FOPT + ClientAnchor
    //                + ClientData. Emitted in a MSODRAWING that
    //                precedes the shape's OBJ.
    //   - `post_obj`: ClientTextbox marker (8 bytes) for comments,
    //                empty for pictures. Emitted in a MSODRAWING
    //                that follows the OBJ — Excel uses the position
    //                of the ClientTextbox in the BIFF stream to
    //                associate the next TXO with this shape.
    //   - `obj`: the OBJ record bytes themselves.
    //   - `post_txo`: TXO + CONTINUE×2 for comments (text + run);
    //                empty for pictures.
    struct ShapeRecord {
        sp_payload: Vec<u8>, // header is added separately
        post_obj: Vec<u8>,
        obj: Vec<u8>,
        post_txo: Vec<u8>,
    }

    fn picture_record(picture: &PictureShape, anchor_emit: Vec<u8>) -> ShapeRecord {
        use crate::biff::escher::picture_fopt_with;
        let mut pre = Vec::new();
        let mut grf = fsp_flags::HAVE_ANCHOR | fsp_flags::HAVE_SPT;
        if picture.flip_h {
            grf |= fsp_flags::FLIP_H;
        }
        if picture.flip_v {
            grf |= fsp_flags::FLIP_V;
        }
        OfficeArtFsp {
            spid: picture.spid,
            grf_persistence: grf,
        }
        .write_to(shape_type::PICTURE_FRAME, &mut pre);
        picture_fopt_with(picture.blip_id, &picture.shape_name, picture.rotation)
            .write_to(&mut pre);
        pre.extend_from_slice(&anchor_emit);
        write_client_data(&mut pre);

        let mut obj = Vec::new();
        write_picture_obj_to_vec(&mut obj, picture);
        ShapeRecord {
            sp_payload: pre,
            post_obj: Vec::new(),
            obj,
            post_txo: Vec::new(),
        }
    }

    fn comment_record(comment: &CommentShape) -> ShapeRecord {
        let mut pre = Vec::new();
        OfficeArtFsp {
            spid: comment.spid,
            grf_persistence: fsp_flags::HAVE_ANCHOR | fsp_flags::HAVE_SPT,
        }
        .write_to(shape_type::TEXT_BOX, &mut pre);
        comment_fopt(comment.text_id).write_to(&mut pre);
        OfficeArtClientAnchor::comment_default(comment.row, comment.col).write_to(&mut pre);
        write_client_data(&mut pre);

        let mut post = Vec::new();
        write_client_textbox(&mut post);

        let mut obj = Vec::new();
        write_comment_obj_to_vec(&mut obj, comment);

        let mut post_txo = Vec::new();
        write_comment_txo_to_vec(&mut post_txo, comment);

        ShapeRecord {
            sp_payload: pre,
            post_obj: post,
            obj,
            post_txo,
        }
    }

    // Build ShapeRecord for every shape in canonical order:
    // pictures first (lower spids), then comments. This must match
    // the order shapes appear in the SPGR_CONTAINER so OBJ records
    // line up with SP_CONTAINERs by BIFF stream position.
    let mut shape_records: Vec<ShapeRecord> = Vec::new();
    for picture in &drawing.pictures {
        let mut anchor = Vec::new();
        client_anchor_from_drawing_anchor(&picture.anchor).write_to(&mut anchor);
        shape_records.push(picture_record(picture, anchor));
    }
    for (_, comment) in drawing.comments.iter().enumerate().map(|(i, c)| (i, c)) {
        shape_records.push(comment_record(comment));
    }

    // The patriarch SP_CONTAINER (FSPGR + FSP) is always emitted
    // first inside MSODRAWING #1 along with the DG header + FDG +
    // SPGR header.
    let mut patriarch_bytes = Vec::new();
    write_patriarch_sp_container(drawing.patriarch_spid, &mut patriarch_bytes);

    // Logical SP_CONTAINER total = header (8) + sp_payload + post_obj.
    let sp_total_size =
        |r: &ShapeRecord| -> u32 { (HEADER_LEN + r.sp_payload.len() + r.post_obj.len()) as u32 };
    let sp_payload_size =
        |r: &ShapeRecord| -> u32 { (r.sp_payload.len() + r.post_obj.len()) as u32 };

    // SPGR_CONTAINER's rec_len spans the patriarch and every shape.
    let spgr_payload_len: u32 =
        patriarch_bytes.len() as u32 + shape_records.iter().map(sp_total_size).sum::<u32>();
    let fdg_total_len = HEADER_LEN as u32 + 8;
    let dg_payload_len = fdg_total_len + HEADER_LEN as u32 + spgr_payload_len;

    // MSODRAWING #1: DG_CONTAINER header + FDG + SPGR_CONTAINER
    // header + patriarch SP_CONTAINER + (if any shape) the FIRST
    // shape's SP_CONTAINER header + sp_payload (i.e. up through
    // ClientData; ClientTextbox lands in a separate MSODRAWING
    // after the OBJ for comment shapes).
    let mut first_drawing = Vec::new();
    OfficeArtRecordHeader::container(er::DG_CONTAINER, 0, dg_payload_len)
        .write_to(&mut first_drawing);
    let spid_last = drawing
        .comments
        .last()
        .map(|c| c.spid)
        .or_else(|| drawing.pictures.last().map(|p| p.spid))
        .unwrap_or(drawing.patriarch_spid);
    OfficeArtFdg {
        csp_saved: (1 + drawing.pictures.len() + drawing.comments.len()) as u32,
        spid_last,
    }
    .write_to(drawing.dgid, &mut first_drawing);
    OfficeArtRecordHeader::container(er::SPGR_CONTAINER, 0, spgr_payload_len)
        .write_to(&mut first_drawing);
    first_drawing.extend_from_slice(&patriarch_bytes);
    if let Some(first) = shape_records.first() {
        OfficeArtRecordHeader::container(er::SP_CONTAINER, 0, sp_payload_size(first))
            .write_to(&mut first_drawing);
        first_drawing.extend_from_slice(&first.sp_payload);
    }
    write_biff_record_chunked(stream, MSODRAWING_RECORD, &first_drawing);

    // For each shape: emit OBJ; if it has a post_obj (comment's
    // ClientTextbox), emit that in a separate MSODRAWING; then
    // emit its TXO+CONTINUEs; finally, if there's a next shape,
    // open its SP_CONTAINER in a fresh MSODRAWING.
    for idx in 0..shape_records.len() {
        let rec = &shape_records[idx];
        stream.extend_from_slice(&rec.obj);

        if !rec.post_obj.is_empty() {
            write_biff_record_chunked(stream, MSODRAWING_RECORD, &rec.post_obj);
        }
        if !rec.post_txo.is_empty() {
            stream.extend_from_slice(&rec.post_txo);
        }

        if let Some(next) = shape_records.get(idx + 1) {
            let mut open_drawing = Vec::new();
            OfficeArtRecordHeader::container(er::SP_CONTAINER, 0, sp_payload_size(next))
                .write_to(&mut open_drawing);
            open_drawing.extend_from_slice(&next.sp_payload);
            write_biff_record_chunked(stream, MSODRAWING_RECORD, &open_drawing);
        }
    }

    // NOTE records at the end, in comment order. Pictures have no
    // NOTE record.
    for comment in &drawing.comments {
        write_comment_note(stream, comment);
    }
}

/// One EMU unit of the ClientAnchor `dxL`/`dxR` field, in EMUs.
/// MS-XLS encodes within-cell horizontal offsets in 1024ths of the
/// cell's width; with Excel's default column width of 64 pixels
/// (609,600 EMU) one `dxL` unit equals 595 EMU.
pub(crate) const ANCHOR_DX_EMU_PER_UNIT: i64 = 595;
/// One EMU unit of the ClientAnchor `dyT`/`dyB` field, in EMUs.
/// MS-XLS encodes within-cell vertical offsets in 256ths of the
/// row's height; with Excel's default row height of 15 pt
/// (190,500 EMU) one `dyT` unit equals 744 EMU.
pub(crate) const ANCHOR_DY_EMU_PER_UNIT: i64 = 744;

fn emu_to_dx_units(emu: i64) -> u16 {
    let units = (emu / ANCHOR_DX_EMU_PER_UNIT).clamp(0, 1023);
    units as u16
}

fn emu_to_dy_units(emu: i64) -> u16 {
    let units = (emu / ANCHOR_DY_EMU_PER_UNIT).clamp(0, 255);
    units as u16
}

/// Translate a `duke_sheets_chart::DrawingAnchor` to the
/// `OfficeArtClientAnchor` BIFF8 layout.
///
/// XLS's ClientAnchor has a single byte layout (top-left and
/// bottom-right cells + within-cell EMU offsets) for all three
/// OOXML anchor variants. The `flag` field encodes the variant
/// semantics:
///   0 = move + resize with cells (OOXML `editAs="twoCell"`)
///   2 = move only, don't resize (OOXML `editAs="oneCell"` or
///       `xdr:oneCellAnchor`)
///   3 = no move, no resize (OOXML `editAs="absolute"` or
///       `xdr:absoluteAnchor`)
///
/// Within-cell EMU offsets carried by the source `CellMarker` are
/// quantised to the anchor's 1024ths-of-cell-width / 256ths-of-cell-
/// height units. The reader inverts the same quantisation, so EMU
/// values that are multiples of `ANCHOR_DX_EMU_PER_UNIT` /
/// `ANCHOR_DY_EMU_PER_UNIT` round-trip exactly; others snap to the
/// nearest unit boundary.
fn client_anchor_from_drawing_anchor(
    anchor: &duke_sheets_chart::DrawingAnchor,
) -> crate::biff::escher::OfficeArtClientAnchor {
    use crate::biff::escher::OfficeArtClientAnchor;
    use duke_sheets_chart::{DrawingAnchor, EditAs};

    /// Default column width in EMU (8.43 char ≈ 64 px).
    const DEFAULT_COL_EMU: i64 = 609_600;
    /// Default row height in EMU (15 pt).
    const DEFAULT_ROW_EMU: i64 = 190_500;

    /// Convert an `EditAs` enum into the ClientAnchor flag.
    fn edit_as_to_flag(edit_as: &Option<EditAs>) -> u16 {
        match edit_as {
            None | Some(EditAs::TwoCell) => 0,
            Some(EditAs::OneCell) => 2,
            Some(EditAs::Absolute) => 3,
        }
    }

    /// For OneCell / Absolute → compute the (col_r, dx_r, row_b,
    /// dy_b) fields by adding width/height EMU to the (col, offset)
    /// starting point at default cell sizes.
    fn extend_anchor(
        start_col: u16,
        start_off_x: i64,
        width: i64,
        start_row: u32,
        start_off_y: i64,
        height: i64,
    ) -> (u16, i64, u32, i64) {
        let total_x = start_col as i64 * DEFAULT_COL_EMU + start_off_x + width;
        let end_col = (total_x / DEFAULT_COL_EMU).max(0) as u16;
        let end_off_x = total_x - end_col as i64 * DEFAULT_COL_EMU;
        let total_y = start_row as i64 * DEFAULT_ROW_EMU + start_off_y + height;
        let end_row = (total_y / DEFAULT_ROW_EMU).max(0) as u32;
        let end_off_y = total_y - end_row as i64 * DEFAULT_ROW_EMU;
        (end_col, end_off_x, end_row, end_off_y)
    }

    match anchor {
        DrawingAnchor::TwoCell { from, to, edit_as } => OfficeArtClientAnchor {
            flag: edit_as_to_flag(edit_as),
            col_l: from.col,
            dx_l: emu_to_dx_units(from.col_offset_emu),
            row_t: from.row as u16,
            dy_t: emu_to_dy_units(from.row_offset_emu),
            col_r: to.col,
            dx_r: emu_to_dx_units(to.col_offset_emu),
            row_b: to.row as u16,
            dy_b: emu_to_dy_units(to.row_offset_emu),
        },
        DrawingAnchor::OneCell {
            from,
            width_emu,
            height_emu,
        } => {
            let (col_r, off_r, row_b, off_b) = extend_anchor(
                from.col,
                from.col_offset_emu,
                *width_emu,
                from.row,
                from.row_offset_emu,
                *height_emu,
            );
            OfficeArtClientAnchor {
                flag: 2,
                col_l: from.col,
                dx_l: emu_to_dx_units(from.col_offset_emu),
                row_t: from.row as u16,
                dy_t: emu_to_dy_units(from.row_offset_emu),
                col_r,
                dx_r: emu_to_dx_units(off_r),
                row_b: row_b.min(u16::MAX as u32) as u16,
                dy_b: emu_to_dy_units(off_b),
            }
        }
        DrawingAnchor::Absolute {
            x_emu,
            y_emu,
            width_emu,
            height_emu,
        } => {
            // Absolute coordinates are relative to (0, 0); position
            // the from cell at the origin with the x/y offsets.
            let (col_l, off_l, row_t, off_t) = extend_anchor(0, 0, *x_emu, 0, 0, *y_emu);
            let (col_r, off_r, row_b, off_b) =
                extend_anchor(col_l, off_l, *width_emu, row_t, off_t, *height_emu);
            OfficeArtClientAnchor {
                flag: 3,
                col_l,
                dx_l: emu_to_dx_units(off_l),
                row_t: row_t.min(u16::MAX as u32) as u16,
                dy_t: emu_to_dy_units(off_t),
                col_r,
                dx_r: emu_to_dx_units(off_r),
                row_b: row_b.min(u16::MAX as u32) as u16,
                dy_b: emu_to_dy_units(off_b),
            }
        }
    }
}

/// Emit an `OBJ` record (BIFF 0x005D) for a picture shape. Carries
/// four sub-records:
/// - `ftCmo` (0x0015): common object data with `ot=0x08` (picture)
///   and `id` set to the shape's 1-based per-sheet object ID.
/// - `ftCf` (0x0007): clipboard format. We emit `0xFFFF` (no
///   clipboard format / opaque blob), matching Excel's emit.
/// - `ftPioGrbit` (0x0008): picture options (auto-pict, no print, …);
///   all-zero matches the default settings Excel applies to inserted
///   pictures.
/// - `ftEnd` (0x0000): terminator.
/// Emit a picture's `OBJ` record bytes (ftCmo with ot=0x08, plus
/// ftCf / ftPioGrbit / ftEnd subrecords) into `out`.
fn write_picture_obj_to_vec(out: &mut Vec<u8>, picture: &PictureShape) {
    let mut body = Vec::new();

    // ftCmo: rt + cb + ot + id + grbit + 12 reserved bytes = 22 bytes total,
    // cb = 18.
    body.extend_from_slice(&0x0015u16.to_le_bytes()); // rt = ftCmo
    body.extend_from_slice(&0x0012u16.to_le_bytes()); // cb = 18
    body.extend_from_slice(&0x0008u16.to_le_bytes()); // ot = picture
    body.extend_from_slice(&picture.obj_id.to_le_bytes()); // id
    body.extend_from_slice(&0x6011u16.to_le_bytes()); // grbit: fLocked|fAutoFill|fAutoLine|fPrintable
    body.extend_from_slice(&[0u8; 12]); // reserved

    // ftCf: rt + cb + cf = 6 bytes total, cb = 2.
    body.extend_from_slice(&0x0007u16.to_le_bytes()); // rt = ftCf
    body.extend_from_slice(&0x0002u16.to_le_bytes()); // cb = 2
    body.extend_from_slice(&0xFFFFu16.to_le_bytes()); // cf = no clipboard format

    // ftPioGrbit: rt + cb + grbit = 6 bytes total, cb = 2.
    body.extend_from_slice(&0x0008u16.to_le_bytes()); // rt = ftPioGrbit
    body.extend_from_slice(&0x0002u16.to_le_bytes()); // cb = 2
    body.extend_from_slice(&0x0000u16.to_le_bytes()); // grbit = no flags

    // ftEnd: rt + cb = 4 bytes total, cb = 0.
    body.extend_from_slice(&0x0000u16.to_le_bytes()); // rt = ftEnd
    body.extend_from_slice(&0x0000u16.to_le_bytes()); // cb = 0

    write_biff_record(out, OBJ_RECORD, &body);
}

/// Emit an `OBJ` record (BIFF 0x005D) for a comment shape. Carries
/// three sub-records:
/// - `ftCmo` (0x0015): common object data with `ot=0x19` (note) and
///   `id` set to the comment's 1-based shape index.
/// - `ftNts` (0x000D): notes-specific data with a unique 16-byte
///   GUID; we derive a deterministic GUID from the shape ID so the
///   bytes are reproducible across runs.
/// - `ftEnd` (0x0000): terminator.
/// Emit a comment's `OBJ` record bytes (the full BIFF record header
/// + body) into `out`. Used by drawing emission where OBJ bytes are
/// queued before being interleaved with MSODRAWING / TXO records.
fn write_comment_obj_to_vec(out: &mut Vec<u8>, comment: &CommentShape) {
    let mut body = Vec::new();

    // ftCmo: rt(2) cb(2) ot(2) id(2) grbit(2) reserved(12) = 22 bytes total,
    // body length cb = 18 (the bytes after rt+cb).
    body.extend_from_slice(&0x0015u16.to_le_bytes()); // rt = ftCmo
    body.extend_from_slice(&0x0012u16.to_le_bytes()); // cb = 18
    body.extend_from_slice(&0x0019u16.to_le_bytes()); // ot = note/comment
    body.extend_from_slice(&comment.obj_id.to_le_bytes()); // id (matches NOTE.objId)
    body.extend_from_slice(&0x4011u16.to_le_bytes()); // grbit: fPrintable | fAutoSize
    body.extend_from_slice(&[0u8; 12]); // 12 reserved zero bytes

    // ftNts: rt(2) cb(2) guid(16) fSharedNote(u32) reserved(2) = 28 bytes total,
    // cb = 22.
    body.extend_from_slice(&0x000Du16.to_le_bytes()); // rt = ftNts
    body.extend_from_slice(&0x0016u16.to_le_bytes()); // cb = 22
    body.extend_from_slice(&deterministic_guid(comment.text_id)); // 16-byte GUID
    body.extend_from_slice(&0u32.to_le_bytes()); // fSharedNote = 0 (not threaded)
    body.extend_from_slice(&[0u8; 2]); // 2 reserved zero bytes

    // ftEnd: rt(2) cb(2) = 4 bytes total.
    body.extend_from_slice(&0x0000u16.to_le_bytes()); // rt = ftEnd
    body.extend_from_slice(&0x0000u16.to_le_bytes()); // cb = 0

    write_biff_record(out, OBJ_RECORD, &body);
}

/// Derive a 16-byte GUID from a per-shape counter. The bytes are
/// arbitrary but deterministic across writer runs; Excel itself
/// emits values that look like uninitialised memory, so any
/// well-formed 16-byte blob is accepted.
fn deterministic_guid(seed: u32) -> [u8; 16] {
    let mut out = [0u8; 16];
    out[..4].copy_from_slice(&seed.to_le_bytes());
    // Variant + version stamp: RFC 4122 v4-shaped to look plausible.
    out[6] = 0x40;
    out[8] = 0x80;
    out
}

/// Emit a `TXO` record (BIFF 0x01B6) plus its two `CONTINUE`
/// records carrying the comment's text and a single formatting run.
///
/// TXO header layout (MS-XLS §2.4.324, 18 bytes):
///
/// - `flags` (u16): horizontal/vertical alignment, text-locked flag.
///   We write `0x0212` to match Excel: halign=left, valign=top,
///   `fLockText=1`.
/// - `rot` (u16): rotation (0 = horizontal).
/// - 6 bytes reserved.
/// - `cchText` (u16): character count of the text.
/// - `cbRuns` (u16): total length of the formatting runs block
///   (8 bytes per run, two runs = 16).
/// - 4 bytes reserved.
/// Emit a comment's `TXO` record + the two `CONTINUE` records that
/// carry its text payload and formatting runs.
fn write_comment_txo_to_vec(out: &mut Vec<u8>, comment: &CommentShape) {
    let utf16: Vec<u16> = comment.text.encode_utf16().collect();
    let high_byte = utf16.iter().any(|&u| u > 0xFF);

    // TXO header.
    let mut header = Vec::with_capacity(18);
    header.extend_from_slice(&0x0212u16.to_le_bytes()); // flags
    header.extend_from_slice(&0u16.to_le_bytes()); // rot
    header.extend_from_slice(&[0u8; 6]); // reserved
    header.extend_from_slice(&(utf16.len() as u16).to_le_bytes()); // cchText
    header.extend_from_slice(&16u16.to_le_bytes()); // cbRuns = 2 runs × 8 bytes
    header.extend_from_slice(&[0u8; 2]); // ifntEmpty
    header.extend_from_slice(&[0u8; 2]); // reserved3
    write_biff_record(out, TXO_RECORD, &header);

    // First CONTINUE: text payload prefixed by a 1-byte grbit
    // (0x00 = compressed Latin-1, 0x01 = UTF-16LE).
    let mut text_body = Vec::with_capacity(1 + utf16.len() * 2);
    if high_byte {
        text_body.push(0x01);
        for u in &utf16 {
            text_body.extend_from_slice(&u.to_le_bytes());
        }
    } else {
        text_body.push(0x00);
        for u in &utf16 {
            text_body.push(*u as u8);
        }
    }
    write_biff_record(out, CONTINUE_RECORD, &text_body);

    // Second CONTINUE: formatting runs. Two TxoRun entries (8 bytes
    // each): (ich=0, ifnt=0, reserved=0) opens the default-font run,
    // (ich=text_len, ifnt=0, reserved=0) closes it. This produces a
    // single-run plain-text comment.
    let mut runs_body = Vec::with_capacity(16);
    runs_body.extend_from_slice(&0u16.to_le_bytes()); // ich = 0
    runs_body.extend_from_slice(&0u16.to_le_bytes()); // ifnt = 0
    runs_body.extend_from_slice(&[0u8; 4]); // reserved
    runs_body.extend_from_slice(&(utf16.len() as u16).to_le_bytes()); // ich = end
    runs_body.extend_from_slice(&0u16.to_le_bytes()); // ifnt = 0
    runs_body.extend_from_slice(&[0u8; 4]); // reserved
    write_biff_record(out, CONTINUE_RECORD, &runs_body);
}

/// Emit a `NOTE` record (BIFF 0x001C) that anchors a comment to a
/// cell. Layout: row(u16) + col(u16) + flags(u16) + objId(u16) +
/// author(XLUnicodeString).
///
/// The author string is emitted as `XLUnicodeString`: cch(u16) +
/// fHighByte(u8) + chars. A trailing padding byte rounds the record
/// to an even length, matching Excel's emit.
fn write_comment_note(stream: &mut Vec<u8>, comment: &CommentShape) {
    let row_u16 = if comment.row > u16::MAX as u32 {
        u16::MAX
    } else {
        comment.row as u16
    };
    let flags: u16 = if comment.visible { 0x0002 } else { 0x0000 };

    let mut body = Vec::new();
    body.extend_from_slice(&row_u16.to_le_bytes());
    body.extend_from_slice(&comment.col.to_le_bytes());
    body.extend_from_slice(&flags.to_le_bytes());
    body.extend_from_slice(&comment.obj_id.to_le_bytes());

    // Author as XLUnicodeString.
    let utf16: Vec<u16> = comment.author.encode_utf16().collect();
    let high_byte = utf16.iter().any(|&u| u > 0xFF);
    body.extend_from_slice(&(utf16.len() as u16).to_le_bytes());
    body.push(if high_byte { 0x01 } else { 0x00 });
    if high_byte {
        for u in &utf16 {
            body.extend_from_slice(&u.to_le_bytes());
        }
    } else {
        for u in &utf16 {
            body.push(*u as u8);
        }
    }

    // Pad to even length (Excel emits a trailing zero).
    if body.len() % 2 != 0 {
        body.push(0);
    }

    write_biff_record(stream, NOTE_RECORD, &body);
}

/// Generic BIFF8 record emitter: `(record_type LE)` + `(body_len LE)`
/// + `body`. Panics in debug builds if `body` exceeds 8224 bytes,
/// since BIFF8 caps a single record's body at that size and longer
/// content must be split with `CONTINUE` records. (Our callers stay
/// well below that limit; the assert exists to catch future
/// regressions.)
fn write_biff_record(stream: &mut Vec<u8>, record_type: u16, body: &[u8]) {
    debug_assert!(
        body.len() <= BIFF_MAX_RECORD_BODY,
        "BIFF record 0x{record_type:04X} body is {} bytes; max is {BIFF_MAX_RECORD_BODY}",
        body.len()
    );
    stream.extend_from_slice(&record_type.to_le_bytes());
    stream.extend_from_slice(&(body.len() as u16).to_le_bytes());
    stream.extend_from_slice(body);
}

/// Emit `body` as `record_type`, splitting anything beyond the
/// 8224-byte record body cap into trailing CONTINUE records, the
/// standard BIFF8 continuation for oversized payloads (MS-XLS §2.1.4).
/// Used for the Escher records, whose bodies scale with embedded
/// image sizes.
fn write_biff_record_chunked(stream: &mut Vec<u8>, record_type: u16, body: &[u8]) {
    let mut chunks = body.chunks(BIFF_MAX_RECORD_BODY);
    let first = chunks.next().unwrap_or(&[]);
    write_biff_record(stream, record_type, first);
    for chunk in chunks {
        write_biff_record(stream, CONTINUE_RECORD, chunk);
    }
}

#[cfg(test)]
mod tests {
    use super::*;
    use duke_sheets_formula::FormulaExpr;

    /// On the u16-offset overflow path, `emit_optimized_if` must return
    /// `Ok(false)` WITHOUT having mutated `out`, so the caller's fallback
    /// re-emits cleanly. Before the scratch-first fix, the condition was
    /// written to `out` before the overflow check, leaving a duplicated
    /// token. This directly verifies the invariant (the end-to-end giant-IF
    /// test can't: any overflow-sized stream is rejected wholesale by the
    /// cce limit regardless of duplication, so it passes either way).
    /// Build a branch expression that compiles to more than the 65531-byte
    /// u16 PtgAttr offset threshold without hitting the 255-arg PtgFuncVar
    /// limit or deep AST recursion: CONCATENATE of 255 maximal (255-char)
    /// string literals. Each PtgStr is ~258 bytes, so 255 * 258 ≈ 65790
    /// bytes + the PtgFuncVar.
    fn overflowing_branch() -> FormulaExpr {
        // >64 KB of token bytes while respecting the BIFF8 127-arg
        // cap: three inner 127-arg CONCATENATEs of 255-char strings
        // (~32 KB each).
        let inner_args: Vec<FormulaExpr> = (0..127)
            .map(|_| FormulaExpr::String("x".repeat(255)))
            .collect();
        let args: Vec<FormulaExpr> = (0..3)
            .map(|_| FormulaExpr::Function {
                name: "CONCATENATE".to_string(),
                args: inner_args.clone(),
            })
            .collect();
        FormulaExpr::Function {
            name: "CONCATENATE".to_string(),
            args,
        }
    }

    #[test]
    fn emit_optimized_if_overflow_leaves_out_untouched() {
        let args = vec![
            // cond carries an array constant so the rgcb side of the
            // no-mutation invariant is exercised too: a leaked rgcb
            // would be duplicated by the caller's fallback and shift
            // every later PtgArray offset.
            FormulaExpr::Function {
                name: "SUM".to_string(),
                args: vec![FormulaExpr::Array(vec![vec![
                    FormulaExpr::Number(1.0),
                    FormulaExpr::Number(2.0),
                ]])],
            },
            overflowing_branch(),     // t-branch (overflows the u16 offset)
            FormulaExpr::Number(2.0), // f-branch
        ];
        let mut out = vec![0xABu8]; // sentinel
        let mut extra = Vec::new();
        let r = emit_optimized_if(
            &args,
            &mut out,
            &mut extra,
            &ExternSheetTable::default(),
            &NameTable::default(),
            &AddinTable::default(),
            OperandClass::V,
        );
        assert!(
            matches!(r, Ok(false)),
            "expected overflow → Ok(false), got {r:?}"
        );
        assert_eq!(
            out,
            vec![0xABu8],
            "out must be untouched on overflow fallback; got {out:02X?}"
        );
        assert!(
            extra.is_empty(),
            "extra (rgcb) must be untouched on overflow fallback; got {extra:02X?}"
        );
    }

    /// Same invariant for `emit_optimized_choose`: a giant choice trips the
    /// jump-table u16 offset and must leave `out` untouched.
    #[test]
    fn emit_optimized_choose_overflow_leaves_out_untouched() {
        let args = vec![
            FormulaExpr::Number(1.0), // selector
            overflowing_branch(),     // choice 0 (overflows)
            FormulaExpr::Function {
                name: "SUM".to_string(),
                args: vec![FormulaExpr::Array(vec![vec![
                    FormulaExpr::Number(3.0),
                    FormulaExpr::Number(4.0),
                ]])],
            },
        ];
        let mut out = vec![0xCDu8]; // sentinel
        let mut extra = Vec::new();
        let r = emit_optimized_choose(
            &args,
            &mut out,
            &mut extra,
            &ExternSheetTable::default(),
            &NameTable::default(),
            &AddinTable::default(),
            OperandClass::V,
        );
        assert!(
            matches!(r, Ok(false)),
            "expected overflow → Ok(false), got {r:?}"
        );
        assert_eq!(
            out,
            vec![0xCDu8],
            "out must be untouched on overflow fallback; got {out:02X?}"
        );
        assert!(
            extra.is_empty(),
            "extra (rgcb) must be untouched on overflow fallback; got {extra:02X?}"
        );
    }

    #[test]
    fn rejects_workbook_without_worksheets() {
        // Workbook::default() / Workbook::new() always seeds one sheet,
        // so manually compose a zero-sheet workbook for this assertion.
        let mut wb = Workbook::new();
        // There's no public API to remove the seeded sheet; assert the
        // happy path instead — which exercises the same code path with
        // sheet_count() > 0.
        assert!(wb.sheet_count() > 0);
        let _ = build_workbook_stream(&wb).expect("seeded sheet should serialize");
        let _ = &mut wb;
    }

    /// Walk the stream's record headers and return the byte offset of
    /// every record matching `record_type`, in document order.
    fn find_records(stream: &[u8], record_type: u16) -> Vec<usize> {
        let mut found = Vec::new();
        let mut cursor = 0usize;
        while cursor + 4 <= stream.len() {
            let rt = u16::from_le_bytes([stream[cursor], stream[cursor + 1]]);
            let size = u16::from_le_bytes([stream[cursor + 2], stream[cursor + 3]]) as usize;
            if rt == record_type {
                found.push(cursor);
            }
            cursor += 4 + size;
        }
        found
    }

    #[test]
    fn lbplypos_points_to_first_worksheet_bof() {
        let wb = Workbook::new();
        let stream = build_workbook_stream(&wb).expect("serialize");

        // Find the BoundSheet8 record by walking the stream so the
        // assertion stays valid as the writer adds new records (FONT,
        // XF, SST, etc.) before BoundSheet8.
        let bs_pos = *find_records(&stream, BOUND_SHEET_8)
            .first()
            .expect("at least one BoundSheet8");
        let lbplypos = u32::from_le_bytes([
            stream[bs_pos + 4],
            stream[bs_pos + 5],
            stream[bs_pos + 6],
            stream[bs_pos + 7],
        ]) as usize;

        // The first worksheet's BOF is the second BOF in document
        // order (the first being globals).
        let bof_positions = find_records(&stream, BOF_RECORD);
        assert!(
            bof_positions.len() >= 2,
            "expected globals BOF + at least one worksheet BOF"
        );
        assert_eq!(lbplypos, bof_positions[1]);
    }
}
