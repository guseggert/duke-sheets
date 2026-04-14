use std::io::{Seek, Write};

use zip::write::SimpleFileOptions;
use zip::ZipWriter;

use crate::biff12::{encode_wide_str, records, RecordWriter};
use crate::error::XlsbResult;
use duke_sheets_core::table::{Table, TotalsRowFunction};

pub(crate) fn write_table_part<W: Write + Seek>(
    zip: &mut ZipWriter<W>,
    options: &SimpleFileOptions,
    table: &Table,
    global_table_num: usize,
) -> XlsbResult<()> {
    let path = format!("xl/tables/table{}.bin", global_table_num);
    zip.start_file(&path, *options)?;

    let mut buf = Vec::new();
    let mut rw = RecordWriter::new(&mut buf);

    write_begin_list(&mut rw, table)?;

    write_auto_filter(&mut rw, table)?;

    rw.write_record(
        records::BRT_BEGIN_LIST_COLS,
        &(table.columns.len() as u32).to_le_bytes(),
    )?;
    for col in &table.columns {
        write_list_col(&mut rw, col)?;
    }
    rw.write_record(records::BRT_END_LIST_COLS, &[])?;

    if let Some(ref style) = table.style_info {
        write_table_style_client(&mut rw, style)?;
    }

    rw.write_record(records::BRT_END_LIST, &[])?;

    drop(rw);
    zip.write_all(&buf)?;
    Ok(())
}

fn write_begin_list<W: Write>(rw: &mut RecordWriter<W>, table: &Table) -> std::io::Result<()> {
    let ref_ = &table.reference;
    let mut payload = Vec::new();

    payload.extend_from_slice(&ref_.start.row.to_le_bytes());
    payload.extend_from_slice(&ref_.end.row.to_le_bytes());
    payload.extend_from_slice(&(ref_.start.col as u32).to_le_bytes());
    payload.extend_from_slice(&(ref_.end.col as u32).to_le_bytes());
    payload.extend_from_slice(&0u32.to_le_bytes()); // lt
    payload.extend_from_slice(&table.id.to_le_bytes());
    payload.extend_from_slice(&table.header_row_count.to_le_bytes());
    payload.extend_from_slice(&table.totals_row_count.to_le_bytes());

    let flags: u32 = if table.totals_row_shown { 1 } else { 0 };
    payload.extend_from_slice(&flags.to_le_bytes());

    payload.extend_from_slice(&0xFFFF_FFFFu32.to_le_bytes()); // nDxfHeader
    payload.extend_from_slice(&0xFFFF_FFFFu32.to_le_bytes()); // nDxfData
    payload.extend_from_slice(&0xFFFF_FFFFu32.to_le_bytes()); // nDxfAgg
    payload.extend_from_slice(&0xFFFF_FFFFu32.to_le_bytes()); // nDxfBorder
    payload.extend_from_slice(&0xFFFF_FFFFu32.to_le_bytes()); // nDxfHeaderBorder
    payload.extend_from_slice(&0xFFFF_FFFFu32.to_le_bytes()); // nDxfAggBorder
    payload.extend_from_slice(&0u32.to_le_bytes()); // dwConnID

    payload.extend_from_slice(&encode_wide_str(&table.name));
    payload.extend_from_slice(&encode_wide_str(&table.display_name));
    payload.extend_from_slice(&encode_wide_str(""));
    payload.extend_from_slice(&encode_wide_str(""));
    payload.extend_from_slice(&encode_wide_str(""));
    payload.extend_from_slice(&encode_wide_str(""));

    rw.write_record(records::BRT_BEGIN_LIST, &payload)
}

fn write_list_col<W: Write>(
    rw: &mut RecordWriter<W>,
    col: &duke_sheets_core::table::TableColumn,
) -> std::io::Result<()> {
    let mut payload = Vec::new();

    payload.extend_from_slice(&col.id.to_le_bytes());
    payload.extend_from_slice(&totals_fn_to_raw(col.totals_row_function).to_le_bytes());
    payload.extend_from_slice(&0xFFFF_FFFFu32.to_le_bytes()); // nDxfHdr
    payload.extend_from_slice(&0xFFFF_FFFFu32.to_le_bytes()); // nDxfInsertRow
    payload.extend_from_slice(&0xFFFF_FFFFu32.to_le_bytes()); // nDxfAgg
    payload.extend_from_slice(&0u32.to_le_bytes()); // idqsif
    payload.extend_from_slice(&0xFFFF_FFFFu32.to_le_bytes()); // nDxfData
    payload.extend_from_slice(&encode_wide_str(&col.name));
    payload.extend_from_slice(&encode_wide_str(""));
    payload.extend_from_slice(&encode_wide_str(""));
    payload.extend_from_slice(&encode_wide_str(""));
    payload.extend_from_slice(&encode_wide_str(""));
    payload.extend_from_slice(&encode_wide_str(""));

    rw.write_record(records::BRT_BEGIN_LIST_COL, &payload)?;
    rw.write_record(records::BRT_END_LIST_COL, &[])
}

fn write_table_style_client<W: Write>(
    rw: &mut RecordWriter<W>,
    style: &duke_sheets_core::table::TableStyleInfo,
) -> std::io::Result<()> {
    let mut flags: u16 = 0;
    if style.show_first_column {
        flags |= 0x01;
    }
    if style.show_last_column {
        flags |= 0x02;
    }
    if style.show_row_stripes {
        flags |= 0x04;
    }
    if style.show_column_stripes {
        flags |= 0x08;
    }

    let mut payload = Vec::new();
    payload.extend_from_slice(&flags.to_le_bytes());
    payload.extend_from_slice(&encode_wide_str(style.name.as_deref().unwrap_or("")));

    rw.write_record(records::BRT_TABLE_STYLE_CLIENT, &payload)
}

fn write_auto_filter<W: Write>(rw: &mut RecordWriter<W>, table: &Table) -> std::io::Result<()> {
    let ref_ = &table.reference;
    let end_row = if table.has_totals_row() {
        ref_.end.row.saturating_sub(table.totals_row_count)
    } else {
        ref_.end.row
    };

    let mut payload = vec![0u8; 16];
    payload[0..4].copy_from_slice(&ref_.start.row.to_le_bytes());
    payload[4..8].copy_from_slice(&end_row.to_le_bytes());
    payload[8..12].copy_from_slice(&(ref_.start.col as u32).to_le_bytes());
    payload[12..16].copy_from_slice(&(ref_.end.col as u32).to_le_bytes());

    rw.write_record(records::BRT_BEGIN_A_FILTER, &payload)?;
    rw.write_record(records::BRT_END_A_FILTER, &[])
}

fn totals_fn_to_raw(f: Option<TotalsRowFunction>) -> u32 {
    match f {
        None => 0,
        Some(TotalsRowFunction::Average) => 1,
        Some(TotalsRowFunction::Count) => 2,
        Some(TotalsRowFunction::CountNums) => 3,
        Some(TotalsRowFunction::Max) => 4,
        Some(TotalsRowFunction::Min) => 5,
        Some(TotalsRowFunction::Sum) => 6,
        Some(TotalsRowFunction::StdDev) => 7,
        Some(TotalsRowFunction::Var) => 8,
        Some(TotalsRowFunction::Custom) => 9,
        Some(TotalsRowFunction::None) => 0,
    }
}
