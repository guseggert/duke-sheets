use std::io::{Read, Seek};

use duke_sheets_core::table::{Table, TableColumn, TableStyleInfo, TotalsRowFunction};
use duke_sheets_core::CellRange;

use crate::biff12::records;
use crate::biff12::{parser, RecordIter};
use crate::error::XlsbResult;

pub(crate) fn read_table<R: Read + Seek>(
    archive: &mut zip::ZipArchive<R>,
    table_path: &str,
) -> XlsbResult<Option<Table>> {
    let file = match archive.by_name(table_path) {
        Ok(f) => f,
        Err(_) => return Ok(None),
    };

    let mut iter = RecordIter::new(file);
    let mut buf = Vec::with_capacity(1024);

    let mut table: Option<Table> = None;
    let mut columns: Vec<TableColumn> = Vec::new();
    let mut style_info: Option<TableStyleInfo> = None;

    loop {
        let (typ, len) = match iter.next_record(&mut buf) {
            Ok(r) => r,
            Err(_) => break,
        };

        match typ {
            records::BRT_BEGIN_LIST => table = parse_table_header(&buf[..len]),
            records::BRT_BEGIN_LIST_COLS => {
            }
            records::BRT_END_LIST_COLS => {
            }
            records::BRT_BEGIN_LIST_COL => {
                if let Some(col) = parse_table_column(&buf[..len]) {
                    columns.push(col);
                }
            }
            records::BRT_TABLE_STYLE_CLIENT => {
                style_info = parse_table_style(&buf[..len]);
            }
            records::BRT_END_LIST => break,
            _ => {}
        }
    }

    if let Some(ref mut t) = table {
        t.columns = columns;
        t.style_info = style_info;
    }

    Ok(table)
}

fn parse_table_header(data: &[u8]) -> Option<Table> {
    if data.len() < 68 {
        return None;
    }
    let first_row = parser::read_u32(data, 0);
    let last_row = parser::read_u32(data, 4);
    let first_col = parser::read_u32(data, 8) as u16;
    let last_col = parser::read_u32(data, 12) as u16;
    let id = parser::read_u32(data, 20);
    let header_row_count = parser::read_u32(data, 24);
    let totals_row_count = parser::read_u32(data, 28);
    let flags = parser::read_u32(data, 32);
    let totals_row_shown = (flags & 1) != 0;

    let mut pos = 64;
    let name = match parser::wide_str(data, pos) {
        Ok((s, consumed)) => {
            pos += consumed;
            s
        }
        Err(_) => return None,
    };
    let display_name = match parser::wide_str(data, pos) {
        Ok((s, consumed)) => {
            pos += consumed;
            s
        }
        Err(_) => name.clone(),
    };
    for _ in 0..4 {
        if let Ok((_, consumed)) = parser::wide_str(data, pos) {
            pos += consumed;
        }
    }

    let reference = CellRange::from_indices(first_row, first_col, last_row, last_col);

    Some(Table {
        id,
        name,
        display_name,
        reference,
        columns: Vec::new(),
        style_info: None,
        header_row_count,
        totals_row_count,
        totals_row_shown,
    })
}

fn parse_table_column(data: &[u8]) -> Option<TableColumn> {
    if data.len() < 28 {
        return None;
    }
    let id = parser::read_u32(data, 0);
    let totals_fn_raw = parser::read_u32(data, 4);

    let totals_row_function = match totals_fn_raw {
        0 => None,
        1 => Some(TotalsRowFunction::Average),
        2 => Some(TotalsRowFunction::Count),
        3 => Some(TotalsRowFunction::CountNums),
        4 => Some(TotalsRowFunction::Max),
        5 => Some(TotalsRowFunction::Min),
        6 => Some(TotalsRowFunction::Sum),
        7 => Some(TotalsRowFunction::StdDev),
        8 => Some(TotalsRowFunction::Var),
        9 => Some(TotalsRowFunction::Custom),
        _ => None,
    };

    let mut pos = 28;
    let st_name = match parser::wide_str(data, pos) {
        Ok((s, consumed)) => {
            pos += consumed;
            s
        }
        Err(_) => return None,
    };
    let totals_label = match parser::wide_str(data, pos) {
        Ok((s, consumed)) => {
            pos += consumed;
            if s.is_empty() {
                None
            } else {
                Some(s)
            }
        }
        Err(_) => None,
    };
    for _ in 0..3 {
        if let Ok((_, consumed)) = parser::wide_str(data, pos) {
            pos += consumed;
        }
    }

    let name = st_name;

    Some(TableColumn {
        id,
        name,
        totals_row_function,
        totals_row_formula: None,
        totals_row_label: totals_label,
        calculated_column_formula: None,
    })
}

fn parse_table_style(data: &[u8]) -> Option<TableStyleInfo> {
    if data.len() < 6 {
        return None;
    }
    let flags = parser::read_u16(data, 0);
    let show_first_column = (flags & 0x01) != 0;
    let show_last_column = (flags & 0x02) != 0;
    let show_row_stripes = (flags & 0x04) != 0;
    let show_column_stripes = (flags & 0x08) != 0;

    let name = match parser::wide_str(data, 2) {
        Ok((s, _)) => {
            if s.is_empty() {
                None
            } else {
                Some(s)
            }
        }
        Err(_) => None,
    };

    Some(TableStyleInfo {
        name,
        show_first_column,
        show_last_column,
        show_row_stripes,
        show_column_stripes,
    })
}
