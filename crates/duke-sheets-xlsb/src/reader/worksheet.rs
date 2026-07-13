use std::collections::HashMap;
use std::io::Read;

use duke_sheets_core::conditional_format::{
    CfColorValue, CfOperator, CfRuleType, CfValue, CfValueType, ConditionalFormatRule,
    IconSetStyle, TimePeriod,
};
use duke_sheets_core::style::{Color, Style};
use duke_sheets_core::validation::{
    DataValidation, ValidationErrorStyle, ValidationOperator, ValidationType,
};
use duke_sheets_core::worksheet::PageOrientation;
use duke_sheets_core::worksheet::{PageBreak, SheetProtection};
use duke_sheets_core::{
    AutoFilter, CellError, CellRange, CellValue, Hyperlink, ProtectedRange, Worksheet,
};

use super::shared_strings::SharedStringEntry;
use duke_sheets_formula::decompile::FormulaContext;

use crate::biff12::records;
use crate::biff12::RecordIter;
use crate::biff12::{parser, token_parser};
use crate::error::XlsbResult;

pub(crate) fn read_worksheet<R: Read>(
    reader: R,
    ws: &mut Worksheet,
    shared_strings: &[SharedStringEntry],
    styles: &[Style],
    formula_ctx: &FormulaContext,
    sheet_rels: &HashMap<String, String>,
) -> XlsbResult<()> {
    let mut iter = RecordIter::new(reader);
    let mut buf = Vec::with_capacity(4096);

    loop {
        let (typ, len) = iter.next_record(&mut buf)?;
        match typ {
            records::BRT_BEGIN_SHEET_DATA => break,

            records::BRT_WS_PROP => {
                parse_ws_prop(&buf[..len], ws);
            }

            records::BRT_BEGIN_SHEET_VIEW => {
                if len >= 4 {
                    let sv_flags = parser::read_u32(&buf, 0);
                    if (sv_flags & 0x02) != 0 {
                        ws.set_selected(true);
                    }
                }
                if len >= 20 {
                    let zoom = parser::read_u16(&buf, 18);
                    if zoom != 100 && zoom > 0 {
                        ws.set_zoom_scale(Some(zoom));
                    }
                }
            }

            records::BRT_PANE => {
                parse_pane(&buf[..len], ws);
            }

            records::BRT_SEL => {
                if len >= 12 {
                    let row = parser::read_u32(&buf, 4);
                    let col = parser::read_u32(&buf, 8) as u16;
                    if row != 0 || col != 0 {
                        ws.set_selection_active_cell(row, col);
                    }
                }
            }

            records::BRT_COL_INFO => {
                parse_col_info(&buf[..len], ws);
            }

            records::BRT_WS_FMT_INFO => {
                // BrtWsFmtInfo: u32 dxGCol + u16 cchDefColWidth +
                // u16 miyDefRwHeight (twips) + u16 flags + 2 outlines.
                if len >= 8 {
                    let cch_def = parser::read_u16(&buf, 4);
                    let miy = parser::read_u16(&buf, 6);
                    if cch_def > 0 {
                        ws.set_default_column_width(cch_def as f64);
                    }
                    if miy > 0 {
                        ws.set_default_row_height(miy as f64 / 20.0);
                    }
                }
            }

            _ => {}
        }
    }

    let mut current_row: u32 = 0;

    let apply_style = |ws: &mut Worksheet, row: u32, col: u16, buf: &[u8]| {
        if buf.len() >= 8 {
            let style_ref = parser::cell_style_ref(buf) as usize;
            if style_ref != 0 {
                if let Some(style) = styles.get(style_ref) {
                    ws.set_cell_style_at(row, col, style).ok();
                }
            }
        }
    };

    loop {
        let (typ, len) = iter.next_record(&mut buf)?;

        match typ {
            records::BRT_ROW_HDR => {
                if len >= 4 {
                    current_row = parser::read_u32(&buf, 0);
                }
                parse_row_hdr(&buf[..len], len, current_row, ws);
            }
            records::BRT_END_SHEET_DATA => break,

            records::BRT_CELL_BLANK => {
                if len >= 8 {
                    let col = parser::read_u32(&buf, 0) as u16;
                    apply_style(ws, current_row, col, &buf);
                }
            }
            records::BRT_CELL_RK => {
                if len >= 12 {
                    let col = parser::read_u32(&buf, 0) as u16;
                    let value = parser::decode_rk(&buf, 8);
                    ws.set_cell_value_at(current_row, col, value).ok();
                    apply_style(ws, current_row, col, &buf);
                }
            }
            records::BRT_CELL_REAL => {
                if len >= 16 {
                    let col = parser::read_u32(&buf, 0) as u16;
                    let value = parser::read_f64(&buf, 8);
                    ws.set_cell_value_at(current_row, col, value).ok();
                    apply_style(ws, current_row, col, &buf);
                }
            }
            records::BRT_CELL_BOOL => {
                if len >= 9 {
                    let col = parser::read_u32(&buf, 0) as u16;
                    let value = buf[8] != 0;
                    ws.set_cell_value_at(current_row, col, value).ok();
                    apply_style(ws, current_row, col, &buf);
                }
            }
            records::BRT_CELL_ERROR => {
                if len >= 9 {
                    let col = parser::read_u32(&buf, 0) as u16;
                    let error = parse_cell_error(buf[8]);
                    ws.set_cell_value_at(current_row, col, CellValue::Error(error))
                        .ok();
                    apply_style(ws, current_row, col, &buf);
                }
            }
            records::BRT_CELL_ST => {
                if len > 8 {
                    let col = parser::read_u32(&buf, 0) as u16;
                    if let Ok((s, _)) = parser::wide_str(&buf, 8) {
                        ws.set_cell_value_at(current_row, col, s.as_str()).ok();
                    }
                    apply_style(ws, current_row, col, &buf);
                }
            }
            records::BRT_CELL_ISST => {
                if len >= 12 {
                    let col = parser::read_u32(&buf, 0) as u16;
                    let isst = parser::read_u32(&buf, 8) as usize;
                    match shared_strings.get(isst) {
                        Some(SharedStringEntry::Plain(s)) => {
                            ws.set_cell_value_at(current_row, col, s.as_str()).ok();
                        }
                        Some(SharedStringEntry::Rich(runs)) => {
                            ws.set_cell_value_at(
                                current_row,
                                col,
                                CellValue::rich_text(runs.clone()),
                            )
                            .ok();
                        }
                        None => {}
                    }
                    apply_style(ws, current_row, col, &buf);
                }
            }

            records::BRT_FMLA_NUM => {
                if len >= 16 {
                    let col = parser::read_u32(&buf, 0) as u16;
                    let value = parser::read_f64(&buf, 8);
                    let cached = CellValue::Number(value);
                    if let Some(formula) = decompile_formula(&buf[..len], 16, formula_ctx) {
                        ws.set_formula_with_cached_value_at(current_row, col, &formula, cached)
                            .ok();
                    } else {
                        ws.set_cell_value_at(current_row, col, value).ok();
                    }
                    apply_style(ws, current_row, col, &buf);
                }
            }
            records::BRT_FMLA_BOOL => {
                if len >= 9 {
                    let col = parser::read_u32(&buf, 0) as u16;
                    let value = buf[8] != 0;
                    let cached = CellValue::Boolean(value);
                    if let Some(formula) = decompile_formula(&buf[..len], 9, formula_ctx) {
                        ws.set_formula_with_cached_value_at(current_row, col, &formula, cached)
                            .ok();
                    } else {
                        ws.set_cell_value_at(current_row, col, value).ok();
                    }
                    apply_style(ws, current_row, col, &buf);
                }
            }
            records::BRT_FMLA_ERROR => {
                if len >= 9 {
                    let col = parser::read_u32(&buf, 0) as u16;
                    let error = parse_cell_error(buf[8]);
                    let cached = CellValue::Error(error);
                    if let Some(formula) = decompile_formula(&buf[..len], 9, formula_ctx) {
                        ws.set_formula_with_cached_value_at(current_row, col, &formula, cached)
                            .ok();
                    } else {
                        ws.set_cell_value_at(current_row, col, CellValue::Error(error))
                            .ok();
                    }
                    apply_style(ws, current_row, col, &buf);
                }
            }
            records::BRT_FMLA_STRING => {
                if len > 8 {
                    let col = parser::read_u32(&buf, 0) as u16;
                    if let Ok((s, str_consumed)) = parser::wide_str(&buf, 8) {
                        let cached = CellValue::String(s.clone().into());
                        let value_end = 8 + str_consumed;
                        if let Some(formula) =
                            decompile_formula(&buf[..len], value_end, formula_ctx)
                        {
                            ws.set_formula_with_cached_value_at(current_row, col, &formula, cached)
                                .ok();
                        } else {
                            ws.set_cell_value_at(current_row, col, s.as_str()).ok();
                        }
                    }
                    apply_style(ws, current_row, col, &buf);
                }
            }

            _ => {}
        }
    }

    let mut cf_pending: Option<CfPendingBase> = None;
    let mut cf_visual: Option<CfVisualAccum> = None;
    let mut cf_ranges: Vec<CellRange> = Vec::new();
    let mut in_row_breaks = false;
    let mut in_col_breaks = false;
    let mut row_breaks_acc: Vec<PageBreak> = Vec::new();
    let mut col_breaks_acc: Vec<PageBreak> = Vec::new();
    // Tracks the open BrtBeginFilterColumn block so child filter
    // records (Top10/Filters/CustomFilters) can attach to a
    // FilterColumn that gets pushed at BrtEndFilterColumn.
    let mut current_filter_column: Option<duke_sheets_core::auto_filter::FilterColumn> = None;
    let mut current_value_filter: Option<duke_sheets_core::auto_filter::ValueFilter> = None;
    let mut current_custom_filters: Option<duke_sheets_core::auto_filter::CustomFilters> = None;

    loop {
        let (typ, len) = match iter.next_record(&mut buf) {
            Ok(r) => r,
            Err(_) => break,
        };

        match typ {
            records::BRT_MERGE_CELL => {
                parse_merge_cell(&buf[..len], ws);
            }

            records::BRT_H_LINK => {
                parse_hyperlink(&buf[..len], ws, sheet_rels);
            }

            records::BRT_BEGIN_A_FILTER => {
                parse_auto_filter(&buf[..len], ws);
            }

            records::BRT_BEGIN_FILTER_COLUMN => {
                if len >= 6 {
                    use duke_sheets_core::auto_filter::{ColumnFilter, FilterColumn, ValueFilter};
                    let col_id = parser::read_u32(&buf, 0);
                    let flags = parser::read_u16(&buf, 4);
                    current_filter_column = Some(FilterColumn {
                        col_id,
                        hidden_button: (flags & 0x01) != 0,
                        show_button: (flags & 0x02) == 0,
                        // Placeholder; replaced when a child filter record arrives.
                        filter: ColumnFilter::Values(ValueFilter {
                            values: Vec::new(),
                            blank: false,
                        }),
                    });
                }
            }
            records::BRT_END_FILTER_COLUMN => {
                if let Some(mut fc) = current_filter_column.take() {
                    use duke_sheets_core::auto_filter::ColumnFilter;
                    if let Some(vf) = current_value_filter.take() {
                        fc.filter = ColumnFilter::Values(vf);
                    } else if let Some(cf) = current_custom_filters.take() {
                        fc.filter = ColumnFilter::Custom(cf);
                    }
                    // Top10 sets fc.filter directly when the record
                    // is parsed, so no merge is needed here.
                    if let Some(af) = ws.auto_filter() {
                        let mut af = af.clone();
                        af.filter_columns.push(fc);
                        ws.set_auto_filter(Some(af));
                    }
                }
            }
            records::BRT_TOP10_FILTER => {
                if len >= 17 {
                    use duke_sheets_core::auto_filter::{ColumnFilter, Top10Filter};
                    if let Some(fc) = current_filter_column.as_mut() {
                        let flags = buf[0];
                        let val = parser::read_f64(&buf, 1);
                        // xNumFilter is meaningful only when fApplied
                        // (bit 2) is set.
                        let filter_val = if (flags & 0x04) != 0 {
                            Some(parser::read_f64(&buf, 9))
                        } else {
                            None
                        };
                        fc.filter = ColumnFilter::Top10(Top10Filter {
                            top: (flags & 0x01) != 0,
                            percent: (flags & 0x02) != 0,
                            val,
                            filter_val,
                        });
                    }
                }
            }
            records::BRT_DYNAMIC_FILTER => {
                if len >= 21 {
                    use duke_sheets_core::auto_filter::{ColumnFilter, DynamicFilter};
                    if let Some(fc) = current_filter_column.as_mut() {
                        let cft = parser::read_u32(&buf, 0);
                        // Skip the 1-byte fApplied flag at offset 4.
                        let val = parser::read_f64(&buf, 5);
                        let max_val = parser::read_f64(&buf, 13);
                        fc.filter = ColumnFilter::Dynamic(DynamicFilter {
                            filter_type: cft_to_dynamic_filter_type(cft),
                            val: if val == 0.0 { None } else { Some(val) },
                            max_val: if max_val == 0.0 { None } else { Some(max_val) },
                        });
                    }
                }
            }
            records::BRT_COLOR_FILTER => {
                if len >= 8 {
                    use duke_sheets_core::auto_filter::{ColorFilter, ColumnFilter};
                    if let Some(fc) = current_filter_column.as_mut() {
                        let dxf_id = parser::read_u32(&buf, 0);
                        let cell_color = parser::read_u32(&buf, 4) != 0;
                        fc.filter = ColumnFilter::Color(ColorFilter {
                            dxf_id: Some(dxf_id),
                            cell_color,
                        });
                    }
                }
            }
            records::BRT_BEGIN_FILTERS => {
                if len >= 4 {
                    use duke_sheets_core::auto_filter::ValueFilter;
                    let blank = parser::read_u32(&buf, 0) != 0;
                    current_value_filter = Some(ValueFilter {
                        values: Vec::new(),
                        blank,
                    });
                }
            }
            records::BRT_END_FILTERS => {
                // Closed by BrtEndFilterColumn merge step.
            }
            records::BRT_FILTER => {
                if let Some(vf) = current_value_filter.as_mut() {
                    // Slice to the record length: the reuse buffer only
                    // grows, so parsing the whole buffer would decode
                    // stale bytes from a previous larger record.
                    if let Ok((s, _)) = parser::wide_str(&buf[..len], 0) {
                        vf.values.push(s);
                    }
                }
            }
            records::BRT_BEGIN_CUSTOM_FILTERS => {
                use duke_sheets_core::auto_filter::CustomFilters;
                // Payload is an i32 AND/OR flag: 0 = AND, nonzero = OR
                // (LO CustomFilter::importRecord).
                let and = if len >= 4 {
                    parser::read_u32(&buf, 0) == 0
                } else {
                    true
                };
                current_custom_filters = Some(CustomFilters {
                    and,
                    conditions: Vec::new(),
                });
            }
            records::BRT_END_CUSTOM_FILTERS => {
                // Closed by BrtEndFilterColumn merge step.
            }
            records::BRT_CUSTOM_FILTER => {
                if let Some(cf) = current_custom_filters.as_mut() {
                    if let Some(cond) = parse_custom_filter(&buf[..len]) {
                        cf.conditions.push(cond);
                    }
                }
            }

            records::BRT_MARGINS => {
                parse_margins(&buf[..len], ws);
            }

            records::BRT_PAGE_SETUP => {
                parse_page_setup(&buf[..len], ws);
            }

            records::BRT_PRINT_OPTIONS => {
                parse_print_options(&buf[..len], ws);
            }

            records::BRT_HEADER_FOOTER => {
                parse_header_footer(&buf[..len], ws);
            }

            records::BRT_SHEET_PROTECTION => {
                parse_sheet_protection(&buf[..len], ws);
            }
            records::BRT_RANGE_PROTECTION => {
                parse_range_protection(&buf[..len], ws);
            }

            records::BRT_BEGIN_RW_BRK => {
                in_row_breaks = true;
                parse_inline_breaks(&buf[..len], &mut row_breaks_acc);
            }

            records::BRT_END_RW_BRK => {
                in_row_breaks = false;
            }

            records::BRT_BEGIN_COL_BRK => {
                in_col_breaks = true;
                parse_inline_breaks(&buf[..len], &mut col_breaks_acc);
            }

            records::BRT_END_COL_BRK => {
                in_col_breaks = false;
            }

            records::BRT_BRK => {
                if len >= 16 {
                    let brk = parse_page_break(&buf[..len]);
                    if in_row_breaks {
                        row_breaks_acc.push(brk);
                    } else if in_col_breaks {
                        col_breaks_acc.push(brk);
                    }
                }
            }

            records::BRT_DVAL => {
                parse_data_validation(&buf[..len], ws, formula_ctx);
            }

            records::BRT_BEGIN_COND_FMT => {
                cf_ranges = parse_cond_fmt_ranges(&buf[..len]);
            }

            records::BRT_BEGIN_CF_RULE => {
                cf_pending = parse_cf_rule_base(&buf[..len], ws, formula_ctx, &cf_ranges);
            }

            records::BRT_BEGIN_COLOR_SCALE => {
                cf_visual = Some(CfVisualAccum::ColorScale {
                    cfvos: Vec::new(),
                    colors: Vec::new(),
                });
            }

            records::BRT_BEGIN_DATA_BAR => {
                if let Some(show_value) = parse_cf_data_bar_begin(&buf[..len]) {
                    cf_visual = Some(CfVisualAccum::DataBar {
                        cfvos: Vec::new(),
                        color: None,
                        show_value,
                    });
                }
            }

            records::BRT_BEGIN_ICON_SET => {
                if let Some((icon_style, reverse, show_value)) =
                    parse_cf_icon_set_begin(&buf[..len])
                {
                    cf_visual = Some(CfVisualAccum::IconSet {
                        icon_style,
                        cfvos: Vec::new(),
                        reverse,
                        show_value,
                    });
                }
            }

            records::BRT_CFVO => {
                if let Some(cfvo) = parse_brt_cfvo(&buf[..len]) {
                    match cf_visual {
                        Some(CfVisualAccum::ColorScale { ref mut cfvos, .. })
                        | Some(CfVisualAccum::DataBar { ref mut cfvos, .. })
                        | Some(CfVisualAccum::IconSet { ref mut cfvos, .. }) => {
                            cfvos.push(cfvo);
                        }
                        None => {}
                    }
                }
            }

            records::BRT_CF_COLOR => {
                let mut cpos = 0;
                let color = read_cf_brt_color(&buf[..len], &mut cpos);
                match cf_visual {
                    Some(CfVisualAccum::ColorScale { ref mut colors, .. }) => colors.push(color),
                    Some(CfVisualAccum::DataBar {
                        color: ref mut bar_color,
                        ..
                    }) => *bar_color = Some(color),
                    _ => {}
                }
            }

            records::BRT_END_COLOR_SCALE => {
                if let (Some(base), Some(CfVisualAccum::ColorScale { cfvos, colors })) =
                    (cf_pending.take(), cf_visual.take())
                {
                    let count = cfvos.len().min(colors.len());
                    let mut cv = Vec::with_capacity(count);
                    for i in 0..count {
                        let (vt, val) = cfvos[i];
                        let value = match vt {
                            CfValueType::Min | CfValueType::Max => None,
                            _ => Some(format_cfvo_value(val)),
                        };
                        cv.push(CfColorValue::new(vt, value, colors[i].clone()));
                    }
                    ws.add_conditional_format(
                        base.into_rule(CfRuleType::ColorScale { colors: cv }),
                    );
                }
            }

            records::BRT_END_DATA_BAR => {
                if let (
                    Some(base),
                    Some(CfVisualAccum::DataBar {
                        cfvos,
                        color: Some(color),
                        show_value,
                    }),
                ) = (cf_pending.take(), cf_visual.take())
                {
                    if cfvos.len() >= 2 {
                        ws.add_conditional_format(base.into_rule(CfRuleType::DataBar {
                            min_value: cf_value_from_pair(cfvos[0]),
                            max_value: cf_value_from_pair(cfvos[1]),
                            color,
                            show_value,
                            gradient: true,
                            border_color: None,
                            negative_color: None,
                        }));
                    }
                }
            }

            records::BRT_END_ICON_SET => {
                if let (
                    Some(base),
                    Some(CfVisualAccum::IconSet {
                        icon_style,
                        cfvos,
                        reverse,
                        show_value,
                    }),
                ) = (cf_pending.take(), cf_visual.take())
                {
                    let values = cfvos.into_iter().map(cf_value_from_pair).collect();
                    ws.add_conditional_format(base.into_rule(CfRuleType::IconSet {
                        icon_style,
                        values,
                        reverse,
                        show_value,
                    }));
                }
            }

            records::BRT_END_CF_RULE => {
                cf_pending = None;
                cf_visual = None;
            }

            records::BRT_END_COND_FMT => {
                cf_ranges.clear();
            }

            _ => {}
        }
    }

    if !row_breaks_acc.is_empty() {
        ws.set_row_breaks(row_breaks_acc);
    }
    if !col_breaks_acc.is_empty() {
        ws.set_col_breaks(col_breaks_acc);
    }

    Ok(())
}

fn parse_page_break(data: &[u8]) -> PageBreak {
    let id = parser::read_u32(data, 0);
    let min = parser::read_u32(data, 4);
    let max = parser::read_u32(data, 8);
    let man = parser::read_u32(data, 12) != 0;
    let pt = if data.len() >= 20 {
        parser::read_u32(data, 16) != 0
    } else {
        false
    };
    PageBreak {
        id,
        min,
        max,
        man,
        pt,
    }
}

fn parse_inline_breaks(data: &[u8], acc: &mut Vec<PageBreak>) {
    let mut off = 0;
    while off + 20 <= data.len() {
        acc.push(parse_page_break(&data[off..]));
        off += 20;
    }
}

fn parse_sheet_protection(data: &[u8], ws: &mut Worksheet) {
    if data.len() < 4 {
        return;
    }

    let prot = if data.len() >= 66 {
        let password_hash = parser::read_u16(data, 0);
        let field = |i: usize| parser::read_u32(data, 2 + i * 4) != 0;
        SheetProtection {
            protected: field(0),
            password_hash: if password_hash != 0 {
                Some(password_hash)
            } else {
                None
            },
            format_cells: field(3),
            format_columns: field(4),
            format_rows: field(5),
            insert_columns: field(6),
            insert_rows: field(7),
            insert_hyperlinks: field(8),
            delete_columns: field(9),
            delete_rows: field(10),
            select_locked_cells: field(11),
            sort: field(12),
            auto_filter: field(13),
            pivot_tables: field(14),
            select_unlocked_cells: field(15),
        }
    } else {
        let flags = parser::read_u16(data, 0);
        let password_hash = parser::read_u16(data, 2);
        SheetProtection {
            protected: (flags & (1 << 0)) != 0,
            password_hash: if password_hash != 0 {
                Some(password_hash)
            } else {
                None
            },
            select_locked_cells: (flags & (1 << 11)) == 0,
            select_unlocked_cells: (flags & (1 << 15)) == 0,
            format_cells: (flags & (1 << 3)) == 0,
            format_columns: (flags & (1 << 4)) == 0,
            format_rows: (flags & (1 << 5)) == 0,
            insert_columns: (flags & (1 << 6)) == 0,
            insert_rows: (flags & (1 << 7)) == 0,
            insert_hyperlinks: (flags & (1 << 8)) == 0,
            delete_columns: (flags & (1 << 9)) == 0,
            delete_rows: (flags & (1 << 10)) == 0,
            sort: (flags & (1 << 12)) == 0,
            auto_filter: (flags & (1 << 13)) == 0,
            pivot_tables: (flags & (1 << 14)) == 0,
        }
    };
    ws.set_protection(Some(prot));
}

fn parse_range_protection(data: &[u8], ws: &mut Worksheet) {
    if data.len() < 6 {
        return;
    }

    let password_hash = parser::read_u16(data, 0);
    let mut pos = 2;
    let ranges = read_sqref(data, &mut pos);
    if ranges.is_empty() || pos >= data.len() {
        return;
    }

    let Ok((name, consumed)) = parser::wide_str(data, pos) else {
        return;
    };
    pos += consumed;
    if name.is_empty() {
        return;
    }

    let security_descriptor = if pos + 4 <= data.len() {
        let len = parser::read_u32(data, pos) as usize;
        pos += 4;
        if len > 0 && pos + len <= data.len() {
            Some(format!("hex:{}", hex_encode(&data[pos..pos + len])))
        } else {
            None
        }
    } else {
        None
    };

    ws.add_protected_range(ProtectedRange {
        name,
        ranges,
        password_hash: if password_hash != 0 {
            Some(password_hash)
        } else {
            None
        },
        security_descriptor,
    });
}

fn hex_encode(bytes: &[u8]) -> String {
    const HEX: &[u8; 16] = b"0123456789ABCDEF";
    let mut out = String::with_capacity(bytes.len() * 2);
    for &byte in bytes {
        out.push(HEX[(byte >> 4) as usize] as char);
        out.push(HEX[(byte & 0x0F) as usize] as char);
    }
    out
}

/// Parse BrtWsProp record for tab color.
///
/// Layout per [MS-XLSB] 2.4.875:
///   Bytes 0-2:  flags + reserved4 (3 bytes)
///   Bytes 3-10: brtcolorTab (BrtColor, 8 bytes)
fn parse_ws_prop(data: &[u8], ws: &mut Worksheet) {
    if data.len() < 11 {
        return;
    }
    let color = parse_brt_color_ws(data, 3);
    if !matches!(color, duke_sheets_core::style::Color::Auto) {
        ws.set_tab_color(Some(color));
    }
}

fn parse_brt_color_ws(buf: &[u8], off: usize) -> duke_sheets_core::style::Color {
    use duke_sheets_core::style::Color;
    if off + 8 > buf.len() {
        return Color::Auto;
    }
    let color_type = buf[off];
    let index = buf[off + 1];
    let tint_raw = i16::from_le_bytes([buf[off + 2], buf[off + 3]]);
    let r = buf[off + 4];
    let g = buf[off + 5];
    let b = buf[off + 6];
    let a = buf[off + 7];
    match color_type {
        0 => Color::Auto,
        1 => Color::Indexed(index),
        2 => {
            if a == 0xFF {
                Color::Rgb { r, g, b }
            } else {
                Color::Argb { a, r, g, b }
            }
        }
        3 => {
            let tint_i8 = if tint_raw == 0 {
                0i8
            } else {
                ((tint_raw as f64 / 32767.0) * 100.0).round() as i8
            };
            Color::Theme {
                index,
                tint: tint_i8,
            }
        }
        _ => Color::Auto,
    }
}

/// Parse BrtRowHdr for row height and hidden state.
///
/// Layout: rw(u32), ixfe(u32), miyRw(u16), grbitRw(3 bytes), ccolspan(u32), spans.
/// grbitRw byte 1 bits: 4=fDyZero(hidden), 5=fUnsynced(custom height).
fn parse_row_hdr(data: &[u8], len: usize, row: u32, ws: &mut Worksheet) {
    if len < 13 {
        return;
    }
    let miy_rw = parser::read_u16(data, 8);
    let flags = data[11];

    let custom_height = (flags & 0x20) != 0;
    let hidden = (flags & 0x10) != 0;

    if custom_height && miy_rw > 0 {
        let height_points = miy_rw as f64 / 20.0;
        ws.set_row_height(row, height_points);
    }
    if hidden {
        ws.set_row_hidden(row, true);
    }
    if len >= 11 {
        let outline_level = data[10] & 0x07;
        if outline_level > 0 {
            ws.set_row_outline_level(row, outline_level);
        }
    }
}

/// Parse BrtColInfo for column width and hidden state.
///
/// Layout: [0..4] first_col (u32), [4..8] last_col (u32),
/// [8..12] coldx (u32, width in 1/256 of char width), [12..16] ixfe (u32),
/// [16..18] flags (u16). Bit 0 = hidden, bit 1 = custom width.
fn parse_col_info(data: &[u8], ws: &mut Worksheet) {
    if data.len() < 18 {
        return;
    }
    let first_col = parser::read_u32(data, 0);
    let last_col = parser::read_u32(data, 4);
    let coldx = parser::read_u32(data, 8);
    let flags = parser::read_u16(data, 16);

    let hidden = (flags & 0x01) != 0;
    let custom_width = (flags & 0x02) != 0;
    let outline_level = ((flags >> 8) & 0x07) as u8;

    let width = coldx as f64 / 256.0;

    for col in first_col..=last_col {
        if col > u16::MAX as u32 {
            break;
        }
        let col_idx = col as u16;
        if custom_width && coldx > 0 {
            ws.set_column_width(col_idx, width);
        }
        if hidden {
            ws.set_column_hidden(col_idx, true);
        }
        if outline_level > 0 {
            ws.set_column_outline_level(col_idx, outline_level);
        }
    }
}

/// Parse BrtMergeCell record.
///
/// Layout: [0..4] first_row (u32), [4..8] last_row (u32),
/// [8..12] first_col (u32), [12..16] last_col (u32).
fn parse_merge_cell(data: &[u8], ws: &mut Worksheet) {
    if data.len() < 16 {
        return;
    }
    let first_row = parser::read_u32(data, 0);
    let last_row = parser::read_u32(data, 4);
    let first_col = parser::read_u32(data, 8) as u16;
    let last_col = parser::read_u32(data, 12) as u16;

    let range = CellRange::from_indices(first_row, first_col, last_row, last_col);
    if let Err(e) = ws.merge_cells(&range) {
        log::warn!("Skipping merge {}: {}", range, e);
    }
}

/// Parse BrtHLink record.
///
/// Layout: [0..4] first_row (u32), [4..8] last_row (u32),
/// [8..12] first_col (u32), [12..16] last_col (u32),
/// then XLWideString: relId, then XLWideString: location,
/// then XLWideString: tooltip, then XLWideString: display.
fn parse_hyperlink(data: &[u8], ws: &mut Worksheet, sheet_rels: &HashMap<String, String>) {
    if data.len() < 16 {
        return;
    }
    let first_row = parser::read_u32(data, 0);
    let first_col = parser::read_u32(data, 8) as u16;

    let mut pos = 16;

    let rel_id = match parser::wide_str(data, pos) {
        Ok((s, consumed)) => {
            pos += consumed;
            s
        }
        Err(_) => return,
    };

    let location = match parser::wide_str(data, pos) {
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

    let tooltip = match parser::wide_str(data, pos) {
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

    let display = match parser::wide_str(data, pos) {
        Ok((s, _)) => {
            if s.is_empty() {
                None
            } else {
                Some(s)
            }
        }
        Err(_) => None,
    };

    let mut target = String::new();
    if !rel_id.is_empty() {
        if let Some(url) = sheet_rels.get(&rel_id) {
            target = url.clone();
        }
    }
    if target.is_empty() {
        if let Some(loc) = &location {
            target = format!("#{}", loc);
        }
    }

    if target.is_empty() && display.is_none() {
        return;
    }

    let hyperlink = Hyperlink {
        target,
        display,
        tooltip,
        location,
    };

    let cell_ref = duke_sheets_core::CellAddress::new(first_row, first_col).to_a1_string();
    ws.set_hyperlink(&cell_ref, hyperlink).ok();
}

/// Parse BrtPane record for freeze panes.
///
/// Layout: [0..8] xSplit (f64), [8..16] ySplit (f64),
/// [16..20] topLeftRow (u32), [20..24] topLeftCol (u32),
/// [24..28] activePane (u32), [28] state flags.
/// State: 0=split, 1=frozen, 2=frozenSplit.
fn parse_pane(data: &[u8], ws: &mut Worksheet) {
    if data.len() < 29 {
        return;
    }
    let x_split = parser::read_f64(data, 0);
    let y_split = parser::read_f64(data, 8);
    let top_left_row = parser::read_u32(data, 16);
    let top_left_col = parser::read_u32(data, 20) as u16;
    let active_pane_raw = parser::read_u32(data, 24);
    let state = data[28];

    match state {
        1 | 2 => {
            let row = y_split.round().max(0.0) as u32;
            let col = x_split.round().max(0.0) as u16;
            if row > 0 || col > 0 {
                ws.set_freeze_panes(row, col);
            }
        }
        0 => {
            if x_split != 0.0 || y_split != 0.0 {
                let active_pane = match active_pane_raw {
                    1 => Some("topRight".to_string()),
                    2 => Some("bottomLeft".to_string()),
                    3 => Some("bottomRight".to_string()),
                    _ => None,
                };
                let top_left = if top_left_row > 0 || top_left_col > 0 {
                    Some((top_left_row, top_left_col))
                } else {
                    None
                };
                ws.set_split_panes(Some(duke_sheets_core::worksheet::SplitPanes {
                    x_split,
                    y_split,
                    top_left,
                    active_pane,
                }));
            }
        }
        _ => {}
    }
}

/// Parse BrtBeginAFilter record for autofilter range.
///
/// Layout: [0..4] first_row (u32), [4..8] last_row (u32),
/// [8..12] first_col (u32), [12..16] last_col (u32).
fn parse_auto_filter(data: &[u8], ws: &mut Worksheet) {
    if data.len() < 16 {
        return;
    }
    let first_row = parser::read_u32(data, 0);
    let last_row = parser::read_u32(data, 4);
    let first_col = parser::read_u32(data, 8) as u16;
    let last_col = parser::read_u32(data, 12) as u16;

    let range = CellRange::from_indices(first_row, first_col, last_row, last_col);
    ws.set_auto_filter(Some(AutoFilter::new(range)));
}

/// Map the BIFF12 cft code to our DynamicFilterType per [MS-XLSB]
/// §2.4.362. Unknown codes map to `Null`.
fn cft_to_dynamic_filter_type(cft: u32) -> duke_sheets_core::auto_filter::DynamicFilterType {
    use duke_sheets_core::auto_filter::DynamicFilterType as D;
    match cft {
        0x01 => D::AboveAverage,
        0x02 => D::BelowAverage,
        0x08 => D::Tomorrow,
        0x09 => D::Today,
        0x0A => D::Yesterday,
        0x0B => D::NextWeek,
        0x0C => D::ThisWeek,
        0x0D => D::LastWeek,
        0x0E => D::NextMonth,
        0x0F => D::ThisMonth,
        0x10 => D::LastMonth,
        0x11 => D::NextQuarter,
        0x12 => D::ThisQuarter,
        0x13 => D::LastQuarter,
        0x14 => D::NextYear,
        0x15 => D::ThisYear,
        0x16 => D::LastYear,
        0x17 => D::YearToDate,
        0x18 => D::Q1,
        0x19 => D::Q2,
        0x1A => D::Q3,
        0x1B => D::Q4,
        0x1C => D::M1,
        0x1D => D::M2,
        0x1E => D::M3,
        0x1F => D::M4,
        0x20 => D::M5,
        0x21 => D::M6,
        0x22 => D::M7,
        0x23 => D::M8,
        0x24 => D::M9,
        0x25 => D::M10,
        0x26 => D::M11,
        0x27 => D::M12,
        _ => D::Null,
    }
}

/// Parse a BrtCustomFilter record per [MS-XLSB] §2.4.348 (cross-checked
/// against LO FilterCriterionModel::readBiffData):
///   vts u8 + operator u8 + 8 bytes xNumOrError + rgch string, present
///   when vts = 6 (string). For vts = 4 (double) the value lives in the
///   8 number bytes.
fn parse_custom_filter(
    data: &[u8],
) -> Option<duke_sheets_core::auto_filter::CustomFilterCondition> {
    use duke_sheets_core::auto_filter::{CustomFilterCondition, FilterOperator};
    if data.len() < 1 + 1 + 8 {
        return None;
    }
    let vts = data[0];
    let operator = match data[1] {
        1 => FilterOperator::LessThan,
        2 => FilterOperator::Equal,
        3 => FilterOperator::LessThanOrEqual,
        4 => FilterOperator::GreaterThan,
        5 => FilterOperator::NotEqual,
        6 => FilterOperator::GreaterThanOrEqual,
        _ => FilterOperator::Equal,
    };
    let value = match vts {
        6 => {
            let (s, _) = parser::wide_str(data, 10).ok()?;
            s
        }
        4 => {
            let v = f64::from_le_bytes(data[2..10].try_into().ok()?);
            if v.is_finite() && v == v.trunc() && v.abs() < 1e15 {
                format!("{}", v as i64)
            } else {
                v.to_string()
            }
        }
        _ => String::new(),
    };
    Some(CustomFilterCondition { operator, value })
}

/// Parse BrtMargins record.
///
/// Layout: 6 × f64 = 48 bytes: left, right, top, bottom, header, footer.
fn parse_margins(data: &[u8], ws: &mut Worksheet) {
    if data.len() < 48 {
        return;
    }
    let mut ps = ws.page_setup().clone();
    ps.left_margin = parser::read_f64(data, 0);
    ps.right_margin = parser::read_f64(data, 8);
    ps.top_margin = parser::read_f64(data, 16);
    ps.bottom_margin = parser::read_f64(data, 24);
    ps.header_margin = parser::read_f64(data, 32);
    ps.footer_margin = parser::read_f64(data, 40);
    ws.set_page_setup(ps);
}

/// Parse BrtPageSetup record.
///
/// Layout: [0..4] iPaperSize (u32), [4..8] iScale (u32),
/// [8..12] iRes (u32), [12..16] iVRes (u32),
/// [16..20] iCopies (u32), [20..24] iPageStart (u32),
/// [24..28] iFitWidth (u32), [28..32] iFitHeight (u32),
/// [32..34] grbit (u16): bit 1 = landscape, bit 2 = not valid paper size,
///   bit 3 = black and white, bit 4 = draft, bit 5 = print notes, etc.
fn parse_page_setup(data: &[u8], ws: &mut Worksheet) {
    if data.len() < 34 {
        return;
    }
    let paper_size = parser::read_u32(data, 0);
    let scale = parser::read_u32(data, 4);
    let fit_width = parser::read_u32(data, 24);
    let fit_height = parser::read_u32(data, 28);
    let grbit = parser::read_u16(data, 32);

    let landscape = (grbit & 0x02) != 0;

    let mut ps = ws.page_setup().clone();
    if paper_size > 0 && paper_size <= 255 {
        ps.paper_size = paper_size as u8;
    }
    if scale >= 10 && scale <= 400 {
        ps.scale = scale as u16;
    }
    ps.orientation = if landscape {
        PageOrientation::Landscape
    } else {
        PageOrientation::Portrait
    };
    if fit_width > 0 && fit_width < 0xFFFF {
        ps.fit_to_width = Some(fit_width as u16);
    }
    if fit_height > 0 && fit_height < 0xFFFF {
        ps.fit_to_height = Some(fit_height as u16);
    }
    ws.set_page_setup(ps);
}

/// Parse BrtPrintOptions record.
///
/// Layout: [0..2] grbit (u16): bit 3 = gridlines, bit 2 = headings.
fn parse_print_options(data: &[u8], ws: &mut Worksheet) {
    if data.len() < 2 {
        return;
    }
    let grbit = parser::read_u16(data, 0);
    let mut ps = ws.page_setup().clone();
    ps.print_headings = (grbit & 0x04) != 0;
    ps.print_gridlines = (grbit & 0x08) != 0;
    ws.set_page_setup(ps);
}

/// Parse BrtHeaderFooter record.
///
/// Layout: [0..2] flags (u16): bit 0 = differentOddEven, bit 1 = differentFirst,
///   bit 2 = scaleWithDoc, bit 3 = alignWithMargins.
/// Then 6 XLWideStrings: oddHeader, oddFooter, evenHeader, evenFooter,
///   firstHeader, firstFooter.
fn parse_header_footer(data: &[u8], ws: &mut Worksheet) {
    if data.len() < 2 {
        return;
    }
    let flags = parser::read_u16(data, 0);
    let mut ps = ws.page_setup().clone();
    ps.different_odd_even = (flags & 0x01) != 0;
    ps.different_first = (flags & 0x02) != 0;
    ps.scale_with_doc = (flags & 0x04) != 0;
    ps.align_with_margins = (flags & 0x08) != 0;

    let mut pos = 2;
    let strings: Vec<Option<String>> = (0..6)
        .map(|_| {
            if pos + 4 > data.len() {
                return None;
            }
            // Each field is an XLNullableWideString: cch 0xFFFFFFFF is
            // a 4-byte null marker that must be skipped, not treated
            // as a parse failure (which would desync every later
            // field).
            if parser::read_u32(data, pos) == 0xFFFFFFFF {
                pos += 4;
                return None;
            }
            match parser::wide_str(data, pos) {
                Ok((s, consumed)) => {
                    pos += consumed;
                    if s.is_empty() {
                        None
                    } else {
                        Some(s)
                    }
                }
                Err(_) => None,
            }
        })
        .collect();

    if let Some(s) = strings.get(0) {
        ps.odd_header = s.clone();
    }
    if let Some(s) = strings.get(1) {
        ps.odd_footer = s.clone();
    }
    if let Some(s) = strings.get(2) {
        ps.even_header = s.clone();
    }
    if let Some(s) = strings.get(3) {
        ps.even_footer = s.clone();
    }
    if let Some(s) = strings.get(4) {
        ps.first_header = s.clone();
    }
    if let Some(s) = strings.get(5) {
        ps.first_footer = s.clone();
    }
    ws.set_page_setup(ps);
}

/// Parse formula tokens from a BRT_FMLA_* record and decompile to text.
///
/// `grbit_offset` is where grbit(2 bytes) starts after the cached value.
/// Layout: grbit(u16) + CellParsedFormula
/// CellParsedFormula: cce(u32) + rgce[cce] + cb(u32) + rgcb[cb]
///
/// When cce > 0: rgce contains the formula tokens, rgcb holds extra data
/// (array constants, etc.).
/// When cce == 0: Excel may store the full token stream in rgcb instead.
fn decompile_formula(data: &[u8], grbit_offset: usize, ctx: &FormulaContext) -> Option<String> {
    let cce_offset = grbit_offset + 2;
    if cce_offset + 4 > data.len() {
        return None;
    }
    let cce = parser::read_u32(data, cce_offset) as usize;
    let tokens_start = cce_offset + 4;
    let tokens_end = tokens_start + cce;
    if tokens_end > data.len() {
        return None;
    }

    let (token_bytes, extra_data) = if cce > 0 {
        let rgce = &data[tokens_start..tokens_end];
        let cb_offset = tokens_end;
        let extra = if cb_offset + 4 <= data.len() {
            let cb = parser::read_u32(data, cb_offset) as usize;
            let rgcb_start = cb_offset + 4;
            let rgcb_end = rgcb_start + cb;
            if cb > 0 && rgcb_end <= data.len() {
                &data[rgcb_start..rgcb_end]
            } else {
                &[] as &[u8]
            }
        } else {
            &[] as &[u8]
        };
        (rgce, extra)
    } else {
        // cce == 0: check if rgcb holds the formula tokens
        let cb_offset = tokens_start; // rgce is empty, cb follows immediately
        if cb_offset + 4 > data.len() {
            return None;
        }
        let cb = parser::read_u32(data, cb_offset) as usize;
        let rgcb_start = cb_offset + 4;
        let rgcb_end = rgcb_start + cb;
        if cb == 0 || rgcb_end > data.len() {
            return None;
        }
        (&data[rgcb_start..rgcb_end], &[] as &[u8])
    };

    if token_bytes.is_empty() {
        return None;
    }
    let tokens = token_parser::parse_tokens_with_extra(token_bytes, extra_data);
    if tokens.is_empty() {
        return None;
    }
    let formula = duke_sheets_formula::decompile::decompiler::decompile(&tokens, ctx);
    if formula.is_empty() {
        None
    } else {
        Some(formula)
    }
}

// BrtDVal layout per [MS-XLSB] §2.4.356:
// [0..4] bit-packed header (u32):
//   bits 0-3:   valType (4)
//   bits 4-6:   errStyle (3)
//   bit  7:     unused (1)
//   bit  8:     fAllowBlank (1)
//   bit  9:     fSuppressCombo (1)
//   bits 10-17: mdImeMode (8)
//   bit  18:    fShowInputMsg (1)
//   bit  19:    fShowErrorMsg (1)
//   bits 20-23: typOperator (4)
//   bits 24-31: reserved (8)
// Then sqrfx (UncheckedSqRfX): cFx u32 + cFx × UncheckedRfX (16 bytes)
// Then DValStrings: 4 XLNullableWideStrings
//   (strErrorTitle, strError, strPromptTitle, strPrompt)
// Then formula1 (DVParsedFormula): cce u32 + rgce + cb u32 + rgcb
// Then formula2 (DVParsedFormula): cce u32 + rgce + cb u32 + rgcb
fn parse_data_validation(data: &[u8], ws: &mut Worksheet, _ctx: &FormulaContext) {
    if data.len() < 4 {
        return;
    }
    let header = parser::read_u32(data, 0);
    let val_type_raw = header & 0xF;
    let err_style_raw = (header >> 4) & 0x7;
    let allow_blank = (header & (1 << 8)) != 0;
    let suppress_dropdown = (header & (1 << 9)) != 0;
    let show_input_message = (header & (1 << 18)) != 0;
    let show_error_alert = (header & (1 << 19)) != 0;
    let operator_raw = (header >> 20) & 0xF;

    let operator = match operator_raw {
        0 => ValidationOperator::Between,
        1 => ValidationOperator::NotBetween,
        2 => ValidationOperator::Equal,
        3 => ValidationOperator::NotEqual,
        4 => ValidationOperator::GreaterThan,
        5 => ValidationOperator::LessThan,
        6 => ValidationOperator::GreaterThanOrEqual,
        7 => ValidationOperator::LessThanOrEqual,
        _ => ValidationOperator::Between,
    };

    let error_style = match err_style_raw {
        0 => ValidationErrorStyle::Stop,
        1 => ValidationErrorStyle::Warning,
        2 => ValidationErrorStyle::Information,
        _ => ValidationErrorStyle::Stop,
    };

    let mut pos = 4;
    let ranges = read_sqref(data, &mut pos);

    let error_title = read_nullable_wide_str(data, &mut pos);
    let error_message = read_nullable_wide_str(data, &mut pos);
    let input_title = read_nullable_wide_str(data, &mut pos);
    let input_message = read_nullable_wide_str(data, &mut pos);

    let formula1 = read_dv_parsed_formula(data, &mut pos);
    let formula2 = read_dv_parsed_formula(data, &mut pos);

    let validation_type = match val_type_raw {
        1 => ValidationType::Whole {
            operator,
            value1: formula1.unwrap_or_default(),
            value2: formula2,
        },
        2 => ValidationType::Decimal {
            operator,
            value1: formula1.unwrap_or_default(),
            value2: formula2,
        },
        3 => ValidationType::List {
            source: unquote_dv_list_source(formula1.unwrap_or_default()),
        },
        4 => ValidationType::Date {
            operator,
            value1: formula1.unwrap_or_default(),
            value2: formula2,
        },
        5 => ValidationType::Time {
            operator,
            value1: formula1.unwrap_or_default(),
            value2: formula2,
        },
        6 => ValidationType::TextLength {
            operator,
            value1: formula1.unwrap_or_default(),
            value2: formula2,
        },
        7 => ValidationType::Custom {
            formula: formula1.unwrap_or_default(),
        },
        _ => ValidationType::None,
    };

    let validation = DataValidation {
        validation_type,
        ranges,
        allow_blank,
        show_dropdown: !suppress_dropdown,
        show_input_message,
        input_title,
        input_message,
        show_error_alert,
        error_style,
        error_title,
        error_message,
    };

    ws.add_data_validation(validation);
}

/// Read a DVParsedFormula: cce u32 + rgce + cb u32 + rgcb.
/// Strip the outer quote pair from a literal list source and unescape
/// doubled quotes. A bare `trim_matches('"')` would also eat quotes
/// that belong to the value and never unescape `""`.
fn unquote_dv_list_source(f: String) -> String {
    if f.len() >= 2 && f.starts_with('"') && f.ends_with('"') {
        f[1..f.len() - 1].replace("\"\"", "\"")
    } else {
        f
    }
}

fn read_dv_parsed_formula(data: &[u8], pos: &mut usize) -> Option<String> {
    if *pos + 4 > data.len() {
        return None;
    }
    let cce = parser::read_u32(data, *pos) as usize;
    *pos += 4;

    let rgce_end = *pos + cce;
    if rgce_end > data.len() {
        *pos = data.len();
        return None;
    }
    let token_bytes = &data[*pos..rgce_end];
    *pos = rgce_end;

    if *pos + 4 > data.len() {
        return None;
    }
    let cb = parser::read_u32(data, *pos) as usize;
    *pos += 4;
    let rgcb_end = *pos + cb;
    if rgcb_end > data.len() {
        *pos = data.len();
        return None;
    }
    let extra = &data[*pos..rgcb_end];
    *pos = rgcb_end;

    if cce == 0 {
        return None;
    }

    let tokens = token_parser::parse_tokens_with_extra(token_bytes, extra);
    if tokens.is_empty() {
        return None;
    }
    let formula = duke_sheets_formula::decompile::decompiler::decompile(
        &tokens,
        &duke_sheets_formula::decompile::FormulaContext::new(Vec::new()),
    );
    if formula.is_empty() {
        None
    } else {
        Some(formula)
    }
}

fn read_sqref(data: &[u8], pos: &mut usize) -> Vec<CellRange> {
    if *pos + 4 > data.len() {
        return Vec::new();
    }
    let count = parser::read_u32(data, *pos) as usize;
    *pos += 4;
    // Untrusted count: each range needs 16 bytes, so clamp the
    // reservation to what the record can actually hold.
    let mut ranges = Vec::with_capacity(count.min((data.len() - *pos) / 16));
    for _ in 0..count {
        if *pos + 16 > data.len() {
            break;
        }
        let first_row = parser::read_u32(data, *pos);
        let last_row = parser::read_u32(data, *pos + 4);
        let first_col = parser::read_u32(data, *pos + 8) as u16;
        let last_col = parser::read_u32(data, *pos + 12) as u16;
        *pos += 16;
        ranges.push(CellRange::from_indices(
            first_row, first_col, last_row, last_col,
        ));
    }
    ranges
}

struct CfPendingBase {
    ranges: Vec<CellRange>,
    priority: u32,
    stop_if_true: bool,
    dxf_id: Option<u32>,
}

enum CfVisualAccum {
    ColorScale {
        cfvos: Vec<(CfValueType, f64)>,
        colors: Vec<Color>,
    },
    DataBar {
        cfvos: Vec<(CfValueType, f64)>,
        color: Option<Color>,
        show_value: bool,
    },
    IconSet {
        icon_style: IconSetStyle,
        cfvos: Vec<(CfValueType, f64)>,
        reverse: bool,
        show_value: bool,
    },
}

impl CfPendingBase {
    fn into_rule(self, rule_type: CfRuleType) -> ConditionalFormatRule {
        ConditionalFormatRule {
            rule_type,
            ranges: self.ranges,
            priority: self.priority,
            stop_if_true: self.stop_if_true,
            format: None,
            dxf_id: self.dxf_id,
        }
    }
}

fn parse_cond_fmt_ranges(data: &[u8]) -> Vec<CellRange> {
    if data.len() < 8 {
        return Vec::new();
    }
    let mut pos = 8; // skip ccf(u32) + fPivot(u32)
    read_sqref(data, &mut pos)
}

fn parse_cf_rule_base(
    data: &[u8],
    ws: &mut Worksheet,
    _ctx: &FormulaContext,
    parent_ranges: &[CellRange],
) -> Option<CfPendingBase> {
    if data.len() < 42 {
        return None;
    }
    let i_type = parser::read_u32(data, 0);
    let _i_template = parser::read_u32(data, 4);
    let dxf_id_raw = parser::read_u32(data, 8);
    let priority = parser::read_u32(data, 12);
    let i_param = parser::read_u32(data, 16);
    // reserved1 at 20, reserved2 at 24
    let flags = parser::read_u16(data, 28);

    let stop_if_true = (flags & 0x02) != 0; // bit1
    let above_average = (flags & 0x04) != 0; // bit2
    let bottom = (flags & 0x08) != 0; // bit3
    let percent = (flags & 0x10) != 0; // bit4

    let cb_fmla1 = parser::read_u32(data, 30);
    let cb_fmla2 = parser::read_u32(data, 34);
    let _cb_fmla3 = parser::read_u32(data, 38);

    let dxf_id = if dxf_id_raw == 0xFFFFFFFF {
        None
    } else {
        Some(dxf_id_raw)
    };

    let mut pos = 42;

    let text = read_nullable_wide_str(data, &mut pos).unwrap_or_default();

    let formula1 = if cb_fmla1 != 0 {
        read_cf_parsed_formula(data, &mut pos)
    } else {
        None
    };
    let formula2 = if cb_fmla2 != 0 {
        read_cf_parsed_formula(data, &mut pos)
    } else {
        None
    };

    let operator = match i_param {
        1 => CfOperator::Between,
        2 => CfOperator::NotBetween,
        3 => CfOperator::Equal,
        4 => CfOperator::NotEqual,
        5 => CfOperator::GreaterThan,
        6 => CfOperator::LessThan,
        7 => CfOperator::GreaterThanOrEqual,
        8 => CfOperator::LessThanOrEqual,
        _ => CfOperator::Equal,
    };

    let base = CfPendingBase {
        ranges: parent_ranges.to_vec(),
        priority,
        stop_if_true,
        dxf_id,
    };

    match i_type {
        3 | 4 | 6 => Some(base),
        _ => {
            let rule_type = match i_type {
                1 => CfRuleType::CellIs {
                    operator,
                    formula1: formula1.unwrap_or_default(),
                    formula2,
                },
                2 => CfRuleType::Expression {
                    formula: formula1.unwrap_or_default(),
                },
                5 => CfRuleType::Top10 {
                    rank: i_param,
                    percent,
                    bottom,
                },
                7 => CfRuleType::UniqueValues,
                8 => CfRuleType::DuplicateValues,
                9 => CfRuleType::ContainsText { text },
                10 => CfRuleType::ContainsBlanks,
                11 => CfRuleType::NotContainsBlanks,
                12 => CfRuleType::ContainsErrors,
                13 => CfRuleType::NotContainsErrors,
                14 => CfRuleType::AboveAverage {
                    above: above_average,
                    equal_average: false,
                    std_dev: None,
                },
                15 => CfRuleType::BeginsWith { text },
                16 => CfRuleType::EndsWith { text },
                17 => CfRuleType::TimePeriod {
                    period: time_period_from_param(i_param),
                },
                _ => {
                    log::warn!("unsupported CF rule type {i_type}");
                    return None;
                }
            };
            ws.add_conditional_format(base.into_rule(rule_type));
            None
        }
    }
}

/// Read an XLNullableWideString. Returns None for a NULL string
/// (cchCharacters == 0xFFFFFFFF) or empty string.
fn read_nullable_wide_str(data: &[u8], pos: &mut usize) -> Option<String> {
    if *pos + 4 > data.len() {
        return None;
    }
    let marker = parser::read_u32(data, *pos);
    if marker == 0xFFFFFFFF {
        *pos += 4;
        return None;
    }
    match parser::wide_str(data, *pos) {
        Ok((s, consumed)) => {
            *pos += consumed;
            if s.is_empty() {
                None
            } else {
                Some(s)
            }
        }
        Err(_) => None,
    }
}

fn read_cf_parsed_formula(data: &[u8], pos: &mut usize) -> Option<String> {
    if *pos + 4 > data.len() {
        return None;
    }
    let cce = parser::read_u32(data, *pos) as usize;
    *pos += 4;
    if cce == 0 {
        return None;
    }
    let rgce_end = *pos + cce;
    if rgce_end > data.len() {
        *pos = data.len();
        return None;
    }
    let token_bytes = &data[*pos..rgce_end];
    *pos = rgce_end;

    // cb + rgcb (skip extra data)
    if *pos + 4 <= data.len() {
        let cb = parser::read_u32(data, *pos) as usize;
        *pos += 4;
        *pos += cb.min(data.len().saturating_sub(*pos));
    }

    let tokens = token_parser::parse_tokens(token_bytes);
    if tokens.is_empty() {
        return None;
    }
    let formula = duke_sheets_formula::decompile::decompiler::decompile(
        &tokens,
        &duke_sheets_formula::decompile::FormulaContext::new(Vec::new()),
    );
    if formula.is_empty() {
        None
    } else {
        Some(formula)
    }
}

fn time_period_from_param(v: u32) -> TimePeriod {
    match v {
        0 => TimePeriod::Today,
        1 => TimePeriod::Yesterday,
        2 => TimePeriod::Tomorrow,
        3 => TimePeriod::Last7Days,
        4 => TimePeriod::ThisWeek,
        5 => TimePeriod::LastWeek,
        6 => TimePeriod::NextWeek,
        7 => TimePeriod::ThisMonth,
        8 => TimePeriod::LastMonth,
        9 => TimePeriod::NextMonth,
        _ => TimePeriod::Today,
    }
}

fn cfvo_type_from_u32(v: u32) -> CfValueType {
    match v {
        1 => CfValueType::Num,
        2 => CfValueType::Min,
        3 => CfValueType::Max,
        4 => CfValueType::Percent,
        5 => CfValueType::Percentile,
        7 => CfValueType::Formula,
        _ => CfValueType::Num,
    }
}

fn parse_brt_cfvo(data: &[u8]) -> Option<(CfValueType, f64)> {
    if data.len() < 24 {
        return None;
    }
    let cfvo_type = parser::read_u32(data, 0);
    let vt = cfvo_type_from_u32(cfvo_type);
    let num_value = parser::read_f64(data, 4);
    Some((vt, num_value))
}

fn read_cf_brt_color(data: &[u8], pos: &mut usize) -> Color {
    if *pos + 8 > data.len() {
        return Color::Auto;
    }
    let color_type = data[*pos] >> 1;
    let index = data[*pos + 1];
    let tint_raw = i16::from_le_bytes([data[*pos + 2], data[*pos + 3]]);
    let r = data[*pos + 4];
    let g = data[*pos + 5];
    let b = data[*pos + 6];
    let a = data[*pos + 7];
    *pos += 8;
    match color_type {
        0 => Color::Auto,
        1 => Color::Indexed(index),
        2 => {
            if a == 0xFF {
                Color::Rgb { r, g, b }
            } else {
                Color::Argb { a, r, g, b }
            }
        }
        3 => {
            let tint_i8 = if tint_raw == 0 {
                0i8
            } else {
                ((tint_raw as f64 / 32767.0) * 100.0).round() as i8
            };
            Color::Theme {
                index,
                tint: tint_i8,
            }
        }
        _ => Color::Auto,
    }
}

fn parse_cf_data_bar_begin(data: &[u8]) -> Option<bool> {
    (data.len() >= 3).then(|| data[2] != 0)
}

fn icon_set_from_u32(v: u32) -> IconSetStyle {
    match v {
        0 => IconSetStyle::Arrows3,
        1 => IconSetStyle::Arrows3Gray,
        2 => IconSetStyle::Flags3,
        3 => IconSetStyle::TrafficLights3,
        4 => IconSetStyle::TrafficLights3Black,
        5 => IconSetStyle::Signs3,
        6 => IconSetStyle::Symbols3,
        7 => IconSetStyle::Symbols3Circled,
        8 => IconSetStyle::Arrows4,
        9 => IconSetStyle::Arrows4Gray,
        10 => IconSetStyle::RedToBlack4,
        11 => IconSetStyle::Rating4,
        12 => IconSetStyle::TrafficLights4,
        13 => IconSetStyle::Arrows5,
        14 => IconSetStyle::Arrows5Gray,
        15 => IconSetStyle::Rating5,
        16 => IconSetStyle::Quarters5,
        17 => IconSetStyle::Stars3,
        18 => IconSetStyle::Triangles3,
        19 => IconSetStyle::Boxes5,
        _ => IconSetStyle::Arrows3,
    }
}

fn parse_cf_icon_set_begin(data: &[u8]) -> Option<(IconSetStyle, bool, bool)> {
    if data.len() < 6 {
        return None;
    }
    let icon_style = icon_set_from_u32(parser::read_u32(data, 0));
    let flags = parser::read_u16(data, 4);
    Some((icon_style, flags & (1 << 2) != 0, flags & (1 << 1) == 0))
}

fn cf_value_from_pair((value_type, value): (CfValueType, f64)) -> CfValue {
    let value = match value_type {
        CfValueType::Min | CfValueType::Max => None,
        _ => Some(format_cfvo_value(value)),
    };
    CfValue::new(value_type, value)
}

fn format_cfvo_value(val: f64) -> String {
    if val == val.floor() && val.abs() < 1e15 {
        format!("{}", val as i64)
    } else {
        format!("{}", val)
    }
}

fn parse_cell_error(code: u8) -> CellError {
    match code {
        0x00 => CellError::Null,
        0x07 => CellError::Div0,
        0x0F => CellError::Value,
        0x17 => CellError::Ref,
        0x1D => CellError::Name,
        0x24 => CellError::Num,
        0x2A => CellError::Na,
        0x2B => CellError::GettingData,
        _ => CellError::Value,
    }
}
