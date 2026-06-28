use crate::prelude::*;
use crate::refresh::*;

use super::RenderedPivot;

pub(crate) fn write_rendered_pivot(
    workbook: &mut Workbook,
    job: &PivotJob,
    rendered: RenderedPivot,
) -> Result<()> {
    let sheet_count = workbook.sheet_count();
    let worksheet = workbook
        .worksheet_mut(job.sheet_index)
        .ok_or_else(|| Error::SheetOutOfBounds(job.sheet_index, sheet_count))?;

    if matches!(
        job.pivot.overwrite_policy,
        PivotOverwritePolicy::FailOnOccupied
    ) {
        ensure_output_range_is_available(worksheet, &job.pivot, rendered.range)?;
    }

    if matches!(
        job.pivot.overwrite_policy,
        PivotOverwritePolicy::ClearOwnedRange
    ) {
        if let Some(range) = job.pivot.rendered_range {
            worksheet.clear_range(&range);
        }
    }
    clear_pivot_merged_ranges(worksheet, job.pivot.rendered_range);

    for (row_offset, row) in rendered.cells.iter().enumerate() {
        for (col_offset, value) in row.iter().enumerate() {
            let row = job.pivot.target.row + row_offset as u32;
            let col = job.pivot.target.col + col_offset as u16;
            if value.is_empty() {
                worksheet.clear_cell_at(row, col);
            } else {
                worksheet.set_cell_value_at(row, col, value.clone())?;
                let cell_format = rendered
                    .cell_number_formats
                    .get(row_offset)
                    .and_then(|row| row.get(col_offset))
                    .and_then(Option::as_deref);
                let column_format = (row_offset >= rendered.data_start_row)
                    .then(|| {
                        rendered
                            .column_number_formats
                            .get(col_offset)
                            .and_then(Option::as_deref)
                    })
                    .flatten();
                if let Some(format) = cell_format.or(column_format) {
                    apply_number_format(worksheet, row, col, format)?;
                }
            }
        }
    }
    write_pivot_merged_ranges(worksheet, &rendered)?;
    write_pivot_row_outline_levels(worksheet, &job.pivot, &rendered);
    write_pivot_column_outline_levels(worksheet, &job.pivot, &rendered);
    write_pivot_row_page_breaks(worksheet, &job.pivot, &rendered);

    if let Some(pivot) = worksheet.pivot_tables_mut().get_mut(job.pivot_index) {
        pivot.rendered_range = Some(rendered.range);
        pivot.refresh_status = PivotRefreshStatus::Succeeded;
        pivot.set_cache_refresh_status(PivotRefreshStatus::Succeeded);
    }

    Ok(())
}

pub(crate) fn clear_pivot_merged_ranges(
    worksheet: &mut Worksheet,
    previous_range: Option<CellRange>,
) {
    let Some(previous_range) = previous_range else {
        return;
    };

    let ranges = worksheet
        .merged_regions()
        .iter()
        .copied()
        .filter(|range| range_contains_range(previous_range, *range))
        .collect::<Vec<_>>();
    for range in ranges {
        worksheet.unmerge_cells(&range);
    }
}

pub(crate) fn write_pivot_merged_ranges(
    worksheet: &mut Worksheet,
    rendered: &RenderedPivot,
) -> Result<()> {
    for range in &rendered.merged_ranges {
        worksheet.merge_cells(range)?;
    }
    Ok(())
}

pub(crate) fn write_pivot_row_outline_levels(
    worksheet: &mut Worksheet,
    pivot: &PivotTable,
    rendered: &RenderedPivot,
) {
    clear_pivot_row_outline_levels(worksheet, pivot.rendered_range);
    clear_pivot_row_outline_levels(worksheet, Some(rendered.range));

    for (offset, level) in rendered.row_outline_levels.iter().copied().enumerate() {
        if level != 0 {
            worksheet.set_row_outline_level(pivot.target.row + offset as u32, level);
        }
        if rendered.row_hidden.get(offset).copied().unwrap_or(false) {
            worksheet.set_row_hidden(pivot.target.row + offset as u32, true);
        }
        if rendered.row_collapsed.get(offset).copied().unwrap_or(false) {
            worksheet.set_row_collapsed(pivot.target.row + offset as u32, true);
        }
    }
}

pub(crate) fn clear_pivot_row_outline_levels(worksheet: &mut Worksheet, range: Option<CellRange>) {
    let Some(range) = range else {
        return;
    };
    for row in range.start.row..=range.end.row {
        worksheet.set_row_outline_level(row, 0);
        worksheet.set_row_collapsed(row, false);
        worksheet.set_row_hidden(row, false);
    }
}

pub(crate) fn write_pivot_column_outline_levels(
    worksheet: &mut Worksheet,
    pivot: &PivotTable,
    rendered: &RenderedPivot,
) {
    clear_pivot_column_outline_levels(worksheet, pivot.rendered_range);
    clear_pivot_column_outline_levels(worksheet, Some(rendered.range));

    for (offset, level) in rendered.column_outline_levels.iter().copied().enumerate() {
        if level != 0 {
            worksheet.set_column_outline_level(pivot.target.col + offset as u16, level);
        }
        if rendered.column_hidden.get(offset).copied().unwrap_or(false) {
            worksheet.set_column_hidden(pivot.target.col + offset as u16, true);
        }
        if rendered
            .column_collapsed
            .get(offset)
            .copied()
            .unwrap_or(false)
        {
            worksheet.set_column_collapsed(pivot.target.col + offset as u16, true);
        }
    }
}

pub(crate) fn clear_pivot_column_outline_levels(
    worksheet: &mut Worksheet,
    range: Option<CellRange>,
) {
    let Some(range) = range else {
        return;
    };
    for col in range.start.col..=range.end.col {
        worksheet.set_column_outline_level(col, 0);
        worksheet.set_column_collapsed(col, false);
        worksheet.set_column_hidden(col, false);
    }
}

pub(crate) fn range_contains_range(outer: CellRange, inner: CellRange) -> bool {
    inner.start.row >= outer.start.row
        && inner.end.row <= outer.end.row
        && inner.start.col >= outer.start.col
        && inner.end.col <= outer.end.col
}

pub(crate) fn write_pivot_row_page_breaks(
    worksheet: &mut Worksheet,
    pivot: &PivotTable,
    rendered: &RenderedPivot,
) {
    let mut row_breaks = worksheet.row_breaks().to_vec();
    row_breaks.retain(|break_| {
        !break_.pt
            || (!row_break_is_in_range(break_.id, pivot.rendered_range)
                && !row_break_is_in_range(break_.id, Some(rendered.range)))
    });

    let mut existing_break_rows = row_breaks
        .iter()
        .map(|break_| break_.id)
        .collect::<AHashSet<_>>();
    for offset in &rendered.row_page_break_offsets {
        let row = pivot.target.row + *offset;
        if existing_break_rows.insert(row) {
            row_breaks.push(PageBreak {
                id: row,
                min: 0,
                max: 16383,
                man: true,
                pt: true,
            });
        }
    }

    worksheet.set_row_breaks(row_breaks);
}

pub(crate) fn row_break_is_in_range(row: u32, range: Option<CellRange>) -> bool {
    range
        .map(|range| row >= range.start.row && row <= range.end.row)
        .unwrap_or(false)
}

pub(crate) fn apply_number_format(
    worksheet: &mut Worksheet,
    row: u32,
    col: u16,
    format: &str,
) -> Result<()> {
    let mut style = worksheet
        .cell_style_at(row, col)
        .cloned()
        .unwrap_or_default();
    style.number_format = NumberFormat::Custom(format.to_string());
    worksheet.set_cell_style_at(row, col, &style)
}

pub(crate) fn mark_pivot_failed(
    workbook: &mut Workbook,
    sheet_index: usize,
    pivot_index: usize,
    message: String,
) {
    if let Some(worksheet) = workbook.worksheet_mut(sheet_index) {
        if let Some(pivot) = worksheet.pivot_tables_mut().get_mut(pivot_index) {
            let status = PivotRefreshStatus::Failed { message };
            pivot.refresh_status = status.clone();
            pivot.set_cache_refresh_status(status);
        }
    }
}

pub(crate) fn mark_pivot_external(workbook: &mut Workbook, sheet_index: usize, pivot_index: usize) {
    if let Some(worksheet) = workbook.worksheet_mut(sheet_index) {
        if let Some(pivot) = worksheet.pivot_tables_mut().get_mut(pivot_index) {
            pivot.refresh_status = PivotRefreshStatus::External;
            pivot.set_cache_refresh_status(PivotRefreshStatus::External);
        }
    }
}

pub(crate) fn ensure_output_range_is_available(
    worksheet: &Worksheet,
    pivot: &PivotTable,
    output_range: CellRange,
) -> Result<()> {
    for address in output_range.cells() {
        if pivot
            .rendered_range
            .is_some_and(|owned| owned.contains(&address))
        {
            continue;
        }

        if !worksheet.get_value_at(address.row, address.col).is_blank() {
            return Err(Error::other(format!(
                "pivot table {} would overwrite non-empty cell {}",
                pivot.name, address
            )));
        }
    }

    Ok(())
}
