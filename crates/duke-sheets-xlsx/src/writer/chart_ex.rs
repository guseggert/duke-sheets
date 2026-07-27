use std::io::{Seek, Write};

use duke_sheets_chart::chart_ex::ChartEx;

use super::XlsxResult;
use crate::error::XlsxError;

pub(super) fn write_chart_ex_part<W: Write + Seek>(
    zip: &mut zip::ZipWriter<W>,
    chart_ex: &ChartEx,
    global_num: usize,
) -> XlsxResult<()> {
    let path = format!("xl/charts/chartEx{}.xml", global_num);
    let bytes = duke_sheets_chart::write::chart_ex_part_bytes(chart_ex)?;
    zip.start_file(&path, zip::write::SimpleFileOptions::default())?;
    zip.write_all(&bytes)?;
    Ok(())
}

/// Write the chart style and chart colour style parts that accompany a
/// chartEx part.
///
/// Both are mandatory in practice: Excel refuses to open a workbook
/// whose chartEx part has no chart style sibling or no chart colour
/// style sibling, so a chart built through the model - which has no raw
/// bytes to replay - gets generated defaults rather than no part at all.
pub(super) fn write_chart_ex_style_color_parts<W: Write + Seek>(
    zip: &mut zip::ZipWriter<W>,
    chart_ex: &ChartEx,
    style_color_num: usize,
) -> XlsxResult<()> {
    let options = zip::write::SimpleFileOptions::default();

    let style = match chart_ex.raw_chart_style {
        Some(ref bytes) => {
            duke_sheets_chart::write::validate_chart_style_part(bytes).map_err(|e| {
                XlsxError::InvalidFormat(format!(
                    "chartEx raw_chart_style is not a part Excel will accept: {e}. \
                     Leave it unset to emit a generated default."
                ))
            })?;
            bytes.clone()
        }
        None => duke_sheets_chart::write::default_chart_style_bytes(),
    };
    zip.start_file(format!("xl/charts/style{style_color_num}.xml"), options)?;
    zip.write_all(&style)?;

    let colors = match chart_ex.raw_chart_color_style {
        Some(ref bytes) => {
            duke_sheets_chart::write::validate_chart_color_style_part(bytes).map_err(|e| {
                XlsxError::InvalidFormat(format!(
                    "chartEx raw_chart_color_style is not a part Excel will accept: {e}. \
                     Leave it unset to emit a generated default."
                ))
            })?;
            bytes.clone()
        }
        None => duke_sheets_chart::write::default_chart_color_style_bytes(),
    };
    zip.start_file(format!("xl/charts/colors{style_color_num}.xml"), options)?;
    zip.write_all(&colors)?;

    Ok(())
}
