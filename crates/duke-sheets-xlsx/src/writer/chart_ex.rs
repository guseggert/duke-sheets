use std::io::{Seek, Write};

use duke_sheets_chart::chart_ex::ChartEx;

use super::XlsxResult;

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

pub(super) fn write_chart_ex_style_color_parts<W: Write + Seek>(
    zip: &mut zip::ZipWriter<W>,
    chart_ex: &ChartEx,
    chart_ex_num: usize,
    style_color_num: usize,
) -> XlsxResult<()> {
    let has_style = chart_ex.raw_chart_style.is_some();
    let has_color = chart_ex.raw_chart_color_style.is_some();
    if !has_style && !has_color {
        return Ok(());
    }
    let options = zip::write::SimpleFileOptions::default();
    let mut rel_id = 1u32;
    let mut rels_xml = String::from(
        r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">"#,
    );
    if let Some(ref bytes) = chart_ex.raw_chart_style {
        let style_path = format!("xl/charts/style{}.xml", style_color_num);
        zip.start_file(&style_path, options)?;
        zip.write_all(bytes)?;
        rels_xml.push_str(&format!(
            r#"<Relationship Id="rId{}" Type="{}" Target="style{}.xml"/>"#,
            rel_id,
            super::RT_CHART_STYLE,
            style_color_num
        ));
        rel_id += 1;
    }
    if let Some(ref bytes) = chart_ex.raw_chart_color_style {
        let color_path = format!("xl/charts/colors{}.xml", style_color_num);
        zip.start_file(&color_path, options)?;
        zip.write_all(bytes)?;
        rels_xml.push_str(&format!(
            r#"<Relationship Id="rId{}" Type="{}" Target="colors{}.xml"/>"#,
            rel_id,
            super::RT_CHART_COLOR_STYLE,
            style_color_num
        ));
    }
    rels_xml.push_str("</Relationships>");
    let rels_path = format!("xl/charts/_rels/chartEx{}.xml.rels", chart_ex_num);
    zip.start_file(&rels_path, options)?;
    zip.write_all(rels_xml.as_bytes())?;
    Ok(())
}
