mod comments;
mod drawing;
pub(crate) mod shared_strings;
pub(crate) mod styles;
mod table;
#[cfg(test)]
mod tests;
pub(crate) mod theme;
pub(crate) mod workbook;
pub(crate) mod worksheet;

use std::collections::HashMap;
use std::io::{BufReader, Cursor, Read, Seek, Write};
use std::path::Path;

use duke_sheets_core::named_range::{NameScope, NamedRange};
use duke_sheets_core::worksheet::SheetVisibility;
use duke_sheets_core::Workbook;
use quick_xml::events::Event;
use quick_xml::reader::Reader;
use quick_xml::Writer;

use crate::error::{XlsbError, XlsbResult};

#[derive(Debug, Clone)]
pub(crate) struct SheetRel {
    pub target: String,
    pub rel_type: String,
}

pub struct XlsbReader;

impl XlsbReader {
    pub fn read_file<P: AsRef<Path>>(path: P) -> XlsbResult<Workbook> {
        let file = std::fs::File::open(path.as_ref())?;
        Self::read(file)
    }

    pub fn read<R: Read + Seek>(reader: R) -> XlsbResult<Workbook> {
        let mut archive = zip::ZipArchive::new(reader)
            .map_err(|e| XlsbError::InvalidFormat(format!("not a valid ZIP: {e}")))?;

        let styles_data = styles::read_styles(&mut archive)?;
        let mut cell_styles = styles_data.styles;
        let dxf_styles = styles_data.dxf_styles;
        let shared_strings = shared_strings::read_shared_strings(&mut archive, &styles_data.fonts)?;
        let relationships = workbook::read_relationships(&mut archive)?;
        let props = workbook::read_workbook(&mut archive, &relationships)?;

        // Load theme palette from xl/theme/theme1.xml (XML even in XLSB)
        let theme_path = relationships
            .theme_path
            .as_deref()
            .unwrap_or("xl/theme/theme1.xml");
        if let Ok(file) = archive.by_name(theme_path) {
            if let Ok(palette) = theme::parse_theme_palette(file) {
                for style in &mut cell_styles {
                    theme::resolve_style_theme_colors(style, &palette);
                }
            }
        }

        let mut wb = Workbook::empty();
        wb.settings_mut().date_1904 = props.date_1904;

        for entry in &props.sheets {
            wb.add_worksheet_with_name_unchecked(&entry.name);
        }

        if props.active_sheet > 0 && props.active_sheet < props.sheets.len() {
            let _ = wb.set_active_sheet(props.active_sheet);
        }

        for (name, itab, refers_to, hidden, comment) in &props.named_ranges {
            let scope = if *itab == 0xFFFFFFFF {
                NameScope::Workbook
            } else {
                NameScope::Sheet(*itab as usize)
            };
            let mut nr = NamedRange::new(name, refers_to.as_str(), scope);
            if *hidden {
                nr.hidden = true;
            }
            if let Some(c) = comment {
                nr.comment = Some(c.clone());
            }
            wb.named_ranges_mut().define_or_update(nr);
        }

        apply_print_settings(&props, &mut wb);

        for (i, entry) in props.sheets.iter().enumerate() {
            if entry.path.is_empty() {
                continue;
            }

            let sheet_rels = read_sheet_rels(&mut archive, &entry.path)?;

            let hyperlink_rels: HashMap<String, String> = sheet_rels
                .iter()
                .filter(|(_, r)| r.rel_type.ends_with("/hyperlink"))
                .map(|(id, r)| (id.clone(), r.target.clone()))
                .collect();

            {
                let file = match archive.by_name(&entry.path) {
                    Ok(f) => f,
                    Err(_) => {
                        log::warn!("sheet file not found in archive: {}", entry.path);
                        continue;
                    }
                };
                let ws = wb.worksheet_mut(i).unwrap();
                match entry.visibility {
                    1 => ws.set_visibility(SheetVisibility::Hidden),
                    2 => ws.set_visibility(SheetVisibility::VeryHidden),
                    _ => {} // 0 = visible (default)
                }
                worksheet::read_worksheet(
                    file,
                    ws,
                    &shared_strings,
                    &cell_styles,
                    &props.formula_ctx,
                    &hyperlink_rels,
                )?;
            }

            if let Ok(Some(mut dr)) =
                drawing::read_drawing_charts(&mut archive, &entry.path, &sheet_rels)
            {
                // Control shapes in the drawing part are the a14
                // twins of the VML controls parsed below; strip them
                // so the raw passthrough doesn't duplicate them.
                for (_, data) in dr.bundle.entries.iter_mut() {
                    *data = strip_control_anchors(data);
                }
                if !dr.bundle.is_empty() {
                    wb.worksheet_mut(i)
                        .unwrap()
                        .raw_drawing_objects
                        .push(dr.bundle.encode());
                }
                let ws = wb.worksheet_mut(i).unwrap();
                for c in dr.charts {
                    ws.add_chart(c);
                }
                for cx in dr.charts_ex {
                    ws.add_chart_ex(cx);
                }
            }

            // Form controls live in the legacy VML drawing part.
            let vml_path = sheet_rels
                .values()
                .find(|r| r.rel_type.ends_with("/vmlDrawing"))
                .map(|r| resolve_rel_path(&entry.path, &r.target));
            if let Some(vml_path) = vml_path {
                if let Ok(mut f) = archive.by_name(&vml_path) {
                    let mut bytes = Vec::new();
                    if std::io::Read::read_to_end(&mut f, &mut bytes).is_ok() {
                        let ws = wb.worksheet_mut(i).unwrap();
                        for shape in duke_sheets_vml::parse_vml_controls(&bytes) {
                            if let Some(control) = shape.to_form_control() {
                                ws.add_form_control(control);
                            }
                        }
                    }
                }
            }

            let ws_num = i + 1;

            let comments_path = sheet_rels
                .values()
                .find(|r| r.rel_type.ends_with("/comments"))
                .map(|r| resolve_rel_path(&entry.path, &r.target))
                .unwrap_or_else(|| format!("xl/comments{}.bin", ws_num));
            comments::read_comments(&mut archive, &comments_path, wb.worksheet_mut(i).unwrap())?;

            let mut table_paths: Vec<String> = sheet_rels
                .values()
                .filter(|r| r.rel_type.ends_with("/table"))
                .map(|r| resolve_rel_path(&entry.path, &r.target))
                .collect();
            table_paths.sort();
            for table_path in &table_paths {
                if let Some(t) = table::read_table(&mut archive, table_path)? {
                    wb.worksheet_mut(i).unwrap().add_table(t);
                }
            }
        }

        if !dxf_styles.is_empty() {
            for i in 0..wb.sheet_count() {
                let ws = wb.worksheet_mut(i).unwrap();
                for rule in ws.conditional_formats_mut() {
                    if let Some(dxf_id) = rule.dxf_id {
                        if let Some(dxf_style) = dxf_styles.get(dxf_id as usize) {
                            rule.format = Some(dxf_style.clone());
                        }
                    }
                }
            }
        }

        if wb.sheet_count() == 0 {
            wb.add_worksheet_with_name_unchecked("Sheet1");
        }

        Ok(wb)
    }
}

fn apply_print_settings(props: &workbook::WorkbookProps, wb: &mut Workbook) {
    for ps in &props.print_areas {
        let idx = ps.sheet_idx as usize;
        if let Some(ws) = wb.worksheet_mut(idx) {
            let sheet_name = ws.name().to_string();
            if let Some(range) = parse_print_area_formula(&ps.refers_to, &sheet_name) {
                ws.set_print_area(range);
            }
        }
    }
    for ps in &props.print_titles {
        let idx = ps.sheet_idx as usize;
        if let Some(ws) = wb.worksheet_mut(idx) {
            let sheet_name = ws.name().to_string();
            let (rows, cols) = parse_print_titles_formula(&ps.refers_to, &sheet_name);
            if let Some((r1, r2)) = rows {
                ws.set_repeat_rows(r1, r2);
            }
            if let Some((c1, c2)) = cols {
                ws.set_repeat_cols(c1, c2);
            }
        }
    }
}

fn parse_print_area_formula(
    formula: &str,
    _sheet_name: &str,
) -> Option<duke_sheets_core::CellRange> {
    let trimmed = formula.trim().trim_start_matches('=');
    let range_part = trimmed.split('!').next_back()?.trim();
    let first_area = range_part.split(',').next()?.trim();
    let clean = first_area.replace('$', "");
    duke_sheets_core::CellRange::parse(&clean).ok()
}

#[allow(clippy::type_complexity)]
fn parse_print_titles_formula(
    formula: &str,
    _sheet_name: &str,
) -> (Option<(u32, u32)>, Option<(u16, u16)>) {
    let mut rows = None;
    let mut cols = None;

    for part in formula.trim().trim_start_matches('=').split(',') {
        let range_part = match part.split('!').next_back() {
            Some(r) => r.trim(),
            None => continue,
        };
        let clean = range_part.replace('$', "");

        if let Some((start, end)) = clean.split_once(':') {
            let start = start.trim();
            let end = end.trim();
            if start.is_empty() || end.is_empty() {
                continue;
            }

            if start.chars().all(|c| c.is_ascii_digit()) && end.chars().all(|c| c.is_ascii_digit())
            {
                if let (Ok(r1), Ok(r2)) = (start.parse::<u32>(), end.parse::<u32>()) {
                    rows = Some((r1.saturating_sub(1), r2.saturating_sub(1)));
                }
            } else if start.chars().all(|c| c.is_ascii_alphabetic())
                && end.chars().all(|c| c.is_ascii_alphabetic())
            {
                if let (Ok(c1), Ok(c2)) = (
                    duke_sheets_core::CellAddress::letters_to_column(start),
                    duke_sheets_core::CellAddress::letters_to_column(end),
                ) {
                    cols = Some((c1, c2));
                }
            } else if let Ok(range) = duke_sheets_core::CellRange::parse(&clean) {
                let is_full_row = range.start.col == 0 && range.end.col >= 16383;
                let is_full_col = range.start.row == 0 && range.end.row >= 1048575;
                if is_full_row && !is_full_col {
                    rows = Some((range.start.row, range.end.row));
                } else if is_full_col && !is_full_row {
                    cols = Some((range.start.col, range.end.col));
                }
            }
        }
    }

    (rows, cols)
}

/// Remove `mc:AlternateContent` blocks carrying an `a14:compatExt`
/// marker from drawing XML. Those blocks are the DrawingML twins of
/// legacy form controls (which round-trip via the VML part instead);
/// replaying them verbatim would duplicate every control. Blocks
/// without a compatExt (e.g. chartEx) are kept.
fn strip_control_anchors(xml: &[u8]) -> Vec<u8> {
    let mut reader = Reader::from_reader(xml);
    reader.config_mut().trim_text(false);
    let mut output = Writer::new(Cursor::new(Vec::with_capacity(xml.len())));
    let mut capture: Option<(Writer<Cursor<Vec<u8>>>, usize, bool)> = None;
    let mut buf = Vec::new();

    loop {
        let event = match reader.read_event_into(&mut buf) {
            Ok(event) => event.into_owned(),
            Err(_) => return xml.to_vec(),
        };
        match event {
            Event::Start(e) if e.local_name().as_ref() == b"AlternateContent" => {
                if let Some((writer, depth, _)) = capture.as_mut() {
                    *depth += 1;
                    if writer.write_event(Event::Start(e)).is_err() {
                        return xml.to_vec();
                    }
                } else {
                    let mut writer = Writer::new(Cursor::new(Vec::new()));
                    if writer.write_event(Event::Start(e)).is_err() {
                        return xml.to_vec();
                    }
                    capture = Some((writer, 1, false));
                }
            }
            Event::End(e) if capture.is_some() => {
                let is_alternate = e.local_name().as_ref() == b"AlternateContent";
                let (writer, depth, _) = capture.as_mut().unwrap();
                if writer.write_event(Event::End(e)).is_err() {
                    return xml.to_vec();
                }
                if is_alternate {
                    *depth -= 1;
                    if *depth == 0 {
                        let (writer, _, has_compat) = capture.take().unwrap();
                        if !has_compat
                            && output
                                .get_mut()
                                .write_all(&writer.into_inner().into_inner())
                                .is_err()
                        {
                            return xml.to_vec();
                        }
                    }
                }
            }
            Event::Eof => {
                if capture.is_some() {
                    return xml.to_vec();
                }
                break;
            }
            event => {
                if let Some((writer, _, has_compat)) = capture.as_mut() {
                    if matches!(&event, Event::Start(e) | Event::Empty(e) if e.local_name().as_ref() == b"compatExt")
                    {
                        *has_compat = true;
                    }
                    if writer.write_event(event).is_err() {
                        return xml.to_vec();
                    }
                } else if output.write_event(event).is_err() {
                    return xml.to_vec();
                }
            }
        }
        buf.clear();
    }
    output.into_inner().into_inner()
}

fn resolve_rel_path(base_path: &str, rel_target: &str) -> String {
    if rel_target.starts_with('/') {
        return rel_target.trim_start_matches('/').to_string();
    }
    let base_dir = base_path.rsplit_once('/').map(|(d, _)| d).unwrap_or("");
    if rel_target.starts_with("../") {
        let parent = base_dir.rsplit_once('/').map(|(d, _)| d).unwrap_or("");
        let rest = rel_target.trim_start_matches("../");
        if parent.is_empty() {
            rest.to_string()
        } else {
            format!("{}/{}", parent, rest)
        }
    } else {
        if base_dir.is_empty() {
            rel_target.to_string()
        } else {
            format!("{}/{}", base_dir, rel_target)
        }
    }
}

fn read_sheet_rels<R: Read + Seek>(
    archive: &mut zip::ZipArchive<R>,
    sheet_path: &str,
) -> XlsbResult<HashMap<String, SheetRel>> {
    let (base_dir, file_name) = sheet_path.rsplit_once('/').unwrap_or(("", sheet_path));
    let rels_path = format!("{}/_rels/{}.rels", base_dir, file_name);

    let file = match archive.by_name(&rels_path) {
        Ok(f) => f,
        Err(_) => return Ok(HashMap::new()),
    };

    let mut reader = Reader::from_reader(BufReader::new(file));
    reader.config_mut().trim_text(true);
    let mut buf = Vec::new();
    let mut rels = HashMap::new();

    loop {
        match reader.read_event_into(&mut buf) {
            Ok(Event::Empty(ref e)) | Ok(Event::Start(ref e))
                if e.name().as_ref() == b"Relationship" =>
            {
                let mut id = String::new();
                let mut target = String::new();
                let mut rel_type = String::new();
                for attr in e.attributes().flatten() {
                    match attr.key.as_ref() {
                        b"Id" => id = String::from_utf8_lossy(&attr.value).into_owned(),
                        b"Target" => target = String::from_utf8_lossy(&attr.value).into_owned(),
                        b"Type" => rel_type = String::from_utf8_lossy(&attr.value).into_owned(),
                        _ => {}
                    }
                }
                if !id.is_empty() && !target.is_empty() {
                    rels.insert(id, SheetRel { target, rel_type });
                }
            }
            Ok(Event::Eof) => break,
            Err(e) => {
                log::warn!("XML error in sheet .rels: {e}");
                break;
            }
            _ => {}
        }
        buf.clear();
    }
    Ok(rels)
}
