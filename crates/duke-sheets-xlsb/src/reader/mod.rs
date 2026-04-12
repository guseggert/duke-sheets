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
use std::io::{BufReader, Read, Seek};
use std::path::Path;

use duke_sheets_core::worksheet::SheetVisibility;
use duke_sheets_core::Workbook;
use quick_xml::events::Event;
use quick_xml::reader::Reader;

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
                if !entry.visible {
                    ws.set_visibility(SheetVisibility::Hidden);
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

            if let Ok(Some(dr)) =
                drawing::read_drawing_charts(&mut archive, &entry.path, &sheet_rels)
            {
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

        if wb.sheet_count() == 0 {
            wb.add_worksheet_with_name_unchecked("Sheet1");
        }

        Ok(wb)
    }
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
