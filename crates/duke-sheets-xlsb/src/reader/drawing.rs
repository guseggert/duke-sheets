use std::collections::HashMap;
use std::io::{BufReader, Cursor, Read, Seek};

use quick_xml::events::Event;
use quick_xml::reader::Reader;

use duke_sheets_chart::{Chart, ChartEx};

use crate::drawing_bundle::DrawingBundle;
use crate::error::XlsbResult;

use super::SheetRel;

pub(crate) struct DrawingResult {
    pub bundle: DrawingBundle,
    pub charts: Vec<Chart>,
    pub charts_ex: Vec<ChartEx>,
}

pub(crate) fn read_drawing_charts<R: Read + Seek>(
    archive: &mut zip::ZipArchive<R>,
    sheet_path: &str,
    sheet_rels: &HashMap<String, SheetRel>,
) -> XlsbResult<Option<DrawingResult>> {
    let drawing_rel = sheet_rels
        .values()
        .find(|r| r.rel_type.ends_with("/drawing"));
    let drawing_rel = match drawing_rel {
        Some(r) => r,
        None => return Ok(None),
    };

    let drawing_path = super::resolve_rel_path(sheet_path, &drawing_rel.target);
    let drawing_bytes = match read_zip_entry(archive, &drawing_path) {
        Some(b) => b,
        None => return Ok(None),
    };

    let mut bundle = DrawingBundle::new();
    bundle.push(drawing_path.clone(), drawing_bytes);

    let mut charts = Vec::new();
    let mut charts_ex = Vec::new();

    let drawing_rels_path = rels_path_for(&drawing_path);
    if let Some(rels_bytes) = read_zip_entry(archive, &drawing_rels_path) {
        let targets = parse_rels_targets(&rels_bytes);
        bundle.push(drawing_rels_path, rels_bytes);

        for (target, rel_type) in &targets {
            let part_path = super::resolve_rel_path(&drawing_path, target);

            let is_chart = rel_type.ends_with("/chart");
            let is_chart_ex = rel_type.ends_with("/chartEx");

            if is_chart || is_chart_ex {
                if let Some(chart_bytes) = read_zip_entry(archive, &part_path) {
                    bundle.push(part_path.clone(), chart_bytes.clone());

                    if is_chart_ex {
                        if let Ok(cx) =
                            duke_sheets_chart::parse::parse_chart_ex_xml(Cursor::new(&chart_bytes))
                        {
                            charts_ex.push(cx);
                        }
                    } else if let Ok(c) =
                        duke_sheets_chart::parse::parse_chart_xml(Cursor::new(&chart_bytes))
                    {
                        charts.push(c);
                    }

                    let part_rels_path = rels_path_for(&part_path);
                    if let Some(part_rels_bytes) = read_zip_entry(archive, &part_rels_path) {
                        let sub_targets = parse_rels_targets(&part_rels_bytes);
                        bundle.push(part_rels_path, part_rels_bytes);

                        for (sub_target, _) in &sub_targets {
                            let sub_path = super::resolve_rel_path(&part_path, sub_target);
                            if let Some(sub_bytes) = read_zip_entry(archive, &sub_path) {
                                bundle.push(sub_path, sub_bytes);
                            }
                        }
                    }
                }
            } else if let Some(part_bytes) = read_zip_entry(archive, &part_path) {
                bundle.push(part_path.clone(), part_bytes);

                let part_rels_path = rels_path_for(&part_path);
                if let Some(part_rels_bytes) = read_zip_entry(archive, &part_rels_path) {
                    let sub_targets = parse_rels_targets(&part_rels_bytes);
                    bundle.push(part_rels_path, part_rels_bytes);

                    for (sub_target, _) in &sub_targets {
                        let sub_path = super::resolve_rel_path(&part_path, sub_target);
                        if let Some(sub_bytes) = read_zip_entry(archive, &sub_path) {
                            bundle.push(sub_path, sub_bytes);
                        }
                    }
                }
            }
        }
    }

    collect_media_entries(archive, &mut bundle);

    Ok(Some(DrawingResult {
        bundle,
        charts,
        charts_ex,
    }))
}

fn read_zip_entry<R: Read + Seek>(archive: &mut zip::ZipArchive<R>, path: &str) -> Option<Vec<u8>> {
    let mut file = archive.by_name(path).ok()?;
    let mut bytes = Vec::new();
    file.read_to_end(&mut bytes).ok()?;
    Some(bytes)
}

fn rels_path_for(path: &str) -> String {
    let (dir, name) = path.rsplit_once('/').unwrap_or(("", path));
    if dir.is_empty() {
        format!("_rels/{}.rels", name)
    } else {
        format!("{}/_rels/{}.rels", dir, name)
    }
}

fn parse_rels_targets(rels_xml: &[u8]) -> Vec<(String, String)> {
    let mut reader = Reader::from_reader(BufReader::new(rels_xml));
    reader.config_mut().trim_text(true);
    let mut buf = Vec::new();
    let mut targets = Vec::new();

    loop {
        match reader.read_event_into(&mut buf) {
            Ok(Event::Empty(ref e)) | Ok(Event::Start(ref e))
                if e.name().as_ref() == b"Relationship" =>
            {
                let mut target = String::new();
                let mut rel_type = String::new();
                for attr in e.attributes().flatten() {
                    match attr.key.as_ref() {
                        b"Target" => target = String::from_utf8_lossy(&attr.value).into_owned(),
                        b"Type" => rel_type = String::from_utf8_lossy(&attr.value).into_owned(),
                        _ => {}
                    }
                }
                if !target.is_empty() {
                    targets.push((target, rel_type));
                }
            }
            Ok(Event::Eof) => break,
            Err(_) => break,
            _ => {}
        }
        buf.clear();
    }
    targets
}

fn collect_media_entries<R: Read + Seek>(
    archive: &mut zip::ZipArchive<R>,
    bundle: &mut DrawingBundle,
) {
    let existing: std::collections::HashSet<String> =
        bundle.entries.iter().map(|(p, _)| p.clone()).collect();

    let media_paths: Vec<String> = (0..archive.len())
        .filter_map(|i| {
            let name = archive.by_index(i).ok()?.name().to_string();
            if name.starts_with("xl/media/") && !existing.contains(&name) {
                Some(name)
            } else {
                None
            }
        })
        .collect();

    for path in media_paths {
        if let Some(bytes) = read_zip_entry(archive, &path) {
            bundle.push(path, bytes);
        }
    }
}
