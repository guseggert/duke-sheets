mod comments;
mod drawing;
mod shared_strings;
mod styles;
mod tables;
#[cfg(test)]
mod tests;
mod vml;
mod workbook;
mod worksheet;

use std::collections::{BTreeMap, BTreeSet, HashMap};
use std::io::{Seek, Write};
use std::path::Path;

use zip::write::SimpleFileOptions;
use zip::ZipWriter;

use crate::biff12::compiler::CompileContext;
use crate::error::XlsbResult;
use duke_sheets_core::Workbook;
use duke_sheets_formula::ast::FormulaExpr;
use duke_sheets_formula::decompile::function_table::{
    function_index, name_body_operand_class, OperandClass,
};
use duke_sheets_formula::parse_formula;

const DEFAULT_THEME_XML: &str = r#"<?xml version="1.0"?>
<a:theme xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" name="Office Theme">
  <a:themeElements>
    <a:clrScheme name="Office">
      <a:dk1>
        <a:sysClr val="windowText" lastClr="000000"/>
      </a:dk1>
      <a:lt1>
        <a:sysClr val="window" lastClr="FFFFFF"/>
      </a:lt1>
      <a:dk2>
        <a:srgbClr val="1F497D"/>
      </a:dk2>
      <a:lt2>
        <a:srgbClr val="EEECE1"/>
      </a:lt2>
      <a:accent1>
        <a:srgbClr val="4F81BD"/>
      </a:accent1>
      <a:accent2>
        <a:srgbClr val="C0504D"/>
      </a:accent2>
      <a:accent3>
        <a:srgbClr val="9BBB59"/>
      </a:accent3>
      <a:accent4>
        <a:srgbClr val="8064A2"/>
      </a:accent4>
      <a:accent5>
        <a:srgbClr val="4BACC6"/>
      </a:accent5>
      <a:accent6>
        <a:srgbClr val="F79646"/>
      </a:accent6>
      <a:hlink>
        <a:srgbClr val="0000FF"/>
      </a:hlink>
      <a:folHlink>
        <a:srgbClr val="800080"/>
      </a:folHlink>
    </a:clrScheme>
    <a:fontScheme name="Office">
      <a:majorFont>
        <a:latin typeface="Cambria"/>
        <a:ea typeface=""/>
        <a:cs typeface=""/>
      </a:majorFont>
      <a:minorFont>
        <a:latin typeface="Calibri"/>
        <a:ea typeface=""/>
        <a:cs typeface=""/>
      </a:minorFont>
    </a:fontScheme>
    <a:fmtScheme name="Office">
      <a:fillStyleLst>
        <a:solidFill><a:schemeClr val="phClr"/></a:solidFill>
        <a:solidFill><a:schemeClr val="phClr"/></a:solidFill>
        <a:solidFill><a:schemeClr val="phClr"/></a:solidFill>
      </a:fillStyleLst>
      <a:lnStyleLst>
        <a:ln w="9525"><a:solidFill><a:schemeClr val="phClr"/></a:solidFill></a:ln>
        <a:ln w="25400"><a:solidFill><a:schemeClr val="phClr"/></a:solidFill></a:ln>
        <a:ln w="38100"><a:solidFill><a:schemeClr val="phClr"/></a:solidFill></a:ln>
      </a:lnStyleLst>
      <a:effectStyleLst>
        <a:effectStyle><a:effectLst/></a:effectStyle>
        <a:effectStyle><a:effectLst/></a:effectStyle>
        <a:effectStyle><a:effectLst/></a:effectStyle>
      </a:effectStyleLst>
      <a:bgFillStyleLst>
        <a:solidFill><a:schemeClr val="phClr"/></a:solidFill>
        <a:solidFill><a:schemeClr val="phClr"/></a:solidFill>
        <a:solidFill><a:schemeClr val="phClr"/></a:solidFill>
      </a:bgFillStyleLst>
    </a:fmtScheme>
  </a:themeElements>
  <a:objectDefaults/>
  <a:extraClrSchemeLst/>
</a:theme>
"#;

pub struct XlsbWriter;

impl XlsbWriter {
    pub fn write_file<P: AsRef<Path>>(workbook: &Workbook, path: P) -> XlsbResult<()> {
        let file = std::fs::File::create(path)?;
        Self::write(workbook, file)
    }

    pub fn write<W: Write + Seek>(workbook: &Workbook, writer: W) -> XlsbResult<()> {
        let mut zip = ZipWriter::new(writer);
        let options =
            SimpleFileOptions::default().compression_method(zip::CompressionMethod::Deflated);

        let sst = shared_strings::build_sst(workbook);

        let sheet_names: Vec<String> = (0..workbook.sheet_count())
            .map(|i| workbook.worksheet(i).unwrap().name().to_string())
            .collect();

        let has_formulas = (0..workbook.sheet_count()).any(|i| {
            let ws = workbook.worksheet(i).unwrap();
            ws.iter_cells()
                .any(|(r, c, _)| ws.formula_data_at(r, c).is_some())
        });

        let xlfn_names = collect_xlfn_names(workbook);
        let external_name_list = collect_external_addin_names(workbook);
        let external_names: HashMap<String, u32> = external_name_list
            .iter()
            .enumerate()
            .map(|(idx, name)| (name.to_ascii_uppercase(), (idx + 1) as u32))
            .collect();
        let external_ixti = (!external_names.is_empty()).then_some(0u16);

        // Names emitted to BrtName must be enumerable from the
        // CompileContext so PtgName can resolve text → ilbl. The list
        // must mirror the records write_user_name_records actually
        // emits: a name whose body fails to compile is skipped there,
        // and including it here would shift every later PtgName index.
        let name_ctx = CompileContext {
            sheet_names: sheet_names.clone(),
            xlfn_names: xlfn_names.clone(),
            defined_names: Vec::new(),
            defined_name_classes: Vec::new(),
            external_names: external_names.clone(),
            external_ixti,
        };
        let mut defined_names: Vec<String> = Vec::new();
        // Body class per name: range-bodied names must keep R-class
        // PtgName tokens even at value positions (implicit
        // intersection otherwise; same rule as name bodies).
        let mut defined_name_classes = Vec::new();
        for nr in workbook.named_ranges().iter() {
            if crate::biff12::compiler::compile_name_body(&nr.refers_to, &name_ctx).is_err() {
                continue;
            }
            let body = nr.refers_to.strip_prefix('=').unwrap_or(&nr.refers_to);
            let class = parse_formula(&format!("={body}"))
                .map(|expr| name_body_operand_class(&expr))
                .unwrap_or(OperandClass::V);
            defined_names.push(nr.name.clone());
            defined_name_classes.push(class);
        }

        Self::write_root_rels(&mut zip, &options)?;
        Self::write_workbook_rels(&mut zip, &options, workbook, &sst)?;
        Self::write_doc_props(&mut zip, &options)?;
        workbook::write_workbook(
            &mut zip,
            &options,
            workbook,
            has_formulas,
            &xlfn_names,
            &external_name_list,
        )?;
        let (style_mapping, rich_font_ids, dxf_mapping) =
            styles::write_styles(&mut zip, &options, workbook, &sst.rich_fonts)?;
        shared_strings::write_sst(&mut zip, &options, &sst, &rich_font_ids)?;
        Self::write_theme_xml(&mut zip, &options)?;

        let mut comment_sheet_indices = Vec::new();
        let mut all_drawing_overrides: Vec<(String, String)> = Vec::new();
        let mut media_default_exts: BTreeSet<String> = BTreeSet::new();
        let mut written_media: std::collections::HashSet<String> = std::collections::HashSet::new();

        // A preserved part keeps the path its own relationship names, so
        // everything the writer numbers itself starts above whatever
        // those have already claimed, or the two collide on one path.
        let mut claimed = drawing::ClaimedPartNumbers::default();
        for i in 0..workbook.sheet_count() {
            let ws = workbook.worksheet(i).unwrap();
            for rel in drawing::sheet_raw_rels(ws) {
                if rel.external || rel.part.is_none() {
                    continue;
                }
                claimed.note(&drawing::resolve_rel_target("xl/drawings", &rel.target));
            }
        }
        let mut next_drawing_num = claimed.drawing + 1;
        let mut next_chart_num = claimed.chart + 1;
        let mut next_chartex_num = claimed.chart_ex + 1;
        let mut next_image_num = claimed.image + 1;
        let total_standard_charts: usize = claimed.style.max(
            (0..workbook.sheet_count())
                .filter_map(|i| workbook.worksheet(i))
                .map(|ws| drawing::sheet_charts(ws).len())
                .sum(),
        );

        let mut global_table_num = 1usize;
        let mut table_global_nums: Vec<Vec<usize>> = Vec::new();
        for i in 0..workbook.sheet_count() {
            let ws = workbook.worksheet(i).unwrap();
            let sheet_tables = ws.tables();
            let mut nums = Vec::new();
            for _ in sheet_tables {
                nums.push(global_table_num);
                global_table_num += 1;
            }
            table_global_nums.push(nums);
        }

        for i in 0..workbook.sheet_count() {
            let ws = workbook.worksheet(i).unwrap();
            let compile_ctx = CompileContext {
                sheet_names: sheet_names.clone(),
                xlfn_names: xlfn_names.clone(),
                defined_names: defined_names.clone(),
                defined_name_classes: defined_name_classes.clone(),
                external_names: external_names.clone(),
                external_ixti,
            };

            let has_drawing = drawing::sheet_has_drawing_content(ws);
            let emit_brt_drawing = has_drawing;

            let mut result = worksheet::write_worksheet(
                &mut zip,
                &options,
                i,
                ws,
                &sst,
                &style_mapping,
                &compile_ctx,
                emit_brt_drawing,
                &dxf_mapping,
                table_global_nums[i].len(),
            )?;

            let sheet_tables = ws.tables();
            for (t_idx, gnum) in table_global_nums[i].iter().enumerate() {
                tables::write_table_part(&mut zip, &options, &sheet_tables[t_idx], *gnum)?;

                let next_rid = result.sheet_rels.len() + 1;
                result.sheet_rels.push(worksheet::SheetRel {
                    id: format!("rId{}", next_rid),
                    rel_type:
                        "http://schemas.openxmlformats.org/officeDocument/2006/relationships/table"
                            .to_string(),
                    target: format!("../tables/table{}.bin", gnum),
                    target_mode: None,
                });
            }

            let drawing_result = if has_drawing {
                let numbering = drawing::DrawingNumbering {
                    drawing_num: next_drawing_num,
                    chart_start: next_chart_num,
                    chartex_start: next_chartex_num,
                    image_start: next_image_num,
                    total_standard_charts,
                };
                next_drawing_num += 1;
                next_chart_num += drawing::sheet_charts(ws).len();
                next_chartex_num += ws.chart_ex_count();
                next_image_num += drawing::sheet_image_payloads(ws).len();
                let dr = drawing::write_drawing_parts(
                    &mut zip,
                    &options,
                    ws,
                    i,
                    &numbering,
                    &mut written_media,
                )?;
                all_drawing_overrides.extend(dr.content_type_overrides.iter().cloned());
                media_default_exts.extend(dr.media_default_exts.iter().cloned());
                Some(dr)
            } else {
                None
            };

            let has_drawing_rel = drawing_result
                .as_ref()
                .and_then(|dr| dr.drawing_path.as_ref())
                .is_some();

            if result.has_comments {
                comments::write_comments(&mut zip, &options, i, ws)?;
                comment_sheet_indices.push(i);
            }
            let has_vml =
                vml::write_legacy_vml(&mut zip, &options, i, ws, &workbook.theme_palette())?;

            if !result.sheet_rels.is_empty() || result.has_comments || has_drawing_rel || has_vml {
                let drawing_path = drawing_result
                    .as_ref()
                    .and_then(|dr| dr.drawing_path.clone());
                Self::write_sheet_rels(
                    &mut zip,
                    &options,
                    i,
                    &result.sheet_rels,
                    result.has_comments,
                    has_vml,
                    drawing_path.as_deref(),
                )?;
            }
        }

        Self::write_content_types(
            &mut zip,
            &options,
            workbook,
            &comment_sheet_indices,
            &all_drawing_overrides,
            &media_default_exts,
            &table_global_nums,
        )?;

        zip.finish()?;
        Ok(())
    }

    #[allow(clippy::too_many_arguments)]
    fn write_content_types<W: Write + Seek>(
        zip: &mut ZipWriter<W>,
        options: &SimpleFileOptions,
        workbook: &Workbook,
        comment_sheets: &[usize],
        drawing_overrides: &[(String, String)],
        media_default_exts: &BTreeSet<String>,
        table_global_nums: &[Vec<usize>],
    ) -> XlsbResult<()> {
        zip.start_file("[Content_Types].xml", *options)?;
        let mut xml = String::from(
            "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>\n\
             <Types xmlns=\"http://schemas.openxmlformats.org/package/2006/content-types\">\
             <Default Extension=\"rels\" ContentType=\"application/vnd.openxmlformats-package.relationships+xml\"/>\
             <Default Extension=\"xml\" ContentType=\"application/xml\"/>\
             <Default Extension=\"vml\" ContentType=\"application/vnd.openxmlformats-officedocument.vmlDrawing\"/>\
             <Override PartName=\"/xl/workbook.bin\" ContentType=\"application/vnd.ms-excel.sheet.binary.macroEnabled.main\"/>\
             <Override PartName=\"/xl/styles.bin\" ContentType=\"application/vnd.ms-excel.styles\"/>\
             <Override PartName=\"/xl/sharedStrings.bin\" ContentType=\"application/vnd.ms-excel.sharedStrings\"/>\
             <Override PartName=\"/xl/theme/theme1.xml\" ContentType=\"application/vnd.openxmlformats-officedocument.theme+xml\"/>\
             <Override PartName=\"/docProps/core.xml\" ContentType=\"application/vnd.openxmlformats-package.core-properties+xml\"/>\
             <Override PartName=\"/docProps/app.xml\" ContentType=\"application/vnd.openxmlformats-officedocument.extended-properties+xml\"/>"
        );
        for i in 0..workbook.sheet_count() {
            xml.push_str(&format!(
                "<Override PartName=\"/xl/worksheets/sheet{}.bin\" ContentType=\"application/vnd.ms-excel.worksheet\"/>",
                i + 1
            ));
        }
        for &i in comment_sheets {
            xml.push_str(&format!(
                "<Override PartName=\"/xl/comments{}.bin\" ContentType=\"application/vnd.ms-excel.comments\"/>",
                i + 1
            ));
        }
        for sheet_nums in table_global_nums {
            for &gnum in sheet_nums {
                xml.push_str(&format!(
                    "<Override PartName=\"/xl/tables/table{}.bin\" ContentType=\"application/vnd.ms-excel.table\"/>",
                    gnum
                ));
            }
        }
        // One Override per part, compared without case as OPC does; a
        // part reached from two sheets is still one part.
        let mut seen_overrides = std::collections::HashSet::new();
        for (part_name, ct) in drawing_overrides {
            if !seen_overrides.insert(part_name.to_ascii_lowercase()) {
                continue;
            }
            xml.push_str(&format!(
                "<Override PartName=\"{}\" ContentType=\"{}\"/>",
                part_name, ct
            ));
        }
        for ext in media_default_exts {
            let mime = duke_sheets_chart::ImageFormat::from_extension(ext)
                .map(duke_sheets_chart::drawing_part::image_format_mime)
                .unwrap_or("application/octet-stream");
            xml.push_str(&format!(
                "<Default Extension=\"{}\" ContentType=\"{}\"/>",
                ext, mime
            ));
        }
        xml.push_str("</Types>");
        zip.write_all(xml.as_bytes())?;
        Ok(())
    }

    fn write_root_rels<W: Write + Seek>(
        zip: &mut ZipWriter<W>,
        options: &SimpleFileOptions,
    ) -> XlsbResult<()> {
        zip.start_file("_rels/.rels", *options)?;
        let xml = "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>\n\
            <Relationships xmlns=\"http://schemas.openxmlformats.org/package/2006/relationships\">\
            <Relationship Id=\"rId1\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument\" Target=\"xl/workbook.bin\"/>\
            <Relationship Id=\"rId2\" Type=\"http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties\" Target=\"docProps/core.xml\"/>\
            <Relationship Id=\"rId3\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/extended-properties\" Target=\"docProps/app.xml\"/>\
            </Relationships>";
        zip.write_all(xml.as_bytes())?;
        Ok(())
    }

    fn write_doc_props<W: Write + Seek>(
        zip: &mut ZipWriter<W>,
        options: &SimpleFileOptions,
    ) -> XlsbResult<()> {
        zip.start_file("docProps/core.xml", *options)?;
        let core = "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>\n\
            <cp:coreProperties \
            xmlns:cp=\"http://schemas.openxmlformats.org/package/2006/metadata/core-properties\" \
            xmlns:dc=\"http://purl.org/dc/elements/1.1/\" \
            xmlns:dcterms=\"http://purl.org/dc/terms/\" \
            xmlns:dcmitype=\"http://purl.org/dc/dcmitype/\" \
            xmlns:xsi=\"http://www.w3.org/2001/XMLSchema-instance\">\
            <dcterms:created xsi:type=\"dcterms:W3CDTF\">2024-01-01T00:00:00Z</dcterms:created>\
            <dcterms:modified xsi:type=\"dcterms:W3CDTF\">2024-01-01T00:00:00Z</dcterms:modified>\
            </cp:coreProperties>";
        zip.write_all(core.as_bytes())?;

        zip.start_file("docProps/app.xml", *options)?;
        let app = "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>\n\
            <Properties xmlns=\"http://schemas.openxmlformats.org/officeDocument/2006/extended-properties\" \
            xmlns:vt=\"http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes\">\
            <Application>Duke Sheets</Application>\
            <DocSecurity>0</DocSecurity>\
            <ScaleCrop>false</ScaleCrop>\
            <LinksUpToDate>false</LinksUpToDate>\
            <SharedDoc>false</SharedDoc>\
            <HyperlinksChanged>false</HyperlinksChanged>\
            <AppVersion>16.0300</AppVersion>\
            </Properties>";
        zip.write_all(app.as_bytes())?;
        Ok(())
    }

    fn write_workbook_rels<W: Write + Seek>(
        zip: &mut ZipWriter<W>,
        options: &SimpleFileOptions,
        workbook: &Workbook,
        sst: &shared_strings::SstMap,
    ) -> XlsbResult<()> {
        zip.start_file("xl/_rels/workbook.bin.rels", *options)?;
        let mut xml = String::from(
            "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>\n\
             <Relationships xmlns=\"http://schemas.openxmlformats.org/package/2006/relationships\">"
        );
        let mut rid = 1;
        for i in 0..workbook.sheet_count() {
            xml.push_str(&format!(
                "<Relationship Id=\"rId{}\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet\" Target=\"worksheets/sheet{}.bin\"/>",
                rid, i + 1
            ));
            rid += 1;
        }
        xml.push_str(&format!(
            "<Relationship Id=\"rId{}\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles\" Target=\"styles.bin\"/>",
            rid
        ));
        rid += 1;
        if !sst.is_empty() {
            xml.push_str(&format!(
                "<Relationship Id=\"rId{}\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/sharedStrings\" Target=\"sharedStrings.bin\"/>",
                rid
            ));
            rid += 1;
        }
        xml.push_str(&format!(
            "<Relationship Id=\"rId{}\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/theme\" Target=\"theme/theme1.xml\"/>",
            rid
        ));
        xml.push_str("</Relationships>");
        zip.write_all(xml.as_bytes())?;
        Ok(())
    }

    fn write_sheet_rels<W: Write + Seek>(
        zip: &mut ZipWriter<W>,
        options: &SimpleFileOptions,
        sheet_index: usize,
        rels: &[worksheet::SheetRel],
        has_comments: bool,
        has_vml: bool,
        drawing_path: Option<&str>,
    ) -> XlsbResult<()> {
        let path = format!("xl/worksheets/_rels/sheet{}.bin.rels", sheet_index + 1);
        zip.start_file(&path, *options)?;
        let mut xml = String::from(
            "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>\n\
             <Relationships xmlns=\"http://schemas.openxmlformats.org/package/2006/relationships\">"
        );

        for rel in rels {
            xml.push_str(&format!(
                "<Relationship Id=\"{}\" Type=\"{}\" Target=\"{}\"",
                rel.id, rel.rel_type, rel.target
            ));
            if let Some(mode) = &rel.target_mode {
                xml.push_str(&format!(" TargetMode=\"{}\"", mode));
            }
            xml.push_str("/>");
        }

        let mut next_rid = rels.len() + 1;

        if has_comments {
            xml.push_str(&format!(
                "<Relationship Id=\"rId{}\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/comments\" Target=\"../comments{}.bin\"/>",
                next_rid, sheet_index + 1
            ));
            next_rid += 1;
        }

        if has_vml {
            xml.push_str(&format!(
                "<Relationship Id=\"rId{}\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/vmlDrawing\" Target=\"../drawings/vmlDrawing{}.vml\"/>",
                next_rid, sheet_index + 1
            ));
            next_rid += 1;
        }

        if let Some(dp) = drawing_path {
            let rel_target = Self::make_relative_path(
                &format!("xl/worksheets/sheet{}.bin", sheet_index + 1),
                dp,
            );
            xml.push_str(&format!(
                "<Relationship Id=\"rId{}\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/drawing\" Target=\"{}\"/>",
                next_rid, rel_target
            ));
        }

        xml.push_str("</Relationships>");
        zip.write_all(xml.as_bytes())?;
        Ok(())
    }

    fn make_relative_path(from: &str, to: &str) -> String {
        let from_dir = from.rsplit_once('/').map(|(d, _)| d).unwrap_or("");
        let to_dir = to.rsplit_once('/').map(|(d, _)| d).unwrap_or("");
        let to_file = to.rsplit_once('/').map(|(_, f)| f).unwrap_or(to);

        if from_dir == to_dir {
            return to_file.to_string();
        }

        let from_parts: Vec<&str> = from_dir.split('/').collect();
        let to_parts: Vec<&str> = to_dir.split('/').collect();

        let common = from_parts
            .iter()
            .zip(to_parts.iter())
            .take_while(|(a, b)| a == b)
            .count();

        let ups = from_parts.len() - common;
        let mut rel = String::new();
        for _ in 0..ups {
            rel.push_str("../");
        }
        for part in &to_parts[common..] {
            rel.push_str(part);
            rel.push('/');
        }
        rel.push_str(to_file);
        rel
    }

    fn write_theme_xml<W: Write + Seek>(
        zip: &mut ZipWriter<W>,
        options: &SimpleFileOptions,
    ) -> XlsbResult<()> {
        zip.start_file("xl/theme/theme1.xml", *options)?;
        zip.write_all(DEFAULT_THEME_XML.as_bytes())?;
        Ok(())
    }
}

fn collect_xlfn_names(workbook: &Workbook) -> HashMap<String, u32> {
    let mut names = BTreeSet::new();
    for i in 0..workbook.sheet_count() {
        let ws = workbook.worksheet(i).unwrap();
        for (r, c, _) in ws.iter_cells() {
            if let Some(fd) = ws.formula_data_at(r, c) {
                let text = if fd.text.starts_with('=') {
                    &fd.text
                } else {
                    &format!("={}", fd.text)
                };
                if let Ok(expr) = parse_formula(text) {
                    collect_xlfn_from_expr(&expr, &mut names);
                }
            }
        }
    }
    let mut map = HashMap::new();
    for (idx, name) in names.into_iter().enumerate() {
        map.insert(name, (idx + 1) as u32);
    }
    map
}

fn collect_xlfn_from_expr(expr: &FormulaExpr, names: &mut BTreeSet<String>) {
    match expr {
        FormulaExpr::Function { name, args } => {
            let lookup = name
                .strip_prefix("_xlfn.")
                .or_else(|| name.strip_prefix("_XLFN."))
                .unwrap_or(name);
            if function_index(lookup).is_none() {
                names.insert(lookup.to_ascii_uppercase());
            }
            for arg in args {
                collect_xlfn_from_expr(arg, names);
            }
        }
        FormulaExpr::BinaryOp { left, right, .. } => {
            collect_xlfn_from_expr(left, names);
            collect_xlfn_from_expr(right, names);
        }
        FormulaExpr::UnaryOp { operand, .. } => {
            collect_xlfn_from_expr(operand, names);
        }
        FormulaExpr::Array(rows) => {
            for row in rows {
                for cell in row {
                    collect_xlfn_from_expr(cell, names);
                }
            }
        }
        FormulaExpr::ExternalFunction { args, .. } => {
            for arg in args {
                collect_xlfn_from_expr(arg, names);
            }
        }
        _ => {}
    }
}

fn collect_external_addin_names(workbook: &Workbook) -> Vec<String> {
    let mut names = BTreeMap::new();
    for i in 0..workbook.sheet_count() {
        let ws = workbook.worksheet(i).unwrap();
        for (r, c, _) in ws.iter_cells() {
            if let Some(fd) = ws.formula_data_at(r, c) {
                let text = if fd.text.starts_with('=') {
                    fd.text.clone()
                } else {
                    format!("={}", fd.text)
                };
                if let Ok(expr) = parse_formula(&text) {
                    collect_external_addin_names_expr(&expr, &mut names);
                }
            }
        }
    }
    names.into_values().collect()
}

fn collect_external_addin_names_expr(expr: &FormulaExpr, names: &mut BTreeMap<String, String>) {
    match expr {
        FormulaExpr::ExternalFunction { name, args, .. } => {
            names
                .entry(name.to_ascii_uppercase())
                .or_insert_with(|| name.clone());
            for arg in args {
                collect_external_addin_names_expr(arg, names);
            }
        }
        FormulaExpr::Function { args, .. } => {
            for arg in args {
                collect_external_addin_names_expr(arg, names);
            }
        }
        FormulaExpr::BinaryOp { left, right, .. } => {
            collect_external_addin_names_expr(left, names);
            collect_external_addin_names_expr(right, names);
        }
        FormulaExpr::UnaryOp { operand, .. } => collect_external_addin_names_expr(operand, names),
        FormulaExpr::Array(rows) => {
            for row in rows {
                for cell in row {
                    collect_external_addin_names_expr(cell, names);
                }
            }
        }
        _ => {}
    }
}
