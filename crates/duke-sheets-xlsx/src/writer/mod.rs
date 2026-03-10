//! XLSX writer — generates OOXML SpreadsheetML using quick-xml Writer API.

use std::collections::{BTreeMap, BTreeSet, HashMap, HashSet};
use std::fs::File;
use std::io::{Cursor, Seek, Write};
use std::path::Path;

use quick_xml::events::{BytesDecl, BytesEnd, BytesStart, BytesText, Event};
use quick_xml::Writer;

use crate::error::{XlsxError, XlsxResult};
use crate::styles::{roundtrip_theme_data_for, XlsxStyleTable};
use duke_sheets_core::style::Color;
use duke_sheets_core::{CellAddress, CellRange, Workbook};

mod comments;
mod conditional_format;
mod data_validation;
mod tables;

// ---------------------------------------------------------------------------
// OOXML namespace URIs
// ---------------------------------------------------------------------------
const NS_SPREADSHEET: &str = "http://schemas.openxmlformats.org/spreadsheetml/2006/main";
const NS_RELATIONSHIPS: &str = "http://schemas.openxmlformats.org/package/2006/relationships";
const NS_DOC_RELS: &str = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
const NS_CONTENT_TYPES: &str = "http://schemas.openxmlformats.org/package/2006/content-types";

// Relationship types
const RT_OFFICE_DOC: &str =
    "http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument";
const RT_WORKSHEET: &str =
    "http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet";
const RT_STYLES: &str =
    "http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles";
const RT_SHARED_STRINGS: &str =
    "http://schemas.openxmlformats.org/officeDocument/2006/relationships/sharedStrings";
const RT_THEME: &str = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/theme";
const RT_COMMENTS: &str =
    "http://schemas.openxmlformats.org/officeDocument/2006/relationships/comments";
const RT_VML_DRAWING: &str =
    "http://schemas.openxmlformats.org/officeDocument/2006/relationships/vmlDrawing";
const RT_HYPERLINK: &str =
    "http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink";
const RT_TABLE: &str = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/table";
const RT_SHEET_METADATA: &str =
    "http://schemas.openxmlformats.org/officeDocument/2006/relationships/sheetMetadata";

// Content types
const CT_WORKBOOK: &str =
    "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml";
const CT_STYLES: &str = "application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml";
const CT_SST: &str =
    "application/vnd.openxmlformats-officedocument.spreadsheetml.sharedStrings+xml";
const CT_WORKSHEET: &str =
    "application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml";
const CT_COMMENTS: &str =
    "application/vnd.openxmlformats-officedocument.spreadsheetml.comments+xml";
const CT_THEME: &str = "application/vnd.openxmlformats-officedocument.theme+xml";
const CT_RELS: &str = "application/vnd.openxmlformats-package.relationships+xml";
const CT_TABLE: &str = "application/vnd.openxmlformats-officedocument.spreadsheetml.table+xml";
const CT_METADATA: &str =
    "application/vnd.openxmlformats-officedocument.spreadsheetml.metadata+xml";

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
        <a:font script="Jpan" typeface="&#xFF2D;&#xFF33; &#xFF30;&#x30B4;&#x30B7;&#x30C3;&#x30AF;"/>
        <a:font script="Hang" typeface="&#xB9D1;&#xC740; &#xACE0;&#xB515;"/>
        <a:font script="Hans" typeface="&#x5B8B;&#x4F53;"/>
        <a:font script="Hant" typeface="&#x65B0;&#x7D30;&#x660E;&#x9AD4;"/>
        <a:font script="Arab" typeface="Times New Roman"/>
        <a:font script="Hebr" typeface="Times New Roman"/>
        <a:font script="Thai" typeface="Tahoma"/>
        <a:font script="Ethi" typeface="Nyala"/>
        <a:font script="Beng" typeface="Vrinda"/>
        <a:font script="Gujr" typeface="Shruti"/>
        <a:font script="Khmr" typeface="MoolBoran"/>
        <a:font script="Knda" typeface="Tunga"/>
        <a:font script="Guru" typeface="Raavi"/>
        <a:font script="Cans" typeface="Euphemia"/>
        <a:font script="Cher" typeface="Plantagenet Cherokee"/>
        <a:font script="Yiii" typeface="Microsoft Yi Baiti"/>
        <a:font script="Tibt" typeface="Microsoft Himalaya"/>
        <a:font script="Thaa" typeface="MV Boli"/>
        <a:font script="Deva" typeface="Mangal"/>
        <a:font script="Telu" typeface="Gautami"/>
        <a:font script="Taml" typeface="Latha"/>
        <a:font script="Syrc" typeface="Estrangelo Edessa"/>
        <a:font script="Orya" typeface="Kalinga"/>
        <a:font script="Mlym" typeface="Kartika"/>
        <a:font script="Laoo" typeface="DokChampa"/>
        <a:font script="Sinh" typeface="Iskoola Pota"/>
        <a:font script="Mong" typeface="Mongolian Baiti"/>
        <a:font script="Viet" typeface="Times New Roman"/>
        <a:font script="Uigh" typeface="Microsoft Uighur"/>
      </a:majorFont>
      <a:minorFont>
        <a:latin typeface="Calibri"/>
        <a:ea typeface=""/>
        <a:cs typeface=""/>
        <a:font script="Jpan" typeface="&#xFF2D;&#xFF33; &#xFF30;&#x30B4;&#x30B7;&#x30C3;&#x30AF;"/>
        <a:font script="Hang" typeface="&#xB9D1;&#xC740; &#xACE0;&#xB515;"/>
        <a:font script="Hans" typeface="&#x5B8B;&#x4F53;"/>
        <a:font script="Hant" typeface="&#x65B0;&#x7D30;&#x660E;&#x9AD4;"/>
        <a:font script="Arab" typeface="Arial"/>
        <a:font script="Hebr" typeface="Arial"/>
        <a:font script="Thai" typeface="Tahoma"/>
        <a:font script="Ethi" typeface="Nyala"/>
        <a:font script="Beng" typeface="Vrinda"/>
        <a:font script="Gujr" typeface="Shruti"/>
        <a:font script="Khmr" typeface="DaunPenh"/>
        <a:font script="Knda" typeface="Tunga"/>
        <a:font script="Guru" typeface="Raavi"/>
        <a:font script="Cans" typeface="Euphemia"/>
        <a:font script="Cher" typeface="Plantagenet Cherokee"/>
        <a:font script="Yiii" typeface="Microsoft Yi Baiti"/>
        <a:font script="Tibt" typeface="Microsoft Himalaya"/>
        <a:font script="Thaa" typeface="MV Boli"/>
        <a:font script="Deva" typeface="Mangal"/>
        <a:font script="Telu" typeface="Gautami"/>
        <a:font script="Taml" typeface="Latha"/>
        <a:font script="Syrc" typeface="Estrangelo Edessa"/>
        <a:font script="Orya" typeface="Kalinga"/>
        <a:font script="Mlym" typeface="Kartika"/>
        <a:font script="Laoo" typeface="DokChampa"/>
        <a:font script="Sinh" typeface="Iskoola Pota"/>
        <a:font script="Mong" typeface="Mongolian Baiti"/>
        <a:font script="Viet" typeface="Arial"/>
        <a:font script="Uigh" typeface="Microsoft Uighur"/>
      </a:minorFont>
    </a:fontScheme>
    <a:fmtScheme name="Office">
      <a:fillStyleLst>
        <a:solidFill>
          <a:schemeClr val="phClr"/>
        </a:solidFill>
        <a:gradFill rotWithShape="1">
          <a:gsLst>
            <a:gs pos="0">
              <a:schemeClr val="phClr">
                <a:tint val="50000"/>
                <a:satMod val="300000"/>
              </a:schemeClr>
            </a:gs>
            <a:gs pos="35000">
              <a:schemeClr val="phClr">
                <a:tint val="37000"/>
                <a:satMod val="300000"/>
              </a:schemeClr>
            </a:gs>
            <a:gs pos="100000">
              <a:schemeClr val="phClr">
                <a:tint val="15000"/>
                <a:satMod val="350000"/>
              </a:schemeClr>
            </a:gs>
          </a:gsLst>
          <a:lin ang="16200000" scaled="1"/>
        </a:gradFill>
        <a:gradFill rotWithShape="1">
          <a:gsLst>
            <a:gs pos="0">
              <a:schemeClr val="phClr">
                <a:shade val="51000"/>
                <a:satMod val="130000"/>
              </a:schemeClr>
            </a:gs>
            <a:gs pos="80000">
              <a:schemeClr val="phClr">
                <a:shade val="93000"/>
                <a:satMod val="130000"/>
              </a:schemeClr>
            </a:gs>
            <a:gs pos="100000">
              <a:schemeClr val="phClr">
                <a:shade val="94000"/>
                <a:satMod val="135000"/>
              </a:schemeClr>
            </a:gs>
          </a:gsLst>
          <a:lin ang="16200000" scaled="0"/>
        </a:gradFill>
      </a:fillStyleLst>
      <a:lnStyleLst>
        <a:ln w="9525" cap="flat" cmpd="sng" algn="ctr">
          <a:solidFill>
            <a:schemeClr val="phClr">
              <a:shade val="95000"/>
              <a:satMod val="105000"/>
            </a:schemeClr>
          </a:solidFill>
          <a:prstDash val="solid"/>
        </a:ln>
        <a:ln w="25400" cap="flat" cmpd="sng" algn="ctr">
          <a:solidFill>
            <a:schemeClr val="phClr"/>
          </a:solidFill>
          <a:prstDash val="solid"/>
        </a:ln>
        <a:ln w="38100" cap="flat" cmpd="sng" algn="ctr">
          <a:solidFill>
            <a:schemeClr val="phClr"/>
          </a:solidFill>
          <a:prstDash val="solid"/>
        </a:ln>
      </a:lnStyleLst>
      <a:effectStyleLst>
        <a:effectStyle>
          <a:effectLst>
            <a:outerShdw blurRad="40000" dist="20000" dir="5400000" rotWithShape="0">
              <a:srgbClr val="000000">
                <a:alpha val="38000"/>
              </a:srgbClr>
            </a:outerShdw>
          </a:effectLst>
        </a:effectStyle>
        <a:effectStyle>
          <a:effectLst>
            <a:outerShdw blurRad="40000" dist="23000" dir="5400000" rotWithShape="0">
              <a:srgbClr val="000000">
                <a:alpha val="35000"/>
              </a:srgbClr>
            </a:outerShdw>
          </a:effectLst>
        </a:effectStyle>
        <a:effectStyle>
          <a:effectLst>
            <a:outerShdw blurRad="40000" dist="23000" dir="5400000" rotWithShape="0">
              <a:srgbClr val="000000">
                <a:alpha val="35000"/>
              </a:srgbClr>
            </a:outerShdw>
          </a:effectLst>
          <a:scene3d>
            <a:camera prst="orthographicFront">
              <a:rot lat="0" lon="0" rev="0"/>
            </a:camera>
            <a:lightRig rig="threePt" dir="t">
              <a:rot lat="0" lon="0" rev="1200000"/>
            </a:lightRig>
          </a:scene3d>
          <a:sp3d>
            <a:bevelT w="63500" h="25400"/>
          </a:sp3d>
        </a:effectStyle>
      </a:effectStyleLst>
      <a:bgFillStyleLst>
        <a:solidFill>
          <a:schemeClr val="phClr"/>
        </a:solidFill>
        <a:gradFill rotWithShape="1">
          <a:gsLst>
            <a:gs pos="0">
              <a:schemeClr val="phClr">
                <a:tint val="40000"/>
                <a:satMod val="350000"/>
              </a:schemeClr>
            </a:gs>
            <a:gs pos="40000">
              <a:schemeClr val="phClr">
                <a:tint val="45000"/>
                <a:shade val="99000"/>
                <a:satMod val="350000"/>
              </a:schemeClr>
            </a:gs>
            <a:gs pos="100000">
              <a:schemeClr val="phClr">
                <a:shade val="20000"/>
                <a:satMod val="255000"/>
              </a:schemeClr>
            </a:gs>
          </a:gsLst>
          <a:path path="circle">
            <a:fillToRect l="50000" t="-80000" r="50000" b="180000"/>
          </a:path>
        </a:gradFill>
        <a:gradFill rotWithShape="1">
          <a:gsLst>
            <a:gs pos="0">
              <a:schemeClr val="phClr">
                <a:tint val="80000"/>
                <a:satMod val="300000"/>
              </a:schemeClr>
            </a:gs>
            <a:gs pos="100000">
              <a:schemeClr val="phClr">
                <a:shade val="30000"/>
                <a:satMod val="200000"/>
              </a:schemeClr>
            </a:gs>
          </a:gsLst>
          <a:path path="circle">
            <a:fillToRect l="50000" t="50000" r="50000" b="50000"/>
          </a:path>
        </a:gradFill>
      </a:bgFillStyleLst>
    </a:fmtScheme>
  </a:themeElements>
  <a:objectDefaults/>
  <a:extraClrSchemeLst/>
</a:theme>
"#;

/// Alias for the XML writer backed by an in-memory buffer.
pub(super) type XmlWriter = Writer<Cursor<Vec<u8>>>;

// ---------------------------------------------------------------------------
// Shared string table
// ---------------------------------------------------------------------------

/// Shared string table — maps string content to SST index.
pub(super) struct SharedStringTable {
    strings: Vec<String>,
    index: HashMap<String, u32>,
}

#[derive(Debug, Clone)]
pub(super) struct WorksheetRelationship {
    id: String,
    rel_type: &'static str,
    target: String,
    target_mode: Option<&'static str>,
}

pub(super) fn write_color_element(w: &mut XmlWriter, tag: &str, color: &Color) -> XlsxResult<()> {
    let mut el = BytesStart::new(tag);
    match color {
        Color::Auto => {
            el.push_attribute(("indexed", "64"));
        }
        Color::Rgb { r, g, b } => {
            let v = format!("FF{:02X}{:02X}{:02X}", r, g, b);
            el.push_attribute(("rgb", v.as_str()));
        }
        Color::Argb { a, r, g, b } => {
            let v = format!("{:02X}{:02X}{:02X}{:02X}", a, r, g, b);
            el.push_attribute(("rgb", v.as_str()));
        }
        Color::Indexed(i) => {
            let v = i.to_string();
            el.push_attribute(("indexed", v.as_str()));
        }
        Color::Theme { index, tint } => {
            let v = index.to_string();
            el.push_attribute(("theme", v.as_str()));
            if *tint != 0 {
                let t = ((*tint as f64) / 100.0).to_string();
                el.push_attribute(("tint", t.as_str()));
            }
        }
    }
    w.write_event(Event::Empty(el))?;
    Ok(())
}

pub(super) fn write_xml_part<W: Write + Seek>(
    zip: &mut zip::ZipWriter<W>,
    path: &str,
    build: impl FnOnce(&mut XmlWriter) -> XlsxResult<()>,
) -> XlsxResult<()> {
    let options = zip::write::SimpleFileOptions::default();
    zip.start_file(path, options)?;
    let mut w = Writer::new(Cursor::new(Vec::new()));
    w.write_event(Event::Decl(BytesDecl::new(
        "1.0",
        Some("UTF-8"),
        Some("yes"),
    )))?;
    build(&mut w)?;
    zip.write_all(&w.into_inner().into_inner())?;
    Ok(())
}

impl SharedStringTable {
    /// Build the SST by scanning all string cells in the workbook.
    fn build(workbook: &Workbook) -> Self {
        let mut strings = Vec::new();
        let mut index = HashMap::new();

        for sheet in workbook.worksheets() {
            for (_row, _col, cell) in sheet.iter_cells() {
                let s = match &cell.value {
                    duke_sheets_core::CellValue::String(s) => s.as_str(),
                    // Rich text uses inline strings, not SST
                    _ => continue,
                };
                if !index.contains_key(s) {
                    let idx = strings.len() as u32;
                    index.insert(s.to_owned(), idx);
                    strings.push(s.to_owned());
                }
            }
        }

        Self { strings, index }
    }

    /// Look up the SST index for a string.
    fn get(&self, s: &str) -> Option<u32> {
        self.index.get(s).copied()
    }

    fn is_empty(&self) -> bool {
        self.strings.is_empty()
    }
}

// ---------------------------------------------------------------------------
// XLSX file writer
// ---------------------------------------------------------------------------

/// XLSX file writer
pub struct XlsxWriter;

impl XlsxWriter {
    /// Write a workbook to a file path
    pub fn write_file<P: AsRef<Path>>(workbook: &Workbook, path: P) -> XlsxResult<()> {
        let file = File::create(path)?;
        Self::write(workbook, file)
    }

    /// Write a workbook to a writer
    pub fn write<W: Write + Seek>(workbook: &Workbook, writer: W) -> XlsxResult<()> {
        let mut zip = zip::ZipWriter::new(writer);
        let needs_metadata = Self::has_dynamic_arrays(workbook);

        // Build a workbook-wide style table.
        let style_table = XlsxStyleTable::build(workbook);

        // Build shared string table (deduplicated across all sheets).
        let sst = SharedStringTable::build(workbook);

        // Determine which sheets have comments
        let sheets_with_comments: Vec<usize> = workbook
            .worksheets()
            .enumerate()
            .filter(|(_, sheet)| sheet.comment_count() > 0)
            .map(|(i, _)| i)
            .collect();

        // Build a mapping: (sheet_index, table_index_in_sheet) → global table number
        // Used for: xl/tables/table{N}.xml paths and relationship IDs.
        let mut table_numbering: Vec<(usize, usize, usize)> = Vec::new(); // (sheet_idx, table_in_sheet_idx, global_num)
        let mut global_table_num = 1usize;
        for (i, sheet) in workbook.worksheets().enumerate() {
            for j in 0..sheet.table_count() {
                table_numbering.push((i, j, global_table_num));
                global_table_num += 1;
            }
        }

        // Write [Content_Types].xml
        Self::write_content_types(
            &mut zip,
            workbook,
            &sheets_with_comments,
            &sst,
            &table_numbering,
            needs_metadata,
        )?;

        // Write _rels/.rels
        Self::write_root_rels(&mut zip)?;

        // Write xl/workbook.xml
        Self::write_workbook_xml(&mut zip, workbook)?;

        // Write xl/_rels/workbook.xml.rels
        Self::write_workbook_rels(&mut zip, workbook, &sst, needs_metadata)?;

        // Write xl/styles.xml
        Self::write_styles_xml(&mut zip, &style_table)?;

        // Write xl/theme/theme1.xml
        Self::write_theme_xml(&mut zip, workbook)?;

        // Write shared string table
        if !sst.is_empty() {
            Self::write_shared_strings(&mut zip, &sst)?;
        }

        if needs_metadata {
            Self::write_metadata_xml(&mut zip)?;
        }

        // Write worksheets and their relationships
        for (i, sheet) in workbook.worksheets().enumerate() {
            // Collect global table numbers for this sheet
            let sheet_table_globals: Vec<usize> = table_numbering
                .iter()
                .filter(|(si, _, _)| *si == i)
                .map(|(_, _, gn)| *gn)
                .collect();

            let rels = Self::write_worksheet(
                &mut zip,
                workbook,
                i,
                &style_table,
                &sst,
                &sheet_table_globals,
            )?;

            if !rels.is_empty() {
                Self::write_worksheet_rels(&mut zip, i, &rels)?;
            }

            if sheet.comment_count() > 0 {
                comments::write_vml_drawing(&mut zip, workbook, i)?;
                comments::write_comments(&mut zip, workbook, i)?;
            }

            // Write table part XML files for this sheet
            for (local_idx, &global_num) in sheet_table_globals.iter().enumerate() {
                tables::write_table_part(&mut zip, sheet, local_idx, global_num)?;
            }
        }

        zip.finish()?;
        Ok(())
    }

    // -----------------------------------------------------------------------
    // [Content_Types].xml
    // -----------------------------------------------------------------------

    fn write_content_types<W: Write + Seek>(
        zip: &mut zip::ZipWriter<W>,
        workbook: &Workbook,
        sheets_with_comments: &[usize],
        sst: &SharedStringTable,
        table_numbering: &[(usize, usize, usize)],
        has_metadata: bool,
    ) -> XlsxResult<()> {
        write_xml_part(zip, "[Content_Types].xml", |w| {
            let mut tag = BytesStart::new("Types");
            tag.push_attribute(("xmlns", NS_CONTENT_TYPES));
            w.write_event(Event::Start(tag))?;

            w.create_element("Default")
                .with_attribute(("Extension", "rels"))
                .with_attribute(("ContentType", CT_RELS))
                .write_empty()?;
            w.create_element("Default")
                .with_attribute(("Extension", "xml"))
                .with_attribute(("ContentType", "application/xml"))
                .write_empty()?;
            if !sheets_with_comments.is_empty() {
                w.create_element("Default")
                    .with_attribute(("Extension", "vml"))
                    .with_attribute((
                        "ContentType",
                        "application/vnd.openxmlformats-officedocument.vmlDrawing",
                    ))
                    .write_empty()?;
            }
            w.create_element("Override")
                .with_attribute(("PartName", "/xl/workbook.xml"))
                .with_attribute(("ContentType", CT_WORKBOOK))
                .write_empty()?;
            w.create_element("Override")
                .with_attribute(("PartName", "/xl/styles.xml"))
                .with_attribute(("ContentType", CT_STYLES))
                .write_empty()?;
            w.create_element("Override")
                .with_attribute(("PartName", "/xl/theme/theme1.xml"))
                .with_attribute(("ContentType", CT_THEME))
                .write_empty()?;

            if !sst.is_empty() {
                w.create_element("Override")
                    .with_attribute(("PartName", "/xl/sharedStrings.xml"))
                    .with_attribute(("ContentType", CT_SST))
                    .write_empty()?;
            }

            for i in 0..workbook.sheet_count() {
                let part = format!("/xl/worksheets/sheet{}.xml", i + 1);
                w.create_element("Override")
                    .with_attribute(("PartName", part.as_str()))
                    .with_attribute(("ContentType", CT_WORKSHEET))
                    .write_empty()?;
            }

            for &i in sheets_with_comments {
                let part = format!("/xl/comments{}.xml", i + 1);
                w.create_element("Override")
                    .with_attribute(("PartName", part.as_str()))
                    .with_attribute(("ContentType", CT_COMMENTS))
                    .write_empty()?;
            }

            for &(_, _, global_num) in table_numbering {
                let part = format!("/xl/tables/table{}.xml", global_num);
                w.create_element("Override")
                    .with_attribute(("PartName", part.as_str()))
                    .with_attribute(("ContentType", CT_TABLE))
                    .write_empty()?;
            }

            if has_metadata {
                w.create_element("Override")
                    .with_attribute(("PartName", "/xl/metadata.xml"))
                    .with_attribute(("ContentType", CT_METADATA))
                    .write_empty()?;
            }

            w.write_event(Event::End(BytesEnd::new("Types")))?;
            Ok(())
        })
    }

    // -----------------------------------------------------------------------
    // _rels/.rels
    // -----------------------------------------------------------------------

    fn write_root_rels<W: Write + Seek>(zip: &mut zip::ZipWriter<W>) -> XlsxResult<()> {
        write_xml_part(zip, "_rels/.rels", |w| {
            let mut tag = BytesStart::new("Relationships");
            tag.push_attribute(("xmlns", NS_RELATIONSHIPS));
            w.write_event(Event::Start(tag))?;

            w.create_element("Relationship")
                .with_attribute(("Id", "rId1"))
                .with_attribute(("Type", RT_OFFICE_DOC))
                .with_attribute(("Target", "xl/workbook.xml"))
                .write_empty()?;

            w.write_event(Event::End(BytesEnd::new("Relationships")))?;
            Ok(())
        })
    }

    // -----------------------------------------------------------------------
    // xl/workbook.xml
    // -----------------------------------------------------------------------

    fn write_workbook_xml<W: Write + Seek>(
        zip: &mut zip::ZipWriter<W>,
        workbook: &Workbook,
    ) -> XlsxResult<()> {
        write_xml_part(zip, "xl/workbook.xml", |w| {
            let mut tag = BytesStart::new("workbook");
            tag.push_attribute(("xmlns", NS_SPREADSHEET));
            tag.push_attribute(("xmlns:r", NS_DOC_RELS));
            w.write_event(Event::Start(tag))?;

            // workbookPr
            let settings = workbook.settings();
            if settings.date_1904 {
                w.create_element("workbookPr")
                    .with_attribute(("date1904", "1"))
                    .write_empty()?;
            }

            // bookViews
            let active = workbook.active_sheet();
            if active > 0 {
                let tab = active.to_string();
                w.write_event(Event::Start(BytesStart::new("bookViews")))?;
                w.create_element("workbookView")
                    .with_attribute(("activeTab", tab.as_str()))
                    .write_empty()?;
                w.write_event(Event::End(BytesEnd::new("bookViews")))?;
            }

            // sheets
            w.write_event(Event::Start(BytesStart::new("sheets")))?;
            for (i, sheet) in workbook.worksheets().enumerate() {
                let sheet_id = (i + 1).to_string();
                let rid = format!("rId{}", i + 1);
                let mut el = w
                    .create_element("sheet")
                    .with_attribute(("name", sheet.name()))
                    .with_attribute(("sheetId", sheet_id.as_str()));
                match sheet.visibility() {
                    duke_sheets_core::SheetVisibility::Hidden => {
                        el = el.with_attribute(("state", "hidden"));
                    }
                    duke_sheets_core::SheetVisibility::VeryHidden => {
                        el = el.with_attribute(("state", "veryHidden"));
                    }
                    duke_sheets_core::SheetVisibility::Visible => {}
                }
                el.with_attribute(("r:id", rid.as_str())).write_empty()?;
            }
            w.write_event(Event::End(BytesEnd::new("sheets")))?;

            let named = workbook.named_ranges();
            let mut print_names: Vec<duke_sheets_core::named_range::NamedRange> = Vec::new();
            let mut generated_keys: HashSet<String> = HashSet::new();
            for (idx, sheet) in workbook.worksheets().enumerate() {
                let ps = sheet.page_setup();

                if let Some(ref range) = ps.print_area {
                    let formula = Self::format_print_area_formula(sheet.name(), range);
                    let mut nr = duke_sheets_core::named_range::NamedRange::new(
                        "_xlnm.Print_Area",
                        formula,
                        duke_sheets_core::named_range::NameScope::Sheet(idx),
                    );
                    nr.hidden = true;
                    generated_keys.insert(Self::defined_name_key(&nr.name, &nr.scope));
                    print_names.push(nr);
                }

                if let Some(formula) =
                    Self::format_print_titles_formula(sheet.name(), ps.repeat_rows, ps.repeat_cols)
                {
                    let mut nr = duke_sheets_core::named_range::NamedRange::new(
                        "_xlnm.Print_Titles",
                        formula,
                        duke_sheets_core::named_range::NameScope::Sheet(idx),
                    );
                    nr.hidden = true;
                    generated_keys.insert(Self::defined_name_key(&nr.name, &nr.scope));
                    print_names.push(nr);
                }
            }

            let existing_non_print: Vec<_> = named
                .iter()
                .filter(|nr| !generated_keys.contains(&Self::defined_name_key(&nr.name, &nr.scope)))
                .collect();

            let total = existing_non_print.len() + print_names.len();
            if total > 0 {
                w.write_event(Event::Start(BytesStart::new("definedNames")))?;

                for nr in &existing_non_print {
                    Self::write_defined_name(w, nr)?;
                }

                for nr in &print_names {
                    Self::write_defined_name(w, nr)?;
                }

                w.write_event(Event::End(BytesEnd::new("definedNames")))?;
            }

            // calcPr
            if settings.calc_on_open {
                w.create_element("calcPr")
                    .with_attribute(("calcId", "191029"))
                    .with_attribute(("fullCalcOnLoad", "1"))
                    .write_empty()?;
            }

            w.write_event(Event::End(BytesEnd::new("workbook")))?;
            Ok(())
        })
    }

    // -----------------------------------------------------------------------
    // xl/_rels/workbook.xml.rels
    // -----------------------------------------------------------------------

    fn write_workbook_rels<W: Write + Seek>(
        zip: &mut zip::ZipWriter<W>,
        workbook: &Workbook,
        sst: &SharedStringTable,
        has_metadata: bool,
    ) -> XlsxResult<()> {
        write_xml_part(zip, "xl/_rels/workbook.xml.rels", |w| {
            let mut tag = BytesStart::new("Relationships");
            tag.push_attribute(("xmlns", NS_RELATIONSHIPS));
            w.write_event(Event::Start(tag))?;

            for i in 0..workbook.sheet_count() {
                let rid = format!("rId{}", i + 1);
                let target = format!("worksheets/sheet{}.xml", i + 1);
                w.create_element("Relationship")
                    .with_attribute(("Id", rid.as_str()))
                    .with_attribute(("Type", RT_WORKSHEET))
                    .with_attribute(("Target", target.as_str()))
                    .write_empty()?;
            }

            // Styles
            let mut next_rid = workbook.sheet_count() + 1;
            let rid = format!("rId{}", next_rid);
            w.create_element("Relationship")
                .with_attribute(("Id", rid.as_str()))
                .with_attribute(("Type", RT_STYLES))
                .with_attribute(("Target", "styles.xml"))
                .write_empty()?;
            next_rid += 1;

            // Shared strings
            if !sst.is_empty() {
                let rid = format!("rId{}", next_rid);
                w.create_element("Relationship")
                    .with_attribute(("Id", rid.as_str()))
                    .with_attribute(("Type", RT_SHARED_STRINGS))
                    .with_attribute(("Target", "sharedStrings.xml"))
                    .write_empty()?;
                next_rid += 1;
            }

            // Theme
            let rid = format!("rId{}", next_rid);
            w.create_element("Relationship")
                .with_attribute(("Id", rid.as_str()))
                .with_attribute(("Type", RT_THEME))
                .with_attribute(("Target", "theme/theme1.xml"))
                .write_empty()?;
            next_rid += 1;

            if has_metadata {
                let rid = format!("rId{}", next_rid);
                w.create_element("Relationship")
                    .with_attribute(("Id", rid.as_str()))
                    .with_attribute(("Type", RT_SHEET_METADATA))
                    .with_attribute(("Target", "metadata.xml"))
                    .write_empty()?;
            }

            w.write_event(Event::End(BytesEnd::new("Relationships")))?;
            Ok(())
        })
    }

    fn write_defined_name(
        w: &mut XmlWriter,
        nr: &duke_sheets_core::named_range::NamedRange,
    ) -> XlsxResult<()> {
        let mut el = w
            .create_element("definedName")
            .with_attribute(("name", nr.name.as_str()));
        let scope_str;
        if let duke_sheets_core::named_range::NameScope::Sheet(idx) = nr.scope {
            scope_str = idx.to_string();
            el = el.with_attribute(("localSheetId", scope_str.as_str()));
        }
        if nr.hidden {
            el = el.with_attribute(("hidden", "1"));
        }
        if let Some(ref comment) = nr.comment {
            el = el.with_attribute(("comment", comment.as_str()));
        }
        el.write_text_content(BytesText::new(&nr.refers_to))?;
        Ok(())
    }

    fn format_print_area_formula(sheet_name: &str, range: &CellRange) -> String {
        let quoted_name = Self::quote_sheet_name(sheet_name);
        let start_col = CellAddress::column_to_letters(range.start.col);
        let end_col = CellAddress::column_to_letters(range.end.col);
        format!(
            "{}!${}${}:${}${}",
            quoted_name,
            start_col,
            range.start.row + 1,
            end_col,
            range.end.row + 1,
        )
    }

    fn format_print_titles_formula(
        sheet_name: &str,
        repeat_rows: Option<(u32, u32)>,
        repeat_cols: Option<(u16, u16)>,
    ) -> Option<String> {
        let quoted_name = Self::quote_sheet_name(sheet_name);
        let mut parts = Vec::new();

        if let Some((r1, r2)) = repeat_rows {
            parts.push(format!("{}!${}:${}", quoted_name, r1 + 1, r2 + 1));
        }

        if let Some((c1, c2)) = repeat_cols {
            let start = CellAddress::column_to_letters(c1);
            let end = CellAddress::column_to_letters(c2);
            parts.push(format!("{}!${}:${}", quoted_name, start, end));
        }

        if parts.is_empty() {
            None
        } else {
            Some(parts.join(","))
        }
    }

    fn quote_sheet_name(name: &str) -> String {
        let needs_quoting = name.contains(' ')
            || name.contains('\'')
            || name.contains('!')
            || name.contains('[')
            || name.contains(']')
            || name.chars().next().is_some_and(|c| c.is_ascii_digit());
        if needs_quoting {
            let escaped = name.replace('\'', "''");
            format!("'{}'", escaped)
        } else {
            name.to_string()
        }
    }

    fn defined_name_key(name: &str, scope: &duke_sheets_core::named_range::NameScope) -> String {
        let name_lower = name.to_ascii_lowercase();
        match scope {
            duke_sheets_core::named_range::NameScope::Workbook => name_lower,
            duke_sheets_core::named_range::NameScope::Sheet(idx) => {
                format!("{}:sheet:{}", name_lower, idx)
            }
        }
    }

    // -----------------------------------------------------------------------
    // xl/sharedStrings.xml
    // -----------------------------------------------------------------------

    fn write_shared_strings<W: Write + Seek>(
        zip: &mut zip::ZipWriter<W>,
        sst: &SharedStringTable,
    ) -> XlsxResult<()> {
        write_xml_part(zip, "xl/sharedStrings.xml", |w| {
            let count = sst.strings.len().to_string();
            let mut tag = BytesStart::new("sst");
            tag.push_attribute(("xmlns", NS_SPREADSHEET));
            tag.push_attribute(("count", count.as_str()));
            tag.push_attribute(("uniqueCount", count.as_str()));
            w.write_event(Event::Start(tag))?;

            for s in &sst.strings {
                w.write_event(Event::Start(BytesStart::new("si")))?;

                let needs_preserve = s.starts_with(|c: char| c.is_ascii_whitespace())
                    || s.ends_with(|c: char| c.is_ascii_whitespace());
                if needs_preserve {
                    let mut t = BytesStart::new("t");
                    t.push_attribute(("xml:space", "preserve"));
                    w.write_event(Event::Start(t))?;
                } else {
                    w.write_event(Event::Start(BytesStart::new("t")))?;
                }
                w.write_event(Event::Text(BytesText::new(s)))?;
                w.write_event(Event::End(BytesEnd::new("t")))?;

                w.write_event(Event::End(BytesEnd::new("si")))?;
            }

            w.write_event(Event::End(BytesEnd::new("sst")))?;
            Ok(())
        })
    }

    fn has_dynamic_arrays(workbook: &Workbook) -> bool {
        for sheet in workbook.worksheets() {
            for (row, col, _) in sheet.formula_cells() {
                if sheet
                    .formula_data_at(row, col)
                    .and_then(|formula| formula.array_result.as_ref())
                    .is_some()
                {
                    return true;
                }
            }
            for (_, _, cell) in sheet.iter_cells() {
                if matches!(&cell.value, duke_sheets_core::CellValue::SpillTarget { .. }) {
                    return true;
                }
            }
        }
        false
    }

    fn write_metadata_xml<W: Write + Seek>(zip: &mut zip::ZipWriter<W>) -> XlsxResult<()> {
        write_xml_part(zip, "xl/metadata.xml", |w| {
            let mut tag = BytesStart::new("metadata");
            tag.push_attribute(("xmlns", NS_SPREADSHEET));
            tag.push_attribute((
                "xmlns:xda",
                "http://schemas.microsoft.com/office/spreadsheetml/2017/dynamicarray",
            ));
            w.write_event(Event::Start(tag))?;

            w.write_event(Event::Start(BytesStart::new("metadataTypes")))?;
            let mut mt = BytesStart::new("metadataType");
            mt.push_attribute(("name", "XLDAPR"));
            mt.push_attribute(("minSupportedVersion", "120000"));
            mt.push_attribute(("copy", "1"));
            mt.push_attribute(("pasteAll", "1"));
            mt.push_attribute(("pasteValues", "1"));
            mt.push_attribute(("merge", "1"));
            mt.push_attribute(("splitFirst", "1"));
            mt.push_attribute(("rowColShift", "1"));
            mt.push_attribute(("clearFormats", "1"));
            mt.push_attribute(("clearComments", "1"));
            mt.push_attribute(("assign", "1"));
            mt.push_attribute(("coerce", "1"));
            mt.push_attribute(("cellMeta", "1"));
            w.write_event(Event::Empty(mt))?;
            w.write_event(Event::End(BytesEnd::new("metadataTypes")))?;

            let mut fm = BytesStart::new("futureMetadata");
            fm.push_attribute(("name", "XLDAPR"));
            fm.push_attribute(("count", "2"));
            w.write_event(Event::Start(fm))?;

            w.write_event(Event::Start(BytesStart::new("bk")))?;
            w.write_event(Event::Start(BytesStart::new("extLst")))?;
            let mut ext1 = BytesStart::new("ext");
            ext1.push_attribute(("uri", "{bdbb8cdc-fa1e-496e-a857-3c3f30c029c3}"));
            w.write_event(Event::Start(ext1))?;
            w.create_element("xda:dynamicArrayProperties")
                .with_attribute(("fDynamic", "1"))
                .with_attribute(("fCollapsed", "0"))
                .write_empty()?;
            w.write_event(Event::End(BytesEnd::new("ext")))?;
            w.write_event(Event::End(BytesEnd::new("extLst")))?;
            w.write_event(Event::End(BytesEnd::new("bk")))?;

            w.write_event(Event::Start(BytesStart::new("bk")))?;
            w.write_event(Event::Start(BytesStart::new("extLst")))?;
            let mut ext2 = BytesStart::new("ext");
            ext2.push_attribute(("uri", "{bdbb8cdc-fa1e-496e-a857-3c3f30c029c3}"));
            w.write_event(Event::Start(ext2))?;
            w.create_element("xda:dynamicArrayProperties")
                .with_attribute(("fDynamic", "0"))
                .with_attribute(("fCollapsed", "1"))
                .write_empty()?;
            w.write_event(Event::End(BytesEnd::new("ext")))?;
            w.write_event(Event::End(BytesEnd::new("extLst")))?;
            w.write_event(Event::End(BytesEnd::new("bk")))?;

            w.write_event(Event::End(BytesEnd::new("futureMetadata")))?;

            let mut cm = BytesStart::new("cellMetadata");
            cm.push_attribute(("count", "2"));
            w.write_event(Event::Start(cm))?;
            w.write_event(Event::Start(BytesStart::new("bk")))?;
            w.create_element("rc")
                .with_attribute(("t", "1"))
                .with_attribute(("v", "0"))
                .write_empty()?;
            w.write_event(Event::End(BytesEnd::new("bk")))?;
            w.write_event(Event::Start(BytesStart::new("bk")))?;
            w.create_element("rc")
                .with_attribute(("t", "1"))
                .with_attribute(("v", "1"))
                .write_empty()?;
            w.write_event(Event::End(BytesEnd::new("bk")))?;
            w.write_event(Event::End(BytesEnd::new("cellMetadata")))?;

            w.write_event(Event::End(BytesEnd::new("metadata")))?;
            Ok(())
        })
    }

    // -----------------------------------------------------------------------
    // xl/styles.xml
    // -----------------------------------------------------------------------

    fn write_styles_xml<W: Write + Seek>(
        zip: &mut zip::ZipWriter<W>,
        style_table: &XlsxStyleTable,
    ) -> XlsxResult<()> {
        write_xml_part(zip, "xl/styles.xml", |w| style_table.write_styles_xml(w))
    }

    fn write_theme_xml<W: Write + Seek>(
        zip: &mut zip::ZipWriter<W>,
        workbook: &Workbook,
    ) -> XlsxResult<()> {
        let options = zip::write::SimpleFileOptions::default();
        zip.start_file("xl/theme/theme1.xml", options)?;
        if let Some(theme_bytes) = roundtrip_theme_data_for(workbook) {
            zip.write_all(&theme_bytes)?;
        } else {
            zip.write_all(DEFAULT_THEME_XML.as_bytes())?;
        }
        Ok(())
    }

    // -----------------------------------------------------------------------
    // xl/worksheets/sheet{N}.xml
    // -----------------------------------------------------------------------

    fn write_worksheet<W: Write + Seek>(
        zip: &mut zip::ZipWriter<W>,
        workbook: &Workbook,
        index: usize,
        style_table: &XlsxStyleTable,
        sst: &SharedStringTable,
        sheet_table_globals: &[usize], // global table numbers for this sheet
    ) -> XlsxResult<Vec<WorksheetRelationship>> {
        let path = format!("xl/worksheets/sheet{}.xml", index + 1);
        let mut rels = Vec::new();
        write_xml_part(zip, &path, |w| {
            let sheet = workbook
                .worksheet(index)
                .ok_or_else(|| XlsxError::InvalidFormat("Sheet not found".into()))?;

            if sheet.comment_count() > 0 {
                let comments_target = format!("../comments{}.xml", index + 1);
                let vml_target = format!("../drawings/vmlDrawing{}.vml", index + 1);
                rels.push(WorksheetRelationship {
                    id: "rId1".to_string(),
                    rel_type: RT_VML_DRAWING,
                    target: vml_target,
                    target_mode: None,
                });
                rels.push(WorksheetRelationship {
                    id: "rId2".to_string(),
                    rel_type: RT_COMMENTS,
                    target: comments_target,
                    target_mode: None,
                });
            }

            let mut tag = BytesStart::new("worksheet");
            tag.push_attribute(("xmlns", NS_SPREADSHEET));
            tag.push_attribute(("xmlns:r", NS_DOC_RELS));
            w.write_event(Event::Start(tag))?;

            // sheetPr (tab color)
            Self::write_sheet_pr(w, sheet)?;

            // dimension (used range)
            Self::write_dimension(w, sheet)?;

            // sheetViews (freeze panes, tab selection)
            Self::write_sheet_views(w, sheet)?;

            // sheetFormatPr (default row height)
            Self::write_sheet_format_pr(w, sheet)?;

            // cols (column widths / hidden columns)
            Self::write_cols(w, sheet)?;

            // sheetData (cell data)
            Self::write_sheet_data(w, sheet, index, style_table, sst)?;

            // sheetProtection
            Self::write_sheet_protection(w, sheet)?;

            Self::write_auto_filter(w, sheet)?;

            // mergeCells
            Self::write_merge_cells(w, sheet)?;

            // conditionalFormatting
            conditional_format::write_conditional_formatting(w, sheet, index, style_table)?;

            // dataValidations
            data_validation::write_data_validations(w, sheet)?;
            Self::write_hyperlinks(w, sheet, &mut rels)?;

            // pageMargins + pageSetup
            Self::write_page_setup(w, sheet)?;
            Self::write_page_breaks(w, sheet)?;

            if sheet.comment_count() > 0 {
                let mut legacy_drawing = BytesStart::new("legacyDrawing");
                legacy_drawing.push_attribute(("r:id", "rId1"));
                w.write_event(Event::Empty(legacy_drawing))?;
            }

            // tableParts (references to xl/tables/tableN.xml)
            if !sheet_table_globals.is_empty() {
                let count_str = sheet_table_globals.len().to_string();
                let mut tp = BytesStart::new("tableParts");
                tp.push_attribute(("count", count_str.as_str()));
                w.write_event(Event::Start(tp))?;
                for &global_num in sheet_table_globals {
                    let rid = format!("rId{}", rels.len() + 1);
                    let target = format!("../tables/table{}.xml", global_num);
                    rels.push(WorksheetRelationship {
                        id: rid.clone(),
                        rel_type: RT_TABLE,
                        target,
                        target_mode: None,
                    });
                    let mut part = BytesStart::new("tablePart");
                    part.push_attribute(("r:id", rid.as_str()));
                    w.write_event(Event::Empty(part))?;
                }
                w.write_event(Event::End(BytesEnd::new("tableParts")))?;
            }

            w.write_event(Event::End(BytesEnd::new("worksheet")))?;
            Ok(())
        })?;

        Ok(rels)
    }

    // -- worksheet sub-sections ---------------------------------------------

    fn write_sheet_pr(w: &mut XmlWriter, sheet: &duke_sheets_core::Worksheet) -> XlsxResult<()> {
        if let Some(color) = sheet.tab_color() {
            w.write_event(Event::Start(BytesStart::new("sheetPr")))?;
            write_color_element(w, "tabColor", &color)?;
            w.write_event(Event::End(BytesEnd::new("sheetPr")))?;
        }
        Ok(())
    }

    fn write_dimension(w: &mut XmlWriter, sheet: &duke_sheets_core::Worksheet) -> XlsxResult<()> {
        let ref_str = if let Some(range) = sheet.used_range() {
            range.to_string()
        } else {
            "A1".to_string()
        };
        w.create_element("dimension")
            .with_attribute(("ref", ref_str.as_str()))
            .write_empty()?;
        Ok(())
    }

    fn write_sheet_format_pr(
        w: &mut XmlWriter,
        _sheet: &duke_sheets_core::Worksheet,
    ) -> XlsxResult<()> {
        // Emit sheetFormatPr with default row height (Excel default is 15)
        w.create_element("sheetFormatPr")
            .with_attribute(("defaultRowHeight", "15"))
            .write_empty()?;
        Ok(())
    }

    fn write_sheet_views(w: &mut XmlWriter, sheet: &duke_sheets_core::Worksheet) -> XlsxResult<()> {
        let freeze = sheet.freeze_panes();
        let split = sheet.split_panes();
        let selected = sheet.is_selected();
        let zoom = sheet.zoom_scale();
        let selections = sheet.selections();

        if freeze.is_none()
            && split.is_none()
            && !selected
            && zoom.is_none()
            && selections.is_empty()
        {
            return Ok(());
        }

        w.write_event(Event::Start(BytesStart::new("sheetViews")))?;

        let mut sv = BytesStart::new("sheetView");
        sv.push_attribute(("workbookViewId", "0"));
        if selected {
            sv.push_attribute(("tabSelected", "1"));
        }
        if let Some(z) = zoom {
            let z_s = z.to_string();
            sv.push_attribute(("zoomScale", z_s.as_str()));
        }
        w.write_event(Event::Start(sv))?;

        if let Some(fp) = freeze {
            let active_pane = match (fp.col > 0, fp.row > 0) {
                (true, true) => "bottomRight",
                (false, true) => "bottomLeft",
                (true, false) => "topRight",
                (false, false) => "bottomLeft",
            };
            let top_left = CellAddress::new(fp.row, fp.col).to_a1_string();

            let mut pane = BytesStart::new("pane");
            if fp.col > 0 {
                let v = fp.col.to_string();
                pane.push_attribute(("xSplit", v.as_str()));
            }
            if fp.row > 0 {
                let v = fp.row.to_string();
                pane.push_attribute(("ySplit", v.as_str()));
            }
            pane.push_attribute(("topLeftCell", top_left.as_str()));
            pane.push_attribute(("activePane", active_pane));
            pane.push_attribute(("state", "frozen"));
            w.write_event(Event::Empty(pane))?;

            if selections.is_empty() {
                // Synthesize a default selection for the active pane
                let default_cell = CellAddress::new(fp.row, fp.col).to_a1_string();
                w.create_element("selection")
                    .with_attribute(("pane", active_pane))
                    .with_attribute(("activeCell", default_cell.as_str()))
                    .with_attribute(("sqref", default_cell.as_str()))
                    .write_empty()?;
            }
        } else if let Some(sp) = split {
            let mut pane = BytesStart::new("pane");
            if sp.x_split != 0.0 {
                let v = sp.x_split.to_string();
                pane.push_attribute(("xSplit", v.as_str()));
            }
            if sp.y_split != 0.0 {
                let v = sp.y_split.to_string();
                pane.push_attribute(("ySplit", v.as_str()));
            }
            if let Some((r, c)) = sp.top_left {
                let top_left = CellAddress::new(r, c).to_a1_string();
                pane.push_attribute(("topLeftCell", top_left.as_str()));
            }
            if let Some(active_pane) = sp.active_pane.as_deref() {
                pane.push_attribute(("activePane", active_pane));
            }
            pane.push_attribute(("state", "split"));
            w.write_event(Event::Empty(pane))?;

            if selections.is_empty() {
                // Synthesize a default selection for the active pane
                let default_active = sp
                    .top_left
                    .map(|(r, c)| CellAddress::new(r, c).to_a1_string())
                    .unwrap_or_else(|| "A1".to_string());
                let mut sel = BytesStart::new("selection");
                if let Some(active_pane) = sp.active_pane.as_deref() {
                    sel.push_attribute(("pane", active_pane));
                }
                sel.push_attribute(("activeCell", default_active.as_str()));
                sel.push_attribute(("sqref", default_active.as_str()));
                w.write_event(Event::Empty(sel))?;
            }
        }

        // Emit all stored selections.
        // If a selection has no pane set but the sheet has freeze/split panes,
        // infer the pane from the active pane of the freeze/split.
        let default_pane: Option<&str> = if freeze.is_some() {
            Some(match (freeze.unwrap().col > 0, freeze.unwrap().row > 0) {
                (true, true) => "bottomRight",
                (false, true) => "bottomLeft",
                (true, false) => "topRight",
                (false, false) => "bottomLeft",
            })
        } else if let Some(sp) = split {
            sp.active_pane.as_deref()
        } else {
            None
        };

        for sel in selections {
            let mut el = BytesStart::new("selection");
            let pane_val = sel.pane.as_deref().or(default_pane);
            if let Some(pane) = pane_val {
                el.push_attribute(("pane", pane));
            }
            if let Some(ac) = &sel.active_cell {
                el.push_attribute(("activeCell", ac.as_str()));
            }
            if let Some(sq) = &sel.sqref {
                el.push_attribute(("sqref", sq.as_str()));
            }
            w.write_event(Event::Empty(el))?;
        }

        w.write_event(Event::End(BytesEnd::new("sheetView")))?;
        w.write_event(Event::End(BytesEnd::new("sheetViews")))?;
        Ok(())
    }

    fn write_cols(w: &mut XmlWriter, sheet: &duke_sheets_core::Worksheet) -> XlsxResult<()> {
        let col_widths = sheet.custom_column_widths();
        let col_hidden = sheet.hidden_columns();
        let col_outline = sheet.column_outline_levels();
        let col_collapsed = sheet.collapsed_columns();

        if col_widths.is_empty()
            && col_hidden.is_empty()
            && col_outline.is_empty()
            && col_collapsed.is_empty()
        {
            return Ok(());
        }

        w.write_event(Event::Start(BytesStart::new("cols")))?;

        let mut cols_to_write: std::collections::BTreeSet<u16> = Default::default();
        for &col in col_widths.keys() {
            cols_to_write.insert(col);
        }
        for &col in col_hidden.keys() {
            cols_to_write.insert(col);
        }
        for &col in col_outline.keys() {
            cols_to_write.insert(col);
        }
        for &col in col_collapsed.keys() {
            cols_to_write.insert(col);
        }

        for col in cols_to_write {
            let col1 = (col as u32 + 1).to_string();
            let width = col_widths.get(&col).copied().unwrap_or(8.43);
            let width_s = format!("{:.2}", width);
            let hidden = col_hidden.get(&col).copied().unwrap_or(false);
            let outline_level = col_outline.get(&col).copied().unwrap_or(0);
            let collapsed = col_collapsed.get(&col).copied().unwrap_or(false);

            let mut el = BytesStart::new("col");
            el.push_attribute(("min", col1.as_str()));
            el.push_attribute(("max", col1.as_str()));
            el.push_attribute(("width", width_s.as_str()));
            el.push_attribute(("customWidth", "1"));
            if hidden {
                el.push_attribute(("hidden", "1"));
            }
            if outline_level > 0 {
                let s = outline_level.to_string();
                el.push_attribute(("outlineLevel", s.as_str()));
            }
            if collapsed {
                el.push_attribute(("collapsed", "1"));
            }
            w.write_event(Event::Empty(el))?;
        }

        w.write_event(Event::End(BytesEnd::new("cols")))?;
        Ok(())
    }

    fn write_sheet_data(
        w: &mut XmlWriter,
        sheet: &duke_sheets_core::Worksheet,
        sheet_index: usize,
        style_table: &XlsxStyleTable,
        sst: &SharedStringTable,
    ) -> XlsxResult<()> {
        w.write_event(Event::Start(BytesStart::new("sheetData")))?;

        let custom_heights = sheet.custom_row_heights();
        let hidden_rows_map = sheet.hidden_rows();
        let row_outline = sheet.row_outline_levels();
        let row_collapsed = sheet.collapsed_rows();
        let mut meta_only_rows: BTreeSet<u32> = Default::default();
        for &r in custom_heights.keys() {
            meta_only_rows.insert(r);
        }
        for &r in hidden_rows_map.keys() {
            meta_only_rows.insert(r);
        }
        for &r in row_outline.keys() {
            meta_only_rows.insert(r);
        }
        for &r in row_collapsed.keys() {
            meta_only_rows.insert(r);
        }

        let mut row_cells: BTreeMap<u32, BTreeSet<u16>> = BTreeMap::new();
        for (row, col, _) in sheet.iter_cells() {
            row_cells.entry(row).or_default().insert(col);
        }

        let mut formula_only_cells: BTreeMap<u32, BTreeSet<u16>> = BTreeMap::new();
        for (row, col, _) in sheet.formula_cells() {
            if sheet.cell_at(row, col).is_none() {
                formula_only_cells.entry(row).or_default().insert(col);
            }
        }

        let mut all_rows = meta_only_rows.clone();
        all_rows.extend(row_cells.keys().copied());
        all_rows.extend(formula_only_cells.keys().copied());

        for row in all_rows {
            let grid_cols = row_cells.get(&row);
            let formula_only_cols = formula_only_cells.get(&row);
            if grid_cols.is_none() && formula_only_cols.is_none() {
                Self::write_meta_row(
                    w,
                    row,
                    custom_heights.get(&row).copied(),
                    hidden_rows_map.get(&row).copied().unwrap_or(false),
                    row_outline.get(&row).copied().unwrap_or(0),
                    row_collapsed.get(&row).copied().unwrap_or(false),
                )?;
                continue;
            }

            let r = (row + 1).to_string();
            let mut row_tag = BytesStart::new("row");
            row_tag.push_attribute(("r", r.as_str()));
            if let Some(&ht) = custom_heights.get(&row) {
                let ht_s = format!("{:.2}", ht);
                row_tag.push_attribute(("ht", ht_s.as_str()));
                row_tag.push_attribute(("customHeight", "1"));
            }
            if sheet.is_row_hidden(row) {
                row_tag.push_attribute(("hidden", "1"));
            }
            let row_outline_level = row_outline.get(&row).copied().unwrap_or(0);
            if row_outline_level > 0 {
                let s = row_outline_level.to_string();
                row_tag.push_attribute(("outlineLevel", s.as_str()));
            }
            if row_collapsed.get(&row).copied().unwrap_or(false) {
                row_tag.push_attribute(("collapsed", "1"));
            }
            w.write_event(Event::Start(row_tag))?;

            let empty_cell = duke_sheets_core::CellData::empty();
            match (grid_cols, formula_only_cols) {
                (Some(cols), None) | (None, Some(cols)) => {
                    for col in cols {
                        let cell = sheet.cell_at(row, *col).unwrap_or(&empty_cell);
                        Self::write_cell(w, row, *col, cell, sheet_index, style_table, sst, sheet)?;
                    }
                }
                (Some(grid_cols), Some(formula_only_cols)) => {
                    let mut grid_iter = grid_cols.iter().peekable();
                    let mut formula_iter = formula_only_cols.iter().peekable();

                    loop {
                        let next_col =
                            match (grid_iter.peek().copied(), formula_iter.peek().copied()) {
                                (Some(&grid_col), Some(&formula_col)) if grid_col < formula_col => {
                                    grid_iter.next();
                                    grid_col
                                }
                                (Some(&grid_col), Some(&formula_col)) if formula_col < grid_col => {
                                    formula_iter.next();
                                    formula_col
                                }
                                (Some(&grid_col), Some(_)) => {
                                    grid_iter.next();
                                    formula_iter.next();
                                    grid_col
                                }
                                (Some(&grid_col), None) => {
                                    grid_iter.next();
                                    grid_col
                                }
                                (None, Some(&formula_col)) => {
                                    formula_iter.next();
                                    formula_col
                                }
                                (None, None) => break,
                            };

                        let cell = sheet.cell_at(row, next_col).unwrap_or(&empty_cell);
                        Self::write_cell(
                            w,
                            row,
                            next_col,
                            cell,
                            sheet_index,
                            style_table,
                            sst,
                            sheet,
                        )?;
                    }
                }
                (None, None) => unreachable!(),
            }

            w.write_event(Event::End(BytesEnd::new("row")))?;
        }

        w.write_event(Event::End(BytesEnd::new("sheetData")))?;
        Ok(())
    }

    /// Write a self-closing `<row>` element with metadata only (height/hidden).
    fn write_meta_row(
        w: &mut XmlWriter,
        row: u32,
        custom_height: Option<f64>,
        hidden: bool,
        outline_level: u8,
        collapsed: bool,
    ) -> XlsxResult<()> {
        let r = (row + 1).to_string();
        let mut tag = BytesStart::new("row");
        tag.push_attribute(("r", r.as_str()));
        if let Some(ht) = custom_height {
            let ht_s = format!("{:.2}", ht);
            tag.push_attribute(("ht", ht_s.as_str()));
            tag.push_attribute(("customHeight", "1"));
        }
        if hidden {
            tag.push_attribute(("hidden", "1"));
        }
        if outline_level > 0 {
            let s = outline_level.to_string();
            tag.push_attribute(("outlineLevel", s.as_str()));
        }
        if collapsed {
            tag.push_attribute(("collapsed", "1"));
        }
        w.write_event(Event::Empty(tag))?;
        Ok(())
    }

    /// Write a single `<c>` cell element.
    fn write_cell(
        w: &mut XmlWriter,
        row: u32,
        col: u16,
        cell: &duke_sheets_core::CellData,
        sheet_index: usize,
        style_table: &XlsxStyleTable,
        sst: &SharedStringTable,
        worksheet: &duke_sheets_core::Worksheet,
    ) -> XlsxResult<()> {
        let addr = CellAddress::new(row, col);
        let cell_ref = addr.to_a1_string();
        let xf_id = style_table.xf_id_for(sheet_index, cell.style_index);
        let xf_str = xf_id.to_string();

        if let Some(formula) = worksheet.formula_data_at(row, col) {
            let formula_text = if formula.text.starts_with('=') {
                &formula.text[1..]
            } else {
                formula.text.as_str()
            };
            let mut c = BytesStart::new("c");
            c.push_attribute(("r", cell_ref.as_str()));
            if xf_id != 0 {
                c.push_attribute(("s", xf_str.as_str()));
            }
            match &cell.value {
                duke_sheets_core::CellValue::String(_)
                | duke_sheets_core::CellValue::RichText(_) => c.push_attribute(("t", "str")),
                duke_sheets_core::CellValue::Boolean(_) => c.push_attribute(("t", "b")),
                duke_sheets_core::CellValue::Error(_) => c.push_attribute(("t", "e")),
                _ => {}
            }
            if formula.array_result.is_some() {
                c.push_attribute(("cm", "1"));
            }
            w.write_event(Event::Start(c))?;
            w.create_element("f")
                .write_text_content(BytesText::new(formula_text))?;
            match &cell.value {
                duke_sheets_core::CellValue::Number(n) => {
                    let v = n.to_string();
                    w.create_element("v")
                        .write_text_content(BytesText::new(&v))?;
                }
                duke_sheets_core::CellValue::String(s) => {
                    w.create_element("v")
                        .write_text_content(BytesText::new(s.as_str()))?;
                }
                duke_sheets_core::CellValue::RichText(runs) => {
                    let text = duke_sheets_core::rich_text_to_plain(runs);
                    w.create_element("v")
                        .write_text_content(BytesText::new(&text))?;
                }
                duke_sheets_core::CellValue::Boolean(b) => {
                    w.create_element("v")
                        .write_text_content(BytesText::new(if *b { "1" } else { "0" }))?;
                }
                duke_sheets_core::CellValue::Error(e) => {
                    w.create_element("v")
                        .write_text_content(BytesText::new(e.as_str()))?;
                }
                duke_sheets_core::CellValue::Empty
                | duke_sheets_core::CellValue::SpillTarget { .. } => {}
            }
            w.write_event(Event::End(BytesEnd::new("c")))?;
            return Ok(());
        }

        match &cell.value {
            duke_sheets_core::CellValue::Number(n) => {
                let mut c = BytesStart::new("c");
                c.push_attribute(("r", cell_ref.as_str()));
                if xf_id != 0 {
                    c.push_attribute(("s", xf_str.as_str()));
                }
                w.write_event(Event::Start(c))?;
                let v = n.to_string();
                w.create_element("v")
                    .write_text_content(BytesText::new(&v))?;
                w.write_event(Event::End(BytesEnd::new("c")))?;
            }
            duke_sheets_core::CellValue::String(s) => {
                let mut c = BytesStart::new("c");
                c.push_attribute(("r", cell_ref.as_str()));
                if xf_id != 0 {
                    c.push_attribute(("s", xf_str.as_str()));
                }
                if let Some(sst_idx) = sst.get(s.as_str()) {
                    c.push_attribute(("t", "s"));
                    w.write_event(Event::Start(c))?;
                    let v = sst_idx.to_string();
                    w.create_element("v")
                        .write_text_content(BytesText::new(&v))?;
                } else {
                    // Fallback to inline string (shouldn't happen)
                    c.push_attribute(("t", "inlineStr"));
                    w.write_event(Event::Start(c))?;
                    w.write_event(Event::Start(BytesStart::new("is")))?;
                    w.create_element("t")
                        .write_text_content(BytesText::new(s.as_str()))?;
                    w.write_event(Event::End(BytesEnd::new("is")))?;
                }
                w.write_event(Event::End(BytesEnd::new("c")))?;
            }
            duke_sheets_core::CellValue::Boolean(b) => {
                let mut c = BytesStart::new("c");
                c.push_attribute(("r", cell_ref.as_str()));
                if xf_id != 0 {
                    c.push_attribute(("s", xf_str.as_str()));
                }
                c.push_attribute(("t", "b"));
                w.write_event(Event::Start(c))?;
                w.create_element("v")
                    .write_text_content(BytesText::new(if *b { "1" } else { "0" }))?;
                w.write_event(Event::End(BytesEnd::new("c")))?;
            }
            duke_sheets_core::CellValue::Error(e) => {
                let mut c = BytesStart::new("c");
                c.push_attribute(("r", cell_ref.as_str()));
                if xf_id != 0 {
                    c.push_attribute(("s", xf_str.as_str()));
                }
                c.push_attribute(("t", "e"));
                w.write_event(Event::Start(c))?;
                w.create_element("v")
                    .write_text_content(BytesText::new(e.as_str()))?;
                w.write_event(Event::End(BytesEnd::new("c")))?;
            }
            duke_sheets_core::CellValue::Empty => {
                // Preserve style-only cells
                if xf_id != 0 {
                    let mut c = BytesStart::new("c");
                    c.push_attribute(("r", cell_ref.as_str()));
                    c.push_attribute(("s", xf_str.as_str()));
                    w.write_event(Event::Empty(c))?;
                }
            }
            duke_sheets_core::CellValue::SpillTarget { .. } => {
                let resolved = worksheet.get_value_at(row, col);
                match &resolved {
                    duke_sheets_core::CellValue::Number(n) => {
                        let mut c = BytesStart::new("c");
                        c.push_attribute(("r", cell_ref.as_str()));
                        if xf_id != 0 {
                            c.push_attribute(("s", xf_str.as_str()));
                        }
                        c.push_attribute(("cm", "2"));
                        w.write_event(Event::Start(c))?;
                        let v = n.to_string();
                        w.create_element("v")
                            .write_text_content(BytesText::new(&v))?;
                        w.write_event(Event::End(BytesEnd::new("c")))?;
                    }
                    duke_sheets_core::CellValue::String(s) => {
                        let mut c = BytesStart::new("c");
                        c.push_attribute(("r", cell_ref.as_str()));
                        if xf_id != 0 {
                            c.push_attribute(("s", xf_str.as_str()));
                        }
                        c.push_attribute(("t", "str"));
                        c.push_attribute(("cm", "2"));
                        w.write_event(Event::Start(c))?;
                        w.create_element("v")
                            .write_text_content(BytesText::new(s.as_str()))?;
                        w.write_event(Event::End(BytesEnd::new("c")))?;
                    }
                    duke_sheets_core::CellValue::Boolean(b) => {
                        let mut c = BytesStart::new("c");
                        c.push_attribute(("r", cell_ref.as_str()));
                        if xf_id != 0 {
                            c.push_attribute(("s", xf_str.as_str()));
                        }
                        c.push_attribute(("t", "b"));
                        c.push_attribute(("cm", "2"));
                        w.write_event(Event::Start(c))?;
                        w.create_element("v")
                            .write_text_content(BytesText::new(if *b { "1" } else { "0" }))?;
                        w.write_event(Event::End(BytesEnd::new("c")))?;
                    }
                    duke_sheets_core::CellValue::Error(e) => {
                        let mut c = BytesStart::new("c");
                        c.push_attribute(("r", cell_ref.as_str()));
                        if xf_id != 0 {
                            c.push_attribute(("s", xf_str.as_str()));
                        }
                        c.push_attribute(("t", "e"));
                        c.push_attribute(("cm", "2"));
                        w.write_event(Event::Start(c))?;
                        w.create_element("v")
                            .write_text_content(BytesText::new(e.as_str()))?;
                        w.write_event(Event::End(BytesEnd::new("c")))?;
                    }
                    _ => {}
                }
            }
            duke_sheets_core::CellValue::RichText(runs) => {
                let mut c = BytesStart::new("c");
                c.push_attribute(("r", cell_ref.as_str()));
                if xf_id != 0 {
                    c.push_attribute(("s", xf_str.as_str()));
                }
                c.push_attribute(("t", "inlineStr"));
                w.write_event(Event::Start(c))?;
                w.write_event(Event::Start(BytesStart::new("is")))?;
                Self::write_rich_text_runs(w, runs)?;
                w.write_event(Event::End(BytesEnd::new("is")))?;
                w.write_event(Event::End(BytesEnd::new("c")))?;
            }
        }
        Ok(())
    }

    /// Write a sequence of rich text runs as `<r>` elements.
    fn write_rich_text_runs(
        w: &mut XmlWriter,
        runs: &[duke_sheets_core::RichTextRun],
    ) -> XlsxResult<()> {
        for run in runs {
            w.write_event(Event::Start(BytesStart::new("r")))?;

            // Write run properties if present
            if let Some(font) = &run.font {
                Self::write_run_properties(w, font)?;
            }

            // Write text with xml:space="preserve" for whitespace
            let needs_preserve = run.text.starts_with(|c: char| c.is_ascii_whitespace())
                || run.text.ends_with(|c: char| c.is_ascii_whitespace());
            if needs_preserve {
                let mut t = BytesStart::new("t");
                t.push_attribute(("xml:space", "preserve"));
                w.write_event(Event::Start(t))?;
            } else {
                w.write_event(Event::Start(BytesStart::new("t")))?;
            }
            w.write_event(Event::Text(BytesText::new(&run.text)))?;
            w.write_event(Event::End(BytesEnd::new("t")))?;

            w.write_event(Event::End(BytesEnd::new("r")))?;
        }
        Ok(())
    }

    /// Write `<rPr>` (run properties) element for a rich text run.
    fn write_run_properties(w: &mut XmlWriter, font: &duke_sheets_core::RunFont) -> XlsxResult<()> {
        w.write_event(Event::Start(BytesStart::new("rPr")))?;

        if let Some(bold) = font.bold {
            if bold {
                w.write_event(Event::Empty(BytesStart::new("b")))?;
            } else {
                let mut tag = BytesStart::new("b");
                tag.push_attribute(("val", "0"));
                w.write_event(Event::Empty(tag))?;
            }
        }
        if let Some(italic) = font.italic {
            if italic {
                w.write_event(Event::Empty(BytesStart::new("i")))?;
            } else {
                let mut tag = BytesStart::new("i");
                tag.push_attribute(("val", "0"));
                w.write_event(Event::Empty(tag))?;
            }
        }
        if let Some(true) = font.strikethrough {
            w.write_event(Event::Empty(BytesStart::new("strike")))?;
        }
        if let Some(underline) = font.underline {
            match underline {
                duke_sheets_core::style::Underline::None => {}
                duke_sheets_core::style::Underline::Single => {
                    w.write_event(Event::Empty(BytesStart::new("u")))?;
                }
                _ => {
                    let val = match underline {
                        duke_sheets_core::style::Underline::Double => "double",
                        duke_sheets_core::style::Underline::SingleAccounting => "singleAccounting",
                        duke_sheets_core::style::Underline::DoubleAccounting => "doubleAccounting",
                        _ => unreachable!(),
                    };
                    let mut tag = BytesStart::new("u");
                    tag.push_attribute(("val", val));
                    w.write_event(Event::Empty(tag))?;
                }
            }
        }
        if let Some(va) = font.vertical_align {
            let val = match va {
                duke_sheets_core::style::FontVerticalAlign::Superscript => "superscript",
                duke_sheets_core::style::FontVerticalAlign::Subscript => "subscript",
                duke_sheets_core::style::FontVerticalAlign::Baseline => "baseline",
            };
            let mut tag = BytesStart::new("vertAlign");
            tag.push_attribute(("val", val));
            w.write_event(Event::Empty(tag))?;
        }
        if let Some(size) = font.size {
            let mut tag = BytesStart::new("sz");
            tag.push_attribute(("val", size.to_string().as_str()));
            w.write_event(Event::Empty(tag))?;
        }
        if let Some(color) = &font.color {
            Self::write_run_color(w, color)?;
        }
        if let Some(name) = &font.name {
            let mut tag = BytesStart::new("rFont");
            tag.push_attribute(("val", name.as_str()));
            w.write_event(Event::Empty(tag))?;
        }
        if let Some(family) = font.family {
            let mut tag = BytesStart::new("family");
            tag.push_attribute(("val", family.to_string().as_str()));
            w.write_event(Event::Empty(tag))?;
        }
        if let Some(charset) = font.charset {
            let mut tag = BytesStart::new("charset");
            tag.push_attribute(("val", charset.to_string().as_str()));
            w.write_event(Event::Empty(tag))?;
        }
        if let Some(scheme) = &font.scheme {
            let mut tag = BytesStart::new("scheme");
            tag.push_attribute(("val", scheme.as_str()));
            w.write_event(Event::Empty(tag))?;
        }

        w.write_event(Event::End(BytesEnd::new("rPr")))?;
        Ok(())
    }

    /// Write a `<color>` element for a run property.
    fn write_run_color(w: &mut XmlWriter, color: &duke_sheets_core::Color) -> XlsxResult<()> {
        let mut tag = BytesStart::new("color");
        match color {
            duke_sheets_core::Color::Rgb { r, g, b } => {
                tag.push_attribute(("rgb", format!("FF{:02X}{:02X}{:02X}", r, g, b).as_str()));
            }
            duke_sheets_core::Color::Argb { a, r, g, b } => {
                tag.push_attribute((
                    "rgb",
                    format!("{:02X}{:02X}{:02X}{:02X}", a, r, g, b).as_str(),
                ));
            }
            duke_sheets_core::Color::Theme { index, tint } => {
                tag.push_attribute(("theme", index.to_string().as_str()));
                if *tint != 0 {
                    let tint_f = *tint as f64 / 100.0;
                    tag.push_attribute(("tint", tint_f.to_string().as_str()));
                }
            }
            duke_sheets_core::Color::Indexed(idx) => {
                tag.push_attribute(("indexed", idx.to_string().as_str()));
            }
            duke_sheets_core::Color::Auto => {
                tag.push_attribute(("auto", "1"));
            }
        }
        w.write_event(Event::Empty(tag))?;
        Ok(())
    }

    fn write_sheet_protection(
        w: &mut XmlWriter,
        sheet: &duke_sheets_core::Worksheet,
    ) -> XlsxResult<()> {
        let prot = match sheet.protection() {
            Some(p) if p.protected => p,
            _ => return Ok(()),
        };

        let mut tag = BytesStart::new("sheetProtection");
        tag.push_attribute(("sheet", "1"));

        if let Some(hash) = prot.password_hash {
            let h = format!("{:04X}", hash);
            tag.push_attribute(("password", h.as_str()));
        }

        // ECMA-376 §18.3.1.85: absent or "1" = NOT allowed.
        // We emit "0" when our model says the action IS allowed.
        macro_rules! prot_allow {
            ($field:expr, $attr:literal) => {
                if $field {
                    tag.push_attribute(($attr, "0"));
                }
            };
        }

        prot_allow!(prot.format_cells, "formatCells");
        prot_allow!(prot.format_columns, "formatColumns");
        prot_allow!(prot.format_rows, "formatRows");
        prot_allow!(prot.insert_columns, "insertColumns");
        prot_allow!(prot.insert_rows, "insertRows");
        prot_allow!(prot.insert_hyperlinks, "insertHyperlinks");
        prot_allow!(prot.delete_columns, "deleteColumns");
        prot_allow!(prot.delete_rows, "deleteRows");
        prot_allow!(prot.sort, "sort");
        prot_allow!(prot.auto_filter, "autoFilter");
        prot_allow!(prot.pivot_tables, "pivotTables");

        // selectLockedCells/selectUnlockedCells: absent = allowed (inverted)
        if !prot.select_locked_cells {
            tag.push_attribute(("selectLockedCells", "1"));
        }
        if !prot.select_unlocked_cells {
            tag.push_attribute(("selectUnlockedCells", "1"));
        }

        w.write_event(Event::Empty(tag))?;
        Ok(())
    }

    fn write_auto_filter(w: &mut XmlWriter, sheet: &duke_sheets_core::Worksheet) -> XlsxResult<()> {
        let Some(auto_filter) = sheet.auto_filter() else {
            return Ok(());
        };

        let range = auto_filter.range.to_string();
        if auto_filter.filter_columns.is_empty() {
            let mut elem = BytesStart::new("autoFilter");
            elem.push_attribute(("ref", range.as_str()));
            w.write_event(Event::Empty(elem))?;
            return Ok(());
        }

        let mut elem = BytesStart::new("autoFilter");
        elem.push_attribute(("ref", range.as_str()));
        w.write_event(Event::Start(elem))?;

        for column in &auto_filter.filter_columns {
            Self::write_filter_column(w, column)?;
        }

        w.write_event(Event::End(BytesEnd::new("autoFilter")))?;
        Ok(())
    }

    fn write_filter_column(
        w: &mut XmlWriter,
        column: &duke_sheets_core::FilterColumn,
    ) -> XlsxResult<()> {
        let col_id = column.col_id.to_string();
        let mut elem = BytesStart::new("filterColumn");
        elem.push_attribute(("colId", col_id.as_str()));
        if column.hidden_button {
            elem.push_attribute(("hiddenButton", "1"));
        }
        if !column.show_button {
            elem.push_attribute(("showButton", "0"));
        }
        w.write_event(Event::Start(elem))?;

        match &column.filter {
            duke_sheets_core::ColumnFilter::Values(values) => {
                let mut filters = BytesStart::new("filters");
                if values.blank {
                    filters.push_attribute(("blank", "1"));
                }
                w.write_event(Event::Start(filters))?;
                for value in &values.values {
                    let mut filter = BytesStart::new("filter");
                    filter.push_attribute(("val", value.as_str()));
                    w.write_event(Event::Empty(filter))?;
                }
                w.write_event(Event::End(BytesEnd::new("filters")))?;
            }
            duke_sheets_core::ColumnFilter::Custom(custom) => {
                let mut custom_filters = BytesStart::new("customFilters");
                if custom.and {
                    custom_filters.push_attribute(("and", "1"));
                }
                w.write_event(Event::Start(custom_filters))?;
                for condition in &custom.conditions {
                    let mut custom_filter = BytesStart::new("customFilter");
                    custom_filter.push_attribute(("operator", condition.operator.to_ooxml()));
                    custom_filter.push_attribute(("val", condition.value.as_str()));
                    w.write_event(Event::Empty(custom_filter))?;
                }
                w.write_event(Event::End(BytesEnd::new("customFilters")))?;
            }
            duke_sheets_core::ColumnFilter::Top10(top10) => {
                let mut top10_el = BytesStart::new("top10");
                top10_el.push_attribute(("top", if top10.top { "1" } else { "0" }));
                top10_el.push_attribute(("percent", if top10.percent { "1" } else { "0" }));
                let val = top10.val.to_string();
                top10_el.push_attribute(("val", val.as_str()));
                let filter_val = top10.filter_val.map(|v| v.to_string());
                if let Some(filter_val) = filter_val.as_deref() {
                    top10_el.push_attribute(("filterVal", filter_val));
                }
                w.write_event(Event::Empty(top10_el))?;
            }
            duke_sheets_core::ColumnFilter::Dynamic(dynamic) => {
                let mut dynamic_el = BytesStart::new("dynamicFilter");
                dynamic_el.push_attribute(("type", dynamic.filter_type.to_ooxml()));
                let val = dynamic.val.map(|v| v.to_string());
                if let Some(val) = val.as_deref() {
                    dynamic_el.push_attribute(("val", val));
                }
                let max_val = dynamic.max_val.map(|v| v.to_string());
                if let Some(max_val) = max_val.as_deref() {
                    dynamic_el.push_attribute(("maxVal", max_val));
                }
                w.write_event(Event::Empty(dynamic_el))?;
            }
            duke_sheets_core::ColumnFilter::Color(color) => {
                let mut color_el = BytesStart::new("colorFilter");
                let dxf_id = color.dxf_id.map(|v| v.to_string());
                if let Some(dxf_id) = dxf_id.as_deref() {
                    color_el.push_attribute(("dxfId", dxf_id));
                }
                color_el.push_attribute(("cellColor", if color.cell_color { "1" } else { "0" }));
                w.write_event(Event::Empty(color_el))?;
            }
        }

        w.write_event(Event::End(BytesEnd::new("filterColumn")))?;
        Ok(())
    }

    fn write_merge_cells(w: &mut XmlWriter, sheet: &duke_sheets_core::Worksheet) -> XlsxResult<()> {
        let merged_regions = sheet.merged_regions();
        if merged_regions.is_empty() {
            return Ok(());
        }

        let count = merged_regions.len().to_string();
        let mut tag = BytesStart::new("mergeCells");
        tag.push_attribute(("count", count.as_str()));
        w.write_event(Event::Start(tag))?;

        for range in merged_regions {
            let r = range.to_string();
            w.create_element("mergeCell")
                .with_attribute(("ref", r.as_str()))
                .write_empty()?;
        }

        w.write_event(Event::End(BytesEnd::new("mergeCells")))?;
        Ok(())
    }

    fn write_page_setup(w: &mut XmlWriter, sheet: &duke_sheets_core::Worksheet) -> XlsxResult<()> {
        let ps = sheet.page_setup();
        let def = duke_sheets_core::PageSetup::default();

        let margins_differ = (ps.left_margin - def.left_margin).abs() > 1e-9
            || (ps.right_margin - def.right_margin).abs() > 1e-9
            || (ps.top_margin - def.top_margin).abs() > 1e-9
            || (ps.bottom_margin - def.bottom_margin).abs() > 1e-9
            || (ps.header_margin - def.header_margin).abs() > 1e-9
            || (ps.footer_margin - def.footer_margin).abs() > 1e-9;

        let setup_differs = ps.paper_size != def.paper_size
            || ps.orientation != def.orientation
            || ps.scale != def.scale
            || ps.fit_to_width.is_some()
            || ps.fit_to_height.is_some();

        let print_options_differ =
            ps.print_gridlines != def.print_gridlines || ps.print_headings != def.print_headings;

        let header_footer_differs = ps.odd_header.is_some()
            || ps.odd_footer.is_some()
            || ps.even_header.is_some()
            || ps.even_footer.is_some()
            || ps.first_header.is_some()
            || ps.first_footer.is_some()
            || ps.different_odd_even
            || ps.different_first
            || !ps.scale_with_doc
            || !ps.align_with_margins;

        if print_options_differ {
            let mut el = BytesStart::new("printOptions");
            if ps.print_gridlines {
                el.push_attribute(("gridLines", "1"));
            }
            if ps.print_headings {
                el.push_attribute(("headings", "1"));
            }
            w.write_event(Event::Empty(el))?;
        }

        if margins_differ {
            let left = ps.left_margin.to_string();
            let right = ps.right_margin.to_string();
            let top = ps.top_margin.to_string();
            let bottom = ps.bottom_margin.to_string();
            let header = ps.header_margin.to_string();
            let footer = ps.footer_margin.to_string();
            w.create_element("pageMargins")
                .with_attribute(("left", left.as_str()))
                .with_attribute(("right", right.as_str()))
                .with_attribute(("top", top.as_str()))
                .with_attribute(("bottom", bottom.as_str()))
                .with_attribute(("header", header.as_str()))
                .with_attribute(("footer", footer.as_str()))
                .write_empty()?;
        }

        if setup_differs {
            let orientation = match ps.orientation {
                duke_sheets_core::PageOrientation::Portrait => "portrait",
                duke_sheets_core::PageOrientation::Landscape => "landscape",
            };
            let paper = ps.paper_size.to_string();
            let mut el = w
                .create_element("pageSetup")
                .with_attribute(("paperSize", paper.as_str()))
                .with_attribute(("orientation", orientation));

            if ps.scale != 100 {
                let s = ps.scale.to_string();
                // Need to bind so the string lives long enough
                el = el.with_attribute(("scale", s.as_str()));
            }
            if let Some(fw) = ps.fit_to_width {
                let s = fw.to_string();
                el = el.with_attribute(("fitToWidth", s.as_str()));
            }
            if let Some(fh) = ps.fit_to_height {
                let s = fh.to_string();
                el = el.with_attribute(("fitToHeight", s.as_str()));
            }
            el.write_empty()?;
        }
        if header_footer_differs {
            let mut el = BytesStart::new("headerFooter");
            // Only write non-default attribute values
            if ps.different_odd_even {
                el.push_attribute(("differentOddEven", "1"));
            }
            if ps.different_first {
                el.push_attribute(("differentFirst", "1"));
            }
            if !ps.scale_with_doc {
                el.push_attribute(("scaleWithDoc", "0"));
            }
            if !ps.align_with_margins {
                el.push_attribute(("alignWithMargins", "0"));
            }
            w.write_event(Event::Start(el))?;
            // Children in spec order: oddHeader, oddFooter, evenHeader, evenFooter, firstHeader, firstFooter
            if let Some(header) = &ps.odd_header {
                w.create_element("oddHeader")
                    .write_text_content(quick_xml::events::BytesText::new(header))?;
            }
            if let Some(footer) = &ps.odd_footer {
                w.create_element("oddFooter")
                    .write_text_content(quick_xml::events::BytesText::new(footer))?;
            }
            if let Some(header) = &ps.even_header {
                w.create_element("evenHeader")
                    .write_text_content(quick_xml::events::BytesText::new(header))?;
            }
            if let Some(footer) = &ps.even_footer {
                w.create_element("evenFooter")
                    .write_text_content(quick_xml::events::BytesText::new(footer))?;
            }
            if let Some(header) = &ps.first_header {
                w.create_element("firstHeader")
                    .write_text_content(quick_xml::events::BytesText::new(header))?;
            }
            if let Some(footer) = &ps.first_footer {
                w.create_element("firstFooter")
                    .write_text_content(quick_xml::events::BytesText::new(footer))?;
            }
            w.write_event(Event::End(BytesEnd::new("headerFooter")))?;
        }

        Ok(())
    }

    fn write_page_breaks(w: &mut XmlWriter, sheet: &duke_sheets_core::Worksheet) -> XlsxResult<()> {
        let row_breaks = sheet.row_breaks();
        if !row_breaks.is_empty() {
            let mut sorted: Vec<_> = row_breaks.to_vec();
            sorted.sort_by_key(|b| b.id);
            let count = sorted.len().to_string();
            let manual_count = sorted.iter().filter(|b| b.man).count().to_string();
            let mut el = BytesStart::new("rowBreaks");
            el.push_attribute(("count", count.as_str()));
            el.push_attribute(("manualBreakCount", manual_count.as_str()));
            w.write_event(Event::Start(el))?;
            for brk in &sorted {
                let id = brk.id.to_string();
                let max = brk.max.to_string();
                let mut be = BytesStart::new("brk");
                be.push_attribute(("id", id.as_str()));
                if brk.min > 0 {
                    let min = brk.min.to_string();
                    be.push_attribute(("min", min.as_str()));
                }
                be.push_attribute(("max", max.as_str()));
                if brk.man {
                    be.push_attribute(("man", "1"));
                }
                if brk.pt {
                    be.push_attribute(("pt", "1"));
                }
                w.write_event(Event::Empty(be))?;
            }
            w.write_event(Event::End(BytesEnd::new("rowBreaks")))?;
        }

        let col_breaks = sheet.col_breaks();
        if !col_breaks.is_empty() {
            let mut sorted: Vec<_> = col_breaks.to_vec();
            sorted.sort_by_key(|b| b.id);
            let count = sorted.len().to_string();
            let manual_count = sorted.iter().filter(|b| b.man).count().to_string();
            let mut el = BytesStart::new("colBreaks");
            el.push_attribute(("count", count.as_str()));
            el.push_attribute(("manualBreakCount", manual_count.as_str()));
            w.write_event(Event::Start(el))?;
            for brk in &sorted {
                let id = brk.id.to_string();
                let max = brk.max.to_string();
                let mut be = BytesStart::new("brk");
                be.push_attribute(("id", id.as_str()));
                if brk.min > 0 {
                    let min = brk.min.to_string();
                    be.push_attribute(("min", min.as_str()));
                }
                be.push_attribute(("max", max.as_str()));
                if brk.man {
                    be.push_attribute(("man", "1"));
                }
                if brk.pt {
                    be.push_attribute(("pt", "1"));
                }
                w.write_event(Event::Empty(be))?;
            }
            w.write_event(Event::End(BytesEnd::new("colBreaks")))?;
        }

        Ok(())
    }

    fn write_hyperlinks(
        w: &mut XmlWriter,
        sheet: &duke_sheets_core::Worksheet,
        rels: &mut Vec<WorksheetRelationship>,
    ) -> XlsxResult<()> {
        let mut hyperlinks: Vec<_> = sheet.hyperlinks().iter().collect();
        hyperlinks.sort_by_key(|(addr, _)| (addr.row, addr.col));

        if hyperlinks.is_empty() {
            return Ok(());
        }

        w.write_event(Event::Start(BytesStart::new("hyperlinks")))?;

        let mut next_rid = Self::next_worksheet_rel_id(rels);

        for (addr, hyperlink) in hyperlinks {
            let cell_ref = addr.to_a1_string();
            let mut tag = BytesStart::new("hyperlink");
            tag.push_attribute(("ref", cell_ref.as_str()));

            let target = hyperlink.target.trim();
            let has_location = hyperlink
                .location
                .as_ref()
                .map(|s| !s.is_empty())
                .unwrap_or(false);
            let is_internal = has_location
                && (target.is_empty()
                    || target.starts_with('#')
                    || !Self::is_external_target(target));

            if is_internal {
                let location = hyperlink
                    .location
                    .as_deref()
                    .filter(|s| !s.is_empty())
                    .or_else(|| target.strip_prefix('#'));
                if let Some(location) = location {
                    tag.push_attribute(("location", location));
                }
            } else if !target.is_empty() {
                let rid = format!("rId{}", next_rid);
                next_rid += 1;
                tag.push_attribute(("r:id", rid.as_str()));
                if let Some(location) = hyperlink.location.as_deref().filter(|s| !s.is_empty()) {
                    tag.push_attribute(("location", location));
                }
                rels.push(WorksheetRelationship {
                    id: rid,
                    rel_type: RT_HYPERLINK,
                    target: target.to_string(),
                    target_mode: Some("External"),
                });
            }

            if let Some(display) = hyperlink.display.as_deref() {
                if !display.is_empty() {
                    tag.push_attribute(("display", display));
                }
            }
            if let Some(tooltip) = hyperlink.tooltip.as_deref() {
                if !tooltip.is_empty() {
                    tag.push_attribute(("tooltip", tooltip));
                }
            }

            w.write_event(Event::Empty(tag))?;
        }

        w.write_event(Event::End(BytesEnd::new("hyperlinks")))?;
        Ok(())
    }

    fn next_worksheet_rel_id(rels: &[WorksheetRelationship]) -> usize {
        rels.iter()
            .filter_map(|rel| {
                rel.id
                    .strip_prefix("rId")
                    .and_then(|s| s.parse::<usize>().ok())
            })
            .max()
            .map_or(1, |max_id| max_id + 1)
    }

    fn is_external_target(target: &str) -> bool {
        target.contains("://") || target.starts_with("mailto:") || target.starts_with("file:")
    }

    // -----------------------------------------------------------------------
    // Worksheet relationships
    // -----------------------------------------------------------------------

    fn write_worksheet_rels<W: Write + Seek>(
        zip: &mut zip::ZipWriter<W>,
        sheet_index: usize,
        rels: &[WorksheetRelationship],
    ) -> XlsxResult<()> {
        let path = format!("xl/worksheets/_rels/sheet{}.xml.rels", sheet_index + 1);
        write_xml_part(zip, &path, |w| {
            let mut tag = BytesStart::new("Relationships");
            tag.push_attribute(("xmlns", NS_RELATIONSHIPS));
            w.write_event(Event::Start(tag))?;

            for rel in rels {
                let mut relationship = w
                    .create_element("Relationship")
                    .with_attribute(("Id", rel.id.as_str()))
                    .with_attribute(("Type", rel.rel_type))
                    .with_attribute(("Target", rel.target.as_str()));
                if let Some(target_mode) = rel.target_mode {
                    relationship = relationship.with_attribute(("TargetMode", target_mode));
                }
                relationship.write_empty()?;
            }

            w.write_event(Event::End(BytesEnd::new("Relationships")))?;
            Ok(())
        })
    }
}

#[cfg(test)]
mod tests {
    use super::*;
    use crate::reader::XlsxReader;
    use duke_sheets_core::{CellRange, ConditionalFormatRule, Hyperlink, SplitPanes};
    use std::io::Read;

    fn read_zip_entry(bytes: Vec<u8>, path: &str) -> String {
        let cursor = Cursor::new(bytes);
        let mut archive = zip::ZipArchive::new(cursor).expect("open zip");
        let mut file = archive.by_name(path).expect("zip entry exists");
        let mut s = String::new();
        file.read_to_string(&mut s).expect("read zip entry utf8");
        s
    }

    #[test]
    fn test_writer_preserves_theme_tab_color_attrs() {
        let mut wb = Workbook::new();
        wb.worksheet_mut(0)
            .unwrap()
            .set_tab_color(Some(Color::theme(4, 50)));

        let mut out = Cursor::new(Vec::new());
        XlsxWriter::write(&wb, &mut out).expect("write workbook");
        let xml = read_zip_entry(out.into_inner(), "xl/worksheets/sheet1.xml");

        assert!(xml.contains("<tabColor theme=\"4\" tint=\"0.5\""));
    }

    #[test]
    fn test_writer_preserves_theme_and_indexed_cf_colors() {
        let mut wb = Workbook::new();
        let sheet = wb.worksheet_mut(0).unwrap();
        sheet.set_cell_value("A1", 1.0).unwrap();
        sheet.set_cell_value("A2", 2.0).unwrap();
        sheet.set_cell_value("A3", 3.0).unwrap();

        let rule = ConditionalFormatRule::color_scale_2(Color::theme(4, -25), Color::Indexed(12))
            .with_range(CellRange::parse("A1:A3").unwrap());
        sheet.add_conditional_format(rule);

        let mut out = Cursor::new(Vec::new());
        XlsxWriter::write(&wb, &mut out).expect("write workbook");
        let xml = read_zip_entry(out.into_inner(), "xl/worksheets/sheet1.xml");

        assert!(xml.contains("<color theme=\"4\" tint=\"-0.25\""));
        assert!(xml.contains("<color indexed=\"12\""));
    }

    #[test]
    fn test_writer_emits_theme_part_and_relationships() {
        let wb = Workbook::new();
        let mut out = Cursor::new(Vec::new());
        XlsxWriter::write(&wb, &mut out).expect("write workbook");
        let bytes = out.into_inner();

        let content_types = read_zip_entry(bytes.clone(), "[Content_Types].xml");
        assert!(content_types.contains("/xl/theme/theme1.xml"));
        assert!(content_types.contains(CT_THEME));

        let rels = read_zip_entry(bytes.clone(), "xl/_rels/workbook.xml.rels");
        assert!(rels.contains(RT_THEME));
        assert!(rels.contains("Target=\"theme/theme1.xml\""));

        let theme = read_zip_entry(bytes, "xl/theme/theme1.xml");
        assert!(theme.contains("<a:theme"));
        assert!(theme.contains("<a:clrScheme"));
    }

    #[test]
    fn test_writer_emits_outline_and_collapsed_attrs() {
        let mut wb = Workbook::new();
        let sheet = wb.worksheet_mut(0).unwrap();
        sheet.set_row_outline_level(1, 2);
        sheet.set_row_collapsed(1, true);
        sheet.set_column_outline_level(2, 3);
        sheet.set_column_collapsed(2, true);

        let mut out = Cursor::new(Vec::new());
        XlsxWriter::write(&wb, &mut out).expect("write workbook");
        let xml = read_zip_entry(out.into_inner(), "xl/worksheets/sheet1.xml");

        assert!(xml.contains("<row r=\"2\" outlineLevel=\"2\" collapsed=\"1\""));
        assert!(xml.contains("<col min=\"3\" max=\"3\""));
        assert!(xml.contains("outlineLevel=\"3\""));
        assert!(xml.contains("collapsed=\"1\""));
    }

    #[test]
    fn test_writer_emits_sheet_view_zoom_and_selection() {
        let mut wb = Workbook::new();
        let sheet = wb.worksheet_mut(0).unwrap();
        sheet.set_selected(true);
        sheet.set_zoom_scale(Some(125));
        sheet.set_selection_active_cell(4, 3); // D5
        sheet.set_selection_range(Some(CellRange::parse("D5:E6").unwrap()));

        let mut out = Cursor::new(Vec::new());
        XlsxWriter::write(&wb, &mut out).expect("write workbook");
        let xml = read_zip_entry(out.into_inner(), "xl/worksheets/sheet1.xml");

        assert!(xml.contains("<sheetView workbookViewId=\"0\" tabSelected=\"1\" zoomScale=\"125\""));
        assert!(xml.contains("<selection activeCell=\"D5\" sqref=\"D5:E6\""));
    }

    #[test]
    fn test_writer_emits_split_pane_sheet_view() {
        let mut wb = Workbook::new();
        let sheet = wb.worksheet_mut(0).unwrap();
        sheet.set_split_panes(Some(SplitPanes {
            x_split: 2000.0,
            y_split: 3000.0,
            top_left: Some((3, 2)), // C4
            active_pane: Some("bottomRight".to_string()),
        }));
        sheet.set_selection_active_cell(4, 3); // D5
        sheet.set_selection_range(Some(CellRange::parse("D5").unwrap()));

        let mut out = Cursor::new(Vec::new());
        XlsxWriter::write(&wb, &mut out).expect("write workbook");
        let xml = read_zip_entry(out.into_inner(), "xl/worksheets/sheet1.xml");

        assert!(xml.contains(
            "<pane xSplit=\"2000\" ySplit=\"3000\" topLeftCell=\"C4\" activePane=\"bottomRight\" state=\"split\""
        ));
        assert!(xml.contains("<selection pane=\"bottomRight\" activeCell=\"D5\" sqref=\"D5\""));
    }

    #[test]
    fn test_writer_emits_multiple_selections() {
        use duke_sheets_core::worksheet::Selection;

        let mut wb = Workbook::new();
        let sheet = wb.worksheet_mut(0).unwrap();
        sheet.set_freeze_panes(1, 0); // Freeze top row
        sheet.set_selections(vec![
            Selection {
                pane: Some("topLeft".to_string()),
                active_cell: Some("A1".to_string()),
                sqref: Some("A1".to_string()),
            },
            Selection {
                pane: Some("bottomLeft".to_string()),
                active_cell: Some("C5".to_string()),
                sqref: Some("C5:D8 F2:G3".to_string()),
            },
        ]);

        let mut out = Cursor::new(Vec::new());
        XlsxWriter::write(&wb, &mut out).expect("write workbook");
        let xml = read_zip_entry(out.into_inner(), "xl/worksheets/sheet1.xml");

        // Both selections should be present
        assert!(
            xml.contains(r#"<selection pane="topLeft" activeCell="A1" sqref="A1""#),
            "missing topLeft selection in: {xml}"
        );
        assert!(
            xml.contains(r#"<selection pane="bottomLeft" activeCell="C5" sqref="C5:D8 F2:G3""#),
            "missing bottomLeft selection in: {xml}"
        );
    }

    #[test]
    fn test_writer_emits_multi_range_sqref() {
        use duke_sheets_core::worksheet::Selection;

        let mut wb = Workbook::new();
        let sheet = wb.worksheet_mut(0).unwrap();
        sheet.add_selection(Selection {
            pane: None,
            active_cell: Some("A1".to_string()),
            sqref: Some("A1:B2 D4:E5 G7".to_string()),
        });

        let mut out = Cursor::new(Vec::new());
        XlsxWriter::write(&wb, &mut out).expect("write workbook");
        let xml = read_zip_entry(out.into_inner(), "xl/worksheets/sheet1.xml");

        assert!(
            xml.contains(r#"sqref="A1:B2 D4:E5 G7""#),
            "multi-range sqref not found in: {xml}"
        );
    }

    #[test]
    fn test_writer_emits_page_setup_print_options_and_header_footer() {
        let mut wb = Workbook::new();
        let sheet = wb.worksheet_mut(0).unwrap();
        let mut ps = sheet.page_setup().clone();
        ps.paper_size = 9;
        ps.orientation = duke_sheets_core::PageOrientation::Landscape;
        ps.scale = 85;
        ps.fit_to_width = Some(1);
        ps.fit_to_height = Some(2);
        ps.left_margin = 0.5;
        ps.right_margin = 0.6;
        ps.top_margin = 0.7;
        ps.bottom_margin = 0.8;
        ps.header_margin = 0.2;
        ps.footer_margin = 0.25;
        ps.print_gridlines = true;
        ps.print_headings = true;
        ps.odd_header = Some("&LLeft&CCenter".to_string());
        ps.odd_footer = Some("&RPage &P".to_string());
        sheet.set_page_setup(ps);

        let mut out = Cursor::new(Vec::new());
        XlsxWriter::write(&wb, &mut out).expect("write workbook");
        let xml = read_zip_entry(out.into_inner(), "xl/worksheets/sheet1.xml");

        assert!(xml.contains("<pageMargins left=\"0.5\" right=\"0.6\" top=\"0.7\" bottom=\"0.8\" header=\"0.2\" footer=\"0.25\""));
        assert!(xml.contains("<pageSetup paperSize=\"9\" orientation=\"landscape\""));
        assert!(xml.contains("scale=\"85\""));
        assert!(xml.contains("fitToWidth=\"1\""));
        assert!(xml.contains("fitToHeight=\"2\""));
        assert!(xml.contains("<printOptions gridLines=\"1\" headings=\"1\""));
        assert!(xml.contains("<headerFooter>"));
        assert!(xml.contains("<oddHeader>&amp;LLeft&amp;CCenter</oddHeader>"));
        assert!(xml.contains("<oddFooter>&amp;RPage &amp;P</oddFooter>"));
    }

    #[test]
    fn test_writer_emits_even_first_header_footer_and_flags() {
        let mut wb = Workbook::new();
        let sheet = wb.worksheet_mut(0).unwrap();
        let mut ps = sheet.page_setup().clone();
        ps.odd_header = Some("&COdd".to_string());
        ps.odd_footer = Some("&COdd Footer".to_string());
        ps.even_header = Some("&CEven".to_string());
        ps.even_footer = Some("&CEven Footer".to_string());
        ps.first_header = Some("&CFirst".to_string());
        ps.first_footer = Some("&CFirst Footer".to_string());
        ps.different_odd_even = true;
        ps.different_first = true;
        ps.scale_with_doc = false;
        ps.align_with_margins = false;
        sheet.set_page_setup(ps);

        let mut out = Cursor::new(Vec::new());
        XlsxWriter::write(&wb, &mut out).expect("write workbook");
        let xml = read_zip_entry(out.into_inner(), "xl/worksheets/sheet1.xml");

        // Attributes on headerFooter element
        assert!(
            xml.contains("differentOddEven=\"1\""),
            "missing differentOddEven"
        );
        assert!(
            xml.contains("differentFirst=\"1\""),
            "missing differentFirst"
        );
        assert!(xml.contains("scaleWithDoc=\"0\""), "missing scaleWithDoc");
        assert!(
            xml.contains("alignWithMargins=\"0\""),
            "missing alignWithMargins"
        );

        // All six child elements in spec order
        assert!(xml.contains("<oddHeader>&amp;COdd</oddHeader>"));
        assert!(xml.contains("<oddFooter>&amp;COdd Footer</oddFooter>"));
        assert!(xml.contains("<evenHeader>&amp;CEven</evenHeader>"));
        assert!(xml.contains("<evenFooter>&amp;CEven Footer</evenFooter>"));
        assert!(xml.contains("<firstHeader>&amp;CFirst</firstHeader>"));
        assert!(xml.contains("<firstFooter>&amp;CFirst Footer</firstFooter>"));

        // Verify element order: oddHeader before evenHeader before firstHeader
        let odd_pos = xml.find("<oddHeader>").unwrap();
        let even_pos = xml.find("<evenHeader>").unwrap();
        let first_pos = xml.find("<firstHeader>").unwrap();
        assert!(odd_pos < even_pos, "oddHeader must come before evenHeader");
        assert!(
            even_pos < first_pos,
            "evenHeader must come before firstHeader"
        );
    }

    #[test]
    fn test_writer_omits_default_header_footer_flags() {
        let mut wb = Workbook::new();
        let sheet = wb.worksheet_mut(0).unwrap();
        let mut ps = sheet.page_setup().clone();
        ps.odd_header = Some("&CTest".to_string());
        // Leave all flags at defaults
        sheet.set_page_setup(ps);

        let mut out = Cursor::new(Vec::new());
        XlsxWriter::write(&wb, &mut out).expect("write workbook");
        let xml = read_zip_entry(out.into_inner(), "xl/worksheets/sheet1.xml");

        // headerFooter element should have no attributes
        assert!(
            xml.contains("<headerFooter>"),
            "should have plain headerFooter tag"
        );
        assert!(
            !xml.contains("differentOddEven"),
            "should not emit default differentOddEven"
        );
        assert!(
            !xml.contains("differentFirst"),
            "should not emit default differentFirst"
        );
        assert!(
            !xml.contains("scaleWithDoc"),
            "should not emit default scaleWithDoc"
        );
        assert!(
            !xml.contains("alignWithMargins"),
            "should not emit default alignWithMargins"
        );
    }

    #[test]
    fn test_hyperlinks_double_roundtrip_preserved() {
        let mut wb = Workbook::new();
        let sheet = wb.worksheet_mut(0).unwrap();

        sheet.set_cell_value("A1", "External").unwrap();
        sheet
            .set_hyperlink(
                "A1",
                Hyperlink {
                    target: "https://example.com".to_string(),
                    display: Some("Example".to_string()),
                    tooltip: Some("Visit site".to_string()),
                    location: None,
                },
            )
            .unwrap();

        sheet.set_cell_value("B2", "Internal").unwrap();
        sheet
            .set_hyperlink(
                "B2",
                Hyperlink {
                    target: "#Sheet1!A1".to_string(),
                    display: Some("Go to A1".to_string()),
                    tooltip: None,
                    location: Some("Sheet1!A1".to_string()),
                },
            )
            .unwrap();

        let mut first_write = Cursor::new(Vec::new());
        XlsxWriter::write(&wb, &mut first_write).unwrap();
        let first_bytes = first_write.into_inner();

        let sheet_rels = read_zip_entry(first_bytes.clone(), "xl/worksheets/_rels/sheet1.xml.rels");
        assert!(sheet_rels.contains(RT_HYPERLINK));
        assert!(sheet_rels.contains("Target=\"https://example.com\""));

        let wb2 = XlsxReader::read(Cursor::new(first_bytes)).unwrap();
        let sheet2 = wb2.worksheet(0).unwrap();
        assert_eq!(sheet2.hyperlink_count(), 2);
        assert_eq!(
            sheet2.hyperlink("A1").unwrap().target,
            "https://example.com".to_string()
        );
        assert_eq!(
            sheet2.hyperlink("B2").unwrap().location.as_deref(),
            Some("Sheet1!A1")
        );

        let mut second_write = Cursor::new(Vec::new());
        XlsxWriter::write(&wb2, &mut second_write).unwrap();
        let second_bytes = second_write.into_inner();

        let wb3 = XlsxReader::read(Cursor::new(second_bytes)).unwrap();
        let sheet3 = wb3.worksheet(0).unwrap();
        assert_eq!(sheet3.hyperlink_count(), 2);

        let a1 = sheet3.hyperlink("A1").unwrap();
        assert_eq!(a1.target, "https://example.com");
        assert_eq!(a1.display.as_deref(), Some("Example"));
        assert_eq!(a1.tooltip.as_deref(), Some("Visit site"));

        let b2 = sheet3.hyperlink("B2").unwrap();
        assert_eq!(b2.target, "#Sheet1!A1");
        assert_eq!(b2.location.as_deref(), Some("Sheet1!A1"));
        assert_eq!(b2.display.as_deref(), Some("Go to A1"));
    }

    #[test]
    fn test_writer_emits_table_parts() {
        use duke_sheets_core::table::{Table, TableColumn, TableStyleInfo};

        let mut wb = Workbook::new();
        let sheet = wb.worksheet_mut(0).unwrap();
        sheet.set_cell_value("A1", "Name").unwrap();
        sheet.set_cell_value("B1", "Score").unwrap();
        sheet.set_cell_value("A2", "Alice").unwrap();
        sheet.set_cell_value("B2", 95.0).unwrap();

        let mut table = Table::new(1, "Scores", CellRange::parse("A1:B3").unwrap());
        table.columns = vec![TableColumn::new(1, "Name"), TableColumn::new(2, "Score")];
        table.style_info = Some(TableStyleInfo {
            name: Some("TableStyleMedium2".into()),
            show_row_stripes: true,
            ..TableStyleInfo::default()
        });
        sheet.add_table(table);

        let mut out = Cursor::new(Vec::new());
        XlsxWriter::write(&wb, &mut out).expect("write workbook");
        let bytes = out.into_inner();

        // Check sheet XML has tableParts
        let sheet_xml = read_zip_entry(bytes.clone(), "xl/worksheets/sheet1.xml");
        assert!(
            sheet_xml.contains("<tableParts count=\"1\">"),
            "missing tableParts"
        );
        assert!(
            sheet_xml.contains("<tablePart r:id="),
            "missing tablePart ref"
        );

        // Check table XML exists and has correct structure
        let table_xml = read_zip_entry(bytes.clone(), "xl/tables/table1.xml");
        assert!(table_xml.contains(r#"name="Scores""#), "wrong name");
        assert!(
            table_xml.contains(r#"displayName="Scores""#),
            "wrong displayName"
        );
        assert!(table_xml.contains(r#"ref="A1:B3""#), "wrong ref");
        assert!(table_xml.contains("<autoFilter"), "missing autoFilter");
        assert!(
            table_xml.contains(r#"<tableColumns count="2""#),
            "wrong column count"
        );
        assert!(table_xml.contains(r#"name="Name""#), "missing col Name");
        assert!(table_xml.contains(r#"name="Score""#), "missing col Score");
        assert!(
            table_xml.contains(r#"name="TableStyleMedium2""#),
            "wrong style"
        );
        assert!(table_xml.contains(r#"showRowStripes="1""#), "wrong stripes");

        // Check content types
        let ct = read_zip_entry(bytes.clone(), "[Content_Types].xml");
        assert!(
            ct.contains("/xl/tables/table1.xml"),
            "missing table content type"
        );
        assert!(
            ct.contains("spreadsheetml.table+xml"),
            "wrong table content type"
        );

        // Check sheet rels
        let rels = read_zip_entry(bytes, "xl/worksheets/_rels/sheet1.xml.rels");
        assert!(
            rels.contains("../tables/table1.xml"),
            "missing table rel target"
        );
        assert!(rels.contains("/table\""), "missing table rel type");
    }

    #[test]
    fn test_writer_emits_table_with_totals_row() {
        use duke_sheets_core::table::{Table, TableColumn, TotalsRowFunction};

        let mut wb = Workbook::new();
        let sheet = wb.worksheet_mut(0).unwrap();
        sheet.set_cell_value("A1", "Item").unwrap();
        sheet.set_cell_value("B1", "Qty").unwrap();

        let mut table = Table::new(1, "Items", CellRange::parse("A1:B4").unwrap());
        let mut col1 = TableColumn::new(1, "Item");
        col1.totals_row_label = Some("Total".into());
        let mut col2 = TableColumn::new(2, "Qty");
        col2.totals_row_function = Some(TotalsRowFunction::Sum);
        table.columns = vec![col1, col2];
        table.totals_row_count = 1;
        sheet.add_table(table);

        let mut out = Cursor::new(Vec::new());
        XlsxWriter::write(&wb, &mut out).expect("write");
        let bytes = out.into_inner();

        let table_xml = read_zip_entry(bytes, "xl/tables/table1.xml");
        assert!(
            table_xml.contains(r#"totalsRowCount="1""#),
            "missing totalsRowCount"
        );
        assert!(
            table_xml.contains(r#"totalsRowLabel="Total""#),
            "missing totalsRowLabel"
        );
        assert!(
            table_xml.contains(r#"totalsRowFunction="sum""#),
            "missing totalsRowFunction"
        );
        // autoFilter ref should exclude the totals row: A1:B3 not A1:B4
        assert!(
            table_xml.contains(r#"<autoFilter ref="A1:B3""#),
            "autoFilter should exclude totals"
        );
    }

    #[test]
    fn test_roundtrip_preserves_custom_theme() {
        // Build a minimal XLSX with a custom theme in memory.
        let custom_theme = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<a:theme xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" name="CustomTest">
  <a:themeElements>
    <a:clrScheme name="CustomColors">
      <a:dk1><a:srgbClr val="111111"/></a:dk1>
      <a:lt1><a:srgbClr val="FEFEFE"/></a:lt1>
      <a:dk2><a:srgbClr val="222222"/></a:dk2>
      <a:lt2><a:srgbClr val="EEEEEE"/></a:lt2>
      <a:accent1><a:srgbClr val="AA0000"/></a:accent1>
      <a:accent2><a:srgbClr val="00BB00"/></a:accent2>
      <a:accent3><a:srgbClr val="0000CC"/></a:accent3>
      <a:accent4><a:srgbClr val="DD00DD"/></a:accent4>
      <a:accent5><a:srgbClr val="00EEDD"/></a:accent5>
      <a:accent6><a:srgbClr val="FFAA00"/></a:accent6>
      <a:hlink><a:srgbClr val="0000FF"/></a:hlink>
      <a:folHlink><a:srgbClr val="FF00FF"/></a:folHlink>
    </a:clrScheme>
    <a:fontScheme name="Test"><a:majorFont><a:latin typeface="Calibri"/></a:majorFont><a:minorFont><a:latin typeface="Calibri"/></a:minorFont></a:fontScheme>
    <a:fmtScheme name="Test"><a:fillStyleLst><a:solidFill><a:schemeClr val="phClr"/></a:solidFill></a:fillStyleLst><a:lnStyleLst><a:ln><a:solidFill><a:schemeClr val="phClr"/></a:solidFill></a:ln></a:lnStyleLst><a:effectStyleLst><a:effectStyle><a:effectLst/></a:effectStyle></a:effectStyleLst><a:bgFillStyleLst><a:solidFill><a:schemeClr val="phClr"/></a:solidFill></a:bgFillStyleLst></a:fmtScheme>
  </a:themeElements>
</a:theme>"#;

        // Create a minimal valid XLSX with the custom theme.
        let mut xlsx_buf = Vec::new();
        {
            let cursor = Cursor::new(&mut xlsx_buf);
            let mut zip = zip::ZipWriter::new(cursor);
            let opts = zip::write::SimpleFileOptions::default();

            // [Content_Types].xml
            zip.start_file("[Content_Types].xml", opts).unwrap();
            zip.write_all(br#"<?xml version="1.0"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/><Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/><Override PartName="/xl/theme/theme1.xml" ContentType="application/vnd.openxmlformats-officedocument.theme+xml"/></Types>"#).unwrap();

            // _rels/.rels
            zip.start_file("_rels/.rels", opts).unwrap();
            zip.write_all(br#"<?xml version="1.0"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/></Relationships>"#).unwrap();

            // xl/workbook.xml
            zip.start_file("xl/workbook.xml", opts).unwrap();
            zip.write_all(br#"<?xml version="1.0"?><workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><sheets><sheet name="Sheet1" sheetId="1" r:id="rId1"/></sheets></workbook>"#).unwrap();

            // xl/_rels/workbook.xml.rels
            zip.start_file("xl/_rels/workbook.xml.rels", opts).unwrap();
            zip.write_all(br#"<?xml version="1.0"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/><Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/theme" Target="theme/theme1.xml"/></Relationships>"#).unwrap();

            // xl/worksheets/sheet1.xml
            zip.start_file("xl/worksheets/sheet1.xml", opts).unwrap();
            zip.write_all(br#"<?xml version="1.0"?><worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><sheetData/></worksheet>"#).unwrap();

            // xl/theme/theme1.xml — the custom theme
            zip.start_file("xl/theme/theme1.xml", opts).unwrap();
            zip.write_all(custom_theme.as_bytes()).unwrap();

            zip.finish().unwrap();
        }

        // Read the handcrafted XLSX.
        let wb = XlsxReader::read(Cursor::new(&xlsx_buf)).unwrap();

        // Write it back out.
        let mut out = Cursor::new(Vec::new());
        XlsxWriter::write(&wb, &mut out).unwrap();
        let out_bytes = out.into_inner();

        // The re-written theme should be our custom theme, not the default.
        let theme_out = read_zip_entry(out_bytes, "xl/theme/theme1.xml");
        assert!(
            theme_out.contains("CustomTest"),
            "theme name 'CustomTest' should survive roundtrip"
        );
        assert!(
            theme_out.contains("AA0000"),
            "custom accent1 color should survive roundtrip"
        );
        assert!(
            !theme_out.contains("Office"),
            "default Office theme should NOT appear"
        );
    }
}
