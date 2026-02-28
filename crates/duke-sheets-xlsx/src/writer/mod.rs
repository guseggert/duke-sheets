//! XLSX writer — generates OOXML SpreadsheetML using quick-xml Writer API.

use std::collections::HashMap;
use std::fs::File;
use std::io::{Cursor, Seek, Write};
use std::path::Path;

use quick_xml::escape::escape;
use quick_xml::events::{BytesDecl, BytesEnd, BytesStart, BytesText, Event};
use quick_xml::Writer;

use crate::error::{XlsxError, XlsxResult};
use crate::styles::XlsxStyleTable;
use duke_sheets_core::style::Color;
use duke_sheets_core::{CellAddress, Workbook};

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
type XmlWriter = Writer<Cursor<Vec<u8>>>;

// ---------------------------------------------------------------------------
// Shared string table
// ---------------------------------------------------------------------------

/// Shared string table — maps string content to SST index.
struct SharedStringTable {
    strings: Vec<String>,
    index: HashMap<String, u32>,
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

        // Write [Content_Types].xml
        Self::write_content_types(&mut zip, workbook, &sheets_with_comments, &sst)?;

        // Write _rels/.rels
        Self::write_root_rels(&mut zip)?;

        // Write xl/workbook.xml
        Self::write_workbook_xml(&mut zip, workbook)?;

        // Write xl/_rels/workbook.xml.rels
        Self::write_workbook_rels(&mut zip, workbook, &sst)?;

        // Write xl/styles.xml
        Self::write_styles_xml(&mut zip, &style_table)?;

        // Write xl/theme/theme1.xml
        Self::write_theme_xml(&mut zip)?;

        // Write shared string table
        if !sst.is_empty() {
            Self::write_shared_strings(&mut zip, &sst)?;
        }

        // Write worksheets and their relationships
        for (i, sheet) in workbook.worksheets().enumerate() {
            Self::write_worksheet(&mut zip, workbook, i, &style_table, &sst)?;

            // Write worksheet relationships if sheet has comments
            if sheet.comment_count() > 0 {
                Self::write_worksheet_rels(&mut zip, i)?;
                Self::write_comments(&mut zip, workbook, i)?;
            }
        }

        zip.finish()?;
        Ok(())
    }

    // -----------------------------------------------------------------------
    // Helper: write an XML part to the zip archive
    // -----------------------------------------------------------------------

    /// Write an XML part to the zip archive.  Creates a `Writer` backed by an
    /// in-memory buffer, writes the XML declaration, calls `build` to produce
    /// the element tree, and flushes the result into the zip entry at `path`.
    fn write_xml_part<W: Write + Seek>(
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

    // -----------------------------------------------------------------------
    // [Content_Types].xml
    // -----------------------------------------------------------------------

    fn write_content_types<W: Write + Seek>(
        zip: &mut zip::ZipWriter<W>,
        workbook: &Workbook,
        sheets_with_comments: &[usize],
        sst: &SharedStringTable,
    ) -> XlsxResult<()> {
        Self::write_xml_part(zip, "[Content_Types].xml", |w| {
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

            w.write_event(Event::End(BytesEnd::new("Types")))?;
            Ok(())
        })
    }

    // -----------------------------------------------------------------------
    // _rels/.rels
    // -----------------------------------------------------------------------

    fn write_root_rels<W: Write + Seek>(zip: &mut zip::ZipWriter<W>) -> XlsxResult<()> {
        Self::write_xml_part(zip, "_rels/.rels", |w| {
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
        Self::write_xml_part(zip, "xl/workbook.xml", |w| {
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
                let name = escape(sheet.name());
                let mut el = w
                    .create_element("sheet")
                    .with_attribute(("name", &*name))
                    .with_attribute(("sheetId", sheet_id.as_str()));
                if !sheet.is_visible() {
                    el = el.with_attribute(("state", "hidden"));
                }
                el.with_attribute(("r:id", rid.as_str())).write_empty()?;
            }
            w.write_event(Event::End(BytesEnd::new("sheets")))?;

            // definedNames
            let named = workbook.named_ranges();
            if named.len() > 0 {
                w.write_event(Event::Start(BytesStart::new("definedNames")))?;
                for nr in named.iter() {
                    let name_esc = escape(&nr.name);
                    let mut el = w
                        .create_element("definedName")
                        .with_attribute(("name", &*name_esc));
                    let scope_str;
                    if let duke_sheets_core::named_range::NameScope::Sheet(idx) = nr.scope {
                        scope_str = idx.to_string();
                        el = el.with_attribute(("localSheetId", scope_str.as_str()));
                    }
                    if nr.hidden {
                        el = el.with_attribute(("hidden", "1"));
                    }
                    if let Some(ref comment) = nr.comment {
                        let c = escape(comment.as_str());
                        el = el.with_attribute(("comment", &*c));
                    }
                    el.write_text_content(BytesText::new(&nr.refers_to))?;
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
    ) -> XlsxResult<()> {
        Self::write_xml_part(zip, "xl/_rels/workbook.xml.rels", |w| {
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

            w.write_event(Event::End(BytesEnd::new("Relationships")))?;
            Ok(())
        })
    }

    // -----------------------------------------------------------------------
    // xl/sharedStrings.xml
    // -----------------------------------------------------------------------

    fn write_shared_strings<W: Write + Seek>(
        zip: &mut zip::ZipWriter<W>,
        sst: &SharedStringTable,
    ) -> XlsxResult<()> {
        Self::write_xml_part(zip, "xl/sharedStrings.xml", |w| {
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

    // -----------------------------------------------------------------------
    // xl/styles.xml
    // -----------------------------------------------------------------------

    fn write_styles_xml<W: Write + Seek>(
        zip: &mut zip::ZipWriter<W>,
        style_table: &XlsxStyleTable,
    ) -> XlsxResult<()> {
        Self::write_xml_part(zip, "xl/styles.xml", |w| style_table.write_styles_xml(w))
    }

    fn write_theme_xml<W: Write + Seek>(zip: &mut zip::ZipWriter<W>) -> XlsxResult<()> {
        let options = zip::write::SimpleFileOptions::default();
        zip.start_file("xl/theme/theme1.xml", options)?;
        zip.write_all(DEFAULT_THEME_XML.as_bytes())?;
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
    ) -> XlsxResult<()> {
        let path = format!("xl/worksheets/sheet{}.xml", index + 1);
        Self::write_xml_part(zip, &path, |w| {
            let sheet = workbook
                .worksheet(index)
                .ok_or_else(|| XlsxError::InvalidFormat("Sheet not found".into()))?;

            let mut tag = BytesStart::new("worksheet");
            tag.push_attribute(("xmlns", NS_SPREADSHEET));
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

            // mergeCells
            Self::write_merge_cells(w, sheet)?;

            // conditionalFormatting
            Self::write_conditional_formatting(w, sheet, index, style_table)?;

            // dataValidations
            Self::write_data_validations(w, sheet)?;

            // pageMargins + pageSetup
            Self::write_page_setup(w, sheet)?;

            w.write_event(Event::End(BytesEnd::new("worksheet")))?;
            Ok(())
        })
    }

    // -- worksheet sub-sections ---------------------------------------------

    fn write_sheet_pr(w: &mut XmlWriter, sheet: &duke_sheets_core::Worksheet) -> XlsxResult<()> {
        if let Some(color) = sheet.tab_color() {
            w.write_event(Event::Start(BytesStart::new("sheetPr")))?;
            Self::write_color_element(w, "tabColor", &color)?;
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

    fn write_color_element(w: &mut XmlWriter, tag: &str, color: &Color) -> XlsxResult<()> {
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

    fn write_sheet_format_pr(w: &mut XmlWriter, _sheet: &duke_sheets_core::Worksheet) -> XlsxResult<()> {
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
        let selection_active = sheet.selection_active_cell();
        let selection_range = sheet.selection_range();

        if freeze.is_none()
            && split.is_none()
            && !selected
            && zoom.is_none()
            && selection_active.is_none()
            && selection_range.is_none()
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

        let active_cell = selection_active.map(|(r, c)| CellAddress::new(r, c).to_a1_string());
        let sqref = selection_range.map(|r| r.to_string());

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

            let default_top_left = CellAddress::new(fp.row, fp.col).to_a1_string();
            let active = active_cell.as_deref().unwrap_or(default_top_left.as_str());
            let sqref_v = sqref.as_deref().unwrap_or(active);

            w.create_element("selection")
                .with_attribute(("pane", active_pane))
                .with_attribute(("activeCell", active))
                .with_attribute(("sqref", sqref_v))
                .write_empty()?;
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

            let default_active = sp
                .top_left
                .map(|(r, c)| CellAddress::new(r, c).to_a1_string())
                .unwrap_or_else(|| "A1".to_string());
            let active = active_cell.as_deref().unwrap_or(default_active.as_str());
            let sqref_v = sqref.as_deref().unwrap_or(active);

            let mut sel = BytesStart::new("selection");
            if let Some(active_pane) = sp.active_pane.as_deref() {
                sel.push_attribute(("pane", active_pane));
            }
            sel.push_attribute(("activeCell", active));
            sel.push_attribute(("sqref", sqref_v));
            w.write_event(Event::Empty(sel))?;
        } else if active_cell.is_some() || sqref.is_some() {
            let active = active_cell.as_deref().unwrap_or("A1");
            let sqref_v = sqref.as_deref().unwrap_or(active);
            w.create_element("selection")
                .with_attribute(("activeCell", active))
                .with_attribute(("sqref", sqref_v))
                .write_empty()?;
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

        // Metadata-only rows (custom height / hidden, no cells).
        let custom_heights = sheet.custom_row_heights();
        let hidden_rows_map = sheet.hidden_rows();
        let row_outline = sheet.row_outline_levels();
        let row_collapsed = sheet.collapsed_rows();
        let mut meta_only_rows: std::collections::BTreeSet<u32> = Default::default();
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

        let mut meta_iter = meta_only_rows.iter().copied().peekable();
        let mut current_row: Option<u32> = None;
        let mut written_rows: std::collections::HashSet<u32> = Default::default();

        for (row, col, cell) in sheet.iter_cells() {
            if current_row != Some(row) {
                // Close previous row
                if current_row.is_some() {
                    w.write_event(Event::End(BytesEnd::new("row")))?;
                }

                // Emit metadata-only rows that come before this data row
                while let Some(&mr) = meta_iter.peek() {
                    if mr >= row {
                        break;
                    }
                    if !written_rows.contains(&mr) {
                        Self::write_meta_row(
                            w,
                            mr,
                            custom_heights.get(&mr).copied(),
                            hidden_rows_map.get(&mr).copied().unwrap_or(false),
                            row_outline.get(&mr).copied().unwrap_or(0),
                            row_collapsed.get(&mr).copied().unwrap_or(false),
                        )?;
                        written_rows.insert(mr);
                    }
                    meta_iter.next();
                }

                // Open new row
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
                current_row = Some(row);
                written_rows.insert(row);
            }

            // Write cell
            Self::write_cell(w, row, col, cell, sheet_index, style_table, sst)?;
        }

        if current_row.is_some() {
            w.write_event(Event::End(BytesEnd::new("row")))?;
        }

        // Emit remaining metadata-only rows after all data rows
        for mr in meta_iter {
            if !written_rows.contains(&mr) {
                Self::write_meta_row(
                    w,
                    mr,
                    custom_heights.get(&mr).copied(),
                    hidden_rows_map.get(&mr).copied().unwrap_or(false),
                    row_outline.get(&mr).copied().unwrap_or(0),
                    row_collapsed.get(&mr).copied().unwrap_or(false),
                )?;
            }
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
    ) -> XlsxResult<()> {
        let addr = CellAddress::new(row, col);
        let cell_ref = addr.to_a1_string();
        let xf_id = style_table.xf_id_for(sheet_index, cell.style_index);
        let xf_str = xf_id.to_string();

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
            duke_sheets_core::CellValue::Formula {
                text, cached_value, ..
            } => {
                let formula_text = if text.starts_with('=') {
                    &text[1..]
                } else {
                    text.as_str()
                };
                let mut c = BytesStart::new("c");
                c.push_attribute(("r", cell_ref.as_str()));
                if xf_id != 0 {
                    c.push_attribute(("s", xf_str.as_str()));
                }
                // Determine type attribute from cached value
                match cached_value.as_deref() {
                    Some(duke_sheets_core::CellValue::String(_)) => {
                        c.push_attribute(("t", "str"));
                    }
                    Some(duke_sheets_core::CellValue::Boolean(_)) => {
                        c.push_attribute(("t", "b"));
                    }
                    Some(duke_sheets_core::CellValue::Error(_)) => {
                        c.push_attribute(("t", "e"));
                    }
                    _ => {}
                }
                w.write_event(Event::Start(c))?;
                w.create_element("f")
                    .write_text_content(BytesText::new(formula_text))?;
                // Write cached value
                match cached_value.as_deref() {
                    Some(duke_sheets_core::CellValue::Number(n)) => {
                        let v = n.to_string();
                        w.create_element("v")
                            .write_text_content(BytesText::new(&v))?;
                    }
                    Some(duke_sheets_core::CellValue::String(s)) => {
                        w.create_element("v")
                            .write_text_content(BytesText::new(s.as_str()))?;
                    }
                    Some(duke_sheets_core::CellValue::Boolean(b)) => {
                        w.create_element("v")
                            .write_text_content(BytesText::new(if *b { "1" } else { "0" }))?;
                    }
                    Some(duke_sheets_core::CellValue::Error(e)) => {
                        w.create_element("v")
                            .write_text_content(BytesText::new(e.as_str()))?;
                    }
                    _ => {}
                }
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
                // SpillTarget cells are not written — computed at runtime.
            }
        }
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

        let header_footer_differs = ps.odd_header.is_some() || ps.odd_footer.is_some();

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
            w.write_event(Event::Start(BytesStart::new("headerFooter")))?;
            if let Some(header) = &ps.odd_header {
                w.create_element("oddHeader")
                    .write_text_content(quick_xml::events::BytesText::new(header))?;
            }
            if let Some(footer) = &ps.odd_footer {
                w.create_element("oddFooter")
                    .write_text_content(quick_xml::events::BytesText::new(footer))?;
            }
            w.write_event(Event::End(BytesEnd::new("headerFooter")))?;
        }

        Ok(())
    }

    // -----------------------------------------------------------------------
    // Conditional formatting
    // -----------------------------------------------------------------------

    fn write_conditional_formatting(
        w: &mut XmlWriter,
        sheet: &duke_sheets_core::Worksheet,
        sheet_index: usize,
        style_table: &XlsxStyleTable,
    ) -> XlsxResult<()> {
        use duke_sheets_core::conditional_format::CfRuleType;

        let rules = sheet.conditional_formats();
        if rules.is_empty() {
            return Ok(());
        }

        for (rule_idx, rule) in rules.iter().enumerate() {
            if rule.ranges.is_empty() {
                continue;
            }

            let sqref: String = rule
                .ranges
                .iter()
                .map(|r| r.to_string())
                .collect::<Vec<_>>()
                .join(" ");

            let mut cf_tag = BytesStart::new("conditionalFormatting");
            cf_tag.push_attribute(("sqref", sqref.as_str()));
            w.write_event(Event::Start(cf_tag))?;

            // Build cfRule attributes
            let rule_type = rule.rule_type.xlsx_type();
            let dxf_id = style_table
                .dxf_id_for(sheet_index, rule_idx)
                .or(rule.dxf_id);
            let priority_val = rule.priority.max(1);
            let priority_s = priority_val.to_string();

            match &rule.rule_type {
                CfRuleType::CellIs {
                    operator,
                    formula1,
                    formula2,
                } => {
                    let mut tag = BytesStart::new("cfRule");
                    tag.push_attribute(("type", rule_type));
                    tag.push_attribute(("operator", operator.xlsx_operator()));
                    tag.push_attribute(("priority", priority_s.as_str()));
                    Self::push_dxf_and_stop(&mut tag, dxf_id, rule.stop_if_true);
                    w.write_event(Event::Start(tag))?;

                    w.create_element("formula")
                        .write_text_content(BytesText::new(formula1))?;
                    if let Some(f2) = formula2 {
                        w.create_element("formula")
                            .write_text_content(BytesText::new(f2))?;
                    }
                    w.write_event(Event::End(BytesEnd::new("cfRule")))?;
                }

                CfRuleType::Expression { formula } => {
                    let mut tag = BytesStart::new("cfRule");
                    tag.push_attribute(("type", rule_type));
                    tag.push_attribute(("priority", priority_s.as_str()));
                    Self::push_dxf_and_stop(&mut tag, dxf_id, rule.stop_if_true);
                    w.write_event(Event::Start(tag))?;

                    w.create_element("formula")
                        .write_text_content(BytesText::new(formula))?;
                    w.write_event(Event::End(BytesEnd::new("cfRule")))?;
                }

                CfRuleType::ColorScale { colors } => {
                    let mut tag = BytesStart::new("cfRule");
                    tag.push_attribute(("type", rule_type));
                    tag.push_attribute(("priority", priority_s.as_str()));
                    if rule.stop_if_true {
                        tag.push_attribute(("stopIfTrue", "1"));
                    }
                    w.write_event(Event::Start(tag))?;

                    w.write_event(Event::Start(BytesStart::new("colorScale")))?;
                    for cv in colors {
                        let mut cfvo = BytesStart::new("cfvo");
                        cfvo.push_attribute(("type", cv.value_type.xlsx_type()));
                        if let Some(ref v) = cv.value {
                            cfvo.push_attribute(("val", v.as_str()));
                        }
                        w.write_event(Event::Empty(cfvo))?;
                    }
                    for cv in colors {
                        Self::write_color_element(w, "color", &cv.color)?;
                    }
                    w.write_event(Event::End(BytesEnd::new("colorScale")))?;
                    w.write_event(Event::End(BytesEnd::new("cfRule")))?;
                }

                CfRuleType::DataBar {
                    min_value,
                    max_value,
                    color,
                    show_value,
                    ..
                } => {
                    let mut tag = BytesStart::new("cfRule");
                    tag.push_attribute(("type", rule_type));
                    tag.push_attribute(("priority", priority_s.as_str()));
                    if rule.stop_if_true {
                        tag.push_attribute(("stopIfTrue", "1"));
                    }
                    w.write_event(Event::Start(tag))?;

                    let mut db = BytesStart::new("dataBar");
                    if !*show_value {
                        db.push_attribute(("showValue", "0"));
                    }
                    w.write_event(Event::Start(db))?;

                    // cfvo for min
                    let mut cfvo_min = BytesStart::new("cfvo");
                    cfvo_min.push_attribute(("type", min_value.value_type.xlsx_type()));
                    if let Some(ref v) = min_value.value {
                        cfvo_min.push_attribute(("val", v.as_str()));
                    }
                    w.write_event(Event::Empty(cfvo_min))?;

                    // cfvo for max
                    let mut cfvo_max = BytesStart::new("cfvo");
                    cfvo_max.push_attribute(("type", max_value.value_type.xlsx_type()));
                    if let Some(ref v) = max_value.value {
                        cfvo_max.push_attribute(("val", v.as_str()));
                    }
                    w.write_event(Event::Empty(cfvo_max))?;

                    Self::write_color_element(w, "color", color)?;

                    w.write_event(Event::End(BytesEnd::new("dataBar")))?;
                    w.write_event(Event::End(BytesEnd::new("cfRule")))?;
                }

                CfRuleType::IconSet {
                    icon_style,
                    values,
                    reverse,
                    show_value,
                } => {
                    let mut tag = BytesStart::new("cfRule");
                    tag.push_attribute(("type", rule_type));
                    tag.push_attribute(("priority", priority_s.as_str()));
                    if rule.stop_if_true {
                        tag.push_attribute(("stopIfTrue", "1"));
                    }
                    w.write_event(Event::Start(tag))?;

                    let mut is_tag = BytesStart::new("iconSet");
                    is_tag.push_attribute(("iconSet", icon_style.xlsx_name()));
                    if *reverse {
                        is_tag.push_attribute(("reverse", "1"));
                    }
                    if !*show_value {
                        is_tag.push_attribute(("showValue", "0"));
                    }
                    w.write_event(Event::Start(is_tag))?;

                    for val in values {
                        let mut cfvo = BytesStart::new("cfvo");
                        cfvo.push_attribute(("type", val.value_type.xlsx_type()));
                        if let Some(ref v) = val.value {
                            cfvo.push_attribute(("val", v.as_str()));
                        }
                        w.write_event(Event::Empty(cfvo))?;
                    }

                    w.write_event(Event::End(BytesEnd::new("iconSet")))?;
                    w.write_event(Event::End(BytesEnd::new("cfRule")))?;
                }

                CfRuleType::Top10 {
                    rank,
                    percent,
                    bottom,
                } => {
                    let mut tag = BytesStart::new("cfRule");
                    tag.push_attribute(("type", rule_type));
                    tag.push_attribute(("priority", priority_s.as_str()));
                    let rank_s = rank.to_string();
                    tag.push_attribute(("rank", rank_s.as_str()));
                    if *percent {
                        tag.push_attribute(("percent", "1"));
                    }
                    if *bottom {
                        tag.push_attribute(("bottom", "1"));
                    }
                    Self::push_dxf_and_stop(&mut tag, dxf_id, rule.stop_if_true);
                    w.write_event(Event::Empty(tag))?;
                }

                CfRuleType::AboveAverage {
                    above,
                    equal_average,
                    std_dev,
                } => {
                    let mut tag = BytesStart::new("cfRule");
                    tag.push_attribute(("type", rule_type));
                    tag.push_attribute(("priority", priority_s.as_str()));
                    if !*above {
                        tag.push_attribute(("aboveAverage", "0"));
                    }
                    if *equal_average {
                        tag.push_attribute(("equalAverage", "1"));
                    }
                    if let Some(s) = std_dev {
                        let v = s.to_string();
                        tag.push_attribute(("stdDev", v.as_str()));
                    }
                    Self::push_dxf_and_stop(&mut tag, dxf_id, rule.stop_if_true);
                    w.write_event(Event::Empty(tag))?;
                }

                CfRuleType::ContainsText { text } => {
                    let mut tag = BytesStart::new("cfRule");
                    tag.push_attribute(("type", rule_type));
                    tag.push_attribute(("priority", priority_s.as_str()));
                    let text_esc = escape(text.as_str());
                    tag.push_attribute(("text", &*text_esc));
                    Self::push_dxf_and_stop(&mut tag, dxf_id, rule.stop_if_true);
                    w.write_event(Event::Start(tag))?;

                    let first_cell = sqref.split(' ').next().unwrap_or("A1");
                    let formula = format!(
                        "NOT(ISERROR(SEARCH(\"{}\",{})))",
                        text.replace('"', "\"\""),
                        first_cell
                    );
                    w.create_element("formula")
                        .write_text_content(BytesText::new(&formula))?;
                    w.write_event(Event::End(BytesEnd::new("cfRule")))?;
                }

                CfRuleType::BeginsWith { text } => {
                    let mut tag = BytesStart::new("cfRule");
                    tag.push_attribute(("type", rule_type));
                    tag.push_attribute(("priority", priority_s.as_str()));
                    let text_esc = escape(text.as_str());
                    tag.push_attribute(("text", &*text_esc));
                    Self::push_dxf_and_stop(&mut tag, dxf_id, rule.stop_if_true);
                    w.write_event(Event::Start(tag))?;

                    let first_cell = sqref
                        .split(' ')
                        .next()
                        .unwrap_or("A1")
                        .split(':')
                        .next()
                        .unwrap_or("A1");
                    let formula = format!(
                        "LEFT({},{})=\"{}\"",
                        first_cell,
                        text.len(),
                        text.replace('"', "\"\"")
                    );
                    w.create_element("formula")
                        .write_text_content(BytesText::new(&formula))?;
                    w.write_event(Event::End(BytesEnd::new("cfRule")))?;
                }

                CfRuleType::EndsWith { text } => {
                    let mut tag = BytesStart::new("cfRule");
                    tag.push_attribute(("type", rule_type));
                    tag.push_attribute(("priority", priority_s.as_str()));
                    let text_esc = escape(text.as_str());
                    tag.push_attribute(("text", &*text_esc));
                    Self::push_dxf_and_stop(&mut tag, dxf_id, rule.stop_if_true);
                    w.write_event(Event::Start(tag))?;

                    let first_cell = sqref
                        .split(' ')
                        .next()
                        .unwrap_or("A1")
                        .split(':')
                        .next()
                        .unwrap_or("A1");
                    let formula = format!(
                        "RIGHT({},{})=\"{}\"",
                        first_cell,
                        text.len(),
                        text.replace('"', "\"\"")
                    );
                    w.create_element("formula")
                        .write_text_content(BytesText::new(&formula))?;
                    w.write_event(Event::End(BytesEnd::new("cfRule")))?;
                }

                CfRuleType::DuplicateValues
                | CfRuleType::UniqueValues
                | CfRuleType::ContainsBlanks
                | CfRuleType::NotContainsBlanks
                | CfRuleType::ContainsErrors
                | CfRuleType::NotContainsErrors => {
                    let mut tag = BytesStart::new("cfRule");
                    tag.push_attribute(("type", rule_type));
                    tag.push_attribute(("priority", priority_s.as_str()));
                    Self::push_dxf_and_stop(&mut tag, dxf_id, rule.stop_if_true);
                    w.write_event(Event::Empty(tag))?;
                }

                CfRuleType::TimePeriod { period } => {
                    let mut tag = BytesStart::new("cfRule");
                    tag.push_attribute(("type", rule_type));
                    tag.push_attribute(("priority", priority_s.as_str()));
                    tag.push_attribute(("timePeriod", period.xlsx_period()));
                    Self::push_dxf_and_stop(&mut tag, dxf_id, rule.stop_if_true);
                    w.write_event(Event::Empty(tag))?;
                }
            }

            w.write_event(Event::End(BytesEnd::new("conditionalFormatting")))?;
        }

        Ok(())
    }

    /// Push optional `dxfId` and `stopIfTrue` attributes onto a `BytesStart`.
    fn push_dxf_and_stop(tag: &mut BytesStart, dxf_id: Option<u32>, stop_if_true: bool) {
        if let Some(id) = dxf_id {
            let s = id.to_string();
            // push_attribute borrows the value, so we need an owned copy in
            // the tag buffer.  BytesStart copies into its internal Vec<u8>.
            tag.push_attribute(("dxfId", s.as_str()));
        }
        if stop_if_true {
            tag.push_attribute(("stopIfTrue", "1"));
        }
    }

    // -----------------------------------------------------------------------
    // Data validations
    // -----------------------------------------------------------------------

    fn write_data_validations(
        w: &mut XmlWriter,
        sheet: &duke_sheets_core::Worksheet,
    ) -> XlsxResult<()> {
        use duke_sheets_core::validation::ValidationType;

        let validations = sheet.data_validations();
        if validations.is_empty() {
            return Ok(());
        }

        let count = validations.len().to_string();
        let mut dv_tag = BytesStart::new("dataValidations");
        dv_tag.push_attribute(("count", count.as_str()));
        w.write_event(Event::Start(dv_tag))?;

        for validation in validations {
            if validation.ranges.is_empty() {
                continue;
            }

            let sqref: String = validation
                .ranges
                .iter()
                .map(|r| r.to_string())
                .collect::<Vec<_>>()
                .join(" ");

            let mut tag = BytesStart::new("dataValidation");

            match &validation.validation_type {
                ValidationType::None => {}
                _ => {
                    tag.push_attribute(("type", validation.validation_type.xlsx_type()));
                }
            }

            // Operator attribute
            match &validation.validation_type {
                ValidationType::Whole { operator, .. }
                | ValidationType::Decimal { operator, .. }
                | ValidationType::Date { operator, .. }
                | ValidationType::Time { operator, .. }
                | ValidationType::TextLength { operator, .. } => {
                    tag.push_attribute(("operator", operator.xlsx_operator()));
                }
                _ => {}
            }

            if validation.allow_blank {
                tag.push_attribute(("allowBlank", "1"));
            }
            if !validation.show_dropdown {
                tag.push_attribute(("showDropDown", "1"));
            }
            if validation.show_input_message {
                tag.push_attribute(("showInputMessage", "1"));
            }
            if validation.show_error_alert {
                tag.push_attribute(("showErrorMessage", "1"));
            }

            match validation.error_style {
                duke_sheets_core::ValidationErrorStyle::Stop => {}
                duke_sheets_core::ValidationErrorStyle::Warning => {
                    tag.push_attribute(("errorStyle", "warning"));
                }
                duke_sheets_core::ValidationErrorStyle::Information => {
                    tag.push_attribute(("errorStyle", "information"));
                }
            }

            if let Some(ref t) = validation.error_title {
                let v = escape(t.as_str());
                tag.push_attribute(("errorTitle", &*v));
            }
            if let Some(ref m) = validation.error_message {
                let v = escape(m.as_str());
                tag.push_attribute(("error", &*v));
            }
            if let Some(ref t) = validation.input_title {
                let v = escape(t.as_str());
                tag.push_attribute(("promptTitle", &*v));
            }
            if let Some(ref m) = validation.input_message {
                let v = escape(m.as_str());
                tag.push_attribute(("prompt", &*v));
            }

            tag.push_attribute(("sqref", sqref.as_str()));
            w.write_event(Event::Start(tag))?;

            // Write formulas based on validation type
            match &validation.validation_type {
                ValidationType::List { source } => {
                    let formula = if source.starts_with('=') {
                        source[1..].to_string()
                    } else if source.contains('!')
                        || source
                            .chars()
                            .all(|c| c.is_ascii_alphanumeric() || c == '$' || c == ':')
                    {
                        source.clone()
                    } else {
                        format!("\"{}\"", source)
                    };
                    w.create_element("formula1")
                        .write_text_content(BytesText::new(&formula))?;
                }
                ValidationType::Whole { value1, value2, .. }
                | ValidationType::Decimal { value1, value2, .. }
                | ValidationType::Date { value1, value2, .. }
                | ValidationType::Time { value1, value2, .. }
                | ValidationType::TextLength { value1, value2, .. } => {
                    w.create_element("formula1")
                        .write_text_content(BytesText::new(value1))?;
                    if let Some(v2) = value2 {
                        w.create_element("formula2")
                            .write_text_content(BytesText::new(v2))?;
                    }
                }
                ValidationType::Custom { formula } => {
                    let f = if formula.starts_with('=') {
                        &formula[1..]
                    } else {
                        formula
                    };
                    w.create_element("formula1")
                        .write_text_content(BytesText::new(f))?;
                }
                ValidationType::None => {}
            }

            w.write_event(Event::End(BytesEnd::new("dataValidation")))?;
        }

        w.write_event(Event::End(BytesEnd::new("dataValidations")))?;
        Ok(())
    }

    // -----------------------------------------------------------------------
    // Worksheet relationships
    // -----------------------------------------------------------------------

    fn write_worksheet_rels<W: Write + Seek>(
        zip: &mut zip::ZipWriter<W>,
        sheet_index: usize,
    ) -> XlsxResult<()> {
        let path = format!("xl/worksheets/_rels/sheet{}.xml.rels", sheet_index + 1);
        Self::write_xml_part(zip, &path, |w| {
            let mut tag = BytesStart::new("Relationships");
            tag.push_attribute(("xmlns", NS_RELATIONSHIPS));
            w.write_event(Event::Start(tag))?;

            let target = format!("../comments{}.xml", sheet_index + 1);
            w.create_element("Relationship")
                .with_attribute(("Id", "rId1"))
                .with_attribute(("Type", RT_COMMENTS))
                .with_attribute(("Target", target.as_str()))
                .write_empty()?;

            w.write_event(Event::End(BytesEnd::new("Relationships")))?;
            Ok(())
        })
    }

    // -----------------------------------------------------------------------
    // Comments
    // -----------------------------------------------------------------------

    fn write_comments<W: Write + Seek>(
        zip: &mut zip::ZipWriter<W>,
        workbook: &Workbook,
        sheet_index: usize,
    ) -> XlsxResult<()> {
        let sheet = workbook
            .worksheet(sheet_index)
            .ok_or_else(|| XlsxError::InvalidFormat("Sheet not found".into()))?;

        if sheet.comment_count() == 0 {
            return Ok(());
        }

        let path = format!("xl/comments{}.xml", sheet_index + 1);
        Self::write_xml_part(zip, &path, |w| {
            let mut tag = BytesStart::new("comments");
            tag.push_attribute(("xmlns", NS_SPREADSHEET));
            w.write_event(Event::Start(tag))?;

            // Authors
            w.write_event(Event::Start(BytesStart::new("authors")))?;
            let authors = sheet.comment_authors();
            for author in authors {
                w.create_element("author")
                    .write_text_content(BytesText::new(author))?;
            }
            if authors.is_empty() {
                // Add empty author for comments without author
                w.create_element("author")
                    .write_text_content(BytesText::new(""))?;
            }
            w.write_event(Event::End(BytesEnd::new("authors")))?;

            // Comment list
            w.write_event(Event::Start(BytesStart::new("commentList")))?;

            let mut comments: Vec<_> = sheet.comments().collect();
            comments.sort_by_key(|((row, col), _)| (*row, *col));

            let author_index: std::collections::HashMap<&str, usize> = authors
                .iter()
                .enumerate()
                .map(|(i, a)| (a.as_str(), i))
                .collect();

            for ((row, col), comment) in comments {
                let cell_ref = CellAddress::new(row, col).to_a1_string();
                let author_id = if comment.author.is_empty() {
                    0
                } else {
                    author_index
                        .get(comment.author.as_str())
                        .copied()
                        .unwrap_or(0)
                };
                let aid = author_id.to_string();

                let mut c_tag = BytesStart::new("comment");
                c_tag.push_attribute(("ref", cell_ref.as_str()));
                c_tag.push_attribute(("authorId", aid.as_str()));
                w.write_event(Event::Start(c_tag))?;

                w.write_event(Event::Start(BytesStart::new("text")))?;
                w.write_event(Event::Start(BytesStart::new("r")))?;
                w.create_element("t")
                    .write_text_content(BytesText::new(&comment.text))?;
                w.write_event(Event::End(BytesEnd::new("r")))?;
                w.write_event(Event::End(BytesEnd::new("text")))?;

                w.write_event(Event::End(BytesEnd::new("comment")))?;
            }

            w.write_event(Event::End(BytesEnd::new("commentList")))?;
            w.write_event(Event::End(BytesEnd::new("comments")))?;
            Ok(())
        })
    }
}

#[cfg(test)]
mod tests {
    use super::*;
    use duke_sheets_core::{CellRange, ConditionalFormatRule, SplitPanes};
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
}
