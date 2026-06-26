//! XLSX styles (styles.xml) read/write helpers

use std::collections::hash_map::DefaultHasher;
use std::collections::HashMap;
use std::hash::{Hash, Hasher};
use std::io::{BufReader, Cursor, Read};
use std::sync::{Mutex, OnceLock};

use quick_xml::escape::escape;
use quick_xml::events::{BytesEnd, BytesStart, Event};
use quick_xml::reader::Reader;
use quick_xml::Writer;

use crate::error::{XlsxError, XlsxResult};
use duke_sheets_core::style::{
    Alignment, BorderEdge, BorderLineStyle, BorderStyle, Color, FillStyle, FontStyle,
    FontVerticalAlign, GradientStop, GradientType, HorizontalAlignment, NumberFormat, PatternType,
    Protection, ReadingOrder, Style, Underline, VerticalAlignment,
};
use duke_sheets_core::Workbook;

/// Alias for the in-memory XML writer (matches writer/mod.rs).
pub(crate) type XmlWriter = Writer<Cursor<Vec<u8>>>;

const NS_SPREADSHEET: &str = "http://schemas.openxmlformats.org/spreadsheetml/2006/main";

#[derive(Debug)]
pub(crate) struct XlsxStyleTable {
    /// Global, deduplicated styles. Index corresponds to the cellXfs index (xfId).
    styles: Vec<Style>,
    /// Per-worksheet mapping: local worksheet style index -> global xfId.
    sheet_maps: Vec<HashMap<u32, u32>>,
    /// DXF (differential format) styles for conditional formatting.
    /// Key: (sheet_index, rule_index), Value: dxf_id
    dxf_styles: Vec<Style>,
    /// Mapping from (sheet_index, cf_rule_index) to dxf_id
    dxf_map: HashMap<(usize, usize), u32>,
    cell_style_xfs: Vec<Style>,
    named_styles: Vec<NamedCellStyle>,
    cell_xf_ids: Vec<u32>,
    custom_num_fmt_ids: HashMap<String, u32>,
}

#[derive(Debug, Clone)]
pub(crate) struct NamedCellStyle {
    pub name: String,
    pub xf_id: u32,
    pub builtin_id: Option<u32>,
}

#[derive(Debug, Clone)]
pub(crate) struct RoundtripStyleData {
    pub cell_styles: Vec<Style>,
    pub cell_style_xfs: Vec<Style>,
    pub named_styles: Vec<NamedCellStyle>,
    pub cell_xf_xf_ids: Vec<u32>,
}

static ROUNDTRIP_STYLE_DATA: OnceLock<Mutex<HashMap<u64, RoundtripStyleData>>> = OnceLock::new();

fn style_data_store() -> &'static Mutex<HashMap<u64, RoundtripStyleData>> {
    ROUNDTRIP_STYLE_DATA.get_or_init(|| Mutex::new(HashMap::new()))
}

fn workbook_style_fingerprint(workbook: &Workbook) -> u64 {
    let mut hasher = DefaultHasher::new();
    workbook.nonce().hash(&mut hasher);
    workbook.sheet_count().hash(&mut hasher);
    for (sheet_idx, sheet) in workbook.worksheets().enumerate() {
        sheet_idx.hash(&mut hasher);
        sheet.name().hash(&mut hasher);
        sheet.cell_count().hash(&mut hasher);
        for (row, col, cell) in sheet.iter_cells() {
            row.hash(&mut hasher);
            col.hash(&mut hasher);
            cell.style_index.hash(&mut hasher);
            if let Some(style) = sheet.style_by_index(cell.style_index) {
                style.hash(&mut hasher);
            }
        }
    }
    hasher.finish()
}

pub(crate) fn register_roundtrip_style_data(workbook: &Workbook, data: RoundtripStyleData) {
    let key = workbook_style_fingerprint(workbook);
    if let Ok(mut store) = style_data_store().lock() {
        store.insert(key, data);
    }
}

fn roundtrip_style_data_for(workbook: &Workbook) -> Option<RoundtripStyleData> {
    let key = workbook_style_fingerprint(workbook);
    style_data_store().lock().ok()?.get(&key).cloned()
}

static ROUNDTRIP_THEME_DATA: OnceLock<Mutex<HashMap<u64, Vec<u8>>>> = OnceLock::new();

fn theme_data_store() -> &'static Mutex<HashMap<u64, Vec<u8>>> {
    ROUNDTRIP_THEME_DATA.get_or_init(|| Mutex::new(HashMap::new()))
}

pub(crate) fn register_roundtrip_theme_data(workbook: &Workbook, theme_xml: Vec<u8>) {
    let key = workbook_style_fingerprint(workbook);
    if let Ok(mut store) = theme_data_store().lock() {
        store.insert(key, theme_xml);
    }
}

pub(crate) fn roundtrip_theme_data_for(workbook: &Workbook) -> Option<Vec<u8>> {
    let key = workbook_style_fingerprint(workbook);
    theme_data_store().lock().ok()?.get(&key).cloned()
}

#[derive(Debug, Clone, Copy)]
struct ResolvedXfIds {
    font_id: u32,
    fill_id: u32,
    border_id: u32,
    num_fmt_id: u32,
}

impl XlsxStyleTable {
    pub(crate) fn build(workbook: &Workbook) -> Self {
        let roundtrip_data = roundtrip_style_data_for(workbook);
        let mut cell_xf_id_by_style: HashMap<Style, u32> = HashMap::new();
        if let Some(data) = &roundtrip_data {
            for (style, xf_id) in data
                .cell_styles
                .iter()
                .cloned()
                .zip(data.cell_xf_xf_ids.iter().copied())
            {
                cell_xf_id_by_style.entry(style).or_insert(xf_id);
            }
        }

        let mut styles: Vec<Style> = Vec::new();
        let mut style_to_xf: HashMap<Style, u32> = HashMap::new();
        let mut cell_xf_ids: Vec<u32> = Vec::new();

        // Index 0 is always default style
        let default = Style::default();
        styles.push(default.clone());
        style_to_xf.insert(default, 0);
        cell_xf_ids.push(0);

        let mut sheet_maps: Vec<HashMap<u32, u32>> = Vec::with_capacity(workbook.sheet_count());

        // DXF styles for conditional formatting
        let mut dxf_styles: Vec<Style> = Vec::new();
        let mut dxf_map: HashMap<(usize, usize), u32> = HashMap::new();
        let mut dxf_style_to_id: HashMap<Style, u32> = HashMap::new();

        for (sheet_idx, sheet) in workbook.worksheets().enumerate() {
            let mut map: HashMap<u32, u32> = HashMap::new();
            map.insert(0, 0);

            for (_row, _col, cell) in sheet.iter_cells() {
                let local_idx = cell.style_index;
                if local_idx == 0 || map.contains_key(&local_idx) {
                    continue;
                }

                let style = sheet
                    .style_by_index(local_idx)
                    .cloned()
                    .unwrap_or_else(Style::default);

                let xf_id = match style_to_xf.get(&style) {
                    Some(&id) => id,
                    None => {
                        let id = styles.len() as u32;
                        styles.push(style.clone());
                        cell_xf_ids.push(cell_xf_id_by_style.get(&style).copied().unwrap_or(0));
                        style_to_xf.insert(style, id);
                        id
                    }
                };

                map.insert(local_idx, xf_id);
            }

            sheet_maps.push(map);

            // Collect DXF styles from conditional formatting rules
            for (rule_idx, rule) in sheet.conditional_formats().iter().enumerate() {
                if let Some(ref format) = rule.format {
                    // Check if we already have this DXF style
                    let dxf_id = match dxf_style_to_id.get(format) {
                        Some(&id) => id,
                        None => {
                            let id = dxf_styles.len() as u32;
                            dxf_styles.push(format.clone());
                            dxf_style_to_id.insert(format.clone(), id);
                            id
                        }
                    };
                    dxf_map.insert((sheet_idx, rule_idx), dxf_id);
                }
            }
        }

        let mut cell_style_xfs = roundtrip_data
            .as_ref()
            .map(|d| d.cell_style_xfs.clone())
            .unwrap_or_else(|| vec![Style::default()]);
        if cell_style_xfs.is_empty() {
            cell_style_xfs.push(Style::default());
        }

        let mut named_styles = roundtrip_data
            .as_ref()
            .map(|d| d.named_styles.clone())
            .unwrap_or_default();
        if named_styles.is_empty() {
            named_styles.push(NamedCellStyle {
                name: "Normal".to_string(),
                xf_id: 0,
                builtin_id: Some(0),
            });
        }

        let custom_num_fmt_ids =
            custom_num_fmt_ids_for_workbook(workbook, &styles, &cell_style_xfs);

        Self {
            styles,
            sheet_maps,
            dxf_styles,
            dxf_map,
            cell_style_xfs,
            named_styles,
            cell_xf_ids,
            custom_num_fmt_ids,
        }
    }

    pub(crate) fn xf_id_for(&self, sheet_index: usize, local_style_index: u32) -> u32 {
        self.sheet_maps
            .get(sheet_index)
            .and_then(|m| m.get(&local_style_index).copied())
            .unwrap_or(0)
    }

    /// Get the DXF ID for a conditional format rule, if it has a format defined
    pub(crate) fn dxf_id_for(&self, sheet_index: usize, rule_index: usize) -> Option<u32> {
        self.dxf_map.get(&(sheet_index, rule_index)).copied()
    }

    pub(crate) fn custom_num_fmt_id(&self, code: &str) -> Option<u32> {
        self.custom_num_fmt_ids.get(code).copied()
    }

    pub(crate) fn write_styles_xml(&self, w: &mut XmlWriter) -> XlsxResult<()> {
        // Build component tables
        let mut font_ids: HashMap<FontStyle, u32> = HashMap::new();
        let mut fonts: Vec<FontStyle> = Vec::new();

        let default_font = FontStyle::default();
        fonts.push(default_font.clone());
        font_ids.insert(default_font, 0);

        let mut fill_ids: HashMap<FillStyle, u32> = HashMap::new();
        let mut fills: Vec<FillStyle> = Vec::new();
        // Excel requires the first two fills to be: none and gray125
        fills.push(FillStyle::None); // id 0
        fills.push(FillStyle::Pattern {
            pattern: PatternType::Gray125,
            foreground: Color::Auto,
            background: Color::Auto,
        }); // id 1
        fill_ids.insert(FillStyle::None, 0);

        let mut border_ids: HashMap<BorderStyle, u32> = HashMap::new();
        let mut borders: Vec<BorderStyle> = Vec::new();
        let default_border = BorderStyle::default();
        borders.push(default_border.clone());
        border_ids.insert(default_border, 0);

        // Resolve component IDs for each style
        let mut resolved: Vec<ResolvedXfIds> = Vec::with_capacity(self.styles.len());
        let mut resolved_cell_style_xfs: Vec<ResolvedXfIds> =
            Vec::with_capacity(self.cell_style_xfs.len());

        let mut resolve_ids_for_style = |style: &Style| -> XlsxResult<ResolvedXfIds> {
            let font_id = match font_ids.get(&style.font) {
                Some(&id) => id,
                None => {
                    let id = fonts.len() as u32;
                    fonts.push(style.font.clone());
                    font_ids.insert(style.font.clone(), id);
                    id
                }
            };

            let fill_id = match &style.fill {
                FillStyle::None => 0,
                other => {
                    if let Some(&id) = fill_ids.get(other) {
                        id
                    } else {
                        let id = fills.len() as u32;
                        fills.push(other.clone());
                        fill_ids.insert(other.clone(), id);
                        id
                    }
                }
            };

            let border_id = match border_ids.get(&style.border) {
                Some(&id) => id,
                None => {
                    let id = borders.len() as u32;
                    borders.push(style.border.clone());
                    border_ids.insert(style.border.clone(), id);
                    id
                }
            };

            let num_fmt_id = match &style.number_format {
                NumberFormat::General => 0,
                NumberFormat::BuiltIn(id) => *id,
                NumberFormat::Custom(code) => self.custom_num_fmt_id(code).ok_or_else(|| {
                    XlsxError::InvalidFormat(format!("custom number format not registered: {code}"))
                })?,
            };

            Ok(ResolvedXfIds {
                font_id,
                fill_id,
                border_id,
                num_fmt_id,
            })
        };

        for style in &self.styles {
            resolved.push(resolve_ids_for_style(style)?);
        }
        for style in &self.cell_style_xfs {
            resolved_cell_style_xfs.push(resolve_ids_for_style(style)?);
        }

        // Write XML
        let mut root = BytesStart::new("styleSheet");
        root.push_attribute(("xmlns", NS_SPREADSHEET));
        w.write_event(Event::Start(root))?;

        // numFmts
        let numfmts = sorted_custom_num_fmts(&self.custom_num_fmt_ids);
        if !numfmts.is_empty() {
            let count = numfmts.len().to_string();
            let mut tag = BytesStart::new("numFmts");
            tag.push_attribute(("count", count.as_str()));
            w.write_event(Event::Start(tag))?;
            for (id, code) in &numfmts {
                let id_s = id.to_string();
                let code_esc = escape(code.as_str());
                w.create_element("numFmt")
                    .with_attribute(("numFmtId", id_s.as_str()))
                    .with_attribute(("formatCode", &*code_esc))
                    .write_empty()?;
            }
            w.write_event(Event::End(BytesEnd::new("numFmts")))?;
        }

        // fonts
        let count = fonts.len().to_string();
        let mut tag = BytesStart::new("fonts");
        tag.push_attribute(("count", count.as_str()));
        w.write_event(Event::Start(tag))?;
        for font in &fonts {
            write_font_xml(w, font)?;
        }
        w.write_event(Event::End(BytesEnd::new("fonts")))?;

        // fills
        let count = fills.len().to_string();
        let mut tag = BytesStart::new("fills");
        tag.push_attribute(("count", count.as_str()));
        w.write_event(Event::Start(tag))?;
        for fill in &fills {
            write_fill_xml(w, fill)?;
        }
        w.write_event(Event::End(BytesEnd::new("fills")))?;

        // borders
        let count = borders.len().to_string();
        let mut tag = BytesStart::new("borders");
        tag.push_attribute(("count", count.as_str()));
        w.write_event(Event::Start(tag))?;
        for border in &borders {
            write_border_xml(w, border)?;
        }
        w.write_event(Event::End(BytesEnd::new("borders")))?;

        // cellStyleXfs (required)
        let count = self.cell_style_xfs.len().to_string();
        let mut tag = BytesStart::new("cellStyleXfs");
        tag.push_attribute(("count", count.as_str()));
        w.write_event(Event::Start(tag))?;
        for (i, ids) in resolved_cell_style_xfs.iter().enumerate() {
            write_cell_style_xf_xml(w, &self.cell_style_xfs[i], *ids)?;
        }
        w.write_event(Event::End(BytesEnd::new("cellStyleXfs")))?;

        // cellXfs
        let count = self.styles.len().to_string();
        let mut tag = BytesStart::new("cellXfs");
        tag.push_attribute(("count", count.as_str()));
        w.write_event(Event::Start(tag))?;
        for (i, ids) in resolved.iter().enumerate() {
            let xf_id = self
                .cell_xf_ids
                .get(i)
                .copied()
                .filter(|xf_id| (*xf_id as usize) < self.cell_style_xfs.len())
                .unwrap_or(0);
            write_xf_xml(w, &self.styles[i], *ids, xf_id)?;
        }
        w.write_event(Event::End(BytesEnd::new("cellXfs")))?;

        // cellStyles (required)
        let count = self.named_styles.len().to_string();
        let mut tag = BytesStart::new("cellStyles");
        tag.push_attribute(("count", count.as_str()));
        w.write_event(Event::Start(tag))?;
        for named_style in &self.named_styles {
            let name_esc = escape(named_style.name.as_str());
            let xf_id = named_style.xf_id.to_string();
            let mut cell_style = BytesStart::new("cellStyle");
            cell_style.push_attribute(("name", &*name_esc));
            cell_style.push_attribute(("xfId", xf_id.as_str()));
            let builtin_id = named_style.builtin_id.map(|v| v.to_string());
            if let Some(builtin_id) = builtin_id.as_deref() {
                cell_style.push_attribute(("builtinId", builtin_id));
            }
            w.write_event(Event::Empty(cell_style))?;
        }
        w.write_event(Event::End(BytesEnd::new("cellStyles")))?;

        // DXFs
        if self.dxf_styles.is_empty() {
            w.create_element("dxfs")
                .with_attribute(("count", "0"))
                .write_empty()?;
        } else {
            let count = self.dxf_styles.len().to_string();
            let mut tag = BytesStart::new("dxfs");
            tag.push_attribute(("count", count.as_str()));
            w.write_event(Event::Start(tag))?;
            for dxf_style in &self.dxf_styles {
                write_dxf_xml(w, dxf_style)?;
            }
            w.write_event(Event::End(BytesEnd::new("dxfs")))?;
        }

        // tableStyles
        w.create_element("tableStyles")
            .with_attribute(("count", "0"))
            .with_attribute(("defaultTableStyle", "TableStyleMedium9"))
            .with_attribute(("defaultPivotStyle", "PivotStyleLight16"))
            .write_empty()?;

        w.write_event(Event::End(BytesEnd::new("styleSheet")))?;
        Ok(())
    }
}

fn custom_num_fmt_ids_for_workbook(
    workbook: &Workbook,
    styles: &[Style],
    cell_style_xfs: &[Style],
) -> HashMap<String, u32> {
    let mut ids = HashMap::new();
    let mut next_id = 164u32;

    for style in styles.iter().chain(cell_style_xfs.iter()) {
        register_custom_num_fmt(&style.number_format, &mut ids, &mut next_id);
    }

    for worksheet in workbook.worksheets() {
        for pivot in worksheet.pivot_tables() {
            for measure in &pivot.measures {
                if let Some(code) = &measure.number_format {
                    register_custom_num_fmt_code(code, &mut ids, &mut next_id);
                }
            }
        }
    }

    ids
}

fn register_custom_num_fmt(
    format: &NumberFormat,
    ids: &mut HashMap<String, u32>,
    next_id: &mut u32,
) {
    if let NumberFormat::Custom(code) = format {
        register_custom_num_fmt_code(code, ids, next_id);
    }
}

fn register_custom_num_fmt_code(code: &str, ids: &mut HashMap<String, u32>, next_id: &mut u32) {
    ids.entry(code.to_string()).or_insert_with(|| {
        let id = *next_id;
        *next_id += 1;
        id
    });
}

fn sorted_custom_num_fmts(ids: &HashMap<String, u32>) -> Vec<(u32, String)> {
    let mut numfmts = ids
        .iter()
        .map(|(code, id)| (*id, code.clone()))
        .collect::<Vec<_>>();
    numfmts.sort_by_key(|(id, _)| *id);
    numfmts
}

/// Write a color element (e.g. `<fgColor>`, `<color>`, `<bgColor>`) to the writer.
fn write_color_xml(w: &mut XmlWriter, tag: &str, color: &Color) -> std::io::Result<()> {
    let mut el = BytesStart::new(tag);
    match color {
        Color::Auto => el.push_attribute(("indexed", "64")),
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
    w.write_event(Event::Empty(el))
}

fn write_font_xml(w: &mut XmlWriter, font: &FontStyle) -> std::io::Result<()> {
    w.write_event(Event::Start(BytesStart::new("font")))?;

    if font.bold {
        w.write_event(Event::Empty(BytesStart::new("b")))?;
    }
    if font.italic {
        w.write_event(Event::Empty(BytesStart::new("i")))?;
    }
    if font.strikethrough {
        w.write_event(Event::Empty(BytesStart::new("strike")))?;
    }
    match font.underline {
        Underline::None => {}
        Underline::Single => w.write_event(Event::Empty(BytesStart::new("u")))?,
        Underline::Double => {
            w.create_element("u")
                .with_attribute(("val", "double"))
                .write_empty()?;
        }
        Underline::SingleAccounting => {
            w.create_element("u")
                .with_attribute(("val", "singleAccounting"))
                .write_empty()?;
        }
        Underline::DoubleAccounting => {
            w.create_element("u")
                .with_attribute(("val", "doubleAccounting"))
                .write_empty()?;
        }
    }
    match font.vertical_align {
        FontVerticalAlign::Baseline => {}
        FontVerticalAlign::Superscript => {
            w.create_element("vertAlign")
                .with_attribute(("val", "superscript"))
                .write_empty()?;
        }
        FontVerticalAlign::Subscript => {
            w.create_element("vertAlign")
                .with_attribute(("val", "subscript"))
                .write_empty()?;
        }
    }

    let size = font.size.to_string();
    w.create_element("sz")
        .with_attribute(("val", size.as_str()))
        .write_empty()?;

    if !matches!(font.color, Color::Auto) {
        write_color_xml(w, "color", &font.color)?;
    }

    let name_esc = escape(&font.name);
    w.create_element("name")
        .with_attribute(("val", &*name_esc))
        .write_empty()?;

    if let Some(family) = font.family {
        let family_s = family.to_string();
        w.create_element("family")
            .with_attribute(("val", family_s.as_str()))
            .write_empty()?;
    }

    if let Some(charset) = font.charset {
        let charset_s = charset.to_string();
        w.create_element("charset")
            .with_attribute(("val", charset_s.as_str()))
            .write_empty()?;
    }

    if let Some(scheme) = &font.scheme {
        w.create_element("scheme")
            .with_attribute(("val", scheme.as_str()))
            .write_empty()?;
    }

    w.write_event(Event::End(BytesEnd::new("font")))?;
    Ok(())
}

fn pattern_type_to_str(p: PatternType) -> &'static str {
    match p {
        PatternType::None => "none",
        PatternType::Solid => "solid",
        PatternType::MediumGray => "mediumGray",
        PatternType::DarkGray => "darkGray",
        PatternType::LightGray => "lightGray",
        PatternType::DarkHorizontal => "darkHorizontal",
        PatternType::DarkVertical => "darkVertical",
        PatternType::DarkDown => "darkDown",
        PatternType::DarkUp => "darkUp",
        PatternType::DarkGrid => "darkGrid",
        PatternType::DarkTrellis => "darkTrellis",
        PatternType::LightHorizontal => "lightHorizontal",
        PatternType::LightVertical => "lightVertical",
        PatternType::LightDown => "lightDown",
        PatternType::LightUp => "lightUp",
        PatternType::LightGrid => "lightGrid",
        PatternType::LightTrellis => "lightTrellis",
        PatternType::Gray125 => "gray125",
        PatternType::Gray0625 => "gray0625",
    }
}

fn write_fill_xml(w: &mut XmlWriter, fill: &FillStyle) -> std::io::Result<()> {
    w.write_event(Event::Start(BytesStart::new("fill")))?;
    match fill {
        FillStyle::None => {
            w.create_element("patternFill")
                .with_attribute(("patternType", "none"))
                .write_empty()?;
        }
        FillStyle::Solid { color } => {
            w.write_event(Event::Start(
                BytesStart::new("patternFill")
                    .with_attributes([("patternType", "solid")].into_iter()),
            ))?;
            write_color_xml(w, "fgColor", color)?;
            w.create_element("bgColor")
                .with_attribute(("indexed", "64"))
                .write_empty()?;
            w.write_event(Event::End(BytesEnd::new("patternFill")))?;
        }
        FillStyle::Pattern {
            pattern,
            foreground,
            background,
        } => {
            w.write_event(Event::Start(
                BytesStart::new("patternFill")
                    .with_attributes([("patternType", pattern_type_to_str(*pattern))].into_iter()),
            ))?;
            write_color_xml(w, "fgColor", foreground)?;
            write_color_xml(w, "bgColor", background)?;
            w.write_event(Event::End(BytesEnd::new("patternFill")))?;
        }
        FillStyle::Gradient {
            gradient_type,
            angle,
            stops,
        } => {
            let mut tag = BytesStart::new("gradientFill");
            let angle_s;
            match gradient_type {
                GradientType::Linear => {
                    if *angle != 0.0 {
                        angle_s = angle.to_string();
                        tag.push_attribute(("degree", angle_s.as_str()));
                    }
                }
                GradientType::Path => tag.push_attribute(("type", "path")),
            }
            if stops.is_empty() {
                w.write_event(Event::Empty(tag))?;
            } else {
                w.write_event(Event::Start(tag))?;
                for stop in stops {
                    let pos = stop.position.to_string();
                    w.write_event(Event::Start(
                        BytesStart::new("stop")
                            .with_attributes([("position", pos.as_str())].into_iter()),
                    ))?;
                    write_color_xml(w, "color", &stop.color)?;
                    w.write_event(Event::End(BytesEnd::new("stop")))?;
                }
                w.write_event(Event::End(BytesEnd::new("gradientFill")))?;
            }
        }
    }
    w.write_event(Event::End(BytesEnd::new("fill")))?;
    Ok(())
}

fn border_style_to_str(s: BorderLineStyle) -> Option<&'static str> {
    match s {
        BorderLineStyle::None => None,
        BorderLineStyle::Thin => Some("thin"),
        BorderLineStyle::Medium => Some("medium"),
        BorderLineStyle::Thick => Some("thick"),
        BorderLineStyle::Dashed => Some("dashed"),
        BorderLineStyle::Dotted => Some("dotted"),
        BorderLineStyle::Double => Some("double"),
        BorderLineStyle::Hair => Some("hair"),
        BorderLineStyle::MediumDashed => Some("mediumDashed"),
        BorderLineStyle::DashDot => Some("dashDot"),
        BorderLineStyle::MediumDashDot => Some("mediumDashDot"),
        BorderLineStyle::DashDotDot => Some("dashDotDot"),
        BorderLineStyle::MediumDashDotDot => Some("mediumDashDotDot"),
        BorderLineStyle::SlantDashDot => Some("slantDashDot"),
    }
}

fn write_border_edge_xml(
    w: &mut XmlWriter,
    tag: &str,
    edge: &Option<BorderEdge>,
) -> std::io::Result<()> {
    match edge {
        None => w.write_event(Event::Empty(BytesStart::new(tag)))?,
        Some(e) => {
            let style_attr = border_style_to_str(e.style);
            if style_attr.is_none() {
                w.write_event(Event::Empty(BytesStart::new(tag)))?;
                return Ok(());
            }
            let mut el = BytesStart::new(tag);
            el.push_attribute(("style", style_attr.unwrap()));
            w.write_event(Event::Start(el))?;
            write_color_xml(w, "color", &e.color)?;
            w.write_event(Event::End(BytesEnd::new(tag)))?;
        }
    }
    Ok(())
}

fn write_border_xml(w: &mut XmlWriter, border: &BorderStyle) -> std::io::Result<()> {
    use duke_sheets_core::style::DiagonalDirection;

    let mut tag = BytesStart::new("border");
    match border.diagonal_direction {
        DiagonalDirection::None => {}
        DiagonalDirection::Down => tag.push_attribute(("diagonalDown", "1")),
        DiagonalDirection::Up => tag.push_attribute(("diagonalUp", "1")),
        DiagonalDirection::Both => {
            tag.push_attribute(("diagonalDown", "1"));
            tag.push_attribute(("diagonalUp", "1"));
        }
    }
    w.write_event(Event::Start(tag))?;
    write_border_edge_xml(w, "left", &border.left)?;
    write_border_edge_xml(w, "right", &border.right)?;
    write_border_edge_xml(w, "top", &border.top)?;
    write_border_edge_xml(w, "bottom", &border.bottom)?;
    write_border_edge_xml(w, "diagonal", &border.diagonal)?;
    w.write_event(Event::End(BytesEnd::new("border")))?;
    Ok(())
}

fn horiz_to_str(h: HorizontalAlignment) -> &'static str {
    match h {
        HorizontalAlignment::General => "general",
        HorizontalAlignment::Left => "left",
        HorizontalAlignment::Center => "center",
        HorizontalAlignment::Right => "right",
        HorizontalAlignment::Fill => "fill",
        HorizontalAlignment::Justify => "justify",
        HorizontalAlignment::CenterContinuous => "centerContinuous",
        HorizontalAlignment::Distributed => "distributed",
    }
}

fn vert_to_str(v: VerticalAlignment) -> &'static str {
    match v {
        VerticalAlignment::Top => "top",
        VerticalAlignment::Center => "center",
        VerticalAlignment::Bottom => "bottom",
        VerticalAlignment::Justify => "justify",
        VerticalAlignment::Distributed => "distributed",
    }
}

/// Returns `true` if alignment was written (i.e. non-default).
fn write_alignment_xml(w: &mut XmlWriter, al: &Alignment) -> std::io::Result<bool> {
    let default = Alignment::default();
    if al == &default {
        return Ok(false);
    }

    let mut el = BytesStart::new("alignment");
    if al.horizontal != default.horizontal {
        el.push_attribute(("horizontal", horiz_to_str(al.horizontal)));
    }
    if al.vertical != default.vertical {
        el.push_attribute(("vertical", vert_to_str(al.vertical)));
    }
    if al.wrap_text {
        el.push_attribute(("wrapText", "1"));
    }
    if al.shrink_to_fit {
        el.push_attribute(("shrinkToFit", "1"));
    }
    if al.indent != 0 {
        let v = al.indent.to_string();
        el.push_attribute(("indent", v.as_str()));
    }
    if al.rotation != 0 {
        let v = al.rotation.to_string();
        el.push_attribute(("textRotation", v.as_str()));
    }
    match al.reading_order {
        ReadingOrder::ContextDependent => {}
        ReadingOrder::LeftToRight => el.push_attribute(("readingOrder", "1")),
        ReadingOrder::RightToLeft => el.push_attribute(("readingOrder", "2")),
    }
    w.write_event(Event::Empty(el))?;
    Ok(true)
}

/// Returns `true` if protection was written (i.e. non-default).
fn write_protection_xml(w: &mut XmlWriter, p: &Protection) -> std::io::Result<bool> {
    let default = Protection::default();
    if p == &default {
        return Ok(false);
    }
    let mut el = BytesStart::new("protection");
    if p.locked != default.locked {
        el.push_attribute(("locked", if p.locked { "1" } else { "0" }));
    }
    if p.hidden != default.hidden {
        el.push_attribute(("hidden", if p.hidden { "1" } else { "0" }));
    }
    w.write_event(Event::Empty(el))?;
    Ok(true)
}

fn write_cell_style_xf_xml(
    w: &mut XmlWriter,
    style: &Style,
    ids: ResolvedXfIds,
) -> std::io::Result<()> {
    write_xf_xml_with_options(w, style, ids, None, false)
}

fn write_xf_xml(
    w: &mut XmlWriter,
    style: &Style,
    ids: ResolvedXfIds,
    xf_id: u32,
) -> std::io::Result<()> {
    write_xf_xml_with_options(w, style, ids, Some(xf_id), true)
}

fn write_xf_xml_with_options(
    w: &mut XmlWriter,
    style: &Style,
    ids: ResolvedXfIds,
    xf_id: Option<u32>,
    include_apply_flags: bool,
) -> std::io::Result<()> {
    let num_fmt_s = ids.num_fmt_id.to_string();
    let font_s = ids.font_id.to_string();
    let fill_s = ids.fill_id.to_string();
    let border_s = ids.border_id.to_string();

    let mut el = BytesStart::new("xf");
    el.push_attribute(("numFmtId", num_fmt_s.as_str()));
    el.push_attribute(("fontId", font_s.as_str()));
    el.push_attribute(("fillId", fill_s.as_str()));
    el.push_attribute(("borderId", border_s.as_str()));
    let xf_id_s = xf_id.map(|v| v.to_string());
    if let Some(xf_id_s) = xf_id_s.as_deref() {
        el.push_attribute(("xfId", xf_id_s));
    }

    if include_apply_flags && ids.num_fmt_id != 0 {
        el.push_attribute(("applyNumberFormat", "1"));
    }
    if include_apply_flags && style.font != FontStyle::default() {
        el.push_attribute(("applyFont", "1"));
    }
    if include_apply_flags && style.fill != FillStyle::None {
        el.push_attribute(("applyFill", "1"));
    }
    if include_apply_flags && style.border != BorderStyle::default() {
        el.push_attribute(("applyBorder", "1"));
    }
    if include_apply_flags && style.alignment != Alignment::default() {
        el.push_attribute(("applyAlignment", "1"));
    }
    if include_apply_flags && style.protection != Protection::default() {
        el.push_attribute(("applyProtection", "1"));
    }

    let has_alignment = style.alignment != Alignment::default();
    let has_protection = style.protection != Protection::default();
    if !has_alignment && !has_protection {
        w.write_event(Event::Empty(el))?;
    } else {
        w.write_event(Event::Start(el))?;
        if has_alignment {
            write_alignment_xml(w, &style.alignment)?;
        }
        if has_protection {
            write_protection_xml(w, &style.protection)?;
        }
        w.write_event(Event::End(BytesEnd::new("xf")))?;
    }
    Ok(())
}

/// Write a DXF (differential format) element for conditional formatting.
fn write_dxf_xml(w: &mut XmlWriter, style: &Style) -> std::io::Result<()> {
    w.write_event(Event::Start(BytesStart::new("dxf")))?;

    // Font (only if non-default)
    if style.font != FontStyle::default() {
        write_font_xml(w, &style.font)?;
    }

    // Number format (inline in DXF with both numFmtId and formatCode)
    if style.number_format != NumberFormat::General {
        let (id, code) = match &style.number_format {
            NumberFormat::General => unreachable!(),
            NumberFormat::BuiltIn(id) => (*id, style.number_format.format_string().to_string()),
            NumberFormat::Custom(code) => (164, code.clone()),
        };
        let id_s = id.to_string();
        let code_esc = escape(&code);
        w.create_element("numFmt")
            .with_attribute(("numFmtId", id_s.as_str()))
            .with_attribute(("formatCode", &*code_esc))
            .write_empty()?;
    }

    // Fill (only if non-default)
    if style.fill != FillStyle::None {
        write_fill_xml(w, &style.fill)?;
    }

    // Alignment (only if non-default)
    write_alignment_xml(w, &style.alignment)?;

    // Border (only if non-default)
    if style.border != BorderStyle::default() {
        write_border_xml(w, &style.border)?;
    }

    w.write_event(Event::End(BytesEnd::new("dxf")))?;
    Ok(())
}

/// Result of reading styles.xml, containing both cell styles and DXF styles
#[derive(Debug)]
pub(crate) struct ParsedStyles {
    pub cell_styles: Vec<Style>,
    pub cell_style_xfs: Vec<Style>,
    pub named_styles: Vec<NamedCellStyle>,
    pub cell_xf_xf_ids: Vec<u32>,
    pub dxf_styles: Vec<Style>,
    pub num_fmts: HashMap<u32, String>,
}

impl ParsedStyles {
    pub(crate) fn roundtrip_data(&self) -> RoundtripStyleData {
        RoundtripStyleData {
            cell_styles: self.cell_styles.clone(),
            cell_style_xfs: self.cell_style_xfs.clone(),
            named_styles: self.named_styles.clone(),
            cell_xf_xf_ids: self.cell_xf_xf_ids.clone(),
        }
    }
}

#[derive(Debug, Clone, Copy, PartialEq, Eq)]
enum XfTarget {
    CellStyle,
    Cell,
}

#[derive(Debug, Clone, Default)]
struct ParsedXf {
    num_fmt_id: u32,
    font_id: u32,
    fill_id: u32,
    border_id: u32,
    xf_id: u32,
    apply_number_format: bool,
    apply_font: bool,
    apply_fill: bool,
    apply_border: bool,
    apply_alignment: bool,
    apply_protection: bool,
    alignment: Alignment,
    protection: Protection,
}

pub(crate) fn read_styles_xml<R: Read>(reader: R) -> XlsxResult<ParsedStyles> {
    let mut xml_reader = Reader::from_reader(BufReader::new(reader));
    xml_reader.config_mut().trim_text(true);

    let mut buf = Vec::new();

    let mut numfmts: HashMap<u32, String> = HashMap::new();
    let mut fonts: Vec<FontStyle> = Vec::new();
    let mut fills: Vec<FillStyle> = Vec::new();
    let mut borders: Vec<BorderStyle> = Vec::new();
    let mut cell_style_xf_defs: Vec<ParsedXf> = Vec::new();
    let mut cell_xf_defs: Vec<ParsedXf> = Vec::new();
    let mut named_styles: Vec<NamedCellStyle> = Vec::new();
    let mut dxf_styles: Vec<Style> = Vec::new();

    // Current objects while parsing
    let mut current_font: Option<FontStyle> = None;
    let mut current_fill_pattern: Option<PatternType> = None;
    let mut current_fill_fg: Color = Color::Auto;
    let mut current_fill_bg: Color = Color::Auto;
    let mut in_fill = false;
    let mut in_gradient_fill = false;
    let mut gradient_type: GradientType = GradientType::Linear;
    let mut gradient_angle: f64 = 0.0;
    let mut gradient_stops: Vec<GradientStop> = Vec::new();
    let mut in_gradient_stop = false;
    let mut current_stop_position: f64 = 0.0;

    let mut current_border: Option<BorderStyle> = None;
    let mut current_border_edge: Option<&'static str> = None;

    // Current xf
    let mut current_xf: Option<ParsedXf> = None;
    let mut current_xf_target: Option<XfTarget> = None;
    let mut in_cell_xfs = false;
    let mut in_cell_style_xfs = false;
    let mut in_cell_styles = false;

    // DXF parsing state
    let mut in_dxfs = false;
    let mut in_dxf = false;
    let mut current_dxf: Option<Style> = None;
    let mut dxf_font: Option<FontStyle> = None;
    let mut dxf_fill_pattern: Option<PatternType> = None;
    let mut dxf_fill_fg: Color = Color::Auto;
    let mut dxf_fill_bg: Color = Color::Auto;
    let mut in_dxf_fill = false;
    let mut dxf_in_gradient_fill = false;
    let mut dxf_gradient_type: GradientType = GradientType::Linear;
    let mut dxf_gradient_angle: f64 = 0.0;
    let mut dxf_gradient_stops: Vec<GradientStop> = Vec::new();
    let mut dxf_in_gradient_stop = false;
    let mut dxf_current_stop_position: f64 = 0.0;
    let mut dxf_border: Option<BorderStyle> = None;
    let mut dxf_border_edge: Option<&'static str> = None;

    loop {
        match xml_reader.read_event_into(&mut buf) {
            Ok(Event::Start(e)) => match e.name().local_name().as_ref() {
                b"numFmts" => {}
                b"fonts" => {}
                b"fills" => {}
                b"borders" => {}

                b"cellXfs" => in_cell_xfs = true,

                b"cellStyleXfs" => in_cell_style_xfs = true,

                b"cellStyles" => in_cell_styles = true,

                b"cellStyle" if in_cell_styles => {
                    named_styles.push(parse_named_cell_style_attrs(&e));
                }

                b"dxfs" => in_dxfs = true,

                b"dxf" if in_dxfs => {
                    in_dxf = true;
                    current_dxf = Some(Style::default());
                    dxf_font = None;
                    dxf_fill_pattern = None;
                    dxf_fill_fg = Color::Auto;
                    dxf_fill_bg = Color::Auto;
                    in_dxf_fill = false;
                    dxf_border = None;
                    dxf_border_edge = None;
                }

                b"font" if in_dxf => dxf_font = Some(FontStyle::default()),

                b"fill" if in_dxf => {
                    in_dxf_fill = true;
                    dxf_fill_pattern = None;
                    dxf_fill_fg = Color::Auto;
                    dxf_fill_bg = Color::Auto;
                }

                b"patternFill" if in_dxf_fill => {
                    for attr in e.attributes().flatten() {
                        if attr.key.local_name().as_ref() == b"patternType" {
                            if let Ok(v) = attr.unescape_value() {
                                dxf_fill_pattern = str_to_pattern_type(&v);
                            }
                        }
                    }
                }

                b"gradientFill" if in_dxf_fill => {
                    dxf_in_gradient_fill = true;
                    dxf_gradient_type = GradientType::Linear;
                    dxf_gradient_angle = 0.0;
                    dxf_gradient_stops.clear();
                    for attr in e.attributes().flatten() {
                        match attr.key.local_name().as_ref() {
                            b"type" => {
                                if let Ok(v) = attr.unescape_value() {
                                    dxf_gradient_type = match v.as_ref() {
                                        "path" => GradientType::Path,
                                        _ => GradientType::Linear,
                                    };
                                }
                            }
                            b"degree" => {
                                if let Ok(v) = attr.unescape_value() {
                                    dxf_gradient_angle = v.parse::<f64>().unwrap_or(0.0);
                                }
                            }
                            _ => {}
                        }
                    }
                }

                b"stop" if dxf_in_gradient_fill => {
                    dxf_in_gradient_stop = true;
                    dxf_current_stop_position = 0.0;
                    for attr in e.attributes().flatten() {
                        if attr.key.local_name().as_ref() == b"position" {
                            if let Ok(v) = attr.unescape_value() {
                                dxf_current_stop_position = v.parse::<f64>().unwrap_or(0.0);
                            }
                        }
                    }
                }

                b"border" if in_dxf => {
                    let mut b = BorderStyle::default();
                    for attr in e.attributes().flatten() {
                        match attr.key.local_name().as_ref() {
                            b"diagonalUp" => {
                                if attr.unescape_value().ok().as_deref() == Some("1") {
                                    b.diagonal_direction =
                                        duke_sheets_core::style::DiagonalDirection::Up;
                                }
                            }
                            b"diagonalDown" => {
                                if attr.unescape_value().ok().as_deref() == Some("1") {
                                    b.diagonal_direction =
                                        duke_sheets_core::style::DiagonalDirection::Down;
                                }
                            }
                            _ => {}
                        }
                    }
                    dxf_border = Some(b);
                }

                b"font" => current_font = Some(FontStyle::default()),

                b"fill" => {
                    in_fill = true;
                    current_fill_pattern = None;
                    current_fill_fg = Color::Auto;
                    current_fill_bg = Color::Auto;
                }

                b"patternFill" if in_fill => {
                    for attr in e.attributes().flatten() {
                        if attr.key.local_name().as_ref() == b"patternType" {
                            if let Ok(v) = attr.unescape_value() {
                                current_fill_pattern = str_to_pattern_type(&v);
                            }
                        }
                    }
                }

                b"gradientFill" if in_fill => {
                    in_gradient_fill = true;
                    gradient_type = GradientType::Linear;
                    gradient_angle = 0.0;
                    gradient_stops.clear();
                    for attr in e.attributes().flatten() {
                        match attr.key.local_name().as_ref() {
                            b"type" => {
                                if let Ok(v) = attr.unescape_value() {
                                    gradient_type = match v.as_ref() {
                                        "path" => GradientType::Path,
                                        _ => GradientType::Linear,
                                    };
                                }
                            }
                            b"degree" => {
                                if let Ok(v) = attr.unescape_value() {
                                    gradient_angle = v.parse::<f64>().unwrap_or(0.0);
                                }
                            }
                            _ => {}
                        }
                    }
                }

                b"stop" if in_gradient_fill => {
                    in_gradient_stop = true;
                    current_stop_position = 0.0;
                    for attr in e.attributes().flatten() {
                        if attr.key.local_name().as_ref() == b"position" {
                            if let Ok(v) = attr.unescape_value() {
                                current_stop_position = v.parse::<f64>().unwrap_or(0.0);
                            }
                        }
                    }
                }

                b"border" => {
                    let mut b = BorderStyle::default();
                    for attr in e.attributes().flatten() {
                        match attr.key.local_name().as_ref() {
                            b"diagonalUp" => {
                                if attr.unescape_value().ok().as_deref() == Some("1") {
                                    b.diagonal_direction =
                                        duke_sheets_core::style::DiagonalDirection::Up;
                                }
                            }
                            b"diagonalDown" => {
                                if attr.unescape_value().ok().as_deref() == Some("1") {
                                    b.diagonal_direction =
                                        duke_sheets_core::style::DiagonalDirection::Down;
                                }
                            }
                            _ => {}
                        }
                    }
                    current_border = Some(b);
                }

                // Border edges
                b"left" | b"right" | b"top" | b"bottom" | b"diagonal" => {
                    let edge_name: &'static str = match e.name().local_name().as_ref() {
                        b"left" => "left",
                        b"right" => "right",
                        b"top" => "top",
                        b"bottom" => "bottom",
                        _ => "diagonal",
                    };

                    // Parse style attribute
                    let mut style: Option<BorderLineStyle> = None;
                    for attr in e.attributes().flatten() {
                        if attr.key.local_name().as_ref() == b"style" {
                            if let Ok(v) = attr.unescape_value() {
                                style = str_to_border_style(&v);
                            }
                        }
                    }

                    if let Some(border) = dxf_border.as_mut() {
                        dxf_border_edge = Some(edge_name);
                        // Create edge with default color; color may be overwritten by nested <color>
                        if let Some(st) = style {
                            if st != BorderLineStyle::None {
                                set_border_edge(
                                    border,
                                    edge_name,
                                    Some(BorderEdge {
                                        style: st,
                                        color: Color::Auto,
                                    }),
                                );
                            }
                        }
                    } else if let Some(border) = current_border.as_mut() {
                        current_border_edge = Some(edge_name);
                        // Create edge with default color; color may be overwritten by nested <color>
                        if let Some(st) = style {
                            if st != BorderLineStyle::None {
                                set_border_edge(
                                    border,
                                    edge_name,
                                    Some(BorderEdge {
                                        style: st,
                                        color: Color::Auto,
                                    }),
                                );
                            }
                        }
                    }
                }

                b"xf" if in_cell_xfs || in_cell_style_xfs => {
                    if in_cell_style_xfs {
                    } else {
                    }
                    current_xf = Some(parse_xf_attrs(&e));
                    current_xf_target = Some(if in_cell_style_xfs {
                        XfTarget::CellStyle
                    } else {
                        XfTarget::Cell
                    });
                }

                b"alignment" => {
                    // DXF alignment takes priority
                    if in_dxf {
                        if let Some(dxf) = current_dxf.as_mut() {
                            parse_alignment_attrs(&e, &mut dxf.alignment);
                        }
                    } else if let Some(xf) = current_xf.as_mut() {
                        parse_alignment_attrs(&e, &mut xf.alignment);
                    }
                }

                b"protection" => {
                    if let Some(xf) = current_xf.as_mut() {
                        for attr in e.attributes().flatten() {
                            let val = match attr.unescape_value() {
                                Ok(v) => v,
                                Err(_) => continue,
                            };
                            match attr.key.local_name().as_ref() {
                                b"locked" => xf.protection.locked = val.as_ref() == "1",
                                b"hidden" => xf.protection.hidden = val.as_ref() == "1",
                                _ => {}
                            }
                        }
                    }
                }

                // font sub-elements (handle both regular fonts and DXF fonts)
                b"sz" => {
                    let font = dxf_font.as_mut().or(current_font.as_mut());
                    if let Some(font) = font {
                        for attr in e.attributes().flatten() {
                            if attr.key.local_name().as_ref() == b"val" {
                                if let Ok(v) = attr.unescape_value() {
                                    font.size = v.parse::<f64>().unwrap_or(font.size);
                                }
                            }
                        }
                    }
                }
                b"name" => {
                    let font = dxf_font.as_mut().or(current_font.as_mut());
                    if let Some(font) = font {
                        for attr in e.attributes().flatten() {
                            if attr.key.local_name().as_ref() == b"val" {
                                if let Ok(v) = attr.unescape_value() {
                                    font.name = v.to_string();
                                }
                            }
                        }
                    }
                }
                b"family" => {
                    let font = dxf_font.as_mut().or(current_font.as_mut());
                    if let Some(font) = font {
                        for attr in e.attributes().flatten() {
                            if attr.key.local_name().as_ref() == b"val" {
                                if let Ok(v) = attr.unescape_value() {
                                    font.family = v.parse::<u8>().ok();
                                }
                            }
                        }
                    }
                }
                b"charset" => {
                    let font = dxf_font.as_mut().or(current_font.as_mut());
                    if let Some(font) = font {
                        for attr in e.attributes().flatten() {
                            if attr.key.local_name().as_ref() == b"val" {
                                if let Ok(v) = attr.unescape_value() {
                                    font.charset = v.parse::<u8>().ok();
                                }
                            }
                        }
                    }
                }
                b"scheme" => {
                    let font = dxf_font.as_mut().or(current_font.as_mut());
                    if let Some(font) = font {
                        for attr in e.attributes().flatten() {
                            if attr.key.local_name().as_ref() == b"val" {
                                if let Ok(v) = attr.unescape_value() {
                                    font.scheme = Some(v.to_string());
                                }
                            }
                        }
                    }
                }
                b"b" => {
                    let font = dxf_font.as_mut().or(current_font.as_mut());
                    if let Some(font) = font {
                        font.bold = true;
                    }
                }
                b"i" => {
                    let font = dxf_font.as_mut().or(current_font.as_mut());
                    if let Some(font) = font {
                        font.italic = true;
                    }
                }
                b"strike" => {
                    let font = dxf_font.as_mut().or(current_font.as_mut());
                    if let Some(font) = font {
                        font.strikethrough = true;
                    }
                }
                b"u" => {
                    let font = dxf_font.as_mut().or(current_font.as_mut());
                    if let Some(font) = font {
                        let mut underline = Underline::Single;
                        for attr in e.attributes().flatten() {
                            if attr.key.local_name().as_ref() == b"val" {
                                if let Ok(v) = attr.unescape_value() {
                                    underline = str_to_underline(&v);
                                }
                            }
                        }
                        font.underline = underline;
                    }
                }

                b"vertAlign" => {
                    let font = dxf_font.as_mut().or(current_font.as_mut());
                    if let Some(font) = font {
                        for attr in e.attributes().flatten() {
                            if attr.key.local_name().as_ref() == b"val" {
                                if let Ok(v) = attr.unescape_value() {
                                    font.vertical_align = match v.as_ref() {
                                        "superscript" => FontVerticalAlign::Superscript,
                                        "subscript" => FontVerticalAlign::Subscript,
                                        _ => FontVerticalAlign::Baseline,
                                    };
                                }
                            }
                        }
                    }
                }

                b"color" => {
                    // Font color, border color, or gradient stop color depending on context
                    let color = parse_color_attrs(&e);
                    // Gradient stop colors take priority
                    if dxf_in_gradient_stop {
                        dxf_gradient_stops
                            .push(GradientStop::new(dxf_current_stop_position, color));
                    } else if in_gradient_stop {
                        gradient_stops.push(GradientStop::new(current_stop_position, color));
                    } else if let Some(font) = dxf_font.as_mut() {
                        font.color = color;
                    } else if let (Some(border), Some(edge_name)) =
                        (dxf_border.as_mut(), dxf_border_edge)
                    {
                        let edge_opt = get_border_edge(border, edge_name).clone();
                        if let Some(mut edge) = edge_opt {
                            edge.color = color;
                            set_border_edge(border, edge_name, Some(edge));
                        }
                    } else if let Some(font) = current_font.as_mut() {
                        font.color = color;
                    } else if let (Some(border), Some(edge_name)) =
                        (current_border.as_mut(), current_border_edge)
                    {
                        // Update border edge color if edge exists
                        let edge_opt = get_border_edge(border, edge_name).clone();
                        if let Some(mut edge) = edge_opt {
                            edge.color = color;
                            set_border_edge(border, edge_name, Some(edge));
                        }
                    }
                }

                b"fgColor" => {
                    if in_dxf_fill {
                        dxf_fill_fg = parse_color_attrs(&e);
                    } else if in_fill {
                        current_fill_fg = parse_color_attrs(&e);
                    }
                }
                b"bgColor" => {
                    if in_dxf_fill {
                        dxf_fill_bg = parse_color_attrs(&e);
                    } else if in_fill {
                        current_fill_bg = parse_color_attrs(&e);
                    }
                }

                _ => {}
            },

            Ok(Event::Empty(e)) => match e.name().local_name().as_ref() {
                b"numFmt" => {
                    let mut id: Option<u32> = None;
                    let mut code: Option<String> = None;
                    for attr in e.attributes().flatten() {
                        match attr.key.local_name().as_ref() {
                            b"numFmtId" => {
                                id = attr.unescape_value().ok().and_then(|s| s.parse().ok())
                            }
                            b"formatCode" => {
                                code = attr.unescape_value().ok().map(|s| s.to_string())
                            }
                            _ => {}
                        }
                    }
                    if in_dxf {
                        // DXF numFmt: apply directly to the current DXF style
                        if let Some(dxf) = current_dxf.as_mut() {
                            dxf.number_format = match (id, code) {
                                (Some(0), _) => NumberFormat::General,
                                (_, Some(c)) => NumberFormat::Custom(c),
                                (Some(id), None) => NumberFormat::BuiltIn(id),
                                _ => NumberFormat::General,
                            };
                        }
                    } else if let (Some(id), Some(code)) = (id, code) {
                        numfmts.insert(id, code);
                    }
                }

                // Font empty tags (handle both regular fonts and DXF fonts)
                b"b" => {
                    let font = dxf_font.as_mut().or(current_font.as_mut());
                    if let Some(font) = font {
                        font.bold = true;
                    }
                }
                b"i" => {
                    let font = dxf_font.as_mut().or(current_font.as_mut());
                    if let Some(font) = font {
                        font.italic = true;
                    }
                }
                b"strike" => {
                    let font = dxf_font.as_mut().or(current_font.as_mut());
                    if let Some(font) = font {
                        font.strikethrough = true;
                    }
                }
                b"u" => {
                    let font = dxf_font.as_mut().or(current_font.as_mut());
                    if let Some(font) = font {
                        let mut underline = Underline::Single;
                        for attr in e.attributes().flatten() {
                            if attr.key.local_name().as_ref() == b"val" {
                                if let Ok(v) = attr.unescape_value() {
                                    underline = str_to_underline(&v);
                                }
                            }
                        }
                        font.underline = underline;
                    }
                }
                b"vertAlign" => {
                    let font = dxf_font.as_mut().or(current_font.as_mut());
                    if let Some(font) = font {
                        for attr in e.attributes().flatten() {
                            if attr.key.local_name().as_ref() == b"val" {
                                if let Ok(v) = attr.unescape_value() {
                                    font.vertical_align = match v.as_ref() {
                                        "superscript" => FontVerticalAlign::Superscript,
                                        "subscript" => FontVerticalAlign::Subscript,
                                        _ => FontVerticalAlign::Baseline,
                                    };
                                }
                            }
                        }
                    }
                }
                b"sz" => {
                    let font = dxf_font.as_mut().or(current_font.as_mut());
                    if let Some(font) = font {
                        for attr in e.attributes().flatten() {
                            if attr.key.local_name().as_ref() == b"val" {
                                if let Ok(v) = attr.unescape_value() {
                                    font.size = v.parse::<f64>().unwrap_or(font.size);
                                }
                            }
                        }
                    }
                }
                b"name" => {
                    let font = dxf_font.as_mut().or(current_font.as_mut());
                    if let Some(font) = font {
                        for attr in e.attributes().flatten() {
                            if attr.key.local_name().as_ref() == b"val" {
                                if let Ok(v) = attr.unescape_value() {
                                    font.name = v.to_string();
                                }
                            }
                        }
                    }
                }
                b"family" => {
                    let font = dxf_font.as_mut().or(current_font.as_mut());
                    if let Some(font) = font {
                        for attr in e.attributes().flatten() {
                            if attr.key.local_name().as_ref() == b"val" {
                                if let Ok(v) = attr.unescape_value() {
                                    font.family = v.parse::<u8>().ok();
                                }
                            }
                        }
                    }
                }
                b"charset" => {
                    let font = dxf_font.as_mut().or(current_font.as_mut());
                    if let Some(font) = font {
                        for attr in e.attributes().flatten() {
                            if attr.key.local_name().as_ref() == b"val" {
                                if let Ok(v) = attr.unescape_value() {
                                    font.charset = v.parse::<u8>().ok();
                                }
                            }
                        }
                    }
                }
                b"scheme" => {
                    let font = dxf_font.as_mut().or(current_font.as_mut());
                    if let Some(font) = font {
                        for attr in e.attributes().flatten() {
                            if attr.key.local_name().as_ref() == b"val" {
                                if let Ok(v) = attr.unescape_value() {
                                    font.scheme = Some(v.to_string());
                                }
                            }
                        }
                    }
                }
                b"color" => {
                    let color = parse_color_attrs(&e);
                    // Gradient stop colors take priority
                    if dxf_in_gradient_stop {
                        dxf_gradient_stops
                            .push(GradientStop::new(dxf_current_stop_position, color));
                    } else if in_gradient_stop {
                        gradient_stops.push(GradientStop::new(current_stop_position, color));
                    } else if let Some(font) = dxf_font.as_mut() {
                        font.color = color;
                    } else if let (Some(border), Some(edge_name)) =
                        (dxf_border.as_mut(), dxf_border_edge)
                    {
                        let edge_opt = get_border_edge(border, edge_name).clone();
                        if let Some(mut edge) = edge_opt {
                            edge.color = color;
                            set_border_edge(border, edge_name, Some(edge));
                        }
                    } else if let Some(font) = current_font.as_mut() {
                        font.color = color;
                    } else if let (Some(border), Some(edge_name)) =
                        (current_border.as_mut(), current_border_edge)
                    {
                        let edge_opt = get_border_edge(border, edge_name).clone();
                        if let Some(mut edge) = edge_opt {
                            edge.color = color;
                            set_border_edge(border, edge_name, Some(edge));
                        }
                    }
                }

                // Fill colors (handle both regular fills and DXF fills)
                b"fgColor" => {
                    if in_dxf_fill {
                        dxf_fill_fg = parse_color_attrs(&e);
                    } else if in_fill {
                        current_fill_fg = parse_color_attrs(&e);
                    }
                }
                b"bgColor" => {
                    if in_dxf_fill {
                        dxf_fill_bg = parse_color_attrs(&e);
                    } else if in_fill {
                        current_fill_bg = parse_color_attrs(&e);
                    }
                }

                // alignment can be self-closing
                b"alignment" => {
                    // DXF alignment takes priority
                    if in_dxf {
                        if let Some(dxf) = current_dxf.as_mut() {
                            parse_alignment_attrs(&e, &mut dxf.alignment);
                        }
                    } else if let Some(xf) = current_xf.as_mut() {
                        parse_alignment_attrs(&e, &mut xf.alignment);
                    }
                }

                // protection can be self-closing
                b"protection" => {
                    if let Some(xf) = current_xf.as_mut() {
                        for attr in e.attributes().flatten() {
                            let val = match attr.unescape_value() {
                                Ok(v) => v,
                                Err(_) => continue,
                            };
                            match attr.key.local_name().as_ref() {
                                b"locked" => xf.protection.locked = val.as_ref() == "1",
                                b"hidden" => xf.protection.hidden = val.as_ref() == "1",
                                _ => {}
                            }
                        }
                    }
                }

                // Border edges can be self-closing (e.g. <left style="dotted"/>)
                b"left" | b"right" | b"top" | b"bottom" | b"diagonal" => {
                    let edge_name: &'static str = match e.name().local_name().as_ref() {
                        b"left" => "left",
                        b"right" => "right",
                        b"top" => "top",
                        b"bottom" => "bottom",
                        _ => "diagonal",
                    };

                    // Parse style attribute
                    let mut style: Option<BorderLineStyle> = None;
                    for attr in e.attributes().flatten() {
                        if attr.key.local_name().as_ref() == b"style" {
                            if let Ok(v) = attr.unescape_value() {
                                style = str_to_border_style(&v);
                            }
                        }
                    }

                    if let Some(border) = dxf_border.as_mut() {
                        if let Some(st) = style {
                            if st != BorderLineStyle::None {
                                set_border_edge(
                                    border,
                                    edge_name,
                                    Some(BorderEdge {
                                        style: st,
                                        color: Color::Auto,
                                    }),
                                );
                            }
                        }
                    } else if let Some(border) = current_border.as_mut() {
                        if let Some(st) = style {
                            if st != BorderLineStyle::None {
                                set_border_edge(
                                    border,
                                    edge_name,
                                    Some(BorderEdge {
                                        style: st,
                                        color: Color::Auto,
                                    }),
                                );
                            }
                        }
                    }
                }

                // patternFill can be self-closing
                b"patternFill" => {
                    if in_dxf_fill {
                        for attr in e.attributes().flatten() {
                            if attr.key.local_name().as_ref() == b"patternType" {
                                if let Ok(v) = attr.unescape_value() {
                                    dxf_fill_pattern = str_to_pattern_type(&v);
                                }
                            }
                        }
                    } else if in_fill {
                        for attr in e.attributes().flatten() {
                            if attr.key.local_name().as_ref() == b"patternType" {
                                if let Ok(v) = attr.unescape_value() {
                                    current_fill_pattern = str_to_pattern_type(&v);
                                }
                            }
                        }
                    }
                }

                // gradientFill can be self-closing (no stops - degenerate case)
                b"gradientFill" => {
                    // Self-closing gradientFill with no stops is a no-op,
                    // but parse attributes in case we need them.
                    if in_dxf_fill {
                        dxf_in_gradient_fill = false;
                    } else if in_fill {
                        in_gradient_fill = false;
                    }
                }

                // xf can be empty (no child elements)
                b"xf" if in_cell_xfs || in_cell_style_xfs => {
                    if in_cell_style_xfs {
                    } else {
                    }
                    let xf = parse_xf_attrs(&e);
                    if in_cell_style_xfs {
                        cell_style_xf_defs.push(xf);
                    } else {
                        cell_xf_defs.push(xf);
                    }
                }

                b"cellStyle" if in_cell_styles => {
                    named_styles.push(parse_named_cell_style_attrs(&e));
                }

                _ => {}
            },

            Ok(Event::End(e)) => match e.name().local_name().as_ref() {
                b"font" => {
                    if in_dxf {
                        // DXF font - apply to current DXF style
                        if let (Some(f), Some(dxf)) = (dxf_font.take(), current_dxf.as_mut()) {
                            dxf.font = f;
                        }
                    } else if let Some(f) = current_font.take() {
                        fonts.push(f);
                    }
                }
                b"stop" => {
                    dxf_in_gradient_stop = false;
                    in_gradient_stop = false;
                }
                b"gradientFill" => {
                    dxf_in_gradient_fill = false;
                    in_gradient_fill = false;
                }
                b"fill" => {
                    if in_dxf_fill {
                        // DXF fill - check if gradient was parsed
                        let fill = if !dxf_gradient_stops.is_empty() {
                            FillStyle::Gradient {
                                gradient_type: dxf_gradient_type,
                                angle: dxf_gradient_angle,
                                stops: std::mem::take(&mut dxf_gradient_stops),
                            }
                        } else {
                            finalize_fill(dxf_fill_pattern, dxf_fill_fg, dxf_fill_bg)
                        };
                        if let Some(dxf) = current_dxf.as_mut() {
                            dxf.fill = fill;
                        }
                        in_dxf_fill = false;
                        dxf_fill_pattern = None;
                        dxf_in_gradient_fill = false;
                    } else if in_fill {
                        // Regular fill - check if gradient was parsed
                        let fill = if !gradient_stops.is_empty() {
                            FillStyle::Gradient {
                                gradient_type,
                                angle: gradient_angle,
                                stops: std::mem::take(&mut gradient_stops),
                            }
                        } else {
                            finalize_fill(current_fill_pattern, current_fill_fg, current_fill_bg)
                        };
                        fills.push(fill);
                        in_fill = false;
                        current_fill_pattern = None;
                        in_gradient_fill = false;
                    }
                }
                b"border" => {
                    if in_dxf {
                        // DXF border - apply to current DXF style
                        if let (Some(b), Some(dxf)) = (dxf_border.take(), current_dxf.as_mut()) {
                            dxf.border = b;
                        }
                        dxf_border_edge = None;
                    } else if let Some(b) = current_border.take() {
                        borders.push(b);
                    }
                    current_border_edge = None;
                }
                b"left" | b"right" | b"top" | b"bottom" | b"diagonal" => {
                    current_border_edge = None;
                    dxf_border_edge = None;
                }
                b"dxf" => {
                    if let Some(dxf) = current_dxf.take() {
                        dxf_styles.push(dxf);
                    }
                    in_dxf = false;
                }
                b"dxfs" => in_dxfs = false,
                b"xf" => {
                    if let Some(xf) = current_xf.take() {
                        match current_xf_target.take() {
                            Some(XfTarget::CellStyle) => cell_style_xf_defs.push(xf),
                            Some(XfTarget::Cell) => cell_xf_defs.push(xf),
                            None => {}
                        }
                    }
                }
                b"cellXfs" => in_cell_xfs = false,
                b"cellStyleXfs" => in_cell_style_xfs = false,
                b"cellStyles" => in_cell_styles = false,
                _ => {}
            },

            Ok(Event::Eof) => break,
            Err(e) => return Err(XlsxError::Xml(e)),
            _ => {}
        }

        buf.clear();
    }

    let cell_style_bases: Vec<Style> = if cell_style_xf_defs.is_empty() {
        Vec::new()
    } else {
        cell_style_xf_defs
            .iter()
            .map(|xf| {
                resolve_style(
                    xf.num_fmt_id,
                    xf.font_id,
                    xf.fill_id,
                    xf.border_id,
                    xf.alignment.clone(),
                    xf.protection,
                    &numfmts,
                    &fonts,
                    &fills,
                    &borders,
                )
            })
            .collect()
    };

    let cell_xf_xf_ids: Vec<u32> = if cell_xf_defs.is_empty() {
        vec![0]
    } else {
        cell_xf_defs.iter().map(|xf| xf.xf_id).collect()
    };

    let cell_styles = if cell_xf_defs.is_empty() {
        vec![Style::default()]
    } else {
        cell_xf_defs
            .iter()
            .map(|xf| {
                let resolved = resolve_style(
                    xf.num_fmt_id,
                    xf.font_id,
                    xf.fill_id,
                    xf.border_id,
                    xf.alignment.clone(),
                    xf.protection,
                    &numfmts,
                    &fonts,
                    &fills,
                    &borders,
                );
                if let Some(base) = cell_style_bases.get(xf.xf_id as usize).cloned() {
                    merge_cell_xf_with_base(base, &resolved, xf)
                } else {
                    resolved
                }
            })
            .collect()
    };

    Ok(ParsedStyles {
        cell_styles,
        cell_style_xfs: cell_style_bases,
        named_styles,
        cell_xf_xf_ids,
        dxf_styles,
        num_fmts: numfmts,
    })
}

fn parse_named_cell_style_attrs(e: &quick_xml::events::BytesStart<'_>) -> NamedCellStyle {
    let mut name = String::from("Normal");
    let mut xf_id = 0;
    let mut builtin_id = None;

    for attr in e.attributes().flatten() {
        let value = match attr.unescape_value() {
            Ok(v) => v,
            Err(_) => continue,
        };
        match attr.key.local_name().as_ref() {
            b"name" => name = value.to_string(),
            b"xfId" => xf_id = value.parse::<u32>().unwrap_or(0),
            b"builtinId" => builtin_id = value.parse::<u32>().ok(),
            _ => {}
        }
    }

    NamedCellStyle {
        name,
        xf_id,
        builtin_id,
    }
}

#[allow(clippy::too_many_arguments)]
fn resolve_style(
    num_fmt_id: u32,
    font_id: u32,
    fill_id: u32,
    border_id: u32,
    alignment: Alignment,
    protection: Protection,
    numfmts: &HashMap<u32, String>,
    fonts: &[FontStyle],
    fills: &[FillStyle],
    borders: &[BorderStyle],
) -> Style {
    let number_format = if num_fmt_id == 0 {
        NumberFormat::General
    } else if let Some(code) = numfmts.get(&num_fmt_id) {
        NumberFormat::Custom(code.clone())
    } else {
        NumberFormat::BuiltIn(num_fmt_id)
    };

    Style {
        font: fonts.get(font_id as usize).cloned().unwrap_or_default(),
        fill: fills.get(fill_id as usize).cloned().unwrap_or_default(),
        border: borders.get(border_id as usize).cloned().unwrap_or_default(),
        alignment,
        protection,
        number_format,
    }
}

fn parse_xf_attrs(e: &quick_xml::events::BytesStart<'_>) -> ParsedXf {
    let mut xf = ParsedXf::default();

    for attr in e.attributes().flatten() {
        let value = match attr.unescape_value() {
            Ok(v) => v,
            Err(_) => continue,
        };
        match attr.key.local_name().as_ref() {
            b"numFmtId" => xf.num_fmt_id = value.parse::<u32>().unwrap_or(0),
            b"fontId" => xf.font_id = value.parse::<u32>().unwrap_or(0),
            b"fillId" => xf.fill_id = value.parse::<u32>().unwrap_or(0),
            b"borderId" => xf.border_id = value.parse::<u32>().unwrap_or(0),
            b"xfId" => xf.xf_id = value.parse::<u32>().unwrap_or(0),
            b"applyNumberFormat" => xf.apply_number_format = value.as_ref() == "1",
            b"applyFont" => xf.apply_font = value.as_ref() == "1",
            b"applyFill" => xf.apply_fill = value.as_ref() == "1",
            b"applyBorder" => xf.apply_border = value.as_ref() == "1",
            b"applyAlignment" => xf.apply_alignment = value.as_ref() == "1",
            b"applyProtection" => xf.apply_protection = value.as_ref() == "1",
            _ => {}
        }
    }

    xf
}

fn merge_cell_xf_with_base(base: Style, resolved_xf: &Style, xf_meta: &ParsedXf) -> Style {
    let mut out = base;

    if xf_meta.apply_number_format || xf_meta.num_fmt_id != 0 {
        out.number_format = resolved_xf.number_format.clone();
    }
    if xf_meta.apply_font || xf_meta.font_id != 0 {
        out.font = resolved_xf.font.clone();
    }
    if xf_meta.apply_fill || xf_meta.fill_id != 0 {
        out.fill = resolved_xf.fill.clone();
    }
    if xf_meta.apply_border || xf_meta.border_id != 0 {
        out.border = resolved_xf.border.clone();
    }
    if xf_meta.apply_alignment || xf_meta.alignment != Alignment::default() {
        out.alignment = resolved_xf.alignment.clone();
    }
    if xf_meta.apply_protection || xf_meta.protection != Protection::default() {
        out.protection = resolved_xf.protection;
    }

    out
}

fn finalize_fill(pattern: Option<PatternType>, fg: Color, bg: Color) -> FillStyle {
    // When patternType is absent but colors are present (common in DXF entries
    // from LibreOffice), infer solid fill.
    let effective_pattern = match pattern {
        Some(p) => p,
        None => {
            let has_fg = !matches!(fg, Color::Auto);
            let has_bg = !matches!(bg, Color::Auto);
            if has_fg || has_bg {
                PatternType::Solid
            } else {
                PatternType::None
            }
        }
    };

    match effective_pattern {
        PatternType::None => FillStyle::None,
        PatternType::Solid => {
            // For solid fills, prefer fg; fall back to bg (LO DXF entries often
            // put the fill color in bgColor rather than fgColor).
            let color = if matches!(fg, Color::Auto) { bg } else { fg };
            FillStyle::Solid { color }
        }
        PatternType::Gray125 => FillStyle::None,
        p => FillStyle::Pattern {
            pattern: p,
            foreground: fg,
            background: bg,
        },
    }
}

fn parse_color_attrs(e: &quick_xml::events::BytesStart<'_>) -> Color {
    // Priority: rgb > theme > indexed > auto
    let mut rgb: Option<String> = None;
    let mut theme: Option<u8> = None;
    let mut tint: Option<f64> = None;
    let mut indexed: Option<u8> = None;
    let mut auto = false;

    for attr in e.attributes().flatten() {
        match attr.key.local_name().as_ref() {
            b"rgb" => rgb = attr.unescape_value().ok().map(|s| s.to_string()),
            b"theme" => {
                theme = attr
                    .unescape_value()
                    .ok()
                    .and_then(|s| s.parse::<u8>().ok());
            }
            b"tint" => {
                tint = attr
                    .unescape_value()
                    .ok()
                    .and_then(|s| s.parse::<f64>().ok());
            }
            b"indexed" => {
                indexed = attr
                    .unescape_value()
                    .ok()
                    .and_then(|s| s.parse::<u8>().ok());
            }
            b"auto" => auto = attr.unescape_value().ok().as_deref() == Some("1"),
            _ => {}
        }
    }

    if let Some(rgb) = rgb {
        let hex = rgb.trim_start_matches('#');
        if hex.len() == 8 {
            if let (Ok(a), Ok(r), Ok(g), Ok(b)) = (
                u8::from_str_radix(&hex[0..2], 16),
                u8::from_str_radix(&hex[2..4], 16),
                u8::from_str_radix(&hex[4..6], 16),
                u8::from_str_radix(&hex[6..8], 16),
            ) {
                return Color::Argb { a, r, g, b };
            }
        } else if hex.len() == 6 {
            if let (Ok(r), Ok(g), Ok(b)) = (
                u8::from_str_radix(&hex[0..2], 16),
                u8::from_str_radix(&hex[2..4], 16),
                u8::from_str_radix(&hex[4..6], 16),
            ) {
                return Color::Rgb { r, g, b };
            }
        }
    }

    if let Some(index) = theme {
        let tint_i8 = tint.map(|t| (t * 100.0).round() as i8).unwrap_or(0);
        return Color::Theme {
            index,
            tint: tint_i8,
        };
    }

    if let Some(i) = indexed {
        return Color::Indexed(i);
    }

    if auto {
        return Color::Auto;
    }

    Color::Auto
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn test_read_styles_cell_xf_inherits_from_cell_style_xf() {
        let xml = r#"<?xml version="1.0" encoding="UTF-8"?>
<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <fonts count="2">
    <font><sz val="11"/><name val="Calibri"/></font>
    <font><b/><sz val="11"/><name val="Calibri"/></font>
  </fonts>
  <fills count="2">
    <fill><patternFill patternType="none"/></fill>
    <fill><patternFill patternType="gray125"/></fill>
  </fills>
  <borders count="1"><border><left/><right/><top/><bottom/><diagonal/></border></borders>
  <cellStyleXfs count="2">
    <xf numFmtId="0" fontId="0" fillId="0" borderId="0"/>
    <xf numFmtId="0" fontId="1" fillId="0" borderId="0" applyFont="1"/>
  </cellStyleXfs>
  <cellXfs count="1">
    <xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="1"/>
  </cellXfs>
  <cellStyles count="1"><cellStyle name="Normal" xfId="0" builtinId="0"/></cellStyles>
</styleSheet>"#;

        let parsed = read_styles_xml(xml.as_bytes()).expect("parse styles");
        assert_eq!(parsed.cell_styles.len(), 1);
        assert!(parsed.cell_styles[0].font.bold);
    }

    #[test]
    fn test_read_styles_cell_xf_can_override_base_with_apply_flag() {
        let xml = r#"<?xml version="1.0" encoding="UTF-8"?>
<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <fonts count="2">
    <font><sz val="11"/><name val="Calibri"/></font>
    <font><b/><sz val="11"/><name val="Calibri"/></font>
  </fonts>
  <fills count="2">
    <fill><patternFill patternType="none"/></fill>
    <fill><patternFill patternType="gray125"/></fill>
  </fills>
  <borders count="1"><border><left/><right/><top/><bottom/><diagonal/></border></borders>
  <cellStyleXfs count="2">
    <xf numFmtId="0" fontId="0" fillId="0" borderId="0"/>
    <xf numFmtId="0" fontId="1" fillId="0" borderId="0" applyFont="1"/>
  </cellStyleXfs>
  <cellXfs count="1">
    <xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="1" applyFont="1"/>
  </cellXfs>
  <cellStyles count="1"><cellStyle name="Normal" xfId="0" builtinId="0"/></cellStyles>
</styleSheet>"#;

        let parsed = read_styles_xml(xml.as_bytes()).expect("parse styles");
        assert_eq!(parsed.cell_styles.len(), 1);
        assert!(!parsed.cell_styles[0].font.bold);
    }

    #[test]
    fn test_read_styles_font_family_charset_scheme() {
        let xml = r#"<?xml version="1.0" encoding="UTF-8"?>
<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <fonts count="1">
    <font>
      <sz val="11"/>
      <name val="Calibri"/>
      <family val="2"/>
      <charset val="1"/>
      <scheme val="minor"/>
    </font>
  </fonts>
  <fills count="2">
    <fill><patternFill patternType="none"/></fill>
    <fill><patternFill patternType="gray125"/></fill>
  </fills>
  <borders count="1"><border><left/><right/><top/><bottom/><diagonal/></border></borders>
  <cellXfs count="1">
    <xf numFmtId="0" fontId="0" fillId="0" borderId="0"/>
  </cellXfs>
</styleSheet>"#;

        let parsed = read_styles_xml(xml.as_bytes()).expect("parse styles");
        let font = &parsed.cell_styles[0].font;
        assert_eq!(font.family, Some(2));
        assert_eq!(font.charset, Some(1));
        assert_eq!(font.scheme.as_deref(), Some("minor"));
    }
}

fn str_to_pattern_type(s: &str) -> Option<PatternType> {
    Some(match s {
        "none" => PatternType::None,
        "solid" => PatternType::Solid,
        "mediumGray" => PatternType::MediumGray,
        "darkGray" => PatternType::DarkGray,
        "lightGray" => PatternType::LightGray,
        "darkHorizontal" => PatternType::DarkHorizontal,
        "darkVertical" => PatternType::DarkVertical,
        "darkDown" => PatternType::DarkDown,
        "darkUp" => PatternType::DarkUp,
        "darkGrid" => PatternType::DarkGrid,
        "darkTrellis" => PatternType::DarkTrellis,
        "lightHorizontal" => PatternType::LightHorizontal,
        "lightVertical" => PatternType::LightVertical,
        "lightDown" => PatternType::LightDown,
        "lightUp" => PatternType::LightUp,
        "lightGrid" => PatternType::LightGrid,
        "lightTrellis" => PatternType::LightTrellis,
        "gray125" => PatternType::Gray125,
        "gray0625" => PatternType::Gray0625,
        _ => return None,
    })
}

fn str_to_border_style(s: &str) -> Option<BorderLineStyle> {
    Some(match s {
        "thin" => BorderLineStyle::Thin,
        "medium" => BorderLineStyle::Medium,
        "thick" => BorderLineStyle::Thick,
        "dashed" => BorderLineStyle::Dashed,
        "dotted" => BorderLineStyle::Dotted,
        "double" => BorderLineStyle::Double,
        "hair" => BorderLineStyle::Hair,
        "mediumDashed" => BorderLineStyle::MediumDashed,
        "dashDot" => BorderLineStyle::DashDot,
        "mediumDashDot" => BorderLineStyle::MediumDashDot,
        "dashDotDot" => BorderLineStyle::DashDotDot,
        "mediumDashDotDot" => BorderLineStyle::MediumDashDotDot,
        "slantDashDot" => BorderLineStyle::SlantDashDot,
        _ => return None,
    })
}

fn str_to_horizontal(s: &str) -> Option<HorizontalAlignment> {
    Some(match s {
        "general" => HorizontalAlignment::General,
        "left" => HorizontalAlignment::Left,
        "center" => HorizontalAlignment::Center,
        "right" => HorizontalAlignment::Right,
        "fill" => HorizontalAlignment::Fill,
        "justify" => HorizontalAlignment::Justify,
        "centerContinuous" => HorizontalAlignment::CenterContinuous,
        "distributed" => HorizontalAlignment::Distributed,
        _ => return None,
    })
}

fn str_to_vertical(s: &str) -> Option<VerticalAlignment> {
    Some(match s {
        "top" => VerticalAlignment::Top,
        "center" => VerticalAlignment::Center,
        "bottom" => VerticalAlignment::Bottom,
        "justify" => VerticalAlignment::Justify,
        "distributed" => VerticalAlignment::Distributed,
        _ => return None,
    })
}

fn str_to_underline(s: &str) -> Underline {
    match s {
        "double" => Underline::Double,
        "singleAccounting" => Underline::SingleAccounting,
        "doubleAccounting" => Underline::DoubleAccounting,
        _ => Underline::Single,
    }
}

fn get_border_edge<'a>(border: &'a BorderStyle, edge: &str) -> &'a Option<BorderEdge> {
    match edge {
        "left" => &border.left,
        "right" => &border.right,
        "top" => &border.top,
        "bottom" => &border.bottom,
        _ => &border.diagonal,
    }
}

fn set_border_edge(border: &mut BorderStyle, edge: &str, val: Option<BorderEdge>) {
    match edge {
        "left" => border.left = val,
        "right" => border.right = val,
        "top" => border.top = val,
        "bottom" => border.bottom = val,
        _ => border.diagonal = val,
    }
}

/// Parse alignment attributes from an XML element into an Alignment struct.
/// Handles both `"1"` and `"true"` for boolean attributes (Excel vs LibreOffice).
fn parse_alignment_attrs(e: &quick_xml::events::BytesStart<'_>, align: &mut Alignment) {
    for attr in e.attributes().flatten() {
        let val = match attr.unescape_value() {
            Ok(v) => v,
            Err(_) => continue,
        };
        match attr.key.local_name().as_ref() {
            b"horizontal" => {
                if let Some(h) = str_to_horizontal(&val) {
                    align.horizontal = h;
                }
            }
            b"vertical" => {
                if let Some(v) = str_to_vertical(&val) {
                    align.vertical = v;
                }
            }
            b"wrapText" => {
                align.wrap_text = val.as_ref() == "1" || val.as_ref() == "true";
            }
            b"shrinkToFit" => {
                align.shrink_to_fit = val.as_ref() == "1" || val.as_ref() == "true";
            }
            b"indent" => align.indent = val.parse::<u8>().unwrap_or(0),
            b"textRotation" => align.rotation = val.parse::<i16>().unwrap_or(0),
            b"readingOrder" => {
                align.reading_order = match val.as_ref() {
                    "1" => ReadingOrder::LeftToRight,
                    "2" => ReadingOrder::RightToLeft,
                    _ => ReadingOrder::ContextDependent,
                };
            }
            _ => {}
        }
    }
}
