use std::collections::HashMap;
use std::io::{Seek, Write};

use zip::write::SimpleFileOptions;
use zip::ZipWriter;

use duke_sheets_core::style::{
    Alignment, BorderEdge, BorderLineStyle, BorderStyle, Color, DiagonalDirection, FillStyle,
    FontStyle, FontVerticalAlign, GradientType, HorizontalAlignment, NumberFormat, PatternType,
    Protection, ReadingOrder, Style, Underline, VerticalAlignment,
};
use duke_sheets_core::Workbook;

use crate::biff12::{encode_wide_str, records, RecordWriter};
use crate::error::XlsbResult;

pub(crate) struct StyleMapping {
    sheet_maps: Vec<HashMap<u32, u32>>,
    custom_numfmt_ids: HashMap<String, u16>,
    #[cfg(test)]
    xf_count: u32,
}

impl StyleMapping {
    pub fn xf_index(&self, sheet_index: usize, local_style_index: u32) -> u32 {
        self.sheet_maps
            .get(sheet_index)
            .and_then(|m| m.get(&local_style_index))
            .copied()
            .unwrap_or(0)
    }

    pub fn custom_numfmt_id(&self, code: &str) -> Option<u16> {
        self.custom_numfmt_ids.get(code).copied()
    }

    #[cfg(test)]
    pub fn xf_count(&self) -> u32 {
        self.xf_count
    }

    #[cfg(test)]
    pub fn max_mapped_xf(&self) -> u32 {
        self.sheet_maps
            .iter()
            .flat_map(|m| m.values())
            .copied()
            .max()
            .unwrap_or(0)
    }

    #[cfg(test)]
    pub fn sheet_map(&self, index: usize) -> Option<&HashMap<u32, u32>> {
        self.sheet_maps.get(index)
    }
}

pub(crate) fn write_styles<W: Write + Seek>(
    zip: &mut ZipWriter<W>,
    options: &SimpleFileOptions,
    workbook: &Workbook,
    extra_fonts: &[FontStyle],
) -> XlsbResult<(StyleMapping, Vec<u16>, DxfMapping)> {
    let (table, mapping, extra_font_ids) = build_style_table(workbook, extra_fonts);

    zip.start_file("xl/styles.bin", *options)?;
    let mut buf = Vec::new();
    let mut rw = RecordWriter::new(&mut buf);

    rw.write_record(0x0116, &[])?;

    write_numfmts(&mut rw, &table)?;
    write_fonts(&mut rw, &table)?;
    write_fills(&mut rw, &table)?;
    write_borders(&mut rw, &table)?;
    write_cell_style_xfs(&mut rw)?;
    write_cell_xfs(&mut rw, &table)?;
    write_named_styles(&mut rw)?;

    let dxf_styles = collect_dxf_styles(workbook);
    write_dxfs(&mut rw, &dxf_styles)?;

    // Empty TableStyles section (required by Excel)
    let mut ts = Vec::new();
    ts.extend_from_slice(&0u32.to_le_bytes()); // count=0
    ts.extend_from_slice(&encode_wide_str("TableStyleMedium2")); // defaultTableStyle
    ts.extend_from_slice(&encode_wide_str("PivotStyleLight16")); // defaultPivotStyle
    rw.write_record(0x01FC, &ts)?; // BrtBeginTableStyles
    rw.write_record(0x01FD, &[])?; // BrtEndTableStyles

    rw.write_record(0x0117, &[])?;

    drop(rw);
    zip.write_all(&buf)?;

    let dxf_mapping = build_dxf_mapping(workbook, &dxf_styles);
    Ok((mapping, extra_font_ids, dxf_mapping))
}

pub(crate) struct StyleTable {
    pub(crate) fonts: Vec<FontStyle>,
    pub(crate) fills: Vec<FillStyle>,
    pub(crate) borders: Vec<BorderStyle>,
    pub(crate) numfmts: Vec<(u16, String)>,
    pub(crate) xfs: Vec<XfEntry>,
}

pub(crate) struct XfEntry {
    pub(crate) font_id: u16,
    pub(crate) fill_id: u16,
    pub(crate) border_id: u16,
    pub(crate) num_fmt_id: u16,
    pub(crate) alignment: Alignment,
    pub(crate) protection: Protection,
}

pub(crate) fn build_style_table(
    workbook: &Workbook,
    extra_fonts: &[FontStyle],
) -> (StyleTable, StyleMapping, Vec<u16>) {
    let mut fonts: Vec<FontStyle> = vec![FontStyle::default()];
    let mut font_map: HashMap<FontKey, u16> = HashMap::new();
    font_map.insert(FontKey::new(&fonts[0]), 0);

    let mut fills: Vec<FillStyle> = vec![FillStyle::None, FillStyle::None];
    let mut fill_map: HashMap<FillKey, u16> = HashMap::new();
    fill_map.insert(FillKey::new(&FillStyle::None), 0);

    let mut borders: Vec<BorderStyle> = vec![BorderStyle::default()];
    let mut border_map: HashMap<BorderKey, u16> = HashMap::new();
    border_map.insert(BorderKey::new(&borders[0]), 0);

    let mut numfmts: Vec<(u16, String)> = Vec::new();
    let mut numfmt_map: HashMap<String, u16> = HashMap::new();
    let mut next_numfmt_id: u16 = 164;

    let mut xfs: Vec<XfEntry> = vec![XfEntry {
        font_id: 0,
        fill_id: 0,
        border_id: 0,
        num_fmt_id: 0,
        alignment: Alignment::default(),
        protection: Protection::new(),
    }];
    let mut style_to_xf: HashMap<Style, u32> = HashMap::new();
    style_to_xf.insert(Style::default(), 0);

    let mut sheet_maps: Vec<HashMap<u32, u32>> = Vec::with_capacity(workbook.sheet_count());

    for sheet in workbook.worksheets() {
        let mut map: HashMap<u32, u32> = HashMap::new();
        map.insert(0, 0);

        for (_row, _col, cell) in sheet.iter_cells() {
            let local_idx = cell.style_index;
            if local_idx == 0 || map.contains_key(&local_idx) {
                continue;
            }

            let style = sheet.style_by_index(local_idx).cloned().unwrap_or_default();

            let xf_id = match style_to_xf.get(&style) {
                Some(&id) => id,
                None => {
                    let font_id = intern_font(&style.font, &mut fonts, &mut font_map);
                    let fill_id = intern_fill(&style.fill, &mut fills, &mut fill_map);
                    let border_id = intern_border(&style.border, &mut borders, &mut border_map);
                    let num_fmt_id = intern_numfmt(
                        &style.number_format,
                        &mut numfmts,
                        &mut numfmt_map,
                        &mut next_numfmt_id,
                    );

                    let id = xfs.len() as u32;
                    xfs.push(XfEntry {
                        font_id,
                        fill_id,
                        border_id,
                        num_fmt_id,
                        alignment: style.alignment.clone(),
                        protection: style.protection,
                    });
                    style_to_xf.insert(style, id);
                    id
                }
            };

            map.insert(local_idx, xf_id);
        }

        sheet_maps.push(map);
    }

    let extra_font_ids: Vec<u16> = extra_fonts
        .iter()
        .map(|f| intern_font(f, &mut fonts, &mut font_map))
        .collect();

    for sheet in workbook.worksheets() {
        for pivot in sheet.pivot_tables() {
            for measure in &pivot.measures {
                let Some(code) = measure.number_format.as_deref() else {
                    continue;
                };
                if is_builtin_number_format_code(code) {
                    continue;
                }
                intern_custom_numfmt_code(code, &mut numfmts, &mut numfmt_map, &mut next_numfmt_id);
            }
        }
    }

    #[cfg(test)]
    let xf_count = xfs.len() as u32;
    let table = StyleTable {
        fonts,
        fills,
        borders,
        numfmts,
        xfs,
    };
    #[cfg(not(test))]
    let mapping = StyleMapping {
        sheet_maps,
        custom_numfmt_ids: numfmt_map,
    };
    #[cfg(test)]
    let mapping = StyleMapping {
        sheet_maps,
        custom_numfmt_ids: numfmt_map,
        xf_count,
    };
    (table, mapping, extra_font_ids)
}

fn intern_font(
    font: &FontStyle,
    fonts: &mut Vec<FontStyle>,
    map: &mut HashMap<FontKey, u16>,
) -> u16 {
    let key = FontKey::new(font);
    if let Some(&id) = map.get(&key) {
        return id;
    }
    let id = fonts.len() as u16;
    fonts.push(font.clone());
    map.insert(key, id);
    id
}

fn intern_fill(
    fill: &FillStyle,
    fills: &mut Vec<FillStyle>,
    map: &mut HashMap<FillKey, u16>,
) -> u16 {
    let key = FillKey::new(fill);
    if let Some(&id) = map.get(&key) {
        return id;
    }
    let id = fills.len() as u16;
    fills.push(fill.clone());
    map.insert(key, id);
    id
}

fn intern_border(
    border: &BorderStyle,
    borders: &mut Vec<BorderStyle>,
    map: &mut HashMap<BorderKey, u16>,
) -> u16 {
    let key = BorderKey::new(border);
    if let Some(&id) = map.get(&key) {
        return id;
    }
    let id = borders.len() as u16;
    borders.push(border.clone());
    map.insert(key, id);
    id
}

fn intern_numfmt(
    nf: &NumberFormat,
    numfmts: &mut Vec<(u16, String)>,
    map: &mut HashMap<String, u16>,
    next_id: &mut u16,
) -> u16 {
    match nf {
        NumberFormat::General => 0,
        NumberFormat::BuiltIn(id) => *id as u16,
        NumberFormat::Custom(code) => intern_custom_numfmt_code(code, numfmts, map, next_id),
    }
}

fn intern_custom_numfmt_code(
    code: &str,
    numfmts: &mut Vec<(u16, String)>,
    map: &mut HashMap<String, u16>,
    next_id: &mut u16,
) -> u16 {
    if let Some(&id) = map.get(code) {
        return id;
    }
    let id = *next_id;
    *next_id += 1;
    numfmts.push((id, code.to_string()));
    map.insert(code.to_string(), id);
    id
}

fn is_builtin_number_format_code(code: &str) -> bool {
    code.eq_ignore_ascii_case("General")
        || (1..=49).any(|id| NumberFormat::BuiltIn(id).format_string() == code)
}

#[derive(PartialEq, Eq, Hash)]
struct FontKey(u64);
impl FontKey {
    fn new(f: &FontStyle) -> Self {
        use std::hash::{Hash, Hasher};
        let mut h = std::collections::hash_map::DefaultHasher::new();
        f.hash(&mut h);
        FontKey(h.finish())
    }
}

#[derive(PartialEq, Eq, Hash)]
struct FillKey(u64);
impl FillKey {
    fn new(f: &FillStyle) -> Self {
        use std::hash::{Hash, Hasher};
        let mut h = std::collections::hash_map::DefaultHasher::new();
        f.hash(&mut h);
        FillKey(h.finish())
    }
}

#[derive(PartialEq, Eq, Hash)]
struct BorderKey(u64);
impl BorderKey {
    fn new(b: &BorderStyle) -> Self {
        use std::hash::{Hash, Hasher};
        let mut h = std::collections::hash_map::DefaultHasher::new();
        b.hash(&mut h);
        BorderKey(h.finish())
    }
}

fn write_numfmts<W: Write>(rw: &mut RecordWriter<W>, table: &StyleTable) -> std::io::Result<()> {
    if table.numfmts.is_empty() {
        return Ok(());
    }
    rw.write_record(
        records::BRT_BEGIN_FMTS,
        &(table.numfmts.len() as u32).to_le_bytes(),
    )?;
    for (id, code) in &table.numfmts {
        let mut payload = Vec::new();
        payload.extend_from_slice(&id.to_le_bytes());
        payload.extend_from_slice(&encode_wide_str(code));
        rw.write_record(records::BRT_FMT, &payload)?;
    }
    rw.write_record(records::BRT_END_FMTS, &[])?;
    Ok(())
}

fn write_fonts<W: Write>(rw: &mut RecordWriter<W>, table: &StyleTable) -> std::io::Result<()> {
    rw.write_record(
        records::BRT_BEGIN_FONTS,
        &(table.fonts.len() as u32).to_le_bytes(),
    )?;
    for font in &table.fonts {
        rw.write_record(records::BRT_FONT, &encode_font(font))?;
    }
    rw.write_record(records::BRT_END_FONTS, &[])?;
    Ok(())
}

fn write_fills<W: Write>(rw: &mut RecordWriter<W>, table: &StyleTable) -> std::io::Result<()> {
    rw.write_record(
        records::BRT_BEGIN_FILLS,
        &(table.fills.len() as u32).to_le_bytes(),
    )?;
    for (i, fill) in table.fills.iter().enumerate() {
        if i < 2 {
            rw.write_record(records::BRT_FILL, &default_fill_payload(i))?;
        } else {
            rw.write_record(records::BRT_FILL, &encode_fill(fill))?;
        }
    }
    rw.write_record(records::BRT_END_FILLS, &[])?;
    Ok(())
}

fn default_fill_payload(index: usize) -> Vec<u8> {
    let mut payload = vec![0u8; 68];
    let fls: u32 = if index == 1 { 17 } else { 0 };
    payload[0..4].copy_from_slice(&fls.to_le_bytes());
    // fgColor: theme 64 (text foreground), A=0xFF
    payload[4] = 3;
    payload[5] = 64;
    payload[11] = 0xFF;
    // bgColor: theme 65 (background), RGBA=0xFFFFFFFF
    payload[12] = 3;
    payload[13] = 65;
    payload[16] = 0xFF;
    payload[17] = 0xFF;
    payload[18] = 0xFF;
    payload[19] = 0xFF;
    payload
}

fn write_borders<W: Write>(rw: &mut RecordWriter<W>, table: &StyleTable) -> std::io::Result<()> {
    rw.write_record(
        records::BRT_BEGIN_BORDERS,
        &(table.borders.len() as u32).to_le_bytes(),
    )?;
    for border in &table.borders {
        rw.write_record(records::BRT_BORDER, &encode_border(border))?;
    }
    rw.write_record(records::BRT_END_BORDERS, &[])?;
    Ok(())
}

fn write_cell_style_xfs<W: Write>(rw: &mut RecordWriter<W>) -> std::io::Result<()> {
    rw.write_record(records::BRT_BEGIN_CELL_STYLE_XFS, &1u32.to_le_bytes())?;
    let mut xf = vec![0u8; 16];
    xf[0..2].copy_from_slice(&0xFFFFu16.to_le_bytes());
    xf[12] = 2 << 3; // alcV = Bottom
    xf[13] = 0x10; // fLocked
    rw.write_record(records::BRT_XF, &xf)?;
    rw.write_record(records::BRT_END_CELL_STYLE_XFS, &[])?;
    Ok(())
}

fn write_cell_xfs<W: Write>(rw: &mut RecordWriter<W>, table: &StyleTable) -> std::io::Result<()> {
    rw.write_record(
        records::BRT_BEGIN_CELL_XFS,
        &(table.xfs.len() as u32).to_le_bytes(),
    )?;
    for xf in &table.xfs {
        rw.write_record(records::BRT_XF, &encode_xf(xf))?;
    }
    rw.write_record(records::BRT_END_CELL_XFS, &[])?;
    Ok(())
}

fn write_named_styles<W: Write>(rw: &mut RecordWriter<W>) -> std::io::Result<()> {
    rw.write_record(records::BRT_BEGIN_STYLES, &1u32.to_le_bytes())?;
    let mut style = Vec::new();
    style.extend_from_slice(&0u32.to_le_bytes()); // ixf
    style.extend_from_slice(&1u32.to_le_bytes()); // builtinId
    style.extend_from_slice(&encode_wide_str("Normal"));
    rw.write_record(records::BRT_STYLE, &style)?;
    rw.write_record(records::BRT_END_STYLES, &[])?;
    Ok(())
}

pub(crate) struct DxfMapping {
    sheet_rule_dxf_ids: Vec<Vec<Option<u32>>>,
    color_filter_dxf_id: Option<u32>,
}

impl DxfMapping {
    pub fn dxf_id_for_rule(&self, sheet_index: usize, rule_index: usize) -> Option<u32> {
        self.sheet_rule_dxf_ids
            .get(sheet_index)
            .and_then(|rules| rules.get(rule_index))
            .and_then(|id| *id)
    }

    /// DXF entry backing color autofilters. `Some` whenever any sheet
    /// carries a `ColumnFilter::Color`.
    pub fn color_filter_dxf_id(&self) -> Option<u32> {
        self.color_filter_dxf_id
    }
}

/// Style synthesized for color autofilters. BrtColorFilter's dxfid
/// MUST reference a real BrtDXF entry (Excel refuses to open the file
/// otherwise), but our `ColorFilter` model carries only the dxf index,
/// not the color, so back every color filter with one solid-fill DXF
/// the way Excel pairs a color filter with a fill DXF.
fn color_filter_dxf_style() -> Style {
    let mut s = Style::default();
    s.fill = duke_sheets_core::style::FillStyle::Solid {
        color: duke_sheets_core::style::Color::Rgb { r: 255, g: 0, b: 0 },
    };
    s
}

fn workbook_has_color_filter(workbook: &Workbook) -> bool {
    use duke_sheets_core::auto_filter::ColumnFilter;
    (0..workbook.sheet_count()).any(|i| {
        workbook
            .worksheet(i)
            .unwrap()
            .auto_filter()
            .is_some_and(|af| {
                af.filter_columns
                    .iter()
                    .any(|fc| matches!(fc.filter, ColumnFilter::Color(_)))
            })
    })
}

fn collect_dxf_styles(workbook: &Workbook) -> Vec<Style> {
    let mut dxf_styles: Vec<Style> = Vec::new();
    let mut seen: HashMap<u64, usize> = HashMap::new();

    for i in 0..workbook.sheet_count() {
        let ws = workbook.worksheet(i).unwrap();
        for rule in ws.conditional_formats() {
            if let Some(ref fmt) = rule.format {
                let key = style_hash(fmt);
                if !seen.contains_key(&key) {
                    seen.insert(key, dxf_styles.len());
                    dxf_styles.push(fmt.clone());
                }
            }
        }
    }
    if workbook_has_color_filter(workbook) {
        let style = color_filter_dxf_style();
        let key = style_hash(&style);
        if !seen.contains_key(&key) {
            seen.insert(key, dxf_styles.len());
            dxf_styles.push(style);
        }
    }
    dxf_styles
}

fn build_dxf_mapping(workbook: &Workbook, dxf_styles: &[Style]) -> DxfMapping {
    let mut style_to_idx: HashMap<u64, u32> = HashMap::new();
    for (i, s) in dxf_styles.iter().enumerate() {
        style_to_idx.insert(style_hash(s), i as u32);
    }

    let mut sheet_rule_dxf_ids = Vec::with_capacity(workbook.sheet_count());
    for i in 0..workbook.sheet_count() {
        let ws = workbook.worksheet(i).unwrap();
        let rules = ws.conditional_formats();
        let mut rule_ids = Vec::with_capacity(rules.len());
        for rule in rules {
            if let Some(ref fmt) = rule.format {
                let key = style_hash(fmt);
                rule_ids.push(style_to_idx.get(&key).copied());
            } else {
                rule_ids.push(None);
            }
        }
        sheet_rule_dxf_ids.push(rule_ids);
    }

    let color_filter_dxf_id = if workbook_has_color_filter(workbook) {
        style_to_idx
            .get(&style_hash(&color_filter_dxf_style()))
            .copied()
    } else {
        None
    };

    DxfMapping {
        sheet_rule_dxf_ids,
        color_filter_dxf_id,
    }
}

fn style_hash(s: &Style) -> u64 {
    use std::hash::{Hash, Hasher};
    let mut h = std::collections::hash_map::DefaultHasher::new();
    s.hash(&mut h);
    h.finish()
}

fn write_dxfs<W: Write>(rw: &mut RecordWriter<W>, dxf_styles: &[Style]) -> std::io::Result<()> {
    rw.write_record(0x01F9, &(dxf_styles.len() as u32).to_le_bytes())?;
    for style in dxf_styles {
        rw.write_record(0x01FB, &encode_dxf(style))?;
    }
    rw.write_record(0x01FA, &[])?;
    Ok(())
}

fn encode_dxf(style: &Style) -> Vec<u8> {
    let default_font = FontStyle::default();
    let default_fill = FillStyle::None;

    let mut props: Vec<Vec<u8>> = Vec::new();

    if style.fill != default_fill {
        if let Some(pat) = fill_pattern_byte(&style.fill) {
            props.push(encode_xfprop(0x0000, &[pat]));
        }
        if let Some(fg) = fill_fg_color(&style.fill) {
            props.push(encode_xfprop(0x0001, &encode_xfprop_color(&fg)));
        }
        props.push(encode_xfprop(0x0002, &encode_xfprop_color(&Color::Auto)));
    }

    if style.font.color != default_font.color {
        props.push(encode_xfprop(
            0x0005,
            &encode_xfprop_color(&style.font.color),
        ));
    }
    if style.font.bold != default_font.bold {
        let bls: u16 = if style.font.bold { 700 } else { 400 };
        props.push(encode_xfprop(0x0019, &bls.to_le_bytes()));
    }
    if style.font.italic != default_font.italic {
        props.push(encode_xfprop(
            0x001C,
            &[if style.font.italic { 1 } else { 0 }],
        ));
    }
    if style.font.strikethrough != default_font.strikethrough {
        props.push(encode_xfprop(
            0x001D,
            &[if style.font.strikethrough { 1 } else { 0 }],
        ));
    }

    let mut payload = Vec::new();
    payload.extend_from_slice(&0x8000u16.to_le_bytes()); // fNewBorder=1
    payload.extend_from_slice(&0u16.to_le_bytes()); // XFProps reserved
    payload.extend_from_slice(&(props.len() as u16).to_le_bytes()); // cprops
    for prop in &props {
        payload.extend_from_slice(prop);
    }
    payload
}

fn encode_xfprop(prop_type: u16, data: &[u8]) -> Vec<u8> {
    let cb = (4 + data.len()) as u16;
    let mut buf = Vec::with_capacity(cb as usize);
    buf.extend_from_slice(&prop_type.to_le_bytes());
    buf.extend_from_slice(&cb.to_le_bytes());
    buf.extend_from_slice(data);
    buf
}

fn encode_xfprop_color(color: &Color) -> [u8; 8] {
    let mut buf = [0u8; 8];
    match color {
        Color::Auto => {
            buf[0] = 0x00;
        }
        Color::Indexed(idx) => {
            buf[0] = 0x03; // fValidRGBA=1, xclrType=1 (palette)
            buf[1] = *idx;
        }
        Color::Rgb { r, g, b } => {
            buf[0] = 0x05; // fValidRGBA=1, xclrType=2 (RGBA)
            buf[1] = 0xFF;
            buf[4] = *r;
            buf[5] = *g;
            buf[6] = *b;
            buf[7] = 0xFF;
        }
        Color::Theme { index, tint } => {
            buf[0] = 0x07; // fValidRGBA=1, xclrType=3 (theme)
            buf[1] = *index;
            let t = (tint.clamp(-1.0, 1.0) * 32767.0).round() as i16;
            buf[2..4].copy_from_slice(&t.to_le_bytes());
        }
        _ => {
            buf[0] = 0x00;
        }
    }
    buf
}

fn fill_pattern_byte(fill: &FillStyle) -> Option<u8> {
    match fill {
        FillStyle::Solid { .. } => Some(1),
        FillStyle::Pattern { pattern, .. } => Some(pattern_to_fls(pattern) as u8),
        _ => None,
    }
}

fn fill_fg_color(fill: &FillStyle) -> Option<Color> {
    match fill {
        FillStyle::Solid { color } => Some(color.clone()),
        FillStyle::Pattern { foreground, .. } => Some(foreground.clone()),
        _ => None,
    }
}

fn encode_brt_color(color: &Color) -> [u8; 8] {
    let mut buf = [0u8; 8];
    match color {
        Color::Auto => {
            buf[0] = 1; // indexed
            buf[1] = 64; // automatic = indexed color 64
        }
        Color::Indexed(idx) => {
            buf[0] = 1;
            buf[1] = *idx;
        }
        Color::Rgb { r, g, b } => {
            buf[0] = 2;
            buf[4] = *r;
            buf[5] = *g;
            buf[6] = *b;
            buf[7] = 0xFF;
        }
        Color::Argb { a, r, g, b } => {
            buf[0] = 2;
            buf[4] = *r;
            buf[5] = *g;
            buf[6] = *b;
            buf[7] = *a;
        }
        Color::Theme { index, tint } => {
            buf[0] = 3;
            buf[1] = *index;
            let tint_i16 = (tint.clamp(-1.0, 1.0) * 32767.0).round() as i16;
            buf[2..4].copy_from_slice(&tint_i16.to_le_bytes());
        }
    }
    buf
}

fn encode_font(font: &FontStyle) -> Vec<u8> {
    let mut payload = vec![0u8; 21];

    let dy_height = (font.size * 20.0).round() as u16;
    payload[0..2].copy_from_slice(&dy_height.to_le_bytes());

    let mut grbit: u16 = 0;
    if font.italic {
        grbit |= 0x02;
    }
    if font.strikethrough {
        grbit |= 0x08;
    }
    payload[2..4].copy_from_slice(&grbit.to_le_bytes());

    let bls: u16 = if font.bold { 700 } else { 400 };
    payload[4..6].copy_from_slice(&bls.to_le_bytes());

    let sss: u16 = match font.vertical_align {
        FontVerticalAlign::Baseline => 0,
        FontVerticalAlign::Superscript => 1,
        FontVerticalAlign::Subscript => 2,
    };
    payload[6..8].copy_from_slice(&sss.to_le_bytes());

    payload[8] = match font.underline {
        Underline::None => 0,
        Underline::Single => 1,
        Underline::Double => 2,
        Underline::SingleAccounting => 0x21,
        Underline::DoubleAccounting => 0x22,
    };

    payload[9] = font.family.unwrap_or(0);
    payload[10] = font.charset.unwrap_or(0);

    let color_bytes = encode_brt_color(&font.color);
    payload[12..20].copy_from_slice(&color_bytes);

    payload[20] = match font.scheme.as_deref() {
        Some("major") => 1,
        Some("minor") => 2,
        _ => 0,
    };

    payload.extend_from_slice(&encode_wide_str(&font.name));

    payload
}

fn encode_fill(fill: &FillStyle) -> Vec<u8> {
    // Full BrtFill: fls(4) + fgColor(8) + bgColor(8) + gradient fields(48) = 68 bytes
    let mut payload = vec![0u8; 68];
    match fill {
        FillStyle::None => {}
        FillStyle::Solid { color } => {
            payload[0..4].copy_from_slice(&1u32.to_le_bytes());
            let color_bytes = encode_brt_color(color);
            payload[4..12].copy_from_slice(&color_bytes);
        }
        FillStyle::Pattern {
            pattern,
            foreground,
            background,
        } => {
            let fls = pattern_to_fls(pattern);
            payload[0..4].copy_from_slice(&fls.to_le_bytes());
            let fg = encode_brt_color(foreground);
            payload[4..12].copy_from_slice(&fg);
            let bg = encode_brt_color(background);
            payload[12..20].copy_from_slice(&bg);
        }
        FillStyle::Gradient {
            gradient_type,
            angle,
            stops,
        } => {
            payload[0..4].copy_from_slice(&40u32.to_le_bytes()); // fls = 40 (gradient)
            let gt: u32 = match gradient_type {
                GradientType::Linear => 0,
                GradientType::Path => 1,
            };
            payload[20..24].copy_from_slice(&gt.to_le_bytes());
            payload[24..32].copy_from_slice(&angle.to_le_bytes());
            payload[64..68].copy_from_slice(&(stops.len() as u32).to_le_bytes());
            for stop in stops {
                payload.extend_from_slice(&stop.position.to_le_bytes());
                payload.extend_from_slice(&encode_brt_color(&stop.color));
            }
        }
    }
    payload
}

fn pattern_to_fls(p: &PatternType) -> u32 {
    match p {
        PatternType::None => 0,
        PatternType::Solid => 1,
        PatternType::MediumGray => 2,
        PatternType::DarkGray => 3,
        PatternType::LightGray => 4,
        PatternType::DarkHorizontal => 5,
        PatternType::DarkVertical => 6,
        PatternType::DarkDown => 7,
        PatternType::DarkUp => 8,
        PatternType::DarkGrid => 9,
        PatternType::DarkTrellis => 10,
        PatternType::LightHorizontal => 11,
        PatternType::LightVertical => 12,
        PatternType::LightDown => 13,
        PatternType::LightUp => 14,
        PatternType::LightGrid => 15,
        PatternType::LightTrellis => 16,
        PatternType::Gray125 => 17,
        PatternType::Gray0625 => 18,
    }
}

fn encode_border(border: &BorderStyle) -> Vec<u8> {
    let edge_size = 10;
    let mut payload = vec![0u8; 1 + edge_size * 5];

    let mut flags: u8 = 0;
    match border.diagonal_direction {
        DiagonalDirection::None => {}
        DiagonalDirection::Down => flags |= 0x01,
        DiagonalDirection::Up => flags |= 0x02,
        DiagonalDirection::Both => flags |= 0x03,
    }
    payload[0] = flags;

    let base = 1;
    encode_border_edge(&border.top, &mut payload[base..base + edge_size]);
    encode_border_edge(
        &border.bottom,
        &mut payload[base + edge_size..base + edge_size * 2],
    );
    encode_border_edge(
        &border.left,
        &mut payload[base + edge_size * 2..base + edge_size * 3],
    );
    encode_border_edge(
        &border.right,
        &mut payload[base + edge_size * 3..base + edge_size * 4],
    );
    encode_border_edge(
        &border.diagonal,
        &mut payload[base + edge_size * 4..base + edge_size * 5],
    );

    payload
}

fn encode_border_edge(edge: &Option<BorderEdge>, buf: &mut [u8]) {
    match edge {
        Some(e) => {
            let style_u16 = border_line_style_byte(&e.style) as u16;
            buf[0..2].copy_from_slice(&style_u16.to_le_bytes());
            let color = encode_brt_color(&e.color);
            buf[2..10].copy_from_slice(&color);
        }
        None => {
            buf[2] = 1; // indexed color type
        }
    }
}

fn border_line_style_byte(style: &BorderLineStyle) -> u8 {
    match style {
        BorderLineStyle::None => 0,
        BorderLineStyle::Thin => 1,
        BorderLineStyle::Medium => 2,
        BorderLineStyle::Dashed => 3,
        BorderLineStyle::Dotted => 4,
        BorderLineStyle::Thick => 5,
        BorderLineStyle::Double => 6,
        BorderLineStyle::Hair => 7,
        BorderLineStyle::MediumDashed => 8,
        BorderLineStyle::DashDot => 9,
        BorderLineStyle::MediumDashDot => 10,
        BorderLineStyle::DashDotDot => 11,
        BorderLineStyle::MediumDashDotDot => 12,
        BorderLineStyle::SlantDashDot => 13,
    }
}

fn encode_xf(xf: &XfEntry) -> Vec<u8> {
    let mut payload = vec![0u8; 16];

    payload[0..2].copy_from_slice(&0u16.to_le_bytes());
    payload[2..4].copy_from_slice(&xf.num_fmt_id.to_le_bytes());
    payload[4..6].copy_from_slice(&xf.font_id.to_le_bytes());
    payload[6..8].copy_from_slice(&xf.fill_id.to_le_bytes());
    payload[8..10].copy_from_slice(&xf.border_id.to_le_bytes());

    let trot: u8 = if xf.alignment.rotation == 255 {
        255
    } else if xf.alignment.rotation >= 0 {
        xf.alignment.rotation as u8
    } else {
        (90 + (-xf.alignment.rotation)) as u8
    };
    payload[10] = trot;

    payload[11] = xf.alignment.indent & 0x0F;

    let alc: u8 = match xf.alignment.horizontal {
        HorizontalAlignment::General => 0,
        HorizontalAlignment::Left => 1,
        HorizontalAlignment::Center => 2,
        HorizontalAlignment::Right => 3,
        HorizontalAlignment::Fill => 4,
        HorizontalAlignment::Justify => 5,
        HorizontalAlignment::CenterContinuous => 6,
        HorizontalAlignment::Distributed => 7,
    };
    let alcv: u8 = match xf.alignment.vertical {
        VerticalAlignment::Top => 0,
        VerticalAlignment::Center => 1,
        VerticalAlignment::Bottom => 2,
        VerticalAlignment::Justify => 3,
        VerticalAlignment::Distributed => 4,
    };
    let f_wrap: u8 = if xf.alignment.wrap_text { 1 } else { 0 };
    payload[12] = alc | (alcv << 3) | (f_wrap << 6);

    let f_shrink: u8 = if xf.alignment.shrink_to_fit { 1 } else { 0 };
    let ro: u8 = match xf.alignment.reading_order {
        ReadingOrder::ContextDependent => 0,
        ReadingOrder::LeftToRight => 1,
        ReadingOrder::RightToLeft => 2,
    };
    let f_locked: u8 = if xf.protection.locked { 1 } else { 0 };
    let f_hidden: u8 = if xf.protection.hidden { 1 } else { 0 };
    payload[13] = f_shrink | (ro << 2) | (f_locked << 4) | (f_hidden << 5);

    let base_prot = Protection::new();
    let mut atr: u8 = 0;
    if xf.num_fmt_id != 0 {
        atr |= 0x01;
    }
    if xf.font_id != 0 {
        atr |= 0x02;
    }
    if xf.alignment != Alignment::default() {
        atr |= 0x04;
    }
    if xf.border_id != 0 {
        atr |= 0x08;
    }
    if xf.fill_id != 0 {
        atr |= 0x10;
    }
    if xf.protection != base_prot {
        atr |= 0x20;
    }
    payload[14] = atr;

    payload
}
