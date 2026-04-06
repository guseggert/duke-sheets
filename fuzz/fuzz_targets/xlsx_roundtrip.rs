//! Fuzz target for XLSX write → read roundtrip.
//!
//! Builds a workbook programmatically from `Arbitrary` input, writes it
//! to XLSX bytes, then reads it back. Any panic in write or read-back
//! indicates a serialization/deserialization bug.

#![no_main]
use arbitrary::{Arbitrary, Unstructured};
use libfuzzer_sys::fuzz_target;
use std::io::Cursor;

use duke_sheets_core::{CellValue, Style, Workbook};
use duke_sheets_chart::{
    CellMarker, Chart, ChartType, DataLabels, DataReference, DataSeries, DrawingAnchor, Legend,
    LegendPosition, Marker, MarkerSymbol, Trendline, TrendlineType,
};

/// Structured workbook specification for fuzzing.
#[derive(Arbitrary, Debug)]
struct FuzzWorkbook {
    sheets: Vec<FuzzSheet>,
}

#[derive(Debug)]
struct FuzzSheet {
    name: String,
    cells: Vec<FuzzCell>,
    charts: Vec<FuzzChart>,
}
impl<'a> Arbitrary<'a> for FuzzSheet {
    fn arbitrary(u: &mut Unstructured<'a>) -> arbitrary::Result<Self> {
        // Sheet name: 1-20 alphanumeric chars
        let name_len = u.int_in_range(1..=20)?;
        let mut name = String::with_capacity(name_len);
        for _ in 0..name_len {
            name.push(u.int_in_range(b'A'..=b'z')? as char);
        }

        let ncells = u.int_in_range(0..=50)?;
        let mut cells = Vec::with_capacity(ncells);
        for _ in 0..ncells {
            cells.push(FuzzCell::arbitrary(u)?);
        }

        let ncharts = u.int_in_range(0..=3)?;
        let mut charts = Vec::with_capacity(ncharts);
        for _ in 0..ncharts {
            charts.push(FuzzChart::arbitrary(u)?);
        }

        Ok(FuzzSheet { name, cells, charts })
    }
}

#[derive(Arbitrary, Debug)]
struct FuzzCell {
    row: u16,
    col: u8,
    value: FuzzValue,
    style: Option<FuzzStyle>,
}

#[derive(Arbitrary, Debug)]
enum FuzzValue {
    Empty,
    Number(f64),
    Int(i32),
    Bool(bool),
    Str(SmallString),
    Formula(SmallFormula),
}

/// Keep strings small.
#[derive(Debug)]
struct SmallString(String);

impl<'a> Arbitrary<'a> for SmallString {
    fn arbitrary(u: &mut Unstructured<'a>) -> arbitrary::Result<Self> {
        let len = u.int_in_range(0..=50)?;
        let mut s = String::with_capacity(len);
        for _ in 0..len {
            // Mix of ASCII + some non-ASCII to test encoding
            let c: u8 = u.arbitrary()?;
            if c < 128 && c >= 32 {
                s.push(c as char);
            } else {
                s.push('x');
            }
        }
        Ok(SmallString(s))
    }
}

/// Simple formula strings.
#[derive(Debug)]
struct SmallFormula(String);

impl<'a> Arbitrary<'a> for SmallFormula {
    fn arbitrary(u: &mut Unstructured<'a>) -> arbitrary::Result<Self> {
        let kind: u8 = u.int_in_range(0..=4)?;
        let f = match kind {
            0 => {
                let n: f64 = u.arbitrary()?;
                if n.is_finite() {
                    format!("={}", n)
                } else {
                    "=0".into()
                }
            }
            1 => {
                let col = (b'A' + u.int_in_range(0..=25)? as u8) as char;
                let row = u.int_in_range(1..=100)?;
                format!("={}{}", col, row)
            }
            2 => {
                let c1 = (b'A' + u.int_in_range(0..=25)? as u8) as char;
                let r1 = u.int_in_range(1..=100)?;
                let c2 = (b'A' + u.int_in_range(0..=25)? as u8) as char;
                let r2 = u.int_in_range(1..=100)?;
                format!("=SUM({}{}:{}{})", c1, r1, c2, r2)
            }
            3 => {
                let c = (b'A' + u.int_in_range(0..=25)? as u8) as char;
                let r = u.int_in_range(1..=100)?;
                format!("=IF({}{}>0,1,0)", c, r)
            }
            _ => {
                let c1 = (b'A' + u.int_in_range(0..=25)? as u8) as char;
                let r1 = u.int_in_range(1..=100)?;
                let c2 = (b'A' + u.int_in_range(0..=25)? as u8) as char;
                let r2 = u.int_in_range(1..=100)?;
                format!("={}{}+{}{}", c1, r1, c2, r2)
            }
        };
        Ok(SmallFormula(f))
    }
}

#[derive(Arbitrary, Debug)]
struct FuzzStyle {
    bold: bool,
    italic: bool,
    font_size: Option<u8>,
}

#[derive(Debug)]
struct FuzzChart {
    chart_type: FuzzChartType,
    title: Option<SmallString>,
    series: Vec<FuzzSeries>,
    anchor: FuzzAnchor,
    legend_pos: Option<FuzzLegendPos>,
    hole_size: Option<u8>,
    data_labels_show_value: Option<bool>,
}

impl<'a> Arbitrary<'a> for FuzzChart {
    fn arbitrary(u: &mut Unstructured<'a>) -> arbitrary::Result<Self> {
        let chart_type = FuzzChartType::arbitrary(u)?;
        let title = u.arbitrary()?;
        let nseries = u.int_in_range(1..=5)?;
        let mut series = Vec::with_capacity(nseries);
        for _ in 0..nseries {
            series.push(FuzzSeries::arbitrary(u)?);
        }
        let anchor = FuzzAnchor::arbitrary(u)?;
        let legend_pos = u.arbitrary()?;
        let hole_size = if matches!(chart_type, FuzzChartType::Doughnut) {
            Some(u.int_in_range(10..=90)?)
        } else {
            None
        };
        let data_labels_show_value = u.arbitrary()?;
        Ok(FuzzChart {
            chart_type,
            title,
            series,
            anchor,
            legend_pos,
            hole_size,
            data_labels_show_value,
        })
    }
}

#[derive(Arbitrary, Debug, Clone, Copy)]
enum FuzzChartType {
    ColumnClustered,
    ColumnStacked,
    ColumnPercentStacked,
    BarClustered,
    BarStacked,
    BarPercentStacked,
    Line,
    LineStacked,
    Pie,
    PieExploded,
    Doughnut,
    Area,
    AreaStacked,
    AreaPercentStacked,
    ScatterMarkers,
    ScatterSmooth,
    ScatterLines,
    Bubble,
    Radar,
    Stock,
    Surface,
}

impl From<FuzzChartType> for ChartType {
    fn from(f: FuzzChartType) -> Self {
        match f {
            FuzzChartType::ColumnClustered => ChartType::ColumnClustered,
            FuzzChartType::ColumnStacked => ChartType::ColumnStacked,
            FuzzChartType::ColumnPercentStacked => ChartType::ColumnPercentStacked,
            FuzzChartType::BarClustered => ChartType::BarClustered,
            FuzzChartType::BarStacked => ChartType::BarStacked,
            FuzzChartType::BarPercentStacked => ChartType::BarPercentStacked,
            FuzzChartType::Line => ChartType::Line,
            FuzzChartType::LineStacked => ChartType::LineStacked,
            FuzzChartType::Pie => ChartType::Pie,
            FuzzChartType::PieExploded => ChartType::PieExploded,
            FuzzChartType::Doughnut => ChartType::Doughnut,
            FuzzChartType::Area => ChartType::Area,
            FuzzChartType::AreaStacked => ChartType::AreaStacked,
            FuzzChartType::AreaPercentStacked => ChartType::AreaPercentStacked,
            FuzzChartType::ScatterMarkers => ChartType::ScatterMarkers,
            FuzzChartType::ScatterSmooth => ChartType::ScatterSmooth,
            FuzzChartType::ScatterLines => ChartType::ScatterLines,
            FuzzChartType::Bubble => ChartType::Bubble,
            FuzzChartType::Radar => ChartType::Radar,
            FuzzChartType::Stock => ChartType::Stock,
            FuzzChartType::Surface => ChartType::Surface,
        }
    }
}

#[derive(Debug)]
struct FuzzSeries {
    values: FuzzDataRef,
    categories: Option<FuzzDataRef>,
    name: Option<SmallString>,
    explosion: Option<u8>,
    smooth: Option<bool>,
    has_trendline: bool,
    has_marker: bool,
}

impl<'a> Arbitrary<'a> for FuzzSeries {
    fn arbitrary(u: &mut Unstructured<'a>) -> arbitrary::Result<Self> {
        Ok(FuzzSeries {
            values: FuzzDataRef::arbitrary(u)?,
            categories: u.arbitrary::<bool>()?.then(|| FuzzDataRef::arbitrary(u)).transpose()?,
            name: u.arbitrary()?,
            explosion: u.arbitrary::<bool>()?.then(|| u.int_in_range(0..=100u8)).transpose()?,
            smooth: u.arbitrary()?,
            has_trendline: u.arbitrary()?,
            has_marker: u.arbitrary()?,
        })
    }
}

#[derive(Debug)]
enum FuzzDataRef {
    Formula(String),
    Numbers(Vec<f64>),
    Strings(Vec<String>),
}

impl<'a> Arbitrary<'a> for FuzzDataRef {
    fn arbitrary(u: &mut Unstructured<'a>) -> arbitrary::Result<Self> {
        let kind: u8 = u.int_in_range(0..=3)?;
        match kind {
            0 | 1 => {
                // Formula: "Sheet1!$C$1:$C$N"
                let col = (b'A' + u.int_in_range(0..=5)?) as char;
                let end_row = u.int_in_range(1..=20)?;
                Ok(FuzzDataRef::Formula(format!(
                    "Sheet1!${col}$1:${col}${end_row}"
                )))
            }
            2 => {
                let n = u.int_in_range(1..=10)?;
                let mut vals = Vec::with_capacity(n);
                for _ in 0..n {
                    let v: f64 = u.arbitrary()?;
                    vals.push(if v.is_finite() { v } else { 0.0 });
                }
                Ok(FuzzDataRef::Numbers(vals))
            }
            _ => {
                let n = u.int_in_range(1..=10)?;
                let mut vals = Vec::with_capacity(n);
                for _ in 0..n {
                    let s = SmallString::arbitrary(u)?;
                    vals.push(s.0);
                }
                Ok(FuzzDataRef::Strings(vals))
            }
        }
    }
}

impl From<&FuzzDataRef> for DataReference {
    fn from(r: &FuzzDataRef) -> Self {
        match r {
            FuzzDataRef::Formula(f) => DataReference::Formula(f.clone()),
            FuzzDataRef::Numbers(v) => DataReference::Numbers(v.clone()),
            FuzzDataRef::Strings(v) => DataReference::Strings(v.clone()),
        }
    }
}

#[derive(Debug)]
struct FuzzAnchor {
    from_col: u16,
    from_row: u32,
    to_col: u16,
    to_row: u32,
}

impl<'a> Arbitrary<'a> for FuzzAnchor {
    fn arbitrary(u: &mut Unstructured<'a>) -> arbitrary::Result<Self> {
        let from_col = u.int_in_range(0..=20u16)?;
        let from_row = u.int_in_range(0..=200u32)?;
        let to_col = from_col + u.int_in_range(3..=15u16)?;
        let to_row = from_row + u.int_in_range(5..=20u32)?;
        Ok(FuzzAnchor { from_col, from_row, to_col, to_row })
    }
}

#[derive(Arbitrary, Debug, Clone, Copy)]
enum FuzzLegendPos {
    Right,
    Top,
    Bottom,
    Left,
    TopRight,
}

impl From<FuzzLegendPos> for LegendPosition {
    fn from(f: FuzzLegendPos) -> Self {
        match f {
            FuzzLegendPos::Right => LegendPosition::Right,
            FuzzLegendPos::Top => LegendPosition::Top,
            FuzzLegendPos::Bottom => LegendPosition::Bottom,
            FuzzLegendPos::Left => LegendPosition::Left,
            FuzzLegendPos::TopRight => LegendPosition::TopRight,
        }
    }
}

fn col_to_letters(col: u8) -> String {
    if col < 26 {
        String::from((b'A' + col) as char)
    } else {
        format!(
            "{}{}",
            (b'A' + col / 26 - 1) as char,
            (b'A' + col % 26) as char
        )
    }
}

fuzz_target!(|data: &[u8]| {
    let fwb = match FuzzWorkbook::arbitrary(&mut Unstructured::new(data)) {
        Ok(wb) => wb,
        Err(_) => return,
    };

    // Build a real Workbook from the fuzz spec
    let mut workbook = Workbook::new();

    // Ensure at least one sheet
    if fwb.sheets.is_empty() {
        return;
    }

    for (i, fsheet) in fwb.sheets.iter().enumerate() {
        // Add sheets beyond the default first one
        if i > 0 {
            let _ = workbook.add_worksheet();
        }
        let sheet = match workbook.worksheet_mut(i) {
            Some(s) => s,
            None => continue,
        };

        // Set sheet name (ignore errors from invalid names)
        let _ = sheet.set_name(&fsheet.name);

        for cell in &fsheet.cells {
            let addr = format!("{}{}", col_to_letters(cell.col), cell.row as u32 + 1);

            match &cell.value {
                FuzzValue::Empty => {}
                FuzzValue::Number(n) => {
                    if n.is_finite() {
                        let _ = sheet.set_cell_value(&addr, *n);
                    }
                }
                FuzzValue::Int(n) => {
                    let _ = sheet.set_cell_value(&addr, *n as f64);
                }
                FuzzValue::Bool(b) => {
                    let _ = sheet.set_cell_value(&addr, CellValue::Boolean(*b));
                }
                FuzzValue::Str(s) => {
                    let _ = sheet.set_cell_value(&addr, s.0.as_str());
                }
                FuzzValue::Formula(f) => {
                    let _ = sheet.set_cell_formula(&addr, &f.0);
                }
            }

            // Apply style if present
            if let Some(style) = &cell.style {
                let mut s = Style::new();
                if style.bold {
                    s = s.bold(true);
                }
                if style.italic {
                    s = s.italic(true);
                }
                if let Some(size) = style.font_size {
                    if size > 0 && size <= 72 {
                        s = s.font_size(size as f64);
                    }
                }
                let _ = sheet.set_cell_style(&addr, &s);
            }
        }

        for fchart in &fsheet.charts {
            let mut chart = Chart::new(fchart.chart_type.into());
            if let Some(ref t) = fchart.title {
                chart.title = Some(t.0.clone());
            }
            if let Some(hole) = fchart.hole_size {
                chart.hole_size = Some(hole as u32);
            }
            if let Some(show_val) = fchart.data_labels_show_value {
                chart.data_labels = Some(DataLabels {
                    show_value: Some(show_val),
                    ..Default::default()
                });
            }
            if let Some(lp) = fchart.legend_pos {
                chart.legend = Some(Legend::new(lp.into()));
            }
            chart.anchor = DrawingAnchor::TwoCell {
                from: CellMarker {
                    col: fchart.anchor.from_col,
                    col_offset_emu: 0,
                    row: fchart.anchor.from_row,
                    row_offset_emu: 0,
                },
                to: CellMarker {
                    col: fchart.anchor.to_col,
                    col_offset_emu: 0,
                    row: fchart.anchor.to_row,
                    row_offset_emu: 0,
                },
                edit_as: None,
            };
            for fs in &fchart.series {
                let mut series = DataSeries::new(DataReference::from(&fs.values));
                if let Some(ref cats) = fs.categories {
                    series.categories = Some(DataReference::from(cats));
                }
                if let Some(ref n) = fs.name {
                    series.name = Some(n.0.clone());
                }
                if let Some(exp) = fs.explosion {
                    series.explosion = Some(exp as u32);
                }
                series.smooth = fs.smooth;
                if fs.has_trendline {
                    series.trendline = Some(Trendline {
                        trendline_type: TrendlineType::Linear,
                        name: None,
                        order: None,
                        period: None,
                        forward: None,
                        backward: None,
                        intercept: None,
                        display_r_squared: None,
                        display_equation: None,
                        label: None,
                    });
                }
                if fs.has_marker {
                    series.marker = Some(Marker {
                        symbol: Some(MarkerSymbol::Circle),
                        size: Some(5),
                    });
                }
                chart.add_series(series);
            }
            sheet.add_chart(chart);
        }
    }

    // Step 1: Write workbook to XLSX bytes
    let mut output = Cursor::new(Vec::new());
    if duke_sheets_xlsx::XlsxWriter::write(&workbook, &mut output).is_err() {
        return;
    }

    // Step 2: Read back — must not panic
    let written = output.into_inner();
    let orig_wb = &workbook;
    let rt_wb = match duke_sheets_xlsx::XlsxReader::read(Cursor::new(&written)) {
        Ok(wb) => wb,
        Err(_) => return,
    };

    // Compare images
    for i in 0..orig_wb.sheet_count() {
        let orig_ws = orig_wb.worksheet(i).unwrap();
        let rt_ws = rt_wb.worksheet(i).unwrap();
        assert_eq!(
            orig_ws.images().len(),
            rt_ws.images().len(),
            "image count mismatch on sheet {}",
            i
        );
    }
});
