//! Chart parity tests: verify duke-sheets chart output against real Excel.
//!
//! - `chart_roundtrip_through_excel`: duke-sheets writes a fat chart workbook,
//!   Excel opens it (no repair), re-saves, duke-sheets reads back, asserts
//!   all chart data survived.
//!
//! - `generate_chart_parity_spreadsheet`: Excel COM creates charts from scratch,
//!   saves to `data/chart-parity.xlsx` for offline verification.

use std::path::PathBuf;

use crate::{
    cleanup_fixture, ensure_vm_temp_dir, excel_bridge, pull_file_from_vm, push_file_to_vm,
    roundtrip_through_excel, temp_fixture,
};
use duke_sheets_chart::{
    Axis, CellMarker, Chart, ChartColor, ChartLine, ChartShapeProperties, ChartType, DataLabels,
    DataPoint, DataReference, DataSeries, DrawingAnchor, Legend, LegendPosition, Marker,
    MarkerSymbol, Trendline, TrendlineType,
};
use duke_sheets_core::{ChartSheet, SheetVisibility, Workbook};
use duke_sheets_xlsx::XlsxWriter;

const REPO_DATA_DIR: &str = "data";

fn default_anchor(from_row: u32, to_row: u32) -> DrawingAnchor {
    DrawingAnchor::TwoCell {
        from: CellMarker {
            col: 5,
            col_offset_emu: 0,
            row: from_row,
            row_offset_emu: 0,
        },
        to: CellMarker {
            col: 15,
            col_offset_emu: 0,
            row: to_row,
            row_offset_emu: 0,
        },
        edit_as: None,
    }
}

fn sample_series(sheet: &str, col: &str) -> DataSeries {
    DataSeries::new(DataReference::formula(format!("{sheet}!${col}$2:${col}$7")))
        .with_name(format!("{sheet}!${col}$1"))
        .with_categories(DataReference::formula(format!("{sheet}!$A$2:$A$7")))
}

fn chart_color(hex: &str) -> ChartColor {
    ChartColor { hex: hex.into() }
}

fn solid_fill(hex: &str) -> ChartShapeProperties {
    ChartShapeProperties {
        solid_fill: Some(chart_color(hex)),
        no_fill: false,
        line: None,
    }
}

fn line_fill(hex: &str) -> ChartShapeProperties {
    ChartShapeProperties {
        solid_fill: None,
        no_fill: false,
        line: Some(ChartLine {
            width: Some(9360),
            solid_fill: Some(chart_color(hex)),
            no_fill: false,
            dash_style: None,
        }),
    }
}

fn fill_and_line(fill_hex: &str, line_hex: &str) -> ChartShapeProperties {
    ChartShapeProperties {
        solid_fill: Some(chart_color(fill_hex)),
        no_fill: false,
        line: Some(ChartLine {
            width: Some(9360),
            solid_fill: Some(chart_color(line_hex)),
            no_fill: false,
            dash_style: None,
        }),
    }
}

fn build_chart_parity_workbook() -> Workbook {
    let mut wb = Workbook::new();

    // Data sheet: categories in A, three value columns in B/C/D
    let sheet = wb.worksheet_mut(0).unwrap();
    sheet.set_name("Data");
    let categories = ["Jan", "Feb", "Mar", "Apr", "May", "Jun"];
    let series_b = [10.0, 25.0, 15.0, 30.0, 20.0, 35.0];
    let series_c = [8.0, 18.0, 22.0, 12.0, 28.0, 16.0];
    let series_d = [5.0, 12.0, 8.0, 20.0, 15.0, 25.0];

    sheet.set_cell_value("A1", "Month").unwrap();
    sheet.set_cell_value("B1", "Revenue").unwrap();
    sheet.set_cell_value("C1", "Profit").unwrap();
    sheet.set_cell_value("D1", "Costs").unwrap();

    for (i, cat) in categories.iter().enumerate() {
        let row = i as u32 + 2;
        let r = format!("{row}");
        sheet.set_cell_value(&format!("A{r}"), *cat).unwrap();
        sheet.set_cell_value(&format!("B{r}"), series_b[i]).unwrap();
        sheet.set_cell_value(&format!("C{r}"), series_c[i]).unwrap();
        sheet.set_cell_value(&format!("D{r}"), series_d[i]).unwrap();
    }

    // Scatter/bubble data in E-G
    sheet.set_cell_value("E1", "X").unwrap();
    sheet.set_cell_value("F1", "Y").unwrap();
    sheet.set_cell_value("G1", "Size").unwrap();
    let scatter_x = [1.0, 2.0, 3.0, 4.0, 5.0, 6.0];
    let scatter_y = [2.1, 4.0, 5.8, 8.1, 10.2, 11.9];
    let bubble_sz = [10.0, 20.0, 15.0, 25.0, 12.0, 30.0];
    for i in 0..6 {
        let row = i + 2;
        let r = format!("{row}");
        sheet
            .set_cell_value(&format!("E{r}"), scatter_x[i as usize])
            .unwrap();
        sheet
            .set_cell_value(&format!("F{r}"), scatter_y[i as usize])
            .unwrap();
        sheet
            .set_cell_value(&format!("G{r}"), bubble_sz[i as usize])
            .unwrap();
    }

    // Stock data in H-K (open, high, low, close)
    sheet.set_cell_value("H1", "Open").unwrap();
    sheet.set_cell_value("I1", "High").unwrap();
    sheet.set_cell_value("J1", "Low").unwrap();
    sheet.set_cell_value("K1", "Close").unwrap();
    let stock = [
        (100.0, 110.0, 95.0, 105.0),
        (105.0, 115.0, 100.0, 108.0),
        (108.0, 120.0, 102.0, 112.0),
        (112.0, 118.0, 106.0, 110.0),
        (110.0, 125.0, 108.0, 122.0),
        (122.0, 130.0, 115.0, 128.0),
    ];
    for (i, (o, h, l, c)) in stock.iter().enumerate() {
        let row = i as u32 + 2;
        let r = format!("{row}");
        sheet.set_cell_value(&format!("H{r}"), *o).unwrap();
        sheet.set_cell_value(&format!("I{r}"), *h).unwrap();
        sheet.set_cell_value(&format!("J{r}"), *l).unwrap();
        sheet.set_cell_value(&format!("K{r}"), *c).unwrap();
    }

    let mut row_cursor = 0u32;
    let row_step = 16;

    // 1. ColumnClustered
    let mut c = Chart::new(ChartType::ColumnClustered);
    c.title = Some("ColumnClustered".into());
    c.shape_properties = Some(fill_and_line("FAFBFC", "D9D9D9"));
    let mut revenue = sample_series("Data", "B");
    revenue.shape_properties = Some(solid_fill("112233"));
    c.add_series(revenue);
    let mut profit = sample_series("Data", "C");
    profit.shape_properties = Some(solid_fill("445566"));
    c.add_series(profit);
    let mut cat_axis = Axis::new().with_title("Month");
    cat_axis.shape_properties = Some(line_fill("878787"));
    c.category_axis = Some(cat_axis);
    let mut val_axis = Axis::new().with_title("Value");
    val_axis.major_gridlines = true;
    val_axis.major_gridlines_shape_properties = Some(line_fill("D9D9D9"));
    val_axis.shape_properties = Some(line_fill("878787"));
    c.value_axis = Some(val_axis);
    c.legend = Some(Legend::new(LegendPosition::Bottom));
    sheet.add_chart(c, default_anchor(row_cursor, row_cursor + row_step - 1));
    row_cursor += row_step;

    // 2. ColumnStacked
    let mut c = Chart::new(ChartType::ColumnStacked);
    c.title = Some("ColumnStacked".into());
    c.add_series(sample_series("Data", "B"));
    c.add_series(sample_series("Data", "C"));
    c.add_series(sample_series("Data", "D"));
    sheet.add_chart(c, default_anchor(row_cursor, row_cursor + row_step - 1));
    row_cursor += row_step;

    // 3. ColumnPercentStacked
    let mut c = Chart::new(ChartType::ColumnPercentStacked);
    c.title = Some("ColumnPercentStacked".into());
    c.add_series(sample_series("Data", "B"));
    c.add_series(sample_series("Data", "C"));
    sheet.add_chart(c, default_anchor(row_cursor, row_cursor + row_step - 1));
    row_cursor += row_step;

    // 4. BarClustered
    let mut c = Chart::new(ChartType::BarClustered);
    c.title = Some("BarClustered".into());
    c.add_series(sample_series("Data", "B"));
    c.add_series(sample_series("Data", "C"));
    c.legend = Some(Legend::new(LegendPosition::Right));
    sheet.add_chart(c, default_anchor(row_cursor, row_cursor + row_step - 1));
    row_cursor += row_step;

    // 5. BarStacked
    let mut c = Chart::new(ChartType::BarStacked);
    c.title = Some("BarStacked".into());
    c.add_series(sample_series("Data", "B"));
    c.add_series(sample_series("Data", "C"));
    sheet.add_chart(c, default_anchor(row_cursor, row_cursor + row_step - 1));
    row_cursor += row_step;

    // 6. BarPercentStacked
    let mut c = Chart::new(ChartType::BarPercentStacked);
    c.title = Some("BarPercentStacked".into());
    c.add_series(sample_series("Data", "B"));
    c.add_series(sample_series("Data", "C"));
    sheet.add_chart(c, default_anchor(row_cursor, row_cursor + row_step - 1));
    row_cursor += row_step;

    // 7. Line
    let mut c = Chart::new(ChartType::Line);
    c.title = Some("Line".into());
    c.add_series(sample_series("Data", "B"));
    c.add_series(sample_series("Data", "C"));
    c.category_axis = Some(Axis::new());
    c.value_axis = Some(Axis::new());
    sheet.add_chart(c, default_anchor(row_cursor, row_cursor + row_step - 1));
    row_cursor += row_step;

    // 8. LineStacked
    let mut c = Chart::new(ChartType::LineStacked);
    c.title = Some("LineStacked".into());
    c.add_series(sample_series("Data", "B"));
    c.add_series(sample_series("Data", "C"));
    sheet.add_chart(c, default_anchor(row_cursor, row_cursor + row_step - 1));
    row_cursor += row_step;

    // 9. Pie
    let mut c = Chart::new(ChartType::Pie);
    c.title = Some("Pie".into());
    let mut s = sample_series("Data", "B");
    s.shape_properties = Some(solid_fill("112233"));
    s.data_points = vec![
        DataPoint {
            index: 1,
            marker: None,
            explosion: None,
            shape_properties: Some(solid_fill("445566")),
        },
        DataPoint {
            index: 2,
            marker: None,
            explosion: None,
            shape_properties: Some(solid_fill("778899")),
        },
        DataPoint {
            index: 3,
            marker: None,
            explosion: None,
            shape_properties: Some(solid_fill("ABCDEF")),
        },
    ];
    c.add_series(s);
    c.legend = Some(Legend::new(LegendPosition::Right));
    sheet.add_chart(c, default_anchor(row_cursor, row_cursor + row_step - 1));
    row_cursor += row_step;

    // 10. PieExploded
    let mut c = Chart::new(ChartType::PieExploded);
    c.title = Some("PieExploded".into());
    let mut s = sample_series("Data", "B");
    s.explosion = Some(25); // must have explosion for Excel to preserve PieExploded
    c.add_series(s);
    sheet.add_chart(c, default_anchor(row_cursor, row_cursor + row_step - 1));
    row_cursor += row_step;

    // 11. Doughnut
    let mut c = Chart::new(ChartType::Doughnut);
    c.title = Some("Doughnut".into());
    c.add_series(sample_series("Data", "B"));
    c.hole_size = Some(50);
    sheet.add_chart(c, default_anchor(row_cursor, row_cursor + row_step - 1));
    row_cursor += row_step;

    // 12. Area
    let mut c = Chart::new(ChartType::Area);
    c.title = Some("Area".into());
    c.add_series(sample_series("Data", "B"));
    c.add_series(sample_series("Data", "C"));
    sheet.add_chart(c, default_anchor(row_cursor, row_cursor + row_step - 1));
    row_cursor += row_step;

    // 13. AreaStacked
    let mut c = Chart::new(ChartType::AreaStacked);
    c.title = Some("AreaStacked".into());
    c.add_series(sample_series("Data", "B"));
    c.add_series(sample_series("Data", "C"));
    sheet.add_chart(c, default_anchor(row_cursor, row_cursor + row_step - 1));
    row_cursor += row_step;

    // 14. AreaPercentStacked
    let mut c = Chart::new(ChartType::AreaPercentStacked);
    c.title = Some("AreaPercentStacked".into());
    c.add_series(sample_series("Data", "B"));
    c.add_series(sample_series("Data", "C"));
    sheet.add_chart(c, default_anchor(row_cursor, row_cursor + row_step - 1));
    row_cursor += row_step;

    // 15. ScatterLines
    let mut c = Chart::new(ChartType::ScatterLines);
    c.title = Some("ScatterLines".into());
    let s = DataSeries::new(DataReference::formula("Data!$F$2:$F$7"))
        .with_name("Data!$F$1")
        .with_categories(DataReference::formula("Data!$E$2:$E$7"));
    c.add_series(s);
    sheet.add_chart(c, default_anchor(row_cursor, row_cursor + row_step - 1));
    row_cursor += row_step;

    // 16. ScatterSmooth
    let mut c = Chart::new(ChartType::ScatterSmooth);
    c.title = Some("ScatterSmooth".into());
    let s = DataSeries::new(DataReference::formula("Data!$F$2:$F$7"))
        .with_name("Data!$F$1")
        .with_categories(DataReference::formula("Data!$E$2:$E$7"));
    c.add_series(s);
    sheet.add_chart(c, default_anchor(row_cursor, row_cursor + row_step - 1));
    row_cursor += row_step;

    // 17. ScatterMarkers
    let mut c = Chart::new(ChartType::ScatterMarkers);
    c.title = Some("ScatterMarkers".into());
    let mut s = DataSeries::new(DataReference::formula("Data!$F$2:$F$7"))
        .with_name("Data!$F$1")
        .with_categories(DataReference::formula("Data!$E$2:$E$7"));
    s.marker = Some(Marker {
        symbol: Some(MarkerSymbol::Circle),
        size: Some(8),
        ..Default::default()
    });
    c.add_series(s);
    sheet.add_chart(c, default_anchor(row_cursor, row_cursor + row_step - 1));
    row_cursor += row_step;

    // 18. Bubble
    let mut c = Chart::new(ChartType::Bubble);
    c.title = Some("Bubble".into());
    let s = DataSeries::new(DataReference::formula("Data!$F$2:$F$7"))
        .with_categories(DataReference::formula("Data!$E$2:$E$7"));
    c.add_series(s);
    sheet.add_chart(c, default_anchor(row_cursor, row_cursor + row_step - 1));
    row_cursor += row_step;

    // 19. Radar
    let mut c = Chart::new(ChartType::Radar);
    c.title = Some("Radar".into());
    c.add_series(sample_series("Data", "B"));
    c.add_series(sample_series("Data", "C"));
    sheet.add_chart(c, default_anchor(row_cursor, row_cursor + row_step - 1));
    row_cursor += row_step;

    // 20. Stock (HLC)
    let mut c = Chart::new(ChartType::Stock);
    c.title = Some("Stock".into());
    for col in ["I", "J", "K"] {
        c.add_series(
            DataSeries::new(DataReference::formula(format!("Data!${col}$2:${col}$7")))
                .with_name(format!("Data!${col}$1"))
                .with_categories(DataReference::formula("Data!$A$2:$A$7".to_string())),
        );
    }
    sheet.add_chart(c, default_anchor(row_cursor, row_cursor + row_step - 1));
    row_cursor += row_step;

    // 21. Line with trendline + data labels
    let mut c = Chart::new(ChartType::Line);
    c.title = Some("Line+Trendline+Labels".into());
    let mut s = sample_series("Data", "B");
    s.trendline = Some(Trendline {
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
    c.add_series(s);
    c.data_labels = Some(DataLabels {
        show_value: Some(true),
        ..Default::default()
    });
    sheet.add_chart(c, default_anchor(row_cursor, row_cursor + row_step - 1));
    row_cursor += row_step;

    // 22. Surface
    let mut c = Chart::new(ChartType::Surface);
    c.title = Some("Surface".into());
    c.add_series(sample_series("Data", "B"));
    c.add_series(sample_series("Data", "C"));
    c.add_series(sample_series("Data", "D"));
    sheet.add_chart(c, default_anchor(row_cursor, row_cursor + row_step - 1));

    let mut cs_chart = Chart::new(ChartType::ColumnClustered);
    cs_chart.title = Some("ChartSheet: Column".into());
    cs_chart.add_series(sample_series("Data", "B"));
    cs_chart.add_series(sample_series("Data", "C"));
    cs_chart.legend = Some(Legend::new(LegendPosition::Bottom));
    wb.add_chartsheet(ChartSheet {
        name: "Chart1".to_string(),
        chart: cs_chart,
        visibility: SheetVisibility::Visible,
        raw_drawing_objects: Vec::new(),
        raw_drawing_rels: Vec::new(),
    })
    .unwrap();

    wb
}

/// duke-sheets writes a fat chart workbook → Excel opens (no repair) →
/// re-saves → duke-sheets reads back → assert chart data survived.
#[test]
fn chart_roundtrip_through_excel() {
    let wb = build_chart_parity_workbook();

    // Count charts before roundtrip
    let ws = wb.worksheet(0).unwrap();
    let chart_count_before = ws.chart_count();
    let chartsheet_count_before = wb.chartsheet_count();
    assert!(
        chart_count_before >= 20,
        "expected >= 20 charts, got {chart_count_before}"
    );
    assert_eq!(chartsheet_count_before, 1);

    // Collect chart types and titles before roundtrip
    let types_before: Vec<(ChartType, Option<String>)> = ws
        .charts()
        .map(|c| (c.payload.chart_type.clone(), c.payload.title.clone()))
        .collect();

    // Roundtrip through real Excel
    let wb2 = roundtrip_through_excel(&wb);

    // Assert chart counts survived
    let ws2 = wb2.worksheet(0).unwrap();
    assert_eq!(
        ws2.chart_count(),
        chart_count_before,
        "worksheet chart count changed after Excel roundtrip"
    );
    assert_eq!(
        wb2.chartsheet_count(),
        chartsheet_count_before,
        "chartsheet count changed after Excel roundtrip"
    );

    // Suppress unused warning for types_before (kept for backward compat)
    let _ = &types_before;

    fn assert_formula(dr: &DataReference, expected: &str, ctx: &str) {
        match dr {
            DataReference::Formula(s) if s == expected => {}
            DataReference::Formula(s) => panic!("{ctx}: expected \"{expected}\", got \"{s}\""),
            other => panic!("{ctx}: expected Formula, got {other:?}"),
        }
    }

    fn assert_legend_pos(chart: &Chart, expected: LegendPosition, ctx: &str) {
        let legend = chart
            .legend
            .as_ref()
            .unwrap_or_else(|| panic!("{ctx}: missing legend"));
        assert!(
            legend.position == expected,
            "{ctx}: expected {expected:?}, got {:?}",
            legend.position
        );
    }

    fn assert_solid_fill(sp: &Option<ChartShapeProperties>, expected: &str, ctx: &str) {
        let actual = sp
            .as_ref()
            .and_then(|sp| sp.solid_fill.as_ref())
            .map(|color| color.hex.as_str())
            .unwrap_or_else(|| panic!("{ctx}: missing solid fill"));
        assert!(
            actual.eq_ignore_ascii_case(expected),
            "{ctx}: expected solid fill {expected}, got {actual}"
        );
    }

    fn assert_line_fill(sp: &Option<ChartShapeProperties>, expected: &str, ctx: &str) {
        let actual = sp
            .as_ref()
            .and_then(|sp| sp.line.as_ref())
            .and_then(|line| line.solid_fill.as_ref())
            .map(|color| color.hex.as_str())
            .unwrap_or_else(|| panic!("{ctx}: missing line fill"));
        assert!(
            actual.eq_ignore_ascii_case(expected),
            "{ctx}: expected line fill {expected}, got {actual}"
        );
    }

    fn assert_data_point_fill(series: &DataSeries, idx: u32, expected: &str, ctx: &str) {
        let point = series
            .data_points
            .iter()
            .find(|point| point.index == idx)
            .unwrap_or_else(|| panic!("{ctx}: missing data point {idx}"));
        assert_solid_fill(&point.shape_properties, expected, ctx);
    }

    let charts: Vec<_> = ws2.charts().collect();
    let orig_charts: Vec<_> = wb.worksheet(0).unwrap().charts().collect();

    for (i, drawn) in charts.iter().enumerate() {
        let c = drawn.payload;
        let orig = orig_charts[i].payload;
        let label = orig.title.as_deref().unwrap_or("?");

        // Chart type (ScatterMarkers↔ScatterLines equivalence)
        assert!(
            chart_type_equivalent(&orig.chart_type, &c.chart_type),
            "chart {i} ({label}) type: expected {:?}, got {:?}",
            orig.chart_type,
            c.chart_type
        );

        // Title
        assert_eq!(&c.title, &orig.title, "chart {i} ({label}) title mismatch");

        // Series count
        assert_eq!(
            c.series.len(),
            orig.series.len(),
            "chart {i} ({label}) series count: expected {}, got {}",
            orig.series.len(),
            c.series.len()
        );

        // Per-series: values formula, name, categories formula
        for (si, s) in c.series.iter().enumerate() {
            let os = &orig.series[si];
            let ctx = format!("chart {i} ({label}) series {si}");

            if let DataReference::Formula(expected) = &os.values {
                assert_formula(&s.values, expected, &format!("{ctx} values"));
            }

            assert_eq!(&s.name, &os.name, "{ctx} name");

            match (&s.categories, &os.categories) {
                (Some(actual), Some(DataReference::Formula(expected))) => {
                    assert_formula(actual, expected, &format!("{ctx} categories"));
                }
                (None, None) => {}
                (actual, expected) => {
                    panic!("{ctx} categories: expected {expected:?}, got {actual:?}");
                }
            }
        }

        // Anchor: col/row positions must survive (EMU offsets may change)
        if let (
            DrawingAnchor::TwoCell { from, to, .. },
            DrawingAnchor::TwoCell {
                from: orig_from,
                to: orig_to,
                ..
            },
        ) = (&drawn.object.anchor, &orig_charts[i].object.anchor)
        {
            assert_eq!(
                from.col, orig_from.col,
                "chart {i} ({label}) anchor from_col"
            );
            assert_eq!(to.col, orig_to.col, "chart {i} ({label}) anchor to_col");
            assert_eq!(
                from.row, orig_from.row,
                "chart {i} ({label}) anchor from_row"
            );
            assert_eq!(to.row, orig_to.row, "chart {i} ({label}) anchor to_row");
        } else {
            assert_eq!(
                std::mem::discriminant(&drawn.object.anchor),
                std::mem::discriminant(&orig_charts[i].object.anchor),
                "chart {i} ({label}) anchor type mismatch"
            );
        }
    }

    // 0: ColumnClustered - axis titles + legend
    {
        let c = charts[0].payload;
        let cat_ax = c
            .category_axis
            .as_ref()
            .expect("chart 0 (ColumnClustered) missing category_axis");
        assert_eq!(
            cat_ax.title.as_deref(),
            Some("Month"),
            "chart 0 catAx title"
        );
        let val_ax = c
            .value_axis
            .as_ref()
            .expect("chart 0 (ColumnClustered) missing value_axis");
        assert_eq!(
            val_ax.title.as_deref(),
            Some("Value"),
            "chart 0 valAx title"
        );
        assert_legend_pos(c, LegendPosition::Bottom, "chart 0 (ColumnClustered)");
        assert_solid_fill(&c.shape_properties, "FAFBFC", "chart 0 chartSpace fill");
        assert_line_fill(&c.shape_properties, "D9D9D9", "chart 0 chartSpace line");
        assert_solid_fill(
            &c.series[0].shape_properties,
            "112233",
            "chart 0 series 0 fill",
        );
        assert_solid_fill(
            &c.series[1].shape_properties,
            "445566",
            "chart 0 series 1 fill",
        );
        assert_line_fill(
            &cat_ax.shape_properties,
            "878787",
            "chart 0 category axis line",
        );
        assert_line_fill(
            &val_ax.shape_properties,
            "878787",
            "chart 0 value axis line",
        );
        assert!(val_ax.major_gridlines, "chart 0 missing major gridlines");
        assert_line_fill(
            &val_ax.major_gridlines_shape_properties,
            "D9D9D9",
            "chart 0 major gridlines line",
        );
    }

    // 3: BarClustered - legend position
    assert_legend_pos(
        charts[3].payload,
        LegendPosition::Right,
        "chart 3 (BarClustered)",
    );

    // 6: Line - axes present (no titles, just existence)
    {
        let c = charts[6].payload;
        assert!(
            c.category_axis.is_some(),
            "chart 6 (Line) missing category_axis"
        );
        assert!(c.value_axis.is_some(), "chart 6 (Line) missing value_axis");
    }

    // 8: Pie - legend=Right, no axes
    {
        let c = charts[8].payload;
        assert_legend_pos(c, LegendPosition::Right, "chart 8 (Pie)");
        assert!(
            c.category_axis.is_none(),
            "chart 8 (Pie) should have no category_axis"
        );
        assert!(
            c.value_axis.is_none(),
            "chart 8 (Pie) should have no value_axis"
        );
        assert_solid_fill(
            &c.series[0].shape_properties,
            "112233",
            "chart 8 series fill",
        );
        assert_data_point_fill(&c.series[0], 1, "445566", "chart 8 point 1 fill");
        assert_data_point_fill(&c.series[0], 2, "778899", "chart 8 point 2 fill");
        assert_data_point_fill(&c.series[0], 3, "ABCDEF", "chart 8 point 3 fill");
    }

    // 9: PieExploded - explosion=25
    {
        let explosion = charts[9].payload.series[0]
            .explosion
            .expect("chart 9 (PieExploded) series[0] missing explosion");
        assert_eq!(explosion, 25, "chart 9 explosion");
    }

    // 10: Doughnut - hole_size=50
    {
        let hole = charts[10]
            .payload
            .hole_size
            .expect("chart 10 (Doughnut) missing hole_size");
        assert_eq!(hole, 50, "chart 10 hole_size");
    }

    // 14-16: Scatter variants - series[0].name = "Data!$F$1"
    for i in [14, 15, 16] {
        let c = charts[i].payload;
        assert_eq!(
            c.series[0].name.as_deref(),
            Some("Data!$F$1"),
            "chart {i} ({:?}) series[0] name",
            c.chart_type
        );
    }

    // 16: ScatterMarkers - marker symbol + size
    {
        let marker = charts[16].payload.series[0]
            .marker
            .as_ref()
            .expect("chart 16 (ScatterMarkers) series[0] missing marker");
        assert_eq!(
            marker.symbol,
            Some(MarkerSymbol::Circle),
            "chart 16 marker symbol"
        );
        assert_eq!(marker.size, Some(8), "chart 16 marker size");
    }

    // 19: Stock - 3 series with specific names
    {
        let c = charts[19].payload;
        let expected_names = ["Data!$I$1", "Data!$J$1", "Data!$K$1"];
        for (si, expected) in expected_names.iter().enumerate() {
            assert_eq!(
                c.series[si].name.as_deref(),
                Some(*expected),
                "chart 19 (Stock) series[{si}] name"
            );
        }
    }

    // 20: Line+Trendline+Labels - trendline type + data_labels.show_value
    {
        let c = charts[20].payload;
        let trendline = c.series[0]
            .trendline
            .as_ref()
            .expect("chart 20 (Line+Trendline) series[0] missing trendline");
        assert_eq!(
            trendline.trendline_type,
            TrendlineType::Linear,
            "chart 20 trendline type"
        );
        let dl = c
            .data_labels
            .as_ref()
            .expect("chart 20 (Line+Trendline) missing data_labels");
        assert_eq!(dl.show_value, Some(true), "chart 20 data_labels.show_value");
    }

    // ChartSheet: full verification
    {
        let cs = wb2.chartsheet(0).unwrap();
        assert_eq!(cs.name, "Chart1", "chartsheet name");
        assert_eq!(
            cs.chart.chart_type,
            ChartType::ColumnClustered,
            "chartsheet chart type"
        );
        assert_eq!(
            cs.chart.title.as_deref(),
            Some("ChartSheet: Column"),
            "chartsheet title"
        );
        assert_eq!(cs.chart.series.len(), 2, "chartsheet series count");
        assert_formula(
            &cs.chart.series[0].values,
            "Data!$B$2:$B$7",
            "chartsheet series 0 values",
        );
        assert_eq!(
            cs.chart.series[0].name.as_deref(),
            Some("Data!$B$1"),
            "chartsheet series 0 name"
        );
        assert_formula(
            &cs.chart.series[1].values,
            "Data!$C$2:$C$7",
            "chartsheet series 1 values",
        );
        assert_eq!(
            cs.chart.series[1].name.as_deref(),
            Some("Data!$C$1"),
            "chartsheet series 1 name"
        );
        assert_legend_pos(&cs.chart, LegendPosition::Bottom, "chartsheet");
    }

    eprintln!(
        "chart_roundtrip_through_excel: all {} worksheet charts + chartsheet verified",
        charts.len()
    );
}

/// Generate a chart parity fixture via Excel COM.
///
/// Excel creates charts from scratch, giving us ground-truth chart XML.
/// Saved to `data/chart-parity.xlsx` for offline verification.
#[test]
fn generate_chart_parity_spreadsheet() {
    let bridge = excel_bridge();
    let fixture = temp_fixture();

    {
        let excel = bridge.lock().unwrap();
        ensure_vm_temp_dir();

        let wb = excel.create_workbook().expect("create workbook");

        // Populate data on Sheet1
        wb.set_cell_value("A1", "Month").unwrap();
        wb.set_cell_value("B1", "Revenue").unwrap();
        wb.set_cell_value("C1", "Profit").unwrap();
        let categories = ["Jan", "Feb", "Mar", "Apr", "May", "Jun"];
        let series_b = [10.0, 25.0, 15.0, 30.0, 20.0, 35.0];
        let series_c = [8.0, 18.0, 22.0, 12.0, 28.0, 16.0];
        for (i, cat) in categories.iter().enumerate() {
            let row = i as u32 + 2;
            wb.set_cell_value(&cell_addr(row, 0), *cat).unwrap();
            wb.set_cell_value(&cell_addr(row, 1), series_b[i]).unwrap();
            wb.set_cell_value(&cell_addr(row, 2), series_c[i]).unwrap();
        }

        // Create charts via Shapes.AddChart2
        // xlColumnClustered=51, xlBarClustered=57, xlLine=4, xlPie=5,
        // xlDoughnut=-4120, xlXYScatter=-4169, xlArea=1, xlRadar=-4151
        let chart_specs: &[(&str, i32)] = &[
            ("ColumnClustered", 51),
            ("BarClustered", 57),
            ("Line", 4),
            ("Pie", 5),
            ("Doughnut", -4120),
            ("XYScatter", -4169),
            ("Area", 1),
            ("Radar", -4151),
        ];

        let ws_handle = excel
            .navigate(
                wb.handle(),
                vec![excel_com_protocol::SheetRef::Index(0).to_chain_step()],
            )
            .expect("navigate to Sheet1");

        for (idx, (name, xl_type)) in chart_specs.iter().enumerate() {
            let top = (idx as f64) * 320.0;
            let shape_result = excel.invoke(
                ws_handle,
                vec![duke_sheets_excel_com::ChainStep::Property(
                    "Shapes".to_string(),
                )],
                "AddChart2",
                vec![
                    serde_json::Value::from(-1),
                    serde_json::Value::from(*xl_type),
                    serde_json::Value::from(350.0),
                    serde_json::Value::from(top),
                    serde_json::Value::from(400.0),
                    serde_json::Value::from(300.0),
                ],
            );

            match shape_result {
                Ok(Some(excel_com_protocol::ResponseData::Handle { handle })) => {
                    // Navigate to Shape.Chart and set source data
                    if let Ok(chart_handle) = excel.navigate(
                        handle,
                        vec![duke_sheets_excel_com::ChainStep::Property(
                            "Chart".to_string(),
                        )],
                    ) {
                        // SetSourceData
                        let range_ref = if *name == "Pie" || *name == "Doughnut" {
                            "A1:B7"
                        } else {
                            "A1:C7"
                        };
                        if let Ok(range_handle) = excel.navigate(
                            wb.handle(),
                            vec![
                                excel_com_protocol::SheetRef::Index(0).to_chain_step(),
                                duke_sheets_excel_com::ChainStep::Indexed(
                                    "Range".to_string(),
                                    serde_json::Value::from(range_ref),
                                ),
                            ],
                        ) {
                            let _ = excel.invoke(
                                chart_handle,
                                vec![],
                                "SetSourceData",
                                vec![serde_json::json!({"$ref": range_handle})],
                            );
                            let _ = excel.release(range_handle);
                        }

                        // Set chart title
                        let _ = excel.set(
                            chart_handle,
                            vec![duke_sheets_excel_com::ChainStep::Property(
                                "ChartTitle".to_string(),
                            )],
                            "Text",
                            serde_json::Value::from(*name),
                        );

                        let _ = excel.release(chart_handle);
                    }
                    let _ = excel.release(handle);
                }
                Ok(_) => eprintln!("AddChart2 for {name}: unexpected response"),
                Err(e) => eprintln!("AddChart2 for {name}: {e}"),
            }
        }

        let _ = excel.release(ws_handle);

        wb.save(&fixture.vm_path).expect("save workbook");
        wb.close().expect("close workbook");
    }

    pull_file_from_vm(&fixture);
    copy_fixture_into_repo(&fixture.host_path).expect("copy fixture into repo");
    cleanup_fixture(&fixture);
}

fn cell_addr(row: u32, col: u32) -> String {
    let col_char = (b'A' + col as u8) as char;
    format!("{col_char}{}", row + 1)
}

/// Check if two chart types are equivalent after Excel roundtrip.
/// Excel normalizes some subtypes on re-save:
/// - ScatterMarkers → ScatterLines (Excel adds default connecting lines)
fn chart_type_equivalent(a: &ChartType, b: &ChartType) -> bool {
    if a == b {
        return true;
    }
    matches!(
        (a, b),
        (ChartType::ScatterMarkers, ChartType::ScatterLines)
            | (ChartType::ScatterLines, ChartType::ScatterMarkers)
    )
}

fn copy_fixture_into_repo(
    host_path: &std::path::Path,
) -> std::result::Result<(), Box<dyn std::error::Error>> {
    let repo_root = PathBuf::from(env!("CARGO_MANIFEST_DIR"))
        .join("../..")
        .canonicalize()?;
    let dest = repo_root.join(REPO_DATA_DIR).join("chart-parity.xlsx");
    std::fs::copy(host_path, &dest)?;
    println!("Copied chart parity fixture to {}", dest.display());
    Ok(())
}

/// Debug: minimal single-chart file → Excel open to narrow down the issue.
#[test]
fn chart_minimal_excel_open() {
    let mut wb = Workbook::new();
    let sheet = wb.worksheet_mut(0).unwrap();
    sheet.set_cell_value("A1", "Cat").unwrap();
    sheet.set_cell_value("B1", "Val").unwrap();
    sheet.set_cell_value("A2", "X").unwrap();
    sheet.set_cell_value("B2", 10.0).unwrap();
    sheet.set_cell_value("A3", "Y").unwrap();
    sheet.set_cell_value("B3", 20.0).unwrap();

    let mut chart = Chart::new(ChartType::ColumnClustered);
    chart.title = Some("Test".into());
    chart.add_series(
        DataSeries::new(DataReference::formula("Sheet1!$B$2:$B$3".to_string()))
            .with_categories(DataReference::formula("Sheet1!$A$2:$A$3".to_string())),
    );
    sheet.add_chart(
        chart,
        DrawingAnchor::TwoCell {
            from: CellMarker {
                col: 3,
                col_offset_emu: 0,
                row: 0,
                row_offset_emu: 0,
            },
            to: CellMarker {
                col: 10,
                col_offset_emu: 0,
                row: 15,
                row_offset_emu: 0,
            },
            edit_as: None,
        },
    );

    let fixture = temp_fixture();
    let mut buf = Vec::new();
    XlsxWriter::write(&wb, std::io::Cursor::new(&mut buf)).expect("write xlsx");
    std::fs::write(&fixture.host_path, &buf).expect("write to disk");
    eprintln!(
        "Wrote {} bytes to {}",
        buf.len(),
        fixture.host_path.display()
    );

    // Also save a copy for manual inspection
    let _ = std::fs::write("/tmp/duke-sheets-excel/chart-debug.xlsx", &buf);

    ensure_vm_temp_dir();
    push_file_to_vm(&fixture);

    let bridge = excel_bridge();
    let excel = bridge.lock().unwrap();
    let opened = excel
        .open_workbook(&fixture.vm_path)
        .expect("Excel should open minimal chart file");

    let name = opened.name().unwrap_or_default();
    eprintln!("Excel opened: {name}");
    assert!(
        !name.contains("Repaired"),
        "Excel repaired the file: {name}"
    );
    assert!(
        !opened.is_read_only().unwrap_or(false),
        "Excel opened read-only"
    );

    opened.close().expect("close");
    cleanup_fixture(&fixture);
}

/// Bisect chart types: create a SEPARATE xlsx for EACH chart type,
/// open each in Excel, and report which ones pass/fail.
#[test]
fn chart_types_bisect() {
    let bridge = excel_bridge();
    ensure_vm_temp_dir();

    // Build shared data workbook once - each iteration clones & adds one chart
    fn make_data_workbook() -> Workbook {
        let mut wb = Workbook::new();
        let sheet = wb.worksheet_mut(0).unwrap();
        sheet.set_name("Data");

        // Categories + 3 value columns
        sheet.set_cell_value("A1", "Month").unwrap();
        sheet.set_cell_value("B1", "Revenue").unwrap();
        sheet.set_cell_value("C1", "Profit").unwrap();
        sheet.set_cell_value("D1", "Costs").unwrap();
        let categories = ["Jan", "Feb", "Mar", "Apr", "May", "Jun"];
        let series_b = [10.0, 25.0, 15.0, 30.0, 20.0, 35.0];
        let series_c = [8.0, 18.0, 22.0, 12.0, 28.0, 16.0];
        let series_d = [5.0, 12.0, 8.0, 20.0, 15.0, 25.0];
        for (i, cat) in categories.iter().enumerate() {
            let row = i as u32 + 2;
            let r = format!("{row}");
            sheet.set_cell_value(&format!("A{r}"), *cat).unwrap();
            sheet.set_cell_value(&format!("B{r}"), series_b[i]).unwrap();
            sheet.set_cell_value(&format!("C{r}"), series_c[i]).unwrap();
            sheet.set_cell_value(&format!("D{r}"), series_d[i]).unwrap();
        }

        // Scatter/bubble data in E-G
        sheet.set_cell_value("E1", "X").unwrap();
        sheet.set_cell_value("F1", "Y").unwrap();
        sheet.set_cell_value("G1", "Size").unwrap();
        let scatter_x = [1.0, 2.0, 3.0, 4.0, 5.0, 6.0];
        let scatter_y = [2.1, 4.0, 5.8, 8.1, 10.2, 11.9];
        let bubble_sz = [10.0, 20.0, 15.0, 25.0, 12.0, 30.0];
        for i in 0..6 {
            let row = i + 2;
            let r = format!("{row}");
            sheet
                .set_cell_value(&format!("E{r}"), scatter_x[i as usize])
                .unwrap();
            sheet
                .set_cell_value(&format!("F{r}"), scatter_y[i as usize])
                .unwrap();
            sheet
                .set_cell_value(&format!("G{r}"), bubble_sz[i as usize])
                .unwrap();
        }

        // Stock data in H-K (open, high, low, close)
        sheet.set_cell_value("H1", "Open").unwrap();
        sheet.set_cell_value("I1", "High").unwrap();
        sheet.set_cell_value("J1", "Low").unwrap();
        sheet.set_cell_value("K1", "Close").unwrap();
        let stock = [
            (100.0, 110.0, 95.0, 105.0),
            (105.0, 115.0, 100.0, 108.0),
            (108.0, 120.0, 102.0, 112.0),
            (112.0, 118.0, 106.0, 110.0),
            (110.0, 125.0, 108.0, 122.0),
            (122.0, 130.0, 115.0, 128.0),
        ];
        for (i, (o, h, l, c)) in stock.iter().enumerate() {
            let row = i as u32 + 2;
            let r = format!("{row}");
            sheet.set_cell_value(&format!("H{r}"), *o).unwrap();
            sheet.set_cell_value(&format!("I{r}"), *h).unwrap();
            sheet.set_cell_value(&format!("J{r}"), *l).unwrap();
            sheet.set_cell_value(&format!("K{r}"), *c).unwrap();
        }

        wb
    }

    struct ChartSpec {
        name: &'static str,
        build: fn(&mut Workbook),
    }

    let specs: Vec<ChartSpec> = vec![
        ChartSpec {
            name: "ColumnClustered",
            build: |wb| {
                let sheet = wb.worksheet_mut(0).unwrap();
                let mut c = Chart::new(ChartType::ColumnClustered);
                c.title = Some("ColumnClustered".into());
                c.add_series(sample_series("Data", "B"));
                c.add_series(sample_series("Data", "C"));
                c.category_axis = Some(Axis::new().with_title("Month"));
                c.value_axis = Some(Axis::new().with_title("Value"));
                c.legend = Some(Legend::new(LegendPosition::Bottom));
                sheet.add_chart(c, default_anchor(0, 15));
            },
        },
        ChartSpec {
            name: "ColumnStacked",
            build: |wb| {
                let sheet = wb.worksheet_mut(0).unwrap();
                let mut c = Chart::new(ChartType::ColumnStacked);
                c.title = Some("ColumnStacked".into());
                c.add_series(sample_series("Data", "B"));
                c.add_series(sample_series("Data", "C"));
                c.add_series(sample_series("Data", "D"));
                sheet.add_chart(c, default_anchor(0, 15));
            },
        },
        ChartSpec {
            name: "ColumnPercentStacked",
            build: |wb| {
                let sheet = wb.worksheet_mut(0).unwrap();
                let mut c = Chart::new(ChartType::ColumnPercentStacked);
                c.title = Some("ColumnPercentStacked".into());
                c.add_series(sample_series("Data", "B"));
                c.add_series(sample_series("Data", "C"));
                sheet.add_chart(c, default_anchor(0, 15));
            },
        },
        ChartSpec {
            name: "BarClustered",
            build: |wb| {
                let sheet = wb.worksheet_mut(0).unwrap();
                let mut c = Chart::new(ChartType::BarClustered);
                c.title = Some("BarClustered".into());
                c.add_series(sample_series("Data", "B"));
                c.add_series(sample_series("Data", "C"));
                c.legend = Some(Legend::new(LegendPosition::Right));
                sheet.add_chart(c, default_anchor(0, 15));
            },
        },
        ChartSpec {
            name: "BarStacked",
            build: |wb| {
                let sheet = wb.worksheet_mut(0).unwrap();
                let mut c = Chart::new(ChartType::BarStacked);
                c.title = Some("BarStacked".into());
                c.add_series(sample_series("Data", "B"));
                c.add_series(sample_series("Data", "C"));
                sheet.add_chart(c, default_anchor(0, 15));
            },
        },
        ChartSpec {
            name: "BarPercentStacked",
            build: |wb| {
                let sheet = wb.worksheet_mut(0).unwrap();
                let mut c = Chart::new(ChartType::BarPercentStacked);
                c.title = Some("BarPercentStacked".into());
                c.add_series(sample_series("Data", "B"));
                c.add_series(sample_series("Data", "C"));
                sheet.add_chart(c, default_anchor(0, 15));
            },
        },
        ChartSpec {
            name: "Line",
            build: |wb| {
                let sheet = wb.worksheet_mut(0).unwrap();
                let mut c = Chart::new(ChartType::Line);
                c.title = Some("Line".into());
                c.add_series(sample_series("Data", "B"));
                c.add_series(sample_series("Data", "C"));
                c.category_axis = Some(Axis::new());
                c.value_axis = Some(Axis::new());
                sheet.add_chart(c, default_anchor(0, 15));
            },
        },
        ChartSpec {
            name: "LineStacked",
            build: |wb| {
                let sheet = wb.worksheet_mut(0).unwrap();
                let mut c = Chart::new(ChartType::LineStacked);
                c.title = Some("LineStacked".into());
                c.add_series(sample_series("Data", "B"));
                c.add_series(sample_series("Data", "C"));
                sheet.add_chart(c, default_anchor(0, 15));
            },
        },
        ChartSpec {
            name: "Pie",
            build: |wb| {
                let sheet = wb.worksheet_mut(0).unwrap();
                let mut c = Chart::new(ChartType::Pie);
                c.title = Some("Pie".into());
                c.add_series(sample_series("Data", "B"));
                c.legend = Some(Legend::new(LegendPosition::Right));
                sheet.add_chart(c, default_anchor(0, 15));
            },
        },
        ChartSpec {
            name: "PieExploded",
            build: |wb| {
                let sheet = wb.worksheet_mut(0).unwrap();
                let mut c = Chart::new(ChartType::PieExploded);
                c.title = Some("PieExploded".into());
                c.add_series(sample_series("Data", "B"));
                sheet.add_chart(c, default_anchor(0, 15));
            },
        },
        ChartSpec {
            name: "Doughnut",
            build: |wb| {
                let sheet = wb.worksheet_mut(0).unwrap();
                let mut c = Chart::new(ChartType::Doughnut);
                c.title = Some("Doughnut".into());
                c.add_series(sample_series("Data", "B"));
                c.hole_size = Some(50);
                sheet.add_chart(c, default_anchor(0, 15));
            },
        },
        ChartSpec {
            name: "Area",
            build: |wb| {
                let sheet = wb.worksheet_mut(0).unwrap();
                let mut c = Chart::new(ChartType::Area);
                c.title = Some("Area".into());
                c.add_series(sample_series("Data", "B"));
                c.add_series(sample_series("Data", "C"));
                sheet.add_chart(c, default_anchor(0, 15));
            },
        },
        ChartSpec {
            name: "AreaStacked",
            build: |wb| {
                let sheet = wb.worksheet_mut(0).unwrap();
                let mut c = Chart::new(ChartType::AreaStacked);
                c.title = Some("AreaStacked".into());
                c.add_series(sample_series("Data", "B"));
                c.add_series(sample_series("Data", "C"));
                sheet.add_chart(c, default_anchor(0, 15));
            },
        },
        ChartSpec {
            name: "AreaPercentStacked",
            build: |wb| {
                let sheet = wb.worksheet_mut(0).unwrap();
                let mut c = Chart::new(ChartType::AreaPercentStacked);
                c.title = Some("AreaPercentStacked".into());
                c.add_series(sample_series("Data", "B"));
                c.add_series(sample_series("Data", "C"));
                sheet.add_chart(c, default_anchor(0, 15));
            },
        },
        ChartSpec {
            name: "ScatterLines",
            build: |wb| {
                let sheet = wb.worksheet_mut(0).unwrap();
                let mut c = Chart::new(ChartType::ScatterLines);
                c.title = Some("ScatterLines".into());
                let s = DataSeries::new(DataReference::formula("Data!$F$2:$F$7"))
                    .with_name("Data!$F$1")
                    .with_categories(DataReference::formula("Data!$E$2:$E$7"));
                c.add_series(s);
                sheet.add_chart(c, default_anchor(0, 15));
            },
        },
        ChartSpec {
            name: "ScatterSmooth",
            build: |wb| {
                let sheet = wb.worksheet_mut(0).unwrap();
                let mut c = Chart::new(ChartType::ScatterSmooth);
                c.title = Some("ScatterSmooth".into());
                let s = DataSeries::new(DataReference::formula("Data!$F$2:$F$7"))
                    .with_name("Data!$F$1")
                    .with_categories(DataReference::formula("Data!$E$2:$E$7"));
                c.add_series(s);
                sheet.add_chart(c, default_anchor(0, 15));
            },
        },
        ChartSpec {
            name: "ScatterMarkers",
            build: |wb| {
                let sheet = wb.worksheet_mut(0).unwrap();
                let mut c = Chart::new(ChartType::ScatterMarkers);
                c.title = Some("ScatterMarkers".into());
                let mut s = DataSeries::new(DataReference::formula("Data!$F$2:$F$7"))
                    .with_name("Data!$F$1")
                    .with_categories(DataReference::formula("Data!$E$2:$E$7"));
                s.marker = Some(Marker {
                    symbol: Some(MarkerSymbol::Circle),
                    size: Some(8),
                    ..Default::default()
                });
                c.add_series(s);
                sheet.add_chart(c, default_anchor(0, 15));
            },
        },
        ChartSpec {
            name: "Bubble",
            build: |wb| {
                let sheet = wb.worksheet_mut(0).unwrap();
                let mut c = Chart::new(ChartType::Bubble);
                c.title = Some("Bubble".into());
                let s = DataSeries::new(DataReference::formula("Data!$F$2:$F$7"))
                    .with_categories(DataReference::formula("Data!$E$2:$E$7"));
                c.add_series(s);
                sheet.add_chart(c, default_anchor(0, 15));
            },
        },
        ChartSpec {
            name: "Radar",
            build: |wb| {
                let sheet = wb.worksheet_mut(0).unwrap();
                let mut c = Chart::new(ChartType::Radar);
                c.title = Some("Radar".into());
                c.add_series(sample_series("Data", "B"));
                c.add_series(sample_series("Data", "C"));
                sheet.add_chart(c, default_anchor(0, 15));
            },
        },
        ChartSpec {
            name: "Stock",
            build: |wb| {
                let sheet = wb.worksheet_mut(0).unwrap();
                let mut c = Chart::new(ChartType::Stock);
                c.title = Some("Stock".into());
                for col in ["I", "J", "K"] {
                    c.add_series(
                        DataSeries::new(DataReference::formula(format!("Data!${col}$2:${col}$7")))
                            .with_name(format!("Data!${col}$1"))
                            .with_categories(DataReference::formula("Data!$A$2:$A$7")),
                    );
                }
                sheet.add_chart(c, default_anchor(0, 15));
            },
        },
        ChartSpec {
            name: "Line+Trendline+Labels",
            build: |wb| {
                let sheet = wb.worksheet_mut(0).unwrap();
                let mut c = Chart::new(ChartType::Line);
                c.title = Some("Line+Trendline+Labels".into());
                let mut s = sample_series("Data", "B");
                s.trendline = Some(Trendline {
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
                c.add_series(s);
                c.data_labels = Some(DataLabels {
                    show_value: Some(true),
                    ..Default::default()
                });
                sheet.add_chart(c, default_anchor(0, 15));
            },
        },
        ChartSpec {
            name: "Surface",
            build: |wb| {
                let sheet = wb.worksheet_mut(0).unwrap();
                let mut c = Chart::new(ChartType::Surface);
                c.title = Some("Surface".into());
                c.add_series(sample_series("Data", "B"));
                c.add_series(sample_series("Data", "C"));
                c.add_series(sample_series("Data", "D"));
                sheet.add_chart(c, default_anchor(0, 15));
            },
        },
    ];

    let mut results: Vec<(&str, bool, String)> = Vec::new();

    for spec in &specs {
        let mut wb = make_data_workbook();
        (spec.build)(&mut wb);

        let fixture = temp_fixture();
        let mut buf = Vec::new();
        XlsxWriter::write(&wb, std::io::Cursor::new(&mut buf)).expect("write xlsx");
        std::fs::write(&fixture.host_path, &buf).expect("write to disk");
        push_file_to_vm(&fixture);

        let excel = bridge.lock().unwrap();
        let result = excel.open_workbook(&fixture.vm_path);

        let (pass, detail) = match result {
            Ok(opened) => {
                let name = opened.name().unwrap_or_default();
                let read_only = opened.is_read_only().unwrap_or(false);
                let _ = opened.close();
                if name.contains("Repaired") {
                    (false, format!("Excel repaired the file: {name}"))
                } else if read_only {
                    (false, "Excel opened read-only".to_string())
                } else {
                    (true, "OK".to_string())
                }
            }
            Err(e) => (false, format!("open_workbook error: {e}")),
        };
        drop(excel);

        results.push((spec.name, pass, detail));
        cleanup_fixture(&fixture);
    }

    eprintln!();
    eprintln!("=== Chart Type Bisect Results ===");
    eprintln!("{:<30} {:<6} {}", "Chart Type", "Result", "Detail");
    eprintln!("{}", "-".repeat(80));
    let mut fail_count = 0;
    for (name, pass, detail) in &results {
        let tag = if *pass {
            "PASS"
        } else {
            fail_count += 1;
            "FAIL"
        };
        eprintln!("{:<30} {:<6} {}", name, tag, detail);
    }
    eprintln!("{}", "-".repeat(80));
    eprintln!(
        "{} passed, {} failed out of {}",
        results.len() - fail_count,
        fail_count,
        results.len()
    );
    eprintln!();

    if fail_count > 0 {
        let failed: Vec<&str> = results
            .iter()
            .filter(|(_, p, _)| !p)
            .map(|(n, _, _)| *n)
            .collect();
        panic!(
            "{fail_count} chart type(s) failed to open in Excel: {}",
            failed.join(", ")
        );
    }
}

/// Read a corpus chartEx file, write it back through duke-sheets, push to
/// VM, verify Excel opens without repair.
#[test]
fn chart_ex_corpus_roundtrip_excel() {
    let corpus = std::env::var("DUKE_CORPUS_DIR").unwrap_or_else(|_| "data/excel-corpus".into());
    let dir = std::path::Path::new(&corpus);
    if !dir.is_dir() {
        eprintln!("corpus not found, skipping");
        return;
    }

    // Find chartEx files dynamically
    let chartex_files: Vec<_> = std::fs::read_dir(dir)
        .unwrap()
        .filter_map(|e| e.ok())
        .map(|e| e.path())
        .filter(|p| p.extension().and_then(|e| e.to_str()) == Some("xlsx") && has_chartex(p))
        .collect();

    assert!(!chartex_files.is_empty(), "no chartEx files in corpus");
    eprintln!("Found {} chartEx files", chartex_files.len());

    // Use the smallest one (waterfall, ~300KB) to avoid slow file transfer
    let smallest = chartex_files
        .iter()
        .min_by_key(|p| p.metadata().map(|m| m.len()).unwrap_or(u64::MAX))
        .unwrap();

    eprintln!(
        "Testing: {}",
        smallest.file_name().unwrap_or_default().to_string_lossy()
    );

    // Read with duke-sheets
    let wb = duke_sheets_core::Workbook::from(
        duke_sheets_xlsx::XlsxReader::read_file(smallest).expect("read"),
    );

    // Write back
    let fixture = temp_fixture();
    let mut buf = Vec::new();
    duke_sheets_xlsx::XlsxWriter::write(&wb, std::io::Cursor::new(&mut buf)).expect("write");
    std::fs::write(&fixture.host_path, &buf).expect("write to disk");
    eprintln!("Wrote {} bytes", buf.len());

    // Push to VM and open in Excel
    ensure_vm_temp_dir();
    push_file_to_vm(&fixture);

    let bridge = excel_bridge();
    let excel = bridge.lock().unwrap();
    let opened = excel
        .open_workbook(&fixture.vm_path)
        .expect("Excel should open chartEx roundtrip file");
    let name = opened.name().unwrap_or_default();
    eprintln!("Excel opened: {name}");
    assert!(
        !name.contains("Repaired"),
        "Excel repaired the file: {name}"
    );
    assert!(!opened.is_read_only().unwrap_or(false), "read-only");
    opened.close().expect("close");
    cleanup_fixture(&fixture);
    eprintln!("PASS: Excel opened chartEx roundtrip without repair");
}

fn has_chartex(path: &std::path::Path) -> bool {
    let Ok(file) = std::fs::File::open(path) else {
        return false;
    };
    let Ok(archive) = zip::ZipArchive::new(file) else {
        return false;
    };
    let result = archive
        .file_names()
        .any(|n| n.starts_with("xl/charts/chartEx") && n.ends_with(".xml"));
    result
}
