//! Writer E2E tests: duke-sheets writes XLSX → Excel opens + re-saves → duke-sheets reads back.
//!
//! Each test builds a `Workbook` in memory, writes it with `XlsxWriter`,
//! pushes it to the Windows VM where real Excel opens it (asserting no
//! repair), re-saves to normalise the XML, pulls it back, and reads it
//! with `XlsxReader` to verify styles/values survived the round trip.

use crate::roundtrip_through_excel;
use duke_sheets_core::auto_filter::{AutoFilter, ColumnFilter, FilterColumn, Top10Filter};
use duke_sheets_core::rich_text::{RichTextRun, RunFont};
use duke_sheets_core::style::{
    BorderLineStyle, BorderStyle, FillStyle, HorizontalAlignment, NumberFormat, Protection,
    Underline, VerticalAlignment,
};
use duke_sheets_core::worksheet::SheetVisibility;
use duke_sheets_core::{
    CellAddress, CellRange, CellValue, Color, ConditionalFormatRule, DataValidation, Hyperlink,
    PivotAggregate, PivotField, PivotFieldRef, PivotFilter, PivotGrouping, PivotManualGroup,
    PivotMeasure, PivotSort, PivotSource, PivotSourceRange, PivotStyle, PivotValue, Style,
    ValidationOperator, Workbook, WorkbookConnection, WorkbookConnectionKind,
};

fn range(start: &str, end: &str) -> CellRange {
    CellRange::new(
        CellAddress::parse(start).unwrap(),
        CellAddress::parse(end).unwrap(),
    )
}

fn basic_pivot_workbook() -> Workbook {
    let mut wb = Workbook::new();
    let sheet = wb.worksheet_mut(0).unwrap();
    sheet.set_cell_value("A1", "Region").unwrap();
    sheet.set_cell_value("B1", "Revenue").unwrap();
    sheet.set_cell_value("A2", "East").unwrap();
    sheet.set_cell_value("B2", 10.0).unwrap();
    sheet.set_cell_value("A3", "West").unwrap();
    sheet.set_cell_value("B3", 20.0).unwrap();

    let pivot = duke_sheets_core::PivotTable::builder("BasicPivot")
        .source_range(CellRange::parse("A1:B3").unwrap())
        .target_address("D1")
        .unwrap()
        .row("Region")
        .measure("Revenue", PivotAggregate::Sum)
        .build()
        .unwrap();
    sheet.add_pivot_table(pivot).unwrap();
    wb
}

fn sort_by_measure_pivot_workbook() -> Workbook {
    let mut wb = Workbook::new();
    let sheet = wb.worksheet_mut(0).unwrap();
    sheet.set_cell_value("A1", "Region").unwrap();
    sheet.set_cell_value("B1", "Quarter").unwrap();
    sheet.set_cell_value("C1", "Revenue").unwrap();
    sheet.set_cell_value("A2", "East").unwrap();
    sheet.set_cell_value("B2", "Q1").unwrap();
    sheet.set_cell_value("C2", 10.0).unwrap();
    sheet.set_cell_value("A3", "West").unwrap();
    sheet.set_cell_value("B3", "Q1").unwrap();
    sheet.set_cell_value("C3", 50.0).unwrap();

    let mut region =
        PivotField::new("Region").with_sort_by(PivotFieldRef::new("Revenue"), PivotAggregate::Sum);
    region.sort = PivotSort::Descending;
    let pivot = duke_sheets_core::PivotTable::builder("ValueSortedPivot")
        .source_range(CellRange::parse("A1:C3").unwrap())
        .target_address("E1")
        .unwrap()
        .row(region)
        .column("Quarter")
        .named_measure("Revenue", PivotAggregate::Sum, "Sum of Revenue")
        .build()
        .unwrap();
    sheet.add_pivot_table(pivot).unwrap();
    wb
}

fn styled_pivot_style() -> PivotStyle {
    PivotStyle {
        name: Some("PivotStyleLight16".to_string()),
        show_row_headers: false,
        show_column_headers: true,
        show_row_stripes: true,
        show_column_stripes: true,
        show_last_column: true,
    }
}

fn styled_pivot_workbook() -> Workbook {
    let mut wb = Workbook::new();
    let sheet = wb.worksheet_mut(0).unwrap();
    sheet.set_cell_value("A1", "Region").unwrap();
    sheet.set_cell_value("B1", "Revenue").unwrap();
    sheet.set_cell_value("A2", "East").unwrap();
    sheet.set_cell_value("B2", 10.0).unwrap();
    sheet.set_cell_value("A3", "West").unwrap();
    sheet.set_cell_value("B3", 20.0).unwrap();

    let pivot = duke_sheets_core::PivotTable::builder("StyledPivot")
        .source_range(CellRange::parse("A1:B3").unwrap())
        .target_address("D1")
        .unwrap()
        .row("Region")
        .named_measure("Revenue", PivotAggregate::Sum, "Total Revenue")
        .style(styled_pivot_style())
        .build()
        .unwrap();
    sheet.add_pivot_table(pivot).unwrap();
    wb
}

fn top_n_pivot_workbook() -> Workbook {
    let mut wb = Workbook::new();
    let sheet = wb.worksheet_mut(0).unwrap();
    sheet.set_cell_value("A1", "Region").unwrap();
    sheet.set_cell_value("B1", "Revenue").unwrap();
    sheet.set_cell_value("A2", "East").unwrap();
    sheet.set_cell_value("B2", 10.0).unwrap();
    sheet.set_cell_value("A3", "West").unwrap();
    sheet.set_cell_value("B3", 20.0).unwrap();
    sheet.set_cell_value("A4", "North").unwrap();
    sheet.set_cell_value("B4", 30.0).unwrap();

    let pivot = duke_sheets_core::PivotTable::builder("TopRegions")
        .source_range(CellRange::parse("A1:B4").unwrap())
        .target_address("D1")
        .unwrap()
        .row("Region")
        .named_measure("Revenue", PivotAggregate::Sum, "Total Revenue")
        .filter(PivotFilter::TopN {
            field: PivotFieldRef::new("Region"),
            measure: PivotMeasure::new("Revenue", PivotAggregate::Sum).with_name("Total Revenue"),
            n: 2,
            top: true,
            percent: false,
        })
        .build()
        .unwrap();
    sheet.add_pivot_table(pivot).unwrap();
    wb
}

fn named_consolidation_pivot_workbook() -> Workbook {
    let mut wb = Workbook::new();
    {
        let sheet = wb.worksheet_mut(0).unwrap();
        sheet.set_cell_value("A1", "Region").unwrap();
        sheet.set_cell_value("B1", "Revenue").unwrap();
        sheet.set_cell_value("A2", "East").unwrap();
        sheet.set_cell_value("B2", 10.0).unwrap();
        sheet.set_cell_value("A3", "West").unwrap();
        sheet.set_cell_value("B3", 20.0).unwrap();
    }
    wb.define_name("NamedSource", "Sheet1!$A$1:$B$3").unwrap();

    let pivot = duke_sheets_core::PivotTable::builder("NamedConsolidation")
        .source(PivotSource::Consolidation {
            ranges: vec![PivotSourceRange::named("NamedSource")],
        })
        .target_address("D1")
        .unwrap()
        .row("Region")
        .measure("Revenue", PivotAggregate::Sum)
        .build()
        .unwrap();
    wb.worksheet_mut(0).unwrap().add_pivot_table(pivot).unwrap();
    wb
}

#[test]
fn test_write_basic_pivot_survives_excel_roundtrip() {
    let result = roundtrip_through_excel(&basic_pivot_workbook());
    let pivot = result
        .worksheet(0)
        .unwrap()
        .pivot_table_by_name("BasicPivot")
        .expect("pivot survives Excel roundtrip");

    assert_eq!(pivot.rows.len(), 1);
    assert_eq!(pivot.rows[0].field.name, "Region");
    assert_eq!(pivot.measures.len(), 1);
    assert_eq!(pivot.measures[0].field.name, "Revenue");
    assert_eq!(pivot.measures[0].aggregate, PivotAggregate::Sum);
}

#[test]
fn test_write_pivot_sort_by_measure_survives_excel_roundtrip() {
    let result = roundtrip_through_excel(&sort_by_measure_pivot_workbook());
    let pivot = result
        .worksheet(0)
        .unwrap()
        .pivot_table_by_name("ValueSortedPivot")
        .expect("pivot survives Excel roundtrip");

    assert_eq!(pivot.rows.len(), 1);
    assert_eq!(pivot.rows[0].field.name, "Region");
    assert_eq!(pivot.rows[0].sort, PivotSort::Descending);
    let measure = pivot.rows[0]
        .sort_by_measure
        .as_ref()
        .expect("sort-by-measure survives Excel roundtrip");
    assert_eq!(measure.field.name, "Revenue");
    assert_eq!(measure.aggregate, PivotAggregate::Sum);
    assert_eq!(measure.name.as_deref(), Some("Sum of Revenue"));
}

#[test]
fn test_write_pivot_style_survives_excel_roundtrip() {
    let result = roundtrip_through_excel(&styled_pivot_workbook());
    let pivot = result
        .worksheet(0)
        .unwrap()
        .pivot_table_by_name("StyledPivot")
        .expect("pivot survives Excel roundtrip");

    assert_eq!(pivot.style, styled_pivot_style());
}

#[test]
fn test_write_pivot_top_n_filter_survives_excel_roundtrip() {
    let result = roundtrip_through_excel(&top_n_pivot_workbook());
    let pivot = result
        .worksheet(0)
        .unwrap()
        .pivot_table_by_name("TopRegions")
        .expect("pivot survives Excel roundtrip");

    assert_eq!(pivot.filters.len(), 1);
    match &pivot.filters[0] {
        PivotFilter::TopN {
            field,
            measure,
            n,
            top,
            percent,
        } => {
            assert_eq!(field.name, "Region");
            assert_eq!(measure.field.name, "Revenue");
            assert_eq!(measure.aggregate, PivotAggregate::Sum);
            assert_eq!(measure.name.as_deref(), Some("Total Revenue"));
            assert_eq!(*n, 2);
            assert!(*top);
            assert!(!*percent);
        }
        other => panic!("unexpected pivot filter after Excel roundtrip: {other:?}"),
    }
}

#[test]
fn test_write_named_consolidation_pivot_survives_excel_roundtrip() {
    let result = roundtrip_through_excel(&named_consolidation_pivot_workbook());
    let pivot = result
        .worksheet(0)
        .unwrap()
        .pivot_table_by_name("NamedConsolidation")
        .expect("pivot survives Excel roundtrip");

    match &pivot.source {
        PivotSource::Consolidation { ranges } => {
            assert_eq!(ranges.len(), 1);
            assert_eq!(ranges[0].name.as_deref(), Some("NamedSource"));
        }
        other => panic!("unexpected pivot source after Excel roundtrip: {other:?}"),
    }
    assert_eq!(pivot.rows.len(), 1);
    assert_eq!(pivot.rows[0].field.name, "Region");
    assert_eq!(pivot.measures.len(), 1);
    assert_eq!(pivot.measures[0].field.name, "Revenue");
}

#[test]
fn test_write_olap_connection_metadata_survives_excel_roundtrip() {
    let mut wb = Workbook::new();
    let mut connection = WorkbookConnection::olap(10, "CubeSales").with_connection_type(5);
    connection.kind = WorkbookConnectionKind::Olap {
        connection: Some("Provider=MSOLAP;Data Source=olapserver;".to_string()),
        command: Some("SalesCube".to_string()),
        command_type: Some(1),
        local: false,
        local_connection: None,
        local_refresh: true,
        send_locale: true,
        row_drill_count: Some(1000),
    };
    wb.add_data_connection(connection).unwrap();

    let result = roundtrip_through_excel(&wb);
    let connection = result
        .data_connection_by_name("CubeSales")
        .expect("CubeSales connection");
    match &connection.kind {
        WorkbookConnectionKind::Olap {
            connection,
            command,
            command_type,
            local,
            local_connection,
            local_refresh,
            send_locale,
            row_drill_count,
        } => {
            assert_eq!(
                connection.as_deref(),
                Some("Provider=MSOLAP;Data Source=olapserver;")
            );
            assert_eq!(command.as_deref(), Some("SalesCube"));
            assert_eq!(*command_type, Some(1));
            assert!(!*local);
            assert_eq!(local_connection.as_deref(), None);
            assert!(*local_refresh);
            assert!(*send_locale);
            assert_eq!(*row_drill_count, Some(1000));
        }
        other => panic!("unexpected connection kind: {other:?}"),
    }
}

#[test]
fn test_write_manual_pivot_grouping_survives_excel_roundtrip() {
    let mut wb = Workbook::new();
    let sheet = wb.worksheet_mut(0).unwrap();
    sheet.set_cell_value("A1", "Region").unwrap();
    sheet.set_cell_value("B1", "Revenue").unwrap();
    sheet.set_cell_value("A2", "East").unwrap();
    sheet.set_cell_value("B2", 10.0).unwrap();
    sheet.set_cell_value("A3", "West").unwrap();
    sheet.set_cell_value("B3", 20.0).unwrap();
    sheet.set_cell_value("A4", "Central").unwrap();
    sheet.set_cell_value("B4", 5.0).unwrap();

    let pivot = duke_sheets_core::PivotTable::builder("ManualGroupedRegions")
        .source_range(CellRange::parse("A1:B4").unwrap())
        .target_address("D1")
        .unwrap()
        .row("Region")
        .measure("Revenue", PivotAggregate::Sum)
        .grouping(PivotGrouping::Manual {
            field: "Region".into(),
            groups: vec![PivotManualGroup::new("Coastal", ["East", "West"])],
        })
        .build()
        .unwrap();
    sheet.add_pivot_table(pivot).unwrap();

    let result = roundtrip_through_excel(&wb);
    let pivot = result
        .worksheet(0)
        .unwrap()
        .pivot_table_by_name("ManualGroupedRegions")
        .expect("pivot survives Excel roundtrip");

    assert_eq!(pivot.groupings.len(), 1);
    match &pivot.groupings[0] {
        PivotGrouping::Manual { field, groups } => {
            assert_eq!(field.name, "Region");
            assert_eq!(groups.len(), 1);
            assert_eq!(groups[0].name, "Coastal");
            assert_eq!(
                groups[0].members,
                vec![
                    PivotValue::String("East".to_string()),
                    PivotValue::String("West".to_string())
                ]
            );
        }
        other => panic!("unexpected grouping after Excel roundtrip: {other:?}"),
    }
}

// Font tests

#[test]
fn test_write_font_bold() {
    let mut wb = Workbook::new();
    let sheet = wb.worksheet_mut(0).unwrap();
    sheet.set_cell_value("A1", "Bold").unwrap();
    sheet
        .set_cell_style("A1", &Style::new().bold(true))
        .unwrap();

    let result = roundtrip_through_excel(&wb);
    let s = result.worksheet(0).unwrap();
    let style = s.cell_style_at(0, 0).expect("A1 should have style");
    assert!(style.font.bold, "Font should be bold");
}

#[test]
fn test_write_font_italic() {
    let mut wb = Workbook::new();
    let sheet = wb.worksheet_mut(0).unwrap();
    sheet.set_cell_value("A1", "Italic").unwrap();
    sheet
        .set_cell_style("A1", &Style::new().italic(true))
        .unwrap();

    let result = roundtrip_through_excel(&wb);
    let s = result.worksheet(0).unwrap();
    let style = s.cell_style_at(0, 0).expect("A1 should have style");
    assert!(style.font.italic, "Font should be italic");
}

#[test]
fn test_write_font_size() {
    let mut wb = Workbook::new();
    let sheet = wb.worksheet_mut(0).unwrap();
    sheet.set_cell_value("A1", "Big").unwrap();
    sheet
        .set_cell_style("A1", &Style::new().font_size(20.0))
        .unwrap();

    let result = roundtrip_through_excel(&wb);
    let s = result.worksheet(0).unwrap();
    let style = s.cell_style_at(0, 0).expect("A1 should have style");
    assert!(
        (style.font.size - 20.0).abs() < 0.5,
        "Expected font size ~20, got {}",
        style.font.size
    );
}

#[test]
fn test_write_font_name() {
    let mut wb = Workbook::new();
    let sheet = wb.worksheet_mut(0).unwrap();
    sheet.set_cell_value("A1", "Courier").unwrap();
    sheet
        .set_cell_style("A1", &Style::new().font_name("Courier New"))
        .unwrap();

    let result = roundtrip_through_excel(&wb);
    let s = result.worksheet(0).unwrap();
    let style = s.cell_style_at(0, 0).expect("A1 should have style");
    assert_eq!(style.font.name, "Courier New");
}

#[test]
fn test_write_font_color() {
    let mut wb = Workbook::new();
    let sheet = wb.worksheet_mut(0).unwrap();
    sheet.set_cell_value("A1", "Red text").unwrap();
    sheet
        .set_cell_style("A1", &Style::new().font_color(Color::rgb(255, 0, 0)))
        .unwrap();

    let result = roundtrip_through_excel(&wb);
    let s = result.worksheet(0).unwrap();
    let style = s.cell_style_at(0, 0).expect("A1 should have style");
    let (r, g, b) = style.font.color.to_rgb();
    assert!(
        r > 200 && g < 50 && b < 50,
        "Expected red font, got ({r}, {g}, {b})"
    );
}

#[test]
fn test_write_font_strikethrough() {
    let mut wb = Workbook::new();
    let sheet = wb.worksheet_mut(0).unwrap();
    sheet.set_cell_value("A1", "Strike").unwrap();
    let mut style = Style::new();
    style.font.strikethrough = true;
    sheet.set_cell_style("A1", &style).unwrap();

    let result = roundtrip_through_excel(&wb);
    let s = result.worksheet(0).unwrap();
    let st = s.cell_style_at(0, 0).expect("A1 should have style");
    assert!(st.font.strikethrough, "Font should be strikethrough");
}

#[test]
fn test_write_font_underline() {
    let mut wb = Workbook::new();
    let sheet = wb.worksheet_mut(0).unwrap();
    sheet.set_cell_value("A1", "Underline").unwrap();
    let mut style = Style::new();
    style.font.underline = Underline::Single;
    sheet.set_cell_style("A1", &style).unwrap();

    let result = roundtrip_through_excel(&wb);
    let s = result.worksheet(0).unwrap();
    let st = s.cell_style_at(0, 0).expect("A1 should have style");
    assert_eq!(st.font.underline, Underline::Single);
}

#[test]
fn test_write_font_combination() {
    let mut wb = Workbook::new();
    let sheet = wb.worksheet_mut(0).unwrap();
    sheet.set_cell_value("A1", "Combo").unwrap();
    let mut style = Style::new()
        .bold(true)
        .italic(true)
        .font_size(14.0)
        .font_color(Color::rgb(0, 0, 255));
    style.font.underline = Underline::Single;
    sheet.set_cell_style("A1", &style).unwrap();

    let result = roundtrip_through_excel(&wb);
    let s = result.worksheet(0).unwrap();
    let st = s.cell_style_at(0, 0).expect("A1 should have style");
    assert!(st.font.bold, "Should be bold");
    assert!(st.font.italic, "Should be italic");
    assert_eq!(st.font.underline, Underline::Single);
    let (r, g, b) = st.font.color.to_rgb();
    assert!(b > 200 && r < 50, "Should be blue, got ({r}, {g}, {b})");
    assert!(
        (st.font.size - 14.0).abs() < 0.5,
        "Expected size ~14, got {}",
        st.font.size
    );
}

// Fill tests

#[test]
fn test_write_solid_fill() {
    let mut wb = Workbook::new();
    let sheet = wb.worksheet_mut(0).unwrap();
    sheet.set_cell_value("A1", "Red fill").unwrap();
    sheet
        .set_cell_style("A1", &Style::new().fill_color(Color::rgb(255, 0, 0)))
        .unwrap();

    let result = roundtrip_through_excel(&wb);
    let s = result.worksheet(0).unwrap();
    let style = s.cell_style_at(0, 0).expect("A1 should have style");
    match &style.fill {
        FillStyle::Solid { color } => {
            let (r, g, b) = color.to_rgb();
            assert!(
                r > 200 && g < 50 && b < 50,
                "Expected red fill, got ({r}, {g}, {b})"
            );
        }
        other => panic!("Expected Solid fill, got {other:?}"),
    }
}

#[test]
fn test_write_fill_with_font_color() {
    let mut wb = Workbook::new();
    let sheet = wb.worksheet_mut(0).unwrap();
    sheet.set_cell_value("A1", "White on Blue").unwrap();
    sheet
        .set_cell_style(
            "A1",
            &Style::new()
                .fill_color(Color::rgb(0, 0, 255))
                .font_color(Color::rgb(255, 255, 255)),
        )
        .unwrap();

    let result = roundtrip_through_excel(&wb);
    let s = result.worksheet(0).unwrap();
    let style = s.cell_style_at(0, 0).expect("A1 should have style");

    match &style.fill {
        FillStyle::Solid { color } => {
            let (_, _, b) = color.to_rgb();
            assert!(b > 200, "Expected blue fill");
        }
        other => panic!("Expected Solid fill, got {other:?}"),
    }
    let (r, g, b) = style.font.color.to_rgb();
    assert!(
        r > 200 && g > 200 && b > 200,
        "Expected white font, got ({r}, {g}, {b})"
    );
}

// Border tests

#[test]
fn test_write_border_thin_all() {
    let mut wb = Workbook::new();
    let sheet = wb.worksheet_mut(0).unwrap();
    sheet.set_cell_value("A1", "Bordered").unwrap();
    sheet
        .set_cell_style(
            "A1",
            &Style {
                border: BorderStyle::all(BorderLineStyle::Thin, Color::Auto),
                ..Default::default()
            },
        )
        .unwrap();

    let result = roundtrip_through_excel(&wb);
    let s = result.worksheet(0).unwrap();
    let style = s.cell_style_at(0, 0).expect("A1 should have style");
    assert!(style.border.left.is_some(), "left border");
    assert!(style.border.right.is_some(), "right border");
    assert!(style.border.top.is_some(), "top border");
    assert!(style.border.bottom.is_some(), "bottom border");
}

#[test]
fn test_write_border_color() {
    let mut wb = Workbook::new();
    let sheet = wb.worksheet_mut(0).unwrap();
    sheet.set_cell_value("A1", "Blue border").unwrap();
    sheet
        .set_cell_style(
            "A1",
            &Style {
                border: BorderStyle::all(BorderLineStyle::Thin, Color::rgb(0, 0, 255)),
                ..Default::default()
            },
        )
        .unwrap();

    let result = roundtrip_through_excel(&wb);
    let s = result.worksheet(0).unwrap();
    let style = s.cell_style_at(0, 0).expect("A1 should have style");
    let edge = style.border.left.as_ref().expect("left border");
    let (_, _, b) = edge.color.to_rgb();
    assert!(b > 200, "Expected blue border color, got b={b}");
}

#[test]
fn test_write_border_individual() {
    let mut wb = Workbook::new();
    let sheet = wb.worksheet_mut(0).unwrap();
    sheet.set_cell_value("A1", "Left only").unwrap();
    sheet
        .set_cell_style(
            "A1",
            &Style {
                border: BorderStyle::new().with_left(BorderLineStyle::Thin, Color::Auto),
                ..Default::default()
            },
        )
        .unwrap();

    let result = roundtrip_through_excel(&wb);
    let s = result.worksheet(0).unwrap();
    let style = s.cell_style_at(0, 0).expect("A1 should have style");
    assert!(style.border.left.is_some(), "left border should be set");
    assert!(style.border.right.is_none(), "right border should be empty");
    assert!(style.border.top.is_none(), "top border should be empty");
    assert!(
        style.border.bottom.is_none(),
        "bottom border should be empty"
    );
}

// Alignment tests

#[test]
fn test_write_alignment_horizontal_center() {
    let mut wb = Workbook::new();
    let sheet = wb.worksheet_mut(0).unwrap();
    sheet.set_cell_value("A1", "Centered").unwrap();
    sheet
        .set_cell_style(
            "A1",
            &Style::new().horizontal_alignment(HorizontalAlignment::Center),
        )
        .unwrap();

    let result = roundtrip_through_excel(&wb);
    let s = result.worksheet(0).unwrap();
    let style = s.cell_style_at(0, 0).expect("A1 should have style");
    assert_eq!(style.alignment.horizontal, HorizontalAlignment::Center);
}

#[test]
fn test_write_alignment_horizontal_right() {
    let mut wb = Workbook::new();
    let sheet = wb.worksheet_mut(0).unwrap();
    sheet.set_cell_value("A1", "Right").unwrap();
    sheet
        .set_cell_style(
            "A1",
            &Style::new().horizontal_alignment(HorizontalAlignment::Right),
        )
        .unwrap();

    let result = roundtrip_through_excel(&wb);
    let s = result.worksheet(0).unwrap();
    let style = s.cell_style_at(0, 0).expect("A1 should have style");
    assert_eq!(style.alignment.horizontal, HorizontalAlignment::Right);
}

#[test]
fn test_write_alignment_vertical_top() {
    let mut wb = Workbook::new();
    let sheet = wb.worksheet_mut(0).unwrap();
    sheet.set_cell_value("A1", "Top").unwrap();
    sheet
        .set_cell_style(
            "A1",
            &Style::new().vertical_alignment(VerticalAlignment::Top),
        )
        .unwrap();

    let result = roundtrip_through_excel(&wb);
    let s = result.worksheet(0).unwrap();
    let style = s.cell_style_at(0, 0).expect("A1 should have style");
    assert_eq!(style.alignment.vertical, VerticalAlignment::Top);
}

#[test]
fn test_write_alignment_wrap() {
    let mut wb = Workbook::new();
    let sheet = wb.worksheet_mut(0).unwrap();
    sheet.set_cell_value("A1", "Wrap\nText").unwrap();
    sheet
        .set_cell_style("A1", &Style::new().wrap_text(true))
        .unwrap();

    let result = roundtrip_through_excel(&wb);
    let s = result.worksheet(0).unwrap();
    let style = s.cell_style_at(0, 0).expect("A1 should have style");
    assert!(style.alignment.wrap_text, "wrap_text should be true");
}

#[test]
fn test_write_alignment_shrink() {
    let mut wb = Workbook::new();
    let sheet = wb.worksheet_mut(0).unwrap();
    sheet.set_cell_value("A1", "Shrink").unwrap();
    let mut style = Style::new();
    style.alignment.shrink_to_fit = true;
    sheet.set_cell_style("A1", &style).unwrap();

    let result = roundtrip_through_excel(&wb);
    let s = result.worksheet(0).unwrap();
    let st = s.cell_style_at(0, 0).expect("A1 should have style");
    assert!(st.alignment.shrink_to_fit, "shrink_to_fit should be true");
}

#[test]
fn test_write_alignment_rotation() {
    let mut wb = Workbook::new();
    let sheet = wb.worksheet_mut(0).unwrap();
    sheet.set_cell_value("A1", "Rotated").unwrap();
    let mut style = Style::new();
    style.alignment.rotation = 45;
    sheet.set_cell_style("A1", &style).unwrap();

    let result = roundtrip_through_excel(&wb);
    let s = result.worksheet(0).unwrap();
    let st = s.cell_style_at(0, 0).expect("A1 should have style");
    assert_eq!(st.alignment.rotation, 45);
}

#[test]
fn test_write_alignment_indent() {
    let mut wb = Workbook::new();
    let sheet = wb.worksheet_mut(0).unwrap();
    sheet.set_cell_value("A1", "Indented").unwrap();
    let mut style = Style::new();
    style.alignment.indent = 2;
    sheet.set_cell_style("A1", &style).unwrap();

    let result = roundtrip_through_excel(&wb);
    let s = result.worksheet(0).unwrap();
    let st = s.cell_style_at(0, 0).expect("A1 should have style");
    assert_eq!(st.alignment.indent, 2);
}

// Number format tests

#[test]
fn test_write_number_format_percent() {
    let mut wb = Workbook::new();
    let sheet = wb.worksheet_mut(0).unwrap();
    sheet.set_cell_value("A1", 0.1234).unwrap();
    sheet
        .set_cell_style("A1", &Style::new().number_format("0.00%"))
        .unwrap();

    let result = roundtrip_through_excel(&wb);
    let s = result.worksheet(0).unwrap();
    let style = s.cell_style_at(0, 0).expect("A1 should have style");
    let fmt = style.number_format.format_string();
    assert!(fmt.contains('%'), "Expected percentage format, got '{fmt}'");
}

#[test]
fn test_write_number_format_date() {
    let mut wb = Workbook::new();
    let sheet = wb.worksheet_mut(0).unwrap();
    sheet.set_cell_value("A1", 45000.0).unwrap(); // a date serial
    sheet
        .set_cell_style(
            "A1",
            &Style {
                number_format: NumberFormat::BuiltIn(14), // m/d/yyyy
                ..Default::default()
            },
        )
        .unwrap();

    let result = roundtrip_through_excel(&wb);
    let s = result.worksheet(0).unwrap();
    let style = s.cell_style_at(0, 0).expect("A1 should have style");
    assert_ne!(
        style.number_format,
        NumberFormat::General,
        "Should have a date format"
    );
}

#[test]
fn test_write_number_format_currency() {
    let mut wb = Workbook::new();
    let sheet = wb.worksheet_mut(0).unwrap();
    sheet.set_cell_value("A1", 1234.56).unwrap();
    sheet
        .set_cell_style("A1", &Style::new().number_format("#,##0.00"))
        .unwrap();

    let result = roundtrip_through_excel(&wb);
    let s = result.worksheet(0).unwrap();
    let style = s.cell_style_at(0, 0).expect("A1 should have style");
    let fmt = style.number_format.format_string();
    assert!(
        fmt.contains("#,##0"),
        "Expected thousands format, got '{fmt}'"
    );
}

#[test]
fn test_write_number_format_custom() {
    let mut wb = Workbook::new();
    let sheet = wb.worksheet_mut(0).unwrap();
    sheet.set_cell_value("A1", 3.14159).unwrap();
    sheet
        .set_cell_style("A1", &Style::new().number_format("0.000"))
        .unwrap();

    let result = roundtrip_through_excel(&wb);
    let s = result.worksheet(0).unwrap();
    let style = s.cell_style_at(0, 0).expect("A1 should have style");
    let fmt = style.number_format.format_string();
    assert!(
        fmt.contains("0.000"),
        "Expected '0.000' format, got '{fmt}'"
    );
}

// Dimension tests

#[test]
fn test_write_row_height() {
    let mut wb = Workbook::new();
    let sheet = wb.worksheet_mut(0).unwrap();
    sheet.set_cell_value("A1", "Tall row").unwrap();
    sheet.set_row_height(0, 40.0);

    let result = roundtrip_through_excel(&wb);
    let s = result.worksheet(0).unwrap();
    let height = s.row_height(0);
    assert!(
        (height - 40.0).abs() < 1.0,
        "Expected row height ~40, got {height}"
    );
}

#[test]
fn test_write_column_width() {
    let mut wb = Workbook::new();
    let sheet = wb.worksheet_mut(0).unwrap();
    sheet.set_cell_value("A1", "Wide column").unwrap();
    sheet.set_column_width(0, 25.0);

    let result = roundtrip_through_excel(&wb);
    let s = result.worksheet(0).unwrap();
    let width = s.column_width(0);
    // Excel may adjust width slightly during re-save
    assert!(
        (width - 25.0).abs() < 2.0,
        "Expected column width ~25, got {width}"
    );
}

#[test]
fn test_write_hidden_row() {
    let mut wb = Workbook::new();
    let sheet = wb.worksheet_mut(0).unwrap();
    sheet.set_cell_value("A1", "Visible").unwrap();
    sheet.set_cell_value("A2", "Hidden").unwrap();
    sheet.set_row_hidden(1, true);

    let result = roundtrip_through_excel(&wb);
    let s = result.worksheet(0).unwrap();
    assert!(!s.is_row_hidden(0), "Row 0 should be visible");
    assert!(s.is_row_hidden(1), "Row 1 should be hidden");
}

// Merged cells

#[test]
fn test_write_merged_cells() {
    let mut wb = Workbook::new();
    let sheet = wb.worksheet_mut(0).unwrap();
    sheet.set_cell_value("A1", "Merged region").unwrap();
    sheet
        .merge_cells(&CellRange::parse("A1:C1").unwrap())
        .unwrap();

    let result = roundtrip_through_excel(&wb);
    let s = result.worksheet(0).unwrap();
    let regions = s.merged_regions();
    assert!(
        !regions.is_empty(),
        "Should have at least one merged region"
    );
    let expected = CellRange::parse("A1:C1").unwrap();
    assert!(
        regions.iter().any(|r| *r == expected),
        "Should contain A1:C1, got {regions:?}"
    );
}

// Conditional formatting

#[test]
fn test_write_conditional_format() {
    let mut wb = Workbook::new();
    let sheet = wb.worksheet_mut(0).unwrap();
    sheet.set_cell_value("A1", 10.0).unwrap();
    sheet.set_cell_value("A2", 60.0).unwrap();
    sheet.set_cell_value("A3", 90.0).unwrap();

    let rule = ConditionalFormatRule::cell_is_greater_than("50")
        .with_range(CellRange::parse("A1:A3").unwrap())
        .with_format(Style::new().fill_color(Color::rgb(0, 255, 0)));
    sheet.add_conditional_format(rule);

    let result = roundtrip_through_excel(&wb);
    let s = result.worksheet(0).unwrap();
    let cfs = s.conditional_formats();
    assert!(
        !cfs.is_empty(),
        "Should have at least one conditional format rule"
    );
}

// Data validation

#[test]
fn test_write_data_validation_list() {
    let mut wb = Workbook::new();
    let sheet = wb.worksheet_mut(0).unwrap();
    sheet.set_cell_value("A1", "Apple").unwrap();

    let dv = DataValidation::list("Apple,Banana,Cherry")
        .with_range(CellRange::parse("A1").unwrap())
        .with_input_message("Pick a fruit", "Choose from the list")
        .with_error_message("Invalid", "Must be a fruit");
    sheet.add_data_validation(dv);

    let result = roundtrip_through_excel(&wb);
    let s = result.worksheet(0).unwrap();
    let dvs = s.data_validations();
    assert!(!dvs.is_empty(), "Should have at least one data validation");
}

#[test]
fn test_write_data_validation_number() {
    let mut wb = Workbook::new();
    let sheet = wb.worksheet_mut(0).unwrap();
    sheet.set_cell_value("A1", 50.0).unwrap();

    let dv = DataValidation::whole_number(ValidationOperator::Between, "1")
        .with_range(CellRange::parse("A1").unwrap());
    sheet.add_data_validation(dv);

    let result = roundtrip_through_excel(&wb);
    let s = result.worksheet(0).unwrap();
    let dvs = s.data_validations();
    assert!(!dvs.is_empty(), "Should have at least one data validation");
}

// Page setup / header-footer tests

#[test]
fn test_write_header_footer_even_first_and_flags() {
    let mut wb = Workbook::new();
    let sheet = wb.worksheet_mut(0).unwrap();
    sheet.set_cell_value("A1", "Page setup test").unwrap();

    let mut ps = sheet.page_setup().clone();
    ps.odd_header = Some("&COdd Header".to_string());
    ps.odd_footer = Some("&COdd Footer".to_string());
    ps.even_header = Some("&CEven Header".to_string());
    ps.even_footer = Some("&CEven Footer".to_string());
    ps.first_header = Some("&CFirst Page".to_string());
    ps.first_footer = Some("&CFirst Footer".to_string());
    ps.different_odd_even = true;
    ps.different_first = true;
    ps.scale_with_doc = false;
    ps.align_with_margins = false;
    sheet.set_page_setup(ps);

    let result = roundtrip_through_excel(&wb);
    let s = result.worksheet(0).unwrap();
    let ps2 = s.page_setup();

    // Verify all header/footer strings survived Excel roundtrip
    assert_eq!(
        ps2.odd_header.as_deref(),
        Some("&COdd Header"),
        "odd_header"
    );
    assert_eq!(
        ps2.odd_footer.as_deref(),
        Some("&COdd Footer"),
        "odd_footer"
    );
    assert_eq!(
        ps2.even_header.as_deref(),
        Some("&CEven Header"),
        "even_header"
    );
    assert_eq!(
        ps2.even_footer.as_deref(),
        Some("&CEven Footer"),
        "even_footer"
    );
    assert_eq!(
        ps2.first_header.as_deref(),
        Some("&CFirst Page"),
        "first_header"
    );
    assert_eq!(
        ps2.first_footer.as_deref(),
        Some("&CFirst Footer"),
        "first_footer"
    );

    // Verify flags
    assert!(ps2.different_odd_even, "different_odd_even");
    assert!(ps2.different_first, "different_first");
    assert!(!ps2.scale_with_doc, "scale_with_doc should be false");
    assert!(
        !ps2.align_with_margins,
        "align_with_margins should be false"
    );
}

#[test]
fn test_read_header_footer_from_excel() {
    let bridge = crate::excel_bridge();
    let fixture = crate::temp_fixture();

    {
        let excel = bridge.lock().unwrap();
        crate::ensure_vm_temp_dir();
        let wb = excel.create_workbook().expect("create workbook");

        // Set headers/footers via PageSetup properties (left/center/right)
        wb.set_page_setup_property("CenterHeader", serde_json::Value::from("Center Header"))
            .expect("set CenterHeader");
        wb.set_page_setup_property("LeftFooter", serde_json::Value::from("Left Footer"))
            .expect("set LeftFooter");
        wb.set_page_setup_property("RightFooter", serde_json::Value::from("Right Footer"))
            .expect("set RightFooter");

        // Enable different first page flag
        wb.set_page_setup_property(
            "DifferentFirstPageHeaderFooter",
            serde_json::Value::from(true),
        )
        .expect("set DifferentFirstPageHeaderFooter");

        // Disable scale with doc
        wb.set_page_setup_property("ScaleWithDocHeaderFooter", serde_json::Value::from(false))
            .expect("set ScaleWithDocHeaderFooter");

        wb.save(&fixture.vm_path).expect("save");
        wb.close().expect("close");
    }

    crate::pull_file_from_vm(&fixture);
    let workbook = duke_sheets_xlsx::XlsxReader::read_file(&fixture.host_path).expect("read");
    let sheet = workbook.worksheet(0).expect("worksheet");
    let ps = sheet.page_setup();

    // Odd header: Excel wraps CenterHeader in &C code
    let odd_header = ps.odd_header.as_deref().unwrap_or("");
    assert!(
        odd_header.contains("Center Header"),
        "odd_header should contain 'Center Header', got: {odd_header:?}"
    );

    // Odd footer: left and right sections combined
    let odd_footer = ps.odd_footer.as_deref().unwrap_or("");
    assert!(
        odd_footer.contains("Left Footer"),
        "odd_footer should contain 'Left Footer', got: {odd_footer:?}"
    );
    assert!(
        odd_footer.contains("Right Footer"),
        "odd_footer should contain 'Right Footer', got: {odd_footer:?}"
    );

    // Flags
    assert!(ps.different_first, "different_first should be true");
    assert!(!ps.scale_with_doc, "scale_with_doc should be false");

    crate::cleanup_fixture(&fixture);
}

// Feature parity tests (XLSX writer): verify Excel preserves the
// feature itself, not just cell values.

#[test]
fn test_write_external_url_hyperlink() {
    let mut wb = Workbook::new();
    let ws = wb.worksheet_mut(0).unwrap();
    ws.set_cell_value("A1", "click").unwrap();
    ws.set_hyperlink(
        "A1",
        Hyperlink {
            target: "https://example.com/path".into(),
            display: Some("click".into()),
            tooltip: None,
            location: None,
        },
    )
    .unwrap();

    let result = roundtrip_through_excel(&wb);
    let s = result.worksheet(0).unwrap();
    assert_eq!(s.get_value_at(0, 0).as_string(), Some("click"));
    let hl = s
        .hyperlink("A1")
        .expect("hyperlink must survive XLSX Excel round-trip");
    assert!(
        hl.target.contains("example.com"),
        "hyperlink target lost: {:?}",
        hl.target
    );
}

#[test]
fn test_write_internal_hyperlink() {
    let mut wb = Workbook::new();
    wb.rename_worksheet(0, "Main").unwrap();
    wb.add_worksheet_with_name("Other").unwrap();
    let ws = wb.worksheet_mut(0).unwrap();
    ws.set_cell_value("A1", "internal").unwrap();
    ws.set_hyperlink(
        "A1",
        Hyperlink {
            target: "#Other!B5".into(),
            display: Some("internal".into()),
            tooltip: None,
            location: None,
        },
    )
    .unwrap();

    let result = roundtrip_through_excel(&wb);
    let s = result.worksheet_by_name("Main").unwrap();
    assert_eq!(s.get_value_at(0, 0).as_string(), Some("internal"));
    let hl = s
        .hyperlink("A1")
        .expect("internal hyperlink must survive XLSX Excel round-trip");
    let combined = format!("{}|{}", hl.target, hl.location.as_deref().unwrap_or(""));
    assert!(
        combined.contains("Other") && combined.contains("B5"),
        "internal target lost: target={:?} location={:?}",
        hl.target,
        hl.location
    );
}

#[test]
fn test_write_cross_sheet_formula() {
    let mut wb = Workbook::new();
    wb.rename_worksheet(0, "Calc").unwrap();
    wb.add_worksheet_with_name("Data").unwrap();
    wb.worksheet_mut(1)
        .unwrap()
        .set_cell_value("A1", 10.0)
        .unwrap();
    wb.worksheet_mut(1)
        .unwrap()
        .set_cell_value("A2", 20.0)
        .unwrap();
    wb.worksheet_mut(1)
        .unwrap()
        .set_cell_value("A3", 30.0)
        .unwrap();
    let calc = wb.worksheet_mut(0).unwrap();
    calc.set_cell_formula("B1", "=Data!A1").unwrap();
    calc.set_formula_result(0, 1, CellValue::Number(10.0))
        .unwrap();
    calc.set_cell_formula("B2", "=SUM(Data!A1:A3)").unwrap();
    calc.set_formula_result(1, 1, CellValue::Number(60.0))
        .unwrap();

    let result = roundtrip_through_excel(&wb);
    let s = result.worksheet_by_name("Calc").unwrap();
    let v1 = s.get_value_at(0, 1);
    match v1.effective_value() {
        CellValue::Number(n) => assert!((n - 10.0).abs() < 1e-9, "B1 = {n}"),
        other => panic!("B1 expected Number(10), got {other:?}"),
    }
    let v2 = s.get_value_at(1, 1);
    match v2.effective_value() {
        CellValue::Number(n) => assert!((n - 60.0).abs() < 1e-9, "B2 = {n}"),
        other => panic!("B2 expected Number(60), got {other:?}"),
    }
}

#[test]
fn test_write_autofilter() {
    let mut wb = Workbook::new();
    let ws = wb.worksheet_mut(0).unwrap();
    ws.set_cell_value("A1", "Name").unwrap();
    ws.set_cell_value("B1", "Value").unwrap();
    ws.set_cell_value("A2", "alpha").unwrap();
    ws.set_cell_value("B2", 10.0).unwrap();
    ws.set_cell_value("A3", "beta").unwrap();
    ws.set_cell_value("B3", 20.0).unwrap();
    let mut af = AutoFilter::new(range("A1", "B3"));
    af.filter_columns.push(FilterColumn::new(
        1,
        ColumnFilter::Top10(Top10Filter {
            top: true,
            percent: false,
            val: 1.0,
            filter_val: None,
        }),
    ));
    ws.set_auto_filter(Some(af));

    let result = roundtrip_through_excel(&wb);
    let s = result.worksheet(0).unwrap();
    let af = s
        .auto_filter()
        .expect("autofilter must survive XLSX Excel round-trip");
    assert_eq!(af.range.start, CellAddress::parse("A1").unwrap());
}

#[test]
fn test_write_rich_text_runs() {
    let mut wb = Workbook::new();
    let ws = wb.worksheet_mut(0).unwrap();
    let bold = RunFont {
        bold: Some(true),
        ..Default::default()
    };
    let italic = RunFont {
        italic: Some(true),
        ..Default::default()
    };
    let red_big = RunFont {
        size: Some(16.0),
        color: Some(Color::Indexed(2)),
        ..Default::default()
    };
    ws.set_cell_value_at(
        0,
        0,
        CellValue::rich_text(vec![
            RichTextRun {
                text: "plain ".into(),
                font: None,
            },
            RichTextRun {
                text: "bold ".into(),
                font: Some(bold),
            },
            RichTextRun {
                text: "italic ".into(),
                font: Some(italic),
            },
            RichTextRun {
                text: "loud".into(),
                font: Some(red_big),
            },
        ]),
    )
    .unwrap();

    let result = roundtrip_through_excel(&wb);
    let s = result.worksheet(0).unwrap();
    let value = s.get_value_at(0, 0);
    assert_eq!(format!("{value}"), "plain bold italic loud");
    let runs = match &value {
        CellValue::RichText(runs) => runs,
        other => panic!("expected RichText after Excel round-trip, got {other:?}"),
    };
    assert!(
        runs.len() >= 4,
        "expected ≥4 runs, got {}: {runs:?}",
        runs.len()
    );
    let has_bold = runs
        .iter()
        .any(|r| matches!(&r.font, Some(f) if f.bold == Some(true)));
    let has_italic = runs
        .iter()
        .any(|r| matches!(&r.font, Some(f) if f.italic == Some(true)));
    let has_big = runs
        .iter()
        .any(|r| matches!(&r.font, Some(f) if matches!(f.size, Some(s) if s >= 14.0)));
    assert!(has_bold, "bold run lost: {runs:?}");
    assert!(has_italic, "italic run lost: {runs:?}");
    assert!(has_big, "size-≥14 run lost: {runs:?}");
}

#[test]
fn test_write_named_range_formula() {
    let mut wb = Workbook::new();
    wb.rename_worksheet(0, "Calc").unwrap();
    wb.add_worksheet_with_name("Data").unwrap();
    wb.worksheet_mut(1)
        .unwrap()
        .set_cell_value("A1", 5.0)
        .unwrap();
    wb.worksheet_mut(1)
        .unwrap()
        .set_cell_value("A2", 10.0)
        .unwrap();
    wb.worksheet_mut(1)
        .unwrap()
        .set_cell_value("A3", 15.0)
        .unwrap();
    wb.define_name("Numbers", "Data!$A$1:$A$3").unwrap();

    let calc = wb.worksheet_mut(0).unwrap();
    calc.set_cell_formula("B1", "=SUM(Numbers)").unwrap();
    calc.set_formula_result(0, 1, CellValue::Number(30.0))
        .unwrap();

    let result = roundtrip_through_excel(&wb);
    let names: Vec<&str> = result
        .named_ranges()
        .iter()
        .map(|nr| nr.name.as_str())
        .collect();
    assert!(
        names.iter().any(|n| n.contains("Numbers")),
        "Numbers name must survive XLSX Excel round-trip; got {names:?}"
    );
    let s = result.worksheet_by_name("Calc").unwrap();
    let v = s.get_value_at(0, 1);
    match v.effective_value() {
        CellValue::Number(n) => assert!((n - 30.0).abs() < 1e-9, "B1 = {n}"),
        other => panic!("B1 expected Number(30), got {other:?}"),
    }
}

#[test]
fn test_write_print_area() {
    let mut wb = Workbook::new();
    let ws = wb.worksheet_mut(0).unwrap();
    ws.set_cell_value("A1", "header").unwrap();
    ws.set_cell_value("B2", 100.0).unwrap();
    let mut ps = ws.page_setup().clone();
    ps.print_area = Some(range("A1", "B2"));
    ws.set_page_setup(ps);

    let result = roundtrip_through_excel(&wb);
    let s = result.worksheet(0).unwrap();
    let print_area = s
        .page_setup()
        .print_area
        .as_ref()
        .expect("print_area must survive XLSX Excel round-trip");
    assert_eq!(print_area.start, CellAddress::parse("A1").unwrap());
    assert_eq!(print_area.end, CellAddress::parse("B2").unwrap());
}

#[test]
fn test_write_sheet_visibility() {
    let mut wb = Workbook::new();
    wb.rename_worksheet(0, "Public").unwrap();
    wb.add_worksheet_with_name("Hidden").unwrap();
    wb.worksheet_mut(0)
        .unwrap()
        .set_cell_value("A1", "p")
        .unwrap();
    wb.worksheet_mut(1)
        .unwrap()
        .set_visibility(SheetVisibility::Hidden);

    let result = roundtrip_through_excel(&wb);
    let hidden = result.worksheet_by_name("Hidden").unwrap();
    assert!(
        hidden.visibility() != SheetVisibility::Visible,
        "Hidden sheet must remain non-visible; got {:?}",
        hidden.visibility()
    );
}

#[test]
fn test_write_freeze_panes() {
    let mut wb = Workbook::new();
    let ws = wb.worksheet_mut(0).unwrap();
    ws.set_cell_value("A1", "header").unwrap();
    ws.set_freeze_panes(1, 1);

    let result = roundtrip_through_excel(&wb);
    let s = result.worksheet(0).unwrap();
    let freeze = s
        .freeze_panes()
        .expect("freeze panes must survive XLSX Excel round-trip");
    assert_eq!((freeze.row, freeze.col), (1, 1));
}

#[test]
fn test_write_sheet_protection() {
    use duke_sheets_core::worksheet::SheetProtection;
    let mut wb = Workbook::new();
    let ws = wb.worksheet_mut(0).unwrap();
    ws.set_cell_value("A1", "locked").unwrap();
    ws.set_protection(Some(SheetProtection {
        protected: true,
        ..Default::default()
    }));

    let result = roundtrip_through_excel(&wb);
    let s = result.worksheet(0).unwrap();
    let prot = s
        .protection()
        .expect("sheet protection must survive XLSX Excel round-trip");
    assert!(prot.protected, "protected flag lost");
}

#[test]
fn test_write_cell_protection_hidden_formula() {
    // Hidden formula bit must survive Excel re-save.
    let mut wb = Workbook::new();
    let ws = wb.worksheet_mut(0).unwrap();
    ws.set_cell_value("A1", "formula").unwrap();
    ws.set_cell_formula("B1", "=A1").unwrap();
    ws.set_formula_result(0, 1, CellValue::Empty).unwrap();
    let mut hidden = Style::new();
    hidden.protection = Protection {
        locked: true,
        hidden: true,
    };
    ws.set_cell_style("B1", &hidden).unwrap();

    let result = roundtrip_through_excel(&wb);
    let s = result.worksheet(0).unwrap();
    let b1 = s
        .cell_style_at(0, 1)
        .expect("B1 should have non-default style after round-trip");
    assert!(
        b1.protection.hidden,
        "B1 hidden flag must survive round-trip, got hidden={}",
        b1.protection.hidden
    );
}

#[test]
fn test_write_cell_protection_unlocked() {
    // User explicitly unlocks a cell. Per ECMA-376 §18.8.32, Excel's
    // effective default is `locked=true`, so the writer must emit
    // `<protection locked="0"/>` for the unlock intent to survive
    // through Excel re-save. This is a regression test against an
    // earlier writer bug where `Protection::default()` matched our
    // Rust `derive(Default)` (= locked=false), causing the writer to
    // skip emission and silently lose the user's unlock intent.
    let mut wb = Workbook::new();
    let ws = wb.worksheet_mut(0).unwrap();
    ws.set_cell_value("A1", "default-locked").unwrap();
    ws.set_cell_value("B1", "explicitly-unlocked").unwrap();
    let mut unlocked = Style::new();
    unlocked.protection = Protection {
        locked: false,
        hidden: false,
    };
    ws.set_cell_style("B1", &unlocked).unwrap();

    let result = roundtrip_through_excel(&wb);
    let s = result.worksheet(0).unwrap();
    let b1 = s
        .cell_style_at(0, 1)
        .expect("B1 must have non-default style for unlocked state");
    assert!(
        !b1.protection.locked,
        "B1 unlock intent lost in round-trip: locked={}",
        b1.protection.locked
    );
}

#[test]
#[ignore = "requires Excel COM bridge on localhost:9876"]
fn excel_can_evaluate_intersection_we_emit() {
    let mut wb = Workbook::new();
    let ws = wb.worksheet_mut(0).unwrap();
    for (row, &(a, b, c)) in [(1, 2, 3), (4, 5, 6), (7, 8, 9)].iter().enumerate() {
        ws.set_cell_value_at(row as u32, 0, a as f64).unwrap();
        ws.set_cell_value_at(row as u32, 1, b as f64).unwrap();
        ws.set_cell_value_at(row as u32, 2, c as f64).unwrap();
    }
    // Intersection of A1:B3 with B2:C3 is the cells at B2, B3 = 5+8 = 13.
    ws.set_cell_formula("E1", "=SUM(A1:B3 B2:C3)").unwrap();
    ws.set_formula_result(0, 4, CellValue::Number(13.0))
        .unwrap();

    let result = roundtrip_through_excel(&wb);
    let s = result.worksheet(0).unwrap();
    let v = s.get_value_at(0, 4);
    match v.effective_value() {
        CellValue::Number(n) => assert!(
            (n - 13.0).abs() < 1e-9,
            "intersection sum drifted: E1 = {n}"
        ),
        other => panic!("E1 expected Number(13), got {other:?}"),
    }
    let formula = s.get_formula_at(0, 4).expect("E1 must still be a formula");
    assert!(
        formula.contains("A1:B3") && formula.contains("B2:C3"),
        "intersection ranges lost from formula: {formula:?}"
    );
}

#[test]
#[ignore = "requires Excel COM bridge on localhost:9876"]
fn excel_can_evaluate_union_we_emit() {
    let mut wb = Workbook::new();
    let ws = wb.worksheet_mut(0).unwrap();
    for (row, &(a, b, c)) in [(1, 2, 3), (4, 5, 6), (7, 8, 9)].iter().enumerate() {
        ws.set_cell_value_at(row as u32, 0, a as f64).unwrap();
        ws.set_cell_value_at(row as u32, 1, b as f64).unwrap();
        ws.set_cell_value_at(row as u32, 2, c as f64).unwrap();
    }
    // Union of A1:A2 and C2:C3 = {1, 4, 6, 9} = 20. Double parens are
    // required so SUM treats the comma as union, not arg separator.
    ws.set_cell_formula("E1", "=SUM((A1:A2,C2:C3))").unwrap();
    ws.set_formula_result(0, 4, CellValue::Number(20.0))
        .unwrap();

    let result = roundtrip_through_excel(&wb);
    let s = result.worksheet(0).unwrap();
    let v = s.get_value_at(0, 4);
    match v.effective_value() {
        CellValue::Number(n) => assert!((n - 20.0).abs() < 1e-9, "union sum drifted: E1 = {n}"),
        other => panic!("E1 expected Number(20), got {other:?}"),
    }
    let formula = s.get_formula_at(0, 4).expect("E1 must still be a formula");
    assert!(
        formula.contains("A1:A2") && formula.contains("C2:C3"),
        "union ranges lost from formula: {formula:?}"
    );
}

/// A valid 68-byte 1x1 transparent PNG used as a deterministic image
/// payload for Excel parity tests.
const TEST_PNG_1X1: &[u8] = &[
    0x89, 0x50, 0x4E, 0x47, 0x0D, 0x0A, 0x1A, 0x0A, 0x00, 0x00, 0x00, 0x0D, 0x49, 0x48, 0x44, 0x52,
    0x00, 0x00, 0x00, 0x01, 0x00, 0x00, 0x00, 0x01, 0x08, 0x06, 0x00, 0x00, 0x00, 0x1F, 0x15, 0xC4,
    0x89, 0x00, 0x00, 0x00, 0x0B, 0x49, 0x44, 0x41, 0x54, 0x78, 0x9C, 0x63, 0x60, 0x00, 0x02, 0x00,
    0x00, 0x05, 0x00, 0x01, 0x7A, 0x5E, 0xAB, 0x3F, 0x00, 0x00, 0x00, 0x00, 0x49, 0x45, 0x4E, 0x44,
    0xAE, 0x42, 0x60, 0x82,
];

/// Excel must accept an XLSX with a `<xdr:oneCellAnchor>` wrapped
/// picture and preserve the OneCell anchor through SaveAs.
#[test]
#[ignore = "requires Excel COM bridge on localhost:9876"]
fn excel_can_read_xlsx_onecell_picture_we_emit() {
    use duke_sheets_chart::{CellMarker, DrawingAnchor, EmbeddedImage, ImageFormat};

    let mut wb = Workbook::new();
    let ws = wb.worksheet_mut(0).unwrap();
    ws.set_cell_value("A1", "anchor").unwrap();
    ws.add_image(EmbeddedImage {
        id: 1,
        name: "OneCellPic".into(),
        description: None,
        anchor: DrawingAnchor::OneCell {
            from: CellMarker {
                col: 2,
                col_offset_emu: 0,
                row: 3,
                row_offset_emu: 0,
            },
            width_emu: 1_500_000,
            height_emu: 800_000,
        },
        format: ImageFormat::Png,
        media_path: String::new(),
        svg_media_path: None,
        width_emu: 1_500_000,
        height_emu: 800_000,
        rotation: None,
        flip_h: false,
        flip_v: false,
        data: TEST_PNG_1X1.to_vec(),
        svg_data: None,
    });

    let result = roundtrip_through_excel(&wb);
    let images = result.worksheet(0).unwrap().images();
    assert_eq!(images.len(), 1, "OneCell picture must survive Excel");
    let img = &images[0];
    match &img.anchor {
        DrawingAnchor::OneCell {
            from,
            width_emu,
            height_emu,
        } => {
            assert_eq!(from.col, 2);
            assert_eq!(from.row, 3);
            assert_eq!(*width_emu, 1_500_000);
            assert_eq!(*height_emu, 800_000);
        }
        other => panic!("expected OneCell anchor after Excel round-trip, got {other:?}"),
    }
}

/// Excel must accept an XLSX with a `<xdr:absoluteAnchor>` wrapped
/// picture and preserve the Absolute anchor through SaveAs.
#[test]
#[ignore = "requires Excel COM bridge on localhost:9876"]
fn excel_can_read_xlsx_absolute_picture_we_emit() {
    use duke_sheets_chart::{DrawingAnchor, EmbeddedImage, ImageFormat};

    let mut wb = Workbook::new();
    let ws = wb.worksheet_mut(0).unwrap();
    ws.set_cell_value("A1", "anchor").unwrap();
    ws.add_image(EmbeddedImage {
        id: 1,
        name: "AbsolutePic".into(),
        description: None,
        anchor: DrawingAnchor::Absolute {
            x_emu: 2_500_000,
            y_emu: 1_200_000,
            width_emu: 1_000_000,
            height_emu: 900_000,
        },
        format: ImageFormat::Png,
        media_path: String::new(),
        svg_media_path: None,
        width_emu: 1_000_000,
        height_emu: 900_000,
        rotation: None,
        flip_h: false,
        flip_v: false,
        data: TEST_PNG_1X1.to_vec(),
        svg_data: None,
    });

    let result = roundtrip_through_excel(&wb);
    let images = result.worksheet(0).unwrap().images();
    assert_eq!(images.len(), 1, "Absolute picture must survive Excel");
    let img = &images[0];
    match &img.anchor {
        DrawingAnchor::Absolute {
            x_emu,
            y_emu,
            width_emu,
            height_emu,
        } => {
            assert_eq!(*x_emu, 2_500_000);
            assert_eq!(*y_emu, 1_200_000);
            assert_eq!(*width_emu, 1_000_000);
            assert_eq!(*height_emu, 900_000);
        }
        other => panic!("expected Absolute anchor after Excel round-trip, got {other:?}"),
    }
}

/// Excel must accept the `<xdr:pic>` + `xl/media/imageN.png` emit
/// (no `Repaired` warning) and round-trip the PNG bytes verbatim
/// through SaveAs.
#[test]
#[ignore = "requires Excel COM bridge on localhost:9876"]
fn excel_can_read_xlsx_png_image_we_emit() {
    use duke_sheets_chart::{CellMarker, DrawingAnchor, EmbeddedImage, ImageFormat};

    let mut wb = Workbook::new();
    let ws = wb.worksheet_mut(0).unwrap();
    ws.set_cell_value("A1", "anchor").unwrap();
    ws.add_image(EmbeddedImage {
        id: 1,
        name: "Pic1".into(),
        description: None,
        anchor: DrawingAnchor::TwoCell {
            from: CellMarker {
                col: 1,
                col_offset_emu: 0,
                row: 2,
                row_offset_emu: 0,
            },
            to: CellMarker {
                col: 5,
                col_offset_emu: 0,
                row: 10,
                row_offset_emu: 0,
            },
            edit_as: None,
        },
        format: ImageFormat::Png,
        media_path: String::new(),
        svg_media_path: None,
        width_emu: 1_000_000,
        height_emu: 2_000_000,
        rotation: None,
        flip_h: false,
        flip_v: false,
        data: TEST_PNG_1X1.to_vec(),
        svg_data: None,
    });

    let result = roundtrip_through_excel(&wb);
    let images = result.worksheet(0).unwrap().images();
    assert_eq!(images.len(), 1, "image must survive Excel re-save");
    let img = &images[0];
    assert_eq!(img.format, ImageFormat::Png);
    assert_eq!(
        img.data, TEST_PNG_1X1,
        "PNG bytes must round-trip through Excel verbatim"
    );
}
