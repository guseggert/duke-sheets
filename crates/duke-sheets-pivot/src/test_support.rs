pub(crate) use duke_sheets_core::{
    CellError, CellRange, CellValue, PivotAggregate, PivotDateGroupUnit, PivotDatePeriod,
    PivotField, PivotFilter, PivotFilterOperator, PivotGrouping, PivotLayout, PivotLayoutKind,
    PivotManualGroup, PivotMeasure, PivotRefreshStatus, PivotShowAs, PivotSort, PivotSource,
    PivotSourceRange, PivotSubtotal, PivotTable, PivotValue, PivotValuesAxis, Table, TableColumn,
    Workbook,
};
pub(crate) use ssfmt::{date_serial::date_to_serial, DateSystem};

#[cfg(feature = "parallel")]
pub(crate) use crate::prelude::PARALLEL_ROW_THRESHOLD;
pub(crate) use crate::runtime_cache::PivotRuntimeCache;
pub(crate) use crate::snapshot::EncodedColumn;
pub(crate) use crate::{FormatPivotSource, PivotRefreshOptions, WorkbookPivotExt};
pub(crate) fn number(workbook: &Workbook, address: &str) -> f64 {
    workbook
        .worksheet(0)
        .unwrap()
        .get_value(address)
        .unwrap()
        .as_number()
        .unwrap()
}

pub(crate) fn text(workbook: &Workbook, address: &str) -> String {
    workbook
        .worksheet(0)
        .unwrap()
        .get_value(address)
        .unwrap()
        .to_string()
}

pub(crate) fn assert_close(actual: f64, expected: f64) {
    assert!(
        (actual - expected).abs() < 1e-9,
        "expected {expected}, got {actual}"
    );
}

pub(crate) fn tabular_layout() -> PivotLayout {
    let mut layout = PivotLayout::default();
    layout.kind = PivotLayoutKind::Tabular;
    layout.repeat_item_labels = true;
    layout
}

pub(crate) fn workbook_with_wrapped_page_fields(page_over_then_down: bool) -> Workbook {
    let mut workbook = Workbook::new();
    let sheet = workbook.worksheet_mut(0).unwrap();
    sheet.set_cell_value("A1", "Region").unwrap();
    sheet.set_cell_value("B1", "Segment").unwrap();
    sheet.set_cell_value("C1", "Channel").unwrap();
    sheet.set_cell_value("D1", "Country").unwrap();
    sheet.set_cell_value("E1", "Revenue").unwrap();
    sheet.set_cell_value("A2", "East").unwrap();
    sheet.set_cell_value("B2", "Retail").unwrap();
    sheet.set_cell_value("C2", "Online").unwrap();
    sheet.set_cell_value("D2", "US").unwrap();
    sheet.set_cell_value("E2", 10.0).unwrap();
    sheet.set_cell_value("A3", "West").unwrap();
    sheet.set_cell_value("B3", "Wholesale").unwrap();
    sheet.set_cell_value("C3", "Store").unwrap();
    sheet.set_cell_value("D3", "CA").unwrap();
    sheet.set_cell_value("E3", 20.0).unwrap();

    let mut layout = PivotLayout::default();
    layout.page_wrap = 2;
    layout.page_over_then_down = page_over_then_down;
    let pivot = PivotTable::builder("SalesPivot")
        .source_range(CellRange::parse("A1:E3").unwrap())
        .target_address("G1")
        .unwrap()
        .page("Segment")
        .page("Channel")
        .page("Country")
        .row("Region")
        .measure("Revenue", PivotAggregate::Sum)
        .layout(layout)
        .build()
        .unwrap();
    sheet.add_pivot_table(pivot).unwrap();

    workbook
}
