use std::cell::RefCell;
use std::io::Cursor;
use std::rc::Rc;
use std::sync::Arc;

use serde::{Deserialize, Serialize};
use wasm_bindgen::prelude::*;
use wasm_bindgen::JsCast;

use duke_sheets::{
    CalculationOptions, ChartType, FormulaValue, WorkbookCalculationExt, WorkbookPivotExt,
};
use duke_sheets_core::{
    CellAddress, CellError, CellRange, CellValue as CoreCellValue, PivotAggregate,
    PivotDateGroupUnit, PivotField, PivotFilter, PivotFilterOperator, PivotGrouping, PivotLayout,
    PivotLayoutKind, PivotManualGroup, PivotMeasure, PivotOverwritePolicy, PivotRefreshPolicy,
    PivotShowAs, PivotSort, PivotSource, PivotSourceRange, PivotStyle, PivotSubtotal, PivotTable,
    PivotValue, Workbook as CoreWorkbook, WorkbookConnection, WorkbookConnectionKind,
};
use duke_sheets_xlsb::XlsbWriter;
use duke_sheets_xlsx::XlsxWriter;

fn parse_xlsx_profile(
    profile: Option<&str>,
    key_bits: Option<u32>,
    spin_count: Option<u32>,
) -> std::result::Result<duke_sheets_xlsx::EncryptionProfile, String> {
    use duke_sheets_xlsx::EncryptionProfile;
    let normalized = profile.map(|p| p.to_lowercase());
    Ok(match normalized.as_deref() {
        None | Some("default") | Some("agile") | Some("ooxml-agile") => EncryptionProfile::Agile {
            key_bits: key_bits.unwrap_or(256),
            spin_count: spin_count.unwrap_or(100_000),
        },
        Some("standard") | Some("ooxml-standard") => EncryptionProfile::Standard {
            key_bits: key_bits.unwrap_or(128),
        },
        Some(other) => return Err(format!("unknown XLSX encryption profile: {other:?}")),
    })
}

fn parse_xls_variant(
    profile: Option<&str>,
    key_bits: Option<u32>,
) -> std::result::Result<duke_sheets_crypto::xls::XlsEncryptionVariant, String> {
    use duke_sheets_crypto::xls::XlsEncryptionVariant;
    let normalized = profile.map(|p| p.to_lowercase());
    Ok(match normalized.as_deref() {
        None | Some("default") | Some("rc4-cryptoapi") | Some("xls-rc4-cryptoapi") => {
            XlsEncryptionVariant::Rc4CryptoApi {
                key_bits: key_bits.unwrap_or(128),
            }
        }
        Some("rc4-legacy") | Some("xls-rc4-legacy") => XlsEncryptionVariant::Rc4Legacy,
        Some("xor") | Some("xls-xor") => XlsEncryptionVariant::Xor,
        Some(other) => return Err(format!("unknown XLS encryption profile: {other:?}")),
    })
}

mod types;
mod workbook_read;
mod worksheet_read;

pub use types::*;

#[wasm_bindgen(typescript_custom_section)]
const ROW_ITERATOR_TYPES: &str = r#"
declare global {
  interface SymbolConstructor {
    readonly dispose: unique symbol;
  }
}

export interface JsRowCell {
  col: number;
  value: string;
  style?: any;
  mergeSpan?: JsMergeSpan;
  isMergedSecondary?: boolean;
  hyperlink?: any;
  comment?: any;
  formula?: string;
  image?: any;
}

export interface JsRow {
  index: number;
  cells: Array<JsRowCell>;
}

export interface JsMergeSpan {
  rowSpan: number;
  colSpan: number;
}

export interface JsRowsOptions {
  useFormattedValues?: boolean;
  useCalculatedValues?: boolean;
  includeStyles?: boolean;
  includeMergeInfo?: boolean;
  includeHyperlinks?: boolean;
  includeComments?: boolean;
  includeFormulas?: boolean;
  includeImages?: boolean;
  skipEmptyValues?: boolean;
  skipBlankValues?: boolean;
}

export interface ColorInput {
  colorType?: "auto" | "rgb" | "argb" | "theme" | "indexed";
  hex?: string;
  r?: number;
  g?: number;
  b?: number;
  a?: number;
  themeIndex?: number;
  tint?: number;
  paletteIndex?: number;
}

export interface FontStyleInput {
  name?: string;
  size?: number;
  bold?: boolean;
  italic?: boolean;
  underline?: "none" | "single" | "double" | "singleAccounting" | "doubleAccounting";
  strikethrough?: boolean;
  color?: ColorInput;
  verticalAlign?: "baseline" | "superscript" | "subscript";
  family?: number;
  charset?: number;
  scheme?: string;
}

export interface GradientStopInput {
  position: number;
  color: ColorInput;
}

export interface FillStyleInput {
  fillType?: "none" | "solid" | "pattern" | "gradient";
  color?: ColorInput;
  pattern?: string;
  foreground?: ColorInput;
  background?: ColorInput;
  gradientType?: "linear" | "path";
  angle?: number;
  stops?: GradientStopInput[];
}

export interface BorderEdgeInput {
  style?: "none" | "thin" | "medium" | "thick" | "dashed" | "dotted" | "double" | "hair" | "mediumDashed" | "dashDot" | "mediumDashDot" | "dashDotDot" | "mediumDashDotDot" | "slantDashDot";
  color?: ColorInput;
}

export interface BorderStyleInput {
  left?: BorderEdgeInput;
  right?: BorderEdgeInput;
  top?: BorderEdgeInput;
  bottom?: BorderEdgeInput;
  diagonal?: BorderEdgeInput;
  diagonalDirection?: "none" | "down" | "up" | "both";
}

export interface AlignmentInput {
  horizontal?: "general" | "left" | "center" | "right" | "fill" | "justify" | "centerContinuous" | "distributed";
  vertical?: "top" | "center" | "bottom" | "justify" | "distributed";
  wrapText?: boolean;
  shrinkToFit?: boolean;
  indent?: number;
  rotation?: number;
  readingOrder?: "contextDependent" | "leftToRight" | "rightToLeft";
}

export interface NumberFormatInput {
  formatType?: "general" | "builtin" | "custom";
  id?: number;
  formatString?: string;
}

export interface CellProtectionInput {
  locked?: boolean;
  hidden?: boolean;
}

export interface StyleInput {
  font?: FontStyleInput;
  fill?: FillStyleInput;
  border?: BorderStyleInput;
  alignment?: AlignmentInput;
  numberFormat?: NumberFormatInput;
  protection?: CellProtectionInput;
}

export class RowIterator implements IterableIterator<JsRow> {
  constructor(ws: Worksheet, opts?: JsRowsOptions, maxRow?: number);
  [Symbol.iterator](): RowIterator;
  next(): IteratorResult<JsRow>;
}

export interface PivotMeasureOptions {
  field: string;
  aggregate?: "sum" | "count" | "countNumbers" | "average" | "max" | "min" | "product" | "stdDev" | "stdDevP" | "var" | "varP";
  name?: string;
  showAs?: "normal" | "percentOfGrandTotal" | "percentOfRowTotal" | "percentOfColumnTotal" | "index" | "runningTotal" | "runTotal" | "differenceFrom" | "difference" | "percentDifferenceFrom" | "percentDiff" | "rankAscending" | "rankDescending";
  baseField?: string;
  baseItem?: string | number | boolean;
  numberFormat?: string;
}

export interface PivotItemFilterOptions {
  kind?: "item" | "items" | "fieldItems" | "field_items";
  field: string;
  items: string[];
}

export type PivotFilterOperator =
  | "equals"
  | "equal"
  | "eq"
  | "notEquals"
  | "notEqual"
  | "ne"
  | "lessThan"
  | "lt"
  | "lessThanOrEqual"
  | "lte"
  | "greaterThan"
  | "gt"
  | "greaterThanOrEqual"
  | "gte"
  | "beginsWith"
  | "doesNotBeginWith"
  | "notBeginsWith"
  | "endsWith"
  | "doesNotEndWith"
  | "notEndsWith"
  | "contains"
  | "doesNotContain"
  | "notContains";

export interface PivotLabelFilterOptions {
  kind: "label";
  field: string;
  operator: PivotFilterOperator;
  text: string;
}

export interface PivotValueFilterOptions {
  kind: "value";
  field: string;
  measure: PivotMeasureOptions;
  operator: PivotFilterOperator;
  value: number;
}

export interface PivotTopNFilterOptions {
  kind: "topN" | "top_n" | "top";
  field: string;
  measure: PivotMeasureOptions;
  n: number;
  top?: boolean;
  percent?: boolean;
}

export type PivotFilterOptions =
  | PivotItemFilterOptions
  | PivotLabelFilterOptions
  | PivotValueFilterOptions
  | PivotTopNFilterOptions;

export interface PivotCalculatedFieldOptions {
  name: string;
  formula: string;
}

export type PivotValueInput = string | number | boolean;

export interface PivotCalculatedItemOptions {
  field: string;
  item: PivotValueInput;
  formula: string;
}

export interface PivotManualGroupOptions {
  name: string;
  members: PivotValueInput[];
}

export interface PivotGroupingOptions {
  field: string;
  kind: "number" | "numeric" | "date" | "manual" | "items" | "item";
  start?: number;
  end?: number;
  interval?: number;
  units?: Array<"seconds" | "minutes" | "hours" | "days" | "months" | "quarters" | "years">;
  groups?: PivotManualGroupOptions[];
}

export interface PivotFieldOptions {
  field: string;
  sort?: "none" | "manual" | "ascending" | "asc" | "descending" | "desc";
  subtotal?:
    | "automatic"
    | "auto"
    | "none"
    | "sum"
    | "count"
    | "count_numbers"
    | "countNumbers"
    | "countnumbers"
    | "count_nums"
    | "countNums"
    | "countnums"
    | "average"
    | "avg"
    | "min"
    | "max"
    | "product"
    | "std_dev"
    | "stdDev"
    | "stddev"
    | "std_dev_p"
    | "stdDevP"
    | "stddevp"
    | "var"
    | "variance"
    | "var_p"
    | "varP"
    | "varp"
    | "variance_p"
    | "varianceP";
  showEmptyItems?: boolean;
  showDropDowns?: boolean;
  subtotalTop?: boolean;
  insertBlankRow?: boolean;
  insertPageBreak?: boolean;
  includeNewItemsInFilter?: boolean;
  itemPageCount?: number;
}

export interface PivotRefreshPolicyOptions {
  refreshOnOpen?: boolean;
  preserveFormatting?: boolean;
  backgroundQuery?: boolean;
  missingItemsLimit?: number;
}

export interface PivotLayoutOptions {
  kind?: "compact" | "outline" | "tabular";
  showRowGrandTotals?: boolean;
  showColumnGrandTotals?: boolean;
  showFieldHeaders?: boolean;
  repeatItemLabels?: boolean;
  showExpandCollapse?: boolean;
  printDrillIndicators?: boolean;
  itemPrintTitles?: boolean;
  fieldPrintTitles?: boolean;
  pageWrap?: number;
  pageOverThenDown?: boolean;
  mergeItemLabels?: boolean;
  dataCaption?: string;
  grandTotalCaption?: string;
  errorCaption?: string;
  showError?: boolean;
  missingCaption?: string;
  showMissing?: boolean;
  asteriskTotals?: boolean;
  showItems?: boolean;
  editData?: boolean;
  disableFieldList?: boolean;
  showCalculatedMembers?: boolean;
  visualTotals?: boolean;
  showMultipleLabel?: boolean;
  showDataDropDown?: boolean;
  showMemberPropertyTips?: boolean;
  showDataTips?: boolean;
  enableWizard?: boolean;
  enableDrill?: boolean;
  enableFieldProperties?: boolean;
  subtotalHiddenItems?: boolean;
  showDropZones?: boolean;
  indent?: number;
  showEmptyRows?: boolean;
  showEmptyColumns?: boolean;
}

export interface PivotStyleOptions {
  name?: string;
  showRowHeaders?: boolean;
  showColumnHeaders?: boolean;
  showRowStripes?: boolean;
  showColumnStripes?: boolean;
  showLastColumn?: boolean;
}

export interface PivotConsolidationRangeOptions {
  sheet: string;
  range: string;
  name?: string;
  pageItems?: string[];
}

export interface PivotTableOptions {
  name: string;
  sourceRange?: string;
  sourceSheet?: string;
  tableName?: string;
  externalConnectionName?: string;
  externalCommandText?: string;
  olapConnectionName?: string;
  consolidationRanges?: PivotConsolidationRangeOptions[];
  target: string;
  rows?: string[];
  columns?: string[];
  pages?: string[];
  rowFields?: PivotFieldOptions[];
  columnFields?: PivotFieldOptions[];
  pageFields?: PivotFieldOptions[];
  measures: PivotMeasureOptions[];
  filters?: PivotFilterOptions[];
  calculatedFields?: PivotCalculatedFieldOptions[];
  calculatedItems?: PivotCalculatedItemOptions[];
  groupings?: PivotGroupingOptions[];
  refreshPolicy?: PivotRefreshPolicyOptions;
  layout?: PivotLayoutOptions;
  style?: PivotStyleOptions;
  overwritePolicy?: "clearOwnedRange" | "clear_owned_range" | "clear" | "overwrite" | "failOnOccupied" | "fail_on_occupied";
}

export interface PivotChartOptions {
  pivotName: string;
  chartType?: "columnClustered" | "columnStacked" | "columnPercentStacked" | "barClustered" | "barStacked" | "barPercentStacked" | "line" | "lineStacked" | "pie" | "pieExploded" | "doughnut" | "area" | "areaStacked" | "areaPercentStacked" | "scatterMarkers" | "scatterSmooth" | "scatterLines" | "bubble" | "radar" | "stock" | "surface" | string;
}

export interface PivotValue {
  kind: "blank" | "boolean" | "number" | "string" | "error";
  number?: number;
  text?: string;
  boolean?: boolean;
  error?: string;
}

export interface PivotSourceRangeDefinition {
  sheet: string;
  range: string;
  name?: string;
  pageItems: string[];
}

export interface PivotSourceDefinition {
  kind: "worksheetRange" | "table" | "external" | "consolidation" | "scenario" | "olap";
  sheet?: string;
  range?: string;
  tableName?: string;
  connectionName?: string;
  commandText?: string;
  ranges?: PivotSourceRangeDefinition[];
  scenarioName?: string;
  cube?: string;
}

export interface PivotFieldDefinition {
  field: string;
  sort: "none" | "ascending" | "descending";
  subtotal: "automatic" | "none" | "sum" | "count" | "countNumbers" | "average" | "min" | "max" | "product" | "stdDev" | "stdDevP" | "var" | "varP";
  showEmptyItems: boolean;
  showDropDowns: boolean;
  subtotalTop: boolean;
  insertBlankRow: boolean;
  insertPageBreak: boolean;
  includeNewItemsInFilter: boolean;
  itemPageCount: number;
}

export interface PivotShowAsDefinition {
  kind: "normal" | "percentOfGrandTotal" | "percentOfRowTotal" | "percentOfColumnTotal" | "index" | "runningTotal" | "differenceFrom" | "percentDifferenceFrom" | "rankAscending" | "rankDescending";
  baseField?: string;
  baseItem?: PivotValue;
}

export interface PivotMeasureDefinition {
  field: string;
  aggregate: "sum" | "count" | "countNumbers" | "average" | "max" | "min" | "product" | "stdDev" | "stdDevP" | "var" | "varP";
  name?: string;
  caption: string;
  showAs: PivotShowAsDefinition;
  numberFormat?: string;
}

export interface PivotFilterDefinition {
  kind: string;
  field?: string;
  items?: PivotValue[];
  operator?: PivotFilterOperator;
  text?: string;
  measure?: PivotMeasureDefinition;
  value?: number;
  n?: number;
  top?: boolean;
  percent?: boolean;
  detail?: string;
}

export interface PivotCalculatedFieldDefinition {
  name: string;
  formula: string;
}

export interface PivotCalculatedItemDefinition {
  field: string;
  item: PivotValue;
  formula: string;
}

export interface PivotManualGroupDefinition {
  name: string;
  members: PivotValue[];
}

export interface PivotGroupingDefinition {
  kind: "number" | "date" | "manual";
  field: string;
  start?: number;
  end?: number;
  interval?: number;
  units?: PivotDateGroupUnit[];
  groups?: PivotManualGroupDefinition[];
}

export interface PivotLayoutDefinition {
  kind: "compact" | "outline" | "tabular";
  showRowGrandTotals: boolean;
  showColumnGrandTotals: boolean;
  showFieldHeaders: boolean;
  repeatItemLabels: boolean;
  showExpandCollapse: boolean;
  printDrillIndicators: boolean;
  itemPrintTitles: boolean;
  fieldPrintTitles: boolean;
  pageWrap: number;
  pageOverThenDown: boolean;
  mergeItemLabels: boolean;
  dataCaption: string;
  grandTotalCaption?: string;
  errorCaption?: string;
  showError: boolean;
  missingCaption?: string;
  showMissing: boolean;
  asteriskTotals: boolean;
  showItems: boolean;
  editData: boolean;
  disableFieldList: boolean;
  showCalculatedMembers: boolean;
  visualTotals: boolean;
  showMultipleLabel: boolean;
  showDataDropDown: boolean;
  showMemberPropertyTips: boolean;
  showDataTips: boolean;
  enableWizard: boolean;
  enableDrill: boolean;
  enableFieldProperties: boolean;
  subtotalHiddenItems: boolean;
  showDropZones: boolean;
  indent: number;
  showEmptyRows: boolean;
  showEmptyColumns: boolean;
}

export interface PivotStyleDefinition {
  name?: string;
  showRowHeaders: boolean;
  showColumnHeaders: boolean;
  showRowStripes: boolean;
  showColumnStripes: boolean;
  showLastColumn: boolean;
}

export interface PivotRefreshPolicyDefinition {
  refreshOnOpen: boolean;
  preserveFormatting: boolean;
  backgroundQuery: boolean;
  missingItemsLimit?: number;
}

export interface PivotRefreshStatusDefinition {
  kind: "notRefreshed" | "succeeded" | "failed" | "external";
  message?: string;
}

export interface PivotTableDefinition {
  id: number;
  name: string;
  source: PivotSourceDefinition;
  target: string;
  rows: PivotFieldDefinition[];
  columns: PivotFieldDefinition[];
  pageFields: PivotFieldDefinition[];
  filters: PivotFilterDefinition[];
  calculatedFields: PivotCalculatedFieldDefinition[];
  calculatedItems: PivotCalculatedItemDefinition[];
  measures: PivotMeasureDefinition[];
  groupings: PivotGroupingDefinition[];
  layout: PivotLayoutDefinition;
  style: PivotStyleDefinition;
  refreshPolicy: PivotRefreshPolicyDefinition;
  overwritePolicy: "clearOwnedRange" | "overwrite" | "failOnOccupied";
  renderedRange?: string;
  refreshStatus: PivotRefreshStatusDefinition;
  extensionCount: number;
}

export interface DataConnectionOptions {
  id: number;
  name: string;
  kind?: "database" | "db" | "web" | "text" | "olap";
  connection?: string;
  command?: string;
  commandType?: number;
  url?: string;
  xml?: boolean;
  sourceData?: boolean;
  htmlTables?: boolean;
  htmlFormat?: string;
  post?: string;
  editPage?: string;
  sourceFile?: string;
  delimiter?: string;
  firstRow?: number;
  delimited?: boolean;
  decimal?: string;
  thousands?: string;
  local?: boolean;
  localConnection?: string;
  localRefresh?: boolean;
  sendLocale?: boolean;
  rowDrillCount?: number;
  refreshOnLoad?: boolean;
  background?: boolean;
  saveData?: boolean;
}

export interface DataConnectionDefinition {
  id: number;
  name: string;
  kind: "database" | "web" | "text" | "olap";
  refreshedVersion: number;
  refreshOnLoad: boolean;
  background: boolean;
  saveData: boolean;
  connection?: string;
  command?: string;
  commandType?: number;
  url?: string;
  xml?: boolean;
  sourceData?: boolean;
  htmlTables?: boolean;
  htmlFormat?: string;
  post?: string;
  editPage?: string;
  sourceFile?: string;
  delimiter?: string;
  firstRow?: number;
  delimited?: boolean;
  decimal?: string;
  thousands?: string;
  local?: boolean;
  localConnection?: string;
  localRefresh?: boolean;
  sendLocale?: boolean;
  rowDrillCount?: number;
}

export interface PivotRefreshStats {
  pivotCount: number;
  pivotsRefreshed: number;
  sourceRows: number;
  outputCells: number;
  cacheHits: number;
  cacheMisses: number;
}

export interface Workbook {
  readonly dataConnectionCount: number;
  readonly dataConnectionNames: string[];
  readonly dataConnections: DataConnectionDefinition[];
  getDataConnection(name: string): DataConnectionDefinition | null;
  getDataConnectionById(id: number): DataConnectionDefinition | null;
  addDataConnection(options: DataConnectionOptions): void;
  refreshPivots(): PivotRefreshStats;
}

export interface Worksheet {
  iterateRows(opts?: JsRowsOptions): RowIterator;
  setCellStyle(address: string, style: StyleInput): void;
  setCellStyleAt(row: number, col: number, style: StyleInput): void;
  setRangeStyle(range: string, style: StyleInput): void;
  readonly pivotCount: number;
  readonly pivotTableNames: string[];
  readonly pivotTables: PivotTableDefinition[];
  getPivotTable(name: string): PivotTableDefinition | null;
  addPivotTable(options: PivotTableOptions): void;
}
"#;

#[derive(Deserialize)]
#[serde(rename_all = "camelCase")]
struct JsCalculationOptions {
    iterative: Option<bool>,
    max_iterations: Option<u32>,
    max_change: Option<f64>,
    force_full_calculation: Option<bool>,
    calculate_volatile: Option<bool>,
    sheets: Option<Vec<usize>>,
    max_threads: Option<usize>,
}

/// Wrapper to make `js_sys::Function` implement Send + Sync.
/// SAFETY: WASM is single-threaded, so Send + Sync are no-ops.
struct SendSyncFunction(js_sys::Function);
unsafe impl Send for SendSyncFunction {}
unsafe impl Sync for SendSyncFunction {}

impl SendSyncFunction {
    fn call1(&self, this: &JsValue, arg: &JsValue) -> Result<JsValue, JsValue> {
        self.0.call1(this, arg)
    }

    fn call2(&self, this: &JsValue, arg1: &JsValue, arg2: &JsValue) -> Result<JsValue, JsValue> {
        self.0.call2(this, arg1, arg2)
    }

    fn call3(
        &self,
        this: &JsValue,
        arg1: &JsValue,
        arg2: &JsValue,
        arg3: &JsValue,
    ) -> Result<JsValue, JsValue> {
        self.0.call3(this, arg1, arg2, arg3)
    }
}

pub(crate) fn to_js_error(e: impl std::fmt::Display) -> JsError {
    JsError::new(&e.to_string())
}

pub(crate) fn to_js_value<T: Serialize>(value: &T) -> Result<JsValue, JsError> {
    serde_wasm_bindgen::to_value(value).map_err(to_js_error)
}

fn build_pivot_table_from_wasm(options: WasmPivotTableOptions) -> Result<PivotTable, JsError> {
    let mut builder = PivotTable::builder(options.name);
    if options.external_command_text.is_some() && options.external_connection_name.is_none() {
        return Err(JsError::new(
            "Pivot options require externalConnectionName when externalCommandText is set",
        ));
    }
    let source_count = usize::from(options.table_name.is_some())
        + usize::from(options.source_range.is_some())
        + usize::from(options.external_connection_name.is_some())
        + usize::from(options.olap_connection_name.is_some())
        + usize::from(options.consolidation_ranges.is_some());
    if source_count != 1 {
        return Err(JsError::new(
            "Pivot options require exactly one of tableName, sourceRange, externalConnectionName, olapConnectionName, or consolidationRanges",
        ));
    }

    match (
        options.table_name,
        options.source_range,
        options.external_connection_name,
        options.olap_connection_name,
        options.consolidation_ranges,
    ) {
        (Some(table_name), None, None, None, None) => {
            builder = builder.table_source(table_name);
        }
        (None, Some(source_range), None, None, None) => {
            let range = CellRange::parse(&source_range)
                .map_err(|e| JsError::new(&format!("Invalid pivot source range: {e}")))?;
            builder = if let Some(sheet) = options.source_sheet {
                builder.source_range_on_sheet(sheet, range)
            } else {
                builder.source_range(range)
            };
        }
        (None, None, Some(connection_name), None, None) => {
            builder = builder.source(PivotSource::External {
                connection_name,
                command_text: options.external_command_text,
            });
        }
        (None, None, None, Some(connection_name), None) => {
            builder = builder.source(PivotSource::Olap {
                connection_name,
                cube: None,
                command_text: None,
            });
        }
        (None, None, None, None, Some(ranges)) => {
            builder = builder.source(PivotSource::Consolidation {
                ranges: build_pivot_consolidation_ranges_from_wasm(ranges)?,
            });
        }
        _ => unreachable!("source_count validation accepts exactly one source"),
    }

    builder = builder
        .target_address(&options.target)
        .map_err(|e| JsError::new(&format!("Invalid pivot target: {e}")))?;
    for field in options.rows.unwrap_or_default() {
        builder = builder.row(field);
    }
    for field in options.columns.unwrap_or_default() {
        builder = builder.column(field);
    }
    for field in options.pages.unwrap_or_default() {
        builder = builder.page(field);
    }
    for field in options.row_fields.unwrap_or_default() {
        builder = builder.row(build_pivot_field_from_wasm(field)?);
    }
    for field in options.column_fields.unwrap_or_default() {
        builder = builder.column(build_pivot_field_from_wasm(field)?);
    }
    for field in options.page_fields.unwrap_or_default() {
        builder = builder.page(build_pivot_field_from_wasm(field)?);
    }
    for measure in options.measures {
        builder = builder.pivot_measure(build_pivot_measure_from_wasm(measure)?);
    }
    for filter in options.filters.unwrap_or_default() {
        builder = builder.filter(build_pivot_filter_from_wasm(filter)?);
    }
    for calculated_field in options.calculated_fields.unwrap_or_default() {
        builder = builder.calculated_field(calculated_field.name, calculated_field.formula);
    }
    for calculated_item in options.calculated_items.unwrap_or_default() {
        builder = builder.calculated_item(
            calculated_item.field,
            pivot_value_from_wasm(calculated_item.item),
            calculated_item.formula,
        );
    }
    for grouping in options.groupings.unwrap_or_default() {
        builder = builder.grouping(build_pivot_grouping_from_wasm(grouping)?);
    }
    if let Some(refresh_policy) = options.refresh_policy {
        builder = builder.refresh_policy(build_pivot_refresh_policy_from_wasm(refresh_policy));
    }
    if let Some(layout) = options.layout {
        builder = builder.layout(build_pivot_layout_from_wasm(layout)?);
    }
    if let Some(style) = options.style {
        builder = builder.style(build_pivot_style_from_wasm(style));
    }
    if let Some(overwrite_policy) = options.overwrite_policy {
        builder = builder.overwrite_policy(parse_pivot_overwrite_policy(&overwrite_policy)?);
    }

    builder.build().map_err(to_js_error)
}

fn build_pivot_consolidation_ranges_from_wasm(
    ranges: Vec<WasmPivotConsolidationRangeOptions>,
) -> Result<Vec<PivotSourceRange>, JsError> {
    if ranges.is_empty() {
        return Err(JsError::new(
            "Pivot consolidationRanges must contain at least one range",
        ));
    }
    ranges
        .into_iter()
        .map(|range| {
            let parsed = CellRange::parse(&range.range)
                .map_err(|e| JsError::new(&format!("Invalid pivot consolidation range: {e}")))?;
            let mut source_range = PivotSourceRange::new(range.sheet, parsed);
            if let Some(name) = range.name {
                source_range = source_range.with_name(name);
            }
            if let Some(page_items) = range.page_items {
                source_range = source_range.with_page_items(page_items);
            }
            Ok(source_range)
        })
        .collect()
}

fn build_workbook_connection_from_wasm(
    options: WasmWorkbookConnectionOptions,
) -> Result<WorkbookConnection, JsError> {
    let kind = options
        .kind
        .as_deref()
        .unwrap_or("database")
        .to_ascii_lowercase();
    let mut connection = match kind.as_str() {
        "database" | "db" => WorkbookConnection::database(
            options.id,
            options.name,
            options
                .connection
                .ok_or_else(|| JsError::new("database data connections require connection"))?,
        ),
        "web" => {
            let mut connection = WorkbookConnection::web(options.id, options.name, "");
            connection.kind = WorkbookConnectionKind::Web {
                url: options.url,
                xml: options.xml.unwrap_or(false),
                source_data: options.source_data.unwrap_or(false),
                html_tables: options.html_tables.unwrap_or(false),
                html_format: options.html_format,
                post: options.post,
                edit_page: options.edit_page,
            };
            connection
        }
        "text" => {
            let mut connection = WorkbookConnection::text(
                options.id,
                options.name,
                options.source_file.clone().unwrap_or_default(),
            );
            connection.kind = WorkbookConnectionKind::Text {
                source_file: options.source_file,
                delimiter: options.delimiter,
                first_row: options.first_row.unwrap_or(1),
                delimited: options.delimited.unwrap_or(true),
                decimal: options.decimal,
                thousands: options.thousands,
            };
            connection
        }
        "olap" => {
            let mut connection = WorkbookConnection::olap(options.id, options.name);
            connection.kind = WorkbookConnectionKind::Olap {
                local: options.local.unwrap_or(false),
                local_connection: options.local_connection,
                local_refresh: options.local_refresh.unwrap_or(true),
                send_locale: options.send_locale.unwrap_or(false),
                row_drill_count: options.row_drill_count,
            };
            connection
        }
        other => {
            return Err(JsError::new(&format!(
                "unknown data connection kind: {other}"
            )))
        }
    };
    if let Some(command) = options.command {
        connection = connection.with_command(command);
    }
    if let Some(command_type) = options.command_type {
        connection = connection.with_command_type(command_type);
    }
    if let Some(refresh_on_load) = options.refresh_on_load {
        connection = connection.with_refresh_on_load(refresh_on_load);
    }
    if let Some(background) = options.background {
        connection = connection.with_background(background);
    }
    if let Some(save_data) = options.save_data {
        connection = connection.with_save_data(save_data);
    }
    Ok(connection)
}

fn build_pivot_field_from_wasm(options: WasmPivotFieldOptions) -> Result<PivotField, JsError> {
    let mut field = PivotField::new(options.field);
    if let Some(sort) = options.sort {
        field.sort = parse_pivot_sort(&sort)?;
    }
    if let Some(subtotal) = options.subtotal {
        field.subtotal = parse_pivot_subtotal(&subtotal)?;
    }
    if let Some(show_empty_items) = options.show_empty_items {
        field.show_empty_items = show_empty_items;
    }
    if let Some(value) = options.show_drop_downs {
        field.show_drop_downs = value;
    }
    if let Some(value) = options.subtotal_top {
        field.subtotal_top = value;
    }
    if let Some(value) = options.insert_blank_row {
        field.insert_blank_row = value;
    }
    if let Some(value) = options.insert_page_break {
        field.insert_page_break = value;
    }
    if let Some(value) = options.include_new_items_in_filter {
        field.include_new_items_in_filter = value;
    }
    if let Some(value) = options.item_page_count {
        field.item_page_count = value;
    }
    Ok(field)
}

fn build_pivot_refresh_policy_from_wasm(
    options: WasmPivotRefreshPolicyOptions,
) -> PivotRefreshPolicy {
    let mut policy = PivotRefreshPolicy::default();
    if let Some(value) = options.refresh_on_open {
        policy.refresh_on_open = value;
    }
    if let Some(value) = options.preserve_formatting {
        policy.preserve_formatting = value;
    }
    if let Some(value) = options.background_query {
        policy.background_query = value;
    }
    policy.missing_items_limit = options.missing_items_limit;
    policy
}

fn build_pivot_layout_from_wasm(options: WasmPivotLayoutOptions) -> Result<PivotLayout, JsError> {
    let mut layout = PivotLayout::default();
    if let Some(kind) = options.kind {
        layout.kind = parse_pivot_layout_kind(&kind)?;
    }
    if let Some(value) = options.show_row_grand_totals {
        layout.show_row_grand_totals = value;
    }
    if let Some(value) = options.show_column_grand_totals {
        layout.show_column_grand_totals = value;
    }
    if let Some(value) = options.show_field_headers {
        layout.show_field_headers = value;
    }
    if let Some(value) = options.repeat_item_labels {
        layout.repeat_item_labels = value;
    }
    if let Some(value) = options.show_expand_collapse {
        layout.show_expand_collapse = value;
    }
    if let Some(value) = options.print_drill_indicators {
        layout.print_drill_indicators = value;
    }
    if let Some(value) = options.item_print_titles {
        layout.item_print_titles = value;
    }
    if let Some(value) = options.field_print_titles {
        layout.field_print_titles = value;
    }
    if let Some(value) = options.page_wrap {
        layout.page_wrap = value;
    }
    if let Some(value) = options.page_over_then_down {
        layout.page_over_then_down = value;
    }
    if let Some(value) = options.merge_item_labels {
        layout.merge_item_labels = value;
    }
    if let Some(value) = options.data_caption {
        layout.data_caption = value;
    }
    if let Some(value) = options.grand_total_caption {
        layout.grand_total_caption = Some(value);
    }
    if let Some(value) = options.error_caption {
        layout.error_caption = Some(value);
    }
    if let Some(value) = options.show_error {
        layout.show_error = value;
    }
    if let Some(value) = options.missing_caption {
        layout.missing_caption = Some(value);
    }
    if let Some(value) = options.show_missing {
        layout.show_missing = value;
    }
    if let Some(value) = options.asterisk_totals {
        layout.asterisk_totals = value;
    }
    if let Some(value) = options.show_items {
        layout.show_items = value;
    }
    if let Some(value) = options.edit_data {
        layout.edit_data = value;
    }
    if let Some(value) = options.disable_field_list {
        layout.disable_field_list = value;
    }
    if let Some(value) = options.show_calculated_members {
        layout.show_calculated_members = value;
    }
    if let Some(value) = options.visual_totals {
        layout.visual_totals = value;
    }
    if let Some(value) = options.show_multiple_label {
        layout.show_multiple_label = value;
    }
    if let Some(value) = options.show_data_drop_down {
        layout.show_data_drop_down = value;
    }
    if let Some(value) = options.show_member_property_tips {
        layout.show_member_property_tips = value;
    }
    if let Some(value) = options.show_data_tips {
        layout.show_data_tips = value;
    }
    if let Some(value) = options.enable_wizard {
        layout.enable_wizard = value;
    }
    if let Some(value) = options.enable_drill {
        layout.enable_drill = value;
    }
    if let Some(value) = options.enable_field_properties {
        layout.enable_field_properties = value;
    }
    if let Some(value) = options.subtotal_hidden_items {
        layout.subtotal_hidden_items = value;
    }
    if let Some(value) = options.show_drop_zones {
        layout.show_drop_zones = value;
    }
    if let Some(value) = options.indent {
        layout.indent = value;
    }
    if let Some(value) = options.show_empty_rows {
        layout.show_empty_rows = value;
    }
    if let Some(value) = options.show_empty_columns {
        layout.show_empty_columns = value;
    }
    Ok(layout)
}

fn build_pivot_style_from_wasm(options: WasmPivotStyleOptions) -> PivotStyle {
    let mut style = PivotStyle::default();
    if let Some(name) = options.name {
        style.name = if name.is_empty() { None } else { Some(name) };
    }
    if let Some(value) = options.show_row_headers {
        style.show_row_headers = value;
    }
    if let Some(value) = options.show_column_headers {
        style.show_column_headers = value;
    }
    if let Some(value) = options.show_row_stripes {
        style.show_row_stripes = value;
    }
    if let Some(value) = options.show_column_stripes {
        style.show_column_stripes = value;
    }
    if let Some(value) = options.show_last_column {
        style.show_last_column = value;
    }
    style
}

fn parse_pivot_layout_kind(value: &str) -> Result<PivotLayoutKind, JsError> {
    Ok(match value {
        "compact" => PivotLayoutKind::Compact,
        "outline" => PivotLayoutKind::Outline,
        "tabular" => PivotLayoutKind::Tabular,
        other => {
            return Err(JsError::new(&format!(
                "Unsupported pivot layout kind: {other}"
            )))
        }
    })
}

fn parse_chart_type(value: Option<&str>) -> Result<ChartType, JsError> {
    match value {
        Some(value) => ChartType::from_name(value)
            .ok_or_else(|| JsError::new(&format!("Unsupported chart type: {value}"))),
        None => Ok(ChartType::ColumnClustered),
    }
}

fn parse_pivot_overwrite_policy(value: &str) -> Result<PivotOverwritePolicy, JsError> {
    Ok(match value {
        "clearOwnedRange" | "clear_owned_range" | "clear" => PivotOverwritePolicy::ClearOwnedRange,
        "overwrite" => PivotOverwritePolicy::Overwrite,
        "failOnOccupied" | "fail_on_occupied" => PivotOverwritePolicy::FailOnOccupied,
        other => {
            return Err(JsError::new(&format!(
                "Unsupported pivot overwrite policy: {other}"
            )));
        }
    })
}

fn parse_pivot_sort(value: &str) -> Result<PivotSort, JsError> {
    Ok(match value {
        "none" | "manual" => PivotSort::None,
        "ascending" | "asc" => PivotSort::Ascending,
        "descending" | "desc" => PivotSort::Descending,
        other => return Err(JsError::new(&format!("Unsupported pivot sort: {other}"))),
    })
}

fn parse_pivot_subtotal(value: &str) -> Result<PivotSubtotal, JsError> {
    Ok(match value {
        "automatic" | "auto" => PivotSubtotal::Automatic,
        "none" => PivotSubtotal::None,
        "sum" => PivotSubtotal::Sum,
        "count" => PivotSubtotal::Count,
        "count_numbers" | "countNumbers" | "countnumbers" | "count_nums" | "countNums"
        | "countnums" => PivotSubtotal::CountNumbers,
        "average" | "avg" => PivotSubtotal::Average,
        "min" => PivotSubtotal::Min,
        "max" => PivotSubtotal::Max,
        "product" => PivotSubtotal::Product,
        "std_dev" | "stdDev" | "stddev" => PivotSubtotal::StdDev,
        "std_dev_p" | "stdDevP" | "stddevp" => PivotSubtotal::StdDevP,
        "var" | "variance" => PivotSubtotal::Var,
        "var_p" | "varP" | "varp" | "variance_p" | "varianceP" => PivotSubtotal::VarP,
        other => {
            return Err(JsError::new(&format!(
                "Unsupported pivot subtotal: {other}"
            )));
        }
    })
}

fn build_pivot_measure_from_wasm(
    options: WasmPivotMeasureOptions,
) -> Result<PivotMeasure, JsError> {
    let aggregate = parse_pivot_aggregate(options.aggregate.as_deref())?;
    let mut measure = PivotMeasure::new(options.field, aggregate);
    if let Some(name) = options.name {
        measure = measure.with_name(name);
    }
    if let Some(show_as) = options.show_as {
        measure = measure.with_show_as(parse_pivot_show_as(
            &show_as,
            options.base_field,
            options.base_item,
        )?);
    }
    if let Some(number_format) = options.number_format {
        measure = measure.with_number_format(number_format);
    }
    Ok(measure)
}

fn build_pivot_filter_from_wasm(options: WasmPivotFilterOptions) -> Result<PivotFilter, JsError> {
    let kind = options.kind.unwrap_or_else(|| {
        if options.items.is_some() {
            "items".to_string()
        } else {
            "item".to_string()
        }
    });
    match kind.as_str() {
        "item" | "items" | "fieldItems" | "field_items" => {
            let items = options
                .items
                .ok_or_else(|| JsError::new("Pivot item filter requires items"))?;
            Ok(PivotFilter::field_items(
                options.field,
                items.into_iter().map(PivotValue::from).collect::<Vec<_>>(),
            ))
        }
        "label" => Ok(PivotFilter::Label {
            field: options.field.into(),
            operator: parse_pivot_filter_operator(options.operator.as_deref())?,
            value: options
                .text
                .ok_or_else(|| JsError::new("Pivot label filter requires text"))?,
        }),
        "value" => Ok(PivotFilter::Value {
            field: options.field.into(),
            measure: build_pivot_measure_from_wasm(
                options
                    .measure
                    .ok_or_else(|| JsError::new("Pivot value filter requires measure"))?,
            )?,
            operator: parse_pivot_filter_operator(options.operator.as_deref())?,
            value: options
                .value
                .ok_or_else(|| JsError::new("Pivot value filter requires value"))?,
        }),
        "topN" | "top_n" | "top" => Ok(PivotFilter::TopN {
            field: options.field.into(),
            measure: build_pivot_measure_from_wasm(
                options
                    .measure
                    .ok_or_else(|| JsError::new("Pivot top-N filter requires measure"))?,
            )?,
            n: options
                .n
                .ok_or_else(|| JsError::new("Pivot top-N filter requires n"))?,
            top: options.top.unwrap_or(true),
            percent: options.percent.unwrap_or(false),
        }),
        other => Err(JsError::new(&format!(
            "Unsupported pivot filter kind: {other}"
        ))),
    }
}

fn build_pivot_grouping_from_wasm(
    options: WasmPivotGroupingOptions,
) -> Result<PivotGrouping, JsError> {
    match options.kind.as_str() {
        "number" | "numeric" => Ok(PivotGrouping::Number {
            field: options.field.into(),
            start: options.start,
            end: options.end,
            interval: options
                .interval
                .ok_or_else(|| JsError::new("Numeric pivot grouping requires interval"))?,
        }),
        "date" => {
            let units = options
                .units
                .ok_or_else(|| JsError::new("Date pivot grouping requires units"))?
                .iter()
                .map(|unit| parse_pivot_date_group_unit(unit))
                .collect::<Result<Vec<_>, _>>()?;
            Ok(PivotGrouping::Date {
                field: options.field.into(),
                units,
            })
        }
        "manual" | "items" | "item" => Ok(PivotGrouping::Manual {
            field: options.field.into(),
            groups: options
                .groups
                .ok_or_else(|| JsError::new("Manual pivot grouping requires groups"))?
                .into_iter()
                .map(|group| PivotManualGroup {
                    name: group.name,
                    members: group
                        .members
                        .into_iter()
                        .map(pivot_value_from_wasm)
                        .collect(),
                })
                .collect(),
        }),
        other => Err(JsError::new(&format!(
            "Unsupported pivot grouping kind: {other}"
        ))),
    }
}

fn parse_pivot_aggregate(value: Option<&str>) -> Result<PivotAggregate, JsError> {
    let Some(value) = value else {
        return Ok(PivotAggregate::Sum);
    };
    Ok(match value {
        "sum" => PivotAggregate::Sum,
        "count" => PivotAggregate::Count,
        "countNumbers" | "countNums" => PivotAggregate::CountNumbers,
        "average" | "avg" => PivotAggregate::Average,
        "max" => PivotAggregate::Max,
        "min" => PivotAggregate::Min,
        "product" => PivotAggregate::Product,
        "stdDev" => PivotAggregate::StdDev,
        "stdDevP" | "stdDevp" => PivotAggregate::StdDevP,
        "var" => PivotAggregate::Var,
        "varP" | "varp" => PivotAggregate::VarP,
        other => {
            return Err(JsError::new(&format!(
                "Unsupported pivot aggregate: {other}"
            )))
        }
    })
}

fn parse_pivot_filter_operator(value: Option<&str>) -> Result<PivotFilterOperator, JsError> {
    let value = value.ok_or_else(|| JsError::new("Pivot filter requires operator"))?;
    Ok(match value {
        "equals" | "equal" | "eq" => PivotFilterOperator::Equals,
        "notEquals" | "notEqual" | "ne" => PivotFilterOperator::NotEquals,
        "lessThan" | "lt" => PivotFilterOperator::LessThan,
        "lessThanOrEqual" | "lte" => PivotFilterOperator::LessThanOrEqual,
        "greaterThan" | "gt" => PivotFilterOperator::GreaterThan,
        "greaterThanOrEqual" | "gte" => PivotFilterOperator::GreaterThanOrEqual,
        "beginsWith" => PivotFilterOperator::BeginsWith,
        "doesNotBeginWith" | "notBeginsWith" => PivotFilterOperator::DoesNotBeginWith,
        "endsWith" => PivotFilterOperator::EndsWith,
        "doesNotEndWith" | "notEndsWith" => PivotFilterOperator::DoesNotEndWith,
        "contains" => PivotFilterOperator::Contains,
        "doesNotContain" | "notContains" => PivotFilterOperator::DoesNotContain,
        other => {
            return Err(JsError::new(&format!(
                "Unsupported pivot filter operator: {other}"
            )));
        }
    })
}

fn parse_pivot_date_group_unit(value: &str) -> Result<PivotDateGroupUnit, JsError> {
    Ok(match value {
        "seconds" => PivotDateGroupUnit::Seconds,
        "minutes" => PivotDateGroupUnit::Minutes,
        "hours" => PivotDateGroupUnit::Hours,
        "days" => PivotDateGroupUnit::Days,
        "months" => PivotDateGroupUnit::Months,
        "quarters" => PivotDateGroupUnit::Quarters,
        "years" => PivotDateGroupUnit::Years,
        other => {
            return Err(JsError::new(&format!(
                "Unsupported pivot date grouping unit: {other}"
            )));
        }
    })
}

fn parse_pivot_show_as(
    value: &str,
    base_field: Option<String>,
    base_item: Option<WasmPivotValueInput>,
) -> Result<PivotShowAs, JsError> {
    Ok(match value {
        "normal" => PivotShowAs::Normal,
        "percentOfGrandTotal" | "percentOfTotal" => PivotShowAs::PercentOfGrandTotal,
        "percentOfRowTotal" | "percentOfRow" => PivotShowAs::PercentOfRowTotal,
        "percentOfColumnTotal" | "percentOfCol" => PivotShowAs::PercentOfColumnTotal,
        "index" => PivotShowAs::Index,
        "runningTotal" | "runTotal" => PivotShowAs::RunningTotal {
            base_field: require_pivot_base_field(value, base_field)?.into(),
        },
        "differenceFrom" | "difference" => PivotShowAs::DifferenceFrom {
            base_field: require_pivot_base_field(value, base_field)?.into(),
            base_item: require_pivot_base_item(value, base_item)?,
        },
        "percentDifferenceFrom" | "percentDiff" => PivotShowAs::PercentDifferenceFrom {
            base_field: require_pivot_base_field(value, base_field)?.into(),
            base_item: require_pivot_base_item(value, base_item)?,
        },
        "rankAscending" => PivotShowAs::RankAscending {
            base_field: require_pivot_base_field(value, base_field)?.into(),
        },
        "rankDescending" => PivotShowAs::RankDescending {
            base_field: require_pivot_base_field(value, base_field)?.into(),
        },
        other => {
            return Err(JsError::new(&format!(
                "Unsupported pivot showAs mode: {other}"
            )))
        }
    })
}

fn require_pivot_base_field(value: &str, base_field: Option<String>) -> Result<String, JsError> {
    base_field.ok_or_else(|| JsError::new(&format!("pivot showAs mode {value} requires baseField")))
}

fn require_pivot_base_item(
    value: &str,
    base_item: Option<WasmPivotValueInput>,
) -> Result<PivotValue, JsError> {
    base_item
        .map(pivot_value_from_wasm)
        .ok_or_else(|| JsError::new(&format!("pivot showAs mode {value} requires baseItem")))
}

fn pivot_value_from_wasm(value: WasmPivotValueInput) -> PivotValue {
    match value {
        WasmPivotValueInput::Number(number) => PivotValue::Number(number),
        WasmPivotValueInput::String(string) => PivotValue::String(string),
        WasmPivotValueInput::Boolean(boolean) => PivotValue::Boolean(boolean),
    }
}

fn cell_error_to_string(e: &CellError) -> &'static str {
    match e {
        CellError::Div0 => "#DIV/0!",
        CellError::Na => "#N/A",
        CellError::Name => "#NAME?",
        CellError::Null => "#NULL!",
        CellError::Num => "#NUM!",
        CellError::Ref => "#REF!",
        CellError::Value => "#VALUE!",
        CellError::GettingData => "#GETTING_DATA",
        CellError::Spill => "#SPILL!",
        CellError::Calc => "#CALC!",
    }
}

#[derive(Serialize)]
#[serde(rename_all = "camelCase")]
struct WasmCalculationStats {
    formula_count: u32,
    cells_calculated: u32,
    errors: u32,
    circular_references: u32,
    volatile_cells: u32,
    converged: bool,
    iterations: u32,
}

impl From<&duke_sheets::CalculationStats> for WasmCalculationStats {
    fn from(stats: &duke_sheets::CalculationStats) -> Self {
        Self {
            formula_count: stats.formula_count as u32,
            cells_calculated: stats.cells_calculated as u32,
            errors: stats.errors as u32,
            circular_references: stats.circular_references as u32,
            volatile_cells: stats.volatile_cells as u32,
            converged: stats.converged,
            iterations: stats.iterations as u32,
        }
    }
}

#[wasm_bindgen]
pub struct CellValue {
    inner: CoreCellValue,
}

#[wasm_bindgen]
impl CellValue {
    #[wasm_bindgen(getter)]
    pub fn is_empty(&self) -> bool {
        matches!(self.inner, CoreCellValue::Empty)
    }

    #[wasm_bindgen(getter)]
    pub fn is_number(&self) -> bool {
        matches!(self.inner, CoreCellValue::Number(_))
    }

    #[wasm_bindgen(getter)]
    pub fn is_text(&self) -> bool {
        matches!(self.inner, CoreCellValue::String(_))
    }

    #[wasm_bindgen(getter)]
    pub fn is_boolean(&self) -> bool {
        matches!(self.inner, CoreCellValue::Boolean(_))
    }

    #[wasm_bindgen(getter)]
    pub fn is_error(&self) -> bool {
        matches!(self.inner, CoreCellValue::Error(_))
    }

    #[wasm_bindgen(js_name = asNumber)]
    pub fn as_number(&self) -> Option<f64> {
        match &self.inner {
            CoreCellValue::Number(n) => Some(*n),
            _ => None,
        }
    }

    #[wasm_bindgen(js_name = asText)]
    pub fn as_text(&self) -> Option<String> {
        match &self.inner {
            CoreCellValue::String(s) => Some(s.to_string()),
            _ => None,
        }
    }

    #[wasm_bindgen(js_name = asBoolean)]
    pub fn as_boolean(&self) -> Option<bool> {
        match &self.inner {
            CoreCellValue::Boolean(b) => Some(*b),
            _ => None,
        }
    }

    #[wasm_bindgen(js_name = asError)]
    pub fn as_error(&self) -> Option<String> {
        match &self.inner {
            CoreCellValue::Error(e) => Some(cell_error_to_string(e).to_string()),
            _ => None,
        }
    }

    #[wasm_bindgen(js_name = toJs)]
    pub fn to_js(&self) -> JsValue {
        match &self.inner {
            CoreCellValue::Empty => JsValue::NULL,
            CoreCellValue::Number(n) => JsValue::from_f64(*n),
            CoreCellValue::String(s) => JsValue::from_str(&s.to_string()),
            CoreCellValue::Boolean(b) => JsValue::from_bool(*b),
            CoreCellValue::Error(e) => JsValue::from_str(cell_error_to_string(e)),
            CoreCellValue::RichText(runs) => {
                let text: String = runs.iter().map(|r| r.text.as_str()).collect();
                JsValue::from_str(&text)
            }
            CoreCellValue::SpillTarget { .. } => JsValue::NULL,
        }
    }

    #[wasm_bindgen(js_name = toString)]
    pub fn to_string_js(&self) -> String {
        match &self.inner {
            CoreCellValue::Empty => String::new(),
            CoreCellValue::Number(n) => n.to_string(),
            CoreCellValue::String(s) => s.to_string(),
            CoreCellValue::Boolean(b) => if *b { "TRUE" } else { "FALSE" }.to_string(),
            CoreCellValue::Error(e) => cell_error_to_string(e).to_string(),
            CoreCellValue::RichText(runs) => runs.iter().map(|r| r.text.as_str()).collect(),
            CoreCellValue::SpillTarget { .. } => String::new(),
        }
    }
}

#[wasm_bindgen]
pub struct Worksheet {
    workbook: Rc<RefCell<CoreWorkbook>>,
    sheet_index: usize,
}

#[wasm_bindgen]
impl Worksheet {
    #[wasm_bindgen(getter)]
    pub fn name(&self) -> Result<String, JsError> {
        let wb = self.workbook.borrow();
        wb.worksheet(self.sheet_index)
            .map(|ws| ws.name().to_string())
            .ok_or_else(|| JsError::new("Worksheet no longer exists"))
    }

    #[wasm_bindgen(js_name = setCell)]
    pub fn set_cell(&self, address: &str, value: JsValue) -> Result<(), JsError> {
        let mut wb = self.workbook.borrow_mut();
        let ws = wb
            .worksheet_mut(self.sheet_index)
            .ok_or_else(|| JsError::new("Worksheet no longer exists"))?;
        let cell_value = js_to_cell_value(value)?;
        let addr = CellAddress::parse(address)
            .map_err(|e| JsError::new(&format!("Invalid cell address: {}", e)))?;
        ws.set_cell_value_at(addr.row, addr.col, cell_value)
            .map_err(to_js_error)
    }

    #[wasm_bindgen(js_name = setFormula)]
    pub fn set_formula(&self, address: &str, formula: &str) -> Result<(), JsError> {
        let mut wb = self.workbook.borrow_mut();
        let ws = wb
            .worksheet_mut(self.sheet_index)
            .ok_or_else(|| JsError::new("Worksheet no longer exists"))?;
        ws.set_cell_formula(address, formula).map_err(to_js_error)
    }

    #[wasm_bindgen(js_name = setCellStyle)]
    pub fn set_cell_style(&self, address: &str, style: JsValue) -> Result<(), JsError> {
        let patch: WasmStylePatch = serde_wasm_bindgen::from_value(style).map_err(to_js_error)?;
        let mut wb = self.workbook.borrow_mut();
        let ws = wb
            .worksheet_mut(self.sheet_index)
            .ok_or_else(|| JsError::new("Worksheet no longer exists"))?;
        let addr = CellAddress::parse(address)
            .map_err(|e| JsError::new(&format!("Invalid cell address: {}", e)))?;
        let mut core_style = ws
            .cell_style_at(addr.row, addr.col)
            .cloned()
            .unwrap_or_default();
        patch
            .apply_to_core_style(&mut core_style)
            .map_err(to_js_error)?;
        ws.set_cell_style_at(addr.row, addr.col, &core_style)
            .map_err(to_js_error)
    }

    #[wasm_bindgen(js_name = setCellStyleAt)]
    pub fn set_cell_style_at(&self, row: u32, col: u32, style: JsValue) -> Result<(), JsError> {
        let patch: WasmStylePatch = serde_wasm_bindgen::from_value(style).map_err(to_js_error)?;
        let mut wb = self.workbook.borrow_mut();
        let ws = wb
            .worksheet_mut(self.sheet_index)
            .ok_or_else(|| JsError::new("Worksheet no longer exists"))?;
        let mut core_style = ws
            .cell_style_at(row, col as u16)
            .cloned()
            .unwrap_or_default();
        patch
            .apply_to_core_style(&mut core_style)
            .map_err(to_js_error)?;
        ws.set_cell_style_at(row, col as u16, &core_style)
            .map_err(to_js_error)
    }

    #[wasm_bindgen(js_name = setRangeStyle)]
    pub fn set_range_style(&self, range_str: &str, style: JsValue) -> Result<(), JsError> {
        let patch: WasmStylePatch = serde_wasm_bindgen::from_value(style).map_err(to_js_error)?;
        let mut wb = self.workbook.borrow_mut();
        let ws = wb
            .worksheet_mut(self.sheet_index)
            .ok_or_else(|| JsError::new("Worksheet no longer exists"))?;
        let range = CellRange::parse(range_str)
            .map_err(|e| JsError::new(&format!("Invalid range: {}", e)))?;
        for addr in range.cells() {
            let mut core_style = ws
                .cell_style_at(addr.row, addr.col)
                .cloned()
                .unwrap_or_default();
            patch
                .apply_to_core_style(&mut core_style)
                .map_err(to_js_error)?;
            ws.set_cell_style_at(addr.row, addr.col, &core_style)
                .map_err(to_js_error)?;
        }
        Ok(())
    }

    #[wasm_bindgen(js_name = getCell)]
    pub fn get_cell(&self, address: &str) -> Result<CellValue, JsError> {
        let wb = self.workbook.borrow();
        let ws = wb
            .worksheet(self.sheet_index)
            .ok_or_else(|| JsError::new("Worksheet no longer exists"))?;
        let addr = CellAddress::parse(address)
            .map_err(|e| JsError::new(&format!("Invalid cell address: {}", e)))?;
        Ok(CellValue {
            inner: ws.get_value_at(addr.row, addr.col),
        })
    }

    #[wasm_bindgen(js_name = getCellAt)]
    pub fn get_cell_at(&self, row: u32, col: u32) -> Result<CellValue, JsError> {
        let wb = self.workbook.borrow();
        let ws = wb
            .worksheet(self.sheet_index)
            .ok_or_else(|| JsError::new("Worksheet no longer exists"))?;
        Ok(CellValue {
            inner: ws.get_value_at(row, col as u16),
        })
    }

    #[wasm_bindgen(js_name = getCalculatedValue)]
    pub fn get_calculated_value(&self, address: &str) -> Result<CellValue, JsError> {
        let wb = self.workbook.borrow();
        let ws = wb
            .worksheet(self.sheet_index)
            .ok_or_else(|| JsError::new("Worksheet no longer exists"))?;
        let addr = CellAddress::parse(address)
            .map_err(|e| JsError::new(&format!("Invalid cell address: {}", e)))?;
        let value = ws
            .get_calculated_value_at(addr.row, addr.col)
            .cloned()
            .unwrap_or(CoreCellValue::Empty);
        Ok(CellValue { inner: value })
    }

    #[wasm_bindgen(js_name = getCalculatedValueAt)]
    pub fn get_calculated_value_at(&self, row: u32, col: u32) -> Result<CellValue, JsError> {
        let wb = self.workbook.borrow();
        let ws = wb
            .worksheet(self.sheet_index)
            .ok_or_else(|| JsError::new("Worksheet no longer exists"))?;
        let value = ws
            .get_calculated_value_at(row, col as u16)
            .cloned()
            .unwrap_or(CoreCellValue::Empty);
        Ok(CellValue { inner: value })
    }

    #[wasm_bindgen(js_name = usedRange)]
    pub fn used_range(&self) -> Result<JsValue, JsError> {
        let wb = self.workbook.borrow();
        let ws = wb
            .worksheet(self.sheet_index)
            .ok_or_else(|| JsError::new("Worksheet no longer exists"))?;
        match ws.used_range() {
            Some(range) => {
                let arr = js_sys::Array::new();
                arr.push(&JsValue::from(range.start.row));
                arr.push(&JsValue::from(range.start.col));
                arr.push(&JsValue::from(range.end.row));
                arr.push(&JsValue::from(range.end.col));
                Ok(arr.into())
            }
            None => Ok(JsValue::NULL),
        }
    }

    #[wasm_bindgen(js_name = setRowHeight)]
    pub fn set_row_height(&self, row: u32, height: f64) -> Result<(), JsError> {
        let mut wb = self.workbook.borrow_mut();
        let ws = wb
            .worksheet_mut(self.sheet_index)
            .ok_or_else(|| JsError::new("Worksheet no longer exists"))?;
        ws.set_row_height(row, height);
        Ok(())
    }

    #[wasm_bindgen(js_name = setColumnWidth)]
    pub fn set_column_width(&self, col: u32, width: f64) -> Result<(), JsError> {
        let mut wb = self.workbook.borrow_mut();
        let ws = wb
            .worksheet_mut(self.sheet_index)
            .ok_or_else(|| JsError::new("Worksheet no longer exists"))?;
        ws.set_column_width(col as u16, width);
        Ok(())
    }

    #[wasm_bindgen(js_name = getRowHeight)]
    pub fn get_row_height(&self, row: u32) -> Result<Option<f64>, JsError> {
        let wb = self.workbook.borrow();
        let ws = wb
            .worksheet(self.sheet_index)
            .ok_or_else(|| JsError::new("Worksheet no longer exists"))?;
        Ok(ws.custom_row_heights().get(&row).copied())
    }

    #[wasm_bindgen(js_name = getColumnWidth)]
    pub fn get_column_width(&self, col: u32) -> Result<Option<f64>, JsError> {
        let wb = self.workbook.borrow();
        let ws = wb
            .worksheet(self.sheet_index)
            .ok_or_else(|| JsError::new("Worksheet no longer exists"))?;
        Ok(ws.custom_column_widths().get(&(col as u16)).copied())
    }

    #[wasm_bindgen(js_name = mergeCells)]
    pub fn merge_cells(&self, range_str: &str) -> Result<(), JsError> {
        let mut wb = self.workbook.borrow_mut();
        let ws = wb
            .worksheet_mut(self.sheet_index)
            .ok_or_else(|| JsError::new("Worksheet no longer exists"))?;
        let range = CellRange::parse(range_str)
            .map_err(|e| JsError::new(&format!("Invalid range: {}", e)))?;
        ws.merge_cells(&range).map_err(to_js_error)
    }

    #[wasm_bindgen(js_name = unmergeCells)]
    pub fn unmerge_cells(&self, range_str: &str) -> Result<bool, JsError> {
        let mut wb = self.workbook.borrow_mut();
        let ws = wb
            .worksheet_mut(self.sheet_index)
            .ok_or_else(|| JsError::new("Worksheet no longer exists"))?;
        let range = CellRange::parse(range_str)
            .map_err(|e| JsError::new(&format!("Invalid range: {}", e)))?;
        Ok(ws.unmerge_cells(&range))
    }

    #[wasm_bindgen(js_name = getImageAt)]
    pub fn get_image_at(&self, row: u32, col: u32) -> Result<JsValue, JsError> {
        let wb = self.workbook.borrow();
        let ws = wb
            .worksheet(self.sheet_index)
            .ok_or_else(|| JsError::new("Worksheet no longer exists"))?;

        match ws.get_image_at(row, col as u16) {
            Some(info) => to_js_value(&WasmImageInfo::from(info)),
            None => Ok(JsValue::NULL),
        }
    }

    #[wasm_bindgen(getter, js_name = pivotCount)]
    pub fn pivot_count(&self) -> Result<usize, JsError> {
        let wb = self.workbook.borrow();
        let ws = wb
            .worksheet(self.sheet_index)
            .ok_or_else(|| JsError::new("Worksheet no longer exists"))?;
        Ok(ws.pivot_table_count())
    }

    #[wasm_bindgen(getter, js_name = pivotTableNames)]
    pub fn pivot_table_names(&self) -> Result<Vec<String>, JsError> {
        let wb = self.workbook.borrow();
        let ws = wb
            .worksheet(self.sheet_index)
            .ok_or_else(|| JsError::new("Worksheet no longer exists"))?;
        Ok(ws
            .pivot_tables()
            .iter()
            .map(|pivot| pivot.name.clone())
            .collect())
    }

    #[wasm_bindgen(getter, js_name = pivotTables)]
    pub fn pivot_tables(&self) -> Result<JsValue, JsError> {
        let wb = self.workbook.borrow();
        let ws = wb
            .worksheet(self.sheet_index)
            .ok_or_else(|| JsError::new("Worksheet no longer exists"))?;
        let pivots = ws
            .pivot_tables()
            .iter()
            .map(WasmPivotTableDefinition::from)
            .collect::<Vec<_>>();
        to_js_value(&pivots)
    }

    #[wasm_bindgen(js_name = getPivotTable)]
    pub fn get_pivot_table(&self, name: &str) -> Result<JsValue, JsError> {
        let wb = self.workbook.borrow();
        let ws = wb
            .worksheet(self.sheet_index)
            .ok_or_else(|| JsError::new("Worksheet no longer exists"))?;
        match ws.pivot_table_by_name(name) {
            Some(pivot) => to_js_value(&WasmPivotTableDefinition::from(pivot)),
            None => Ok(JsValue::NULL),
        }
    }

    #[wasm_bindgen(js_name = addPivotTable)]
    pub fn add_pivot_table(&self, options: JsValue) -> Result<(), JsError> {
        let options: WasmPivotTableOptions =
            serde_wasm_bindgen::from_value(options).map_err(to_js_error)?;
        let pivot = build_pivot_table_from_wasm(options)?;
        let mut wb = self.workbook.borrow_mut();
        let ws = wb
            .worksheet_mut(self.sheet_index)
            .ok_or_else(|| JsError::new("Worksheet no longer exists"))?;
        ws.add_pivot_table(pivot).map_err(to_js_error)
    }

    #[wasm_bindgen(js_name = addPivotChart)]
    pub fn add_pivot_chart(&self, options: JsValue) -> Result<JsValue, JsError> {
        let options: WasmPivotChartOptions =
            serde_wasm_bindgen::from_value(options).map_err(to_js_error)?;
        let chart_type = parse_chart_type(options.chart_type.as_deref())?;
        let mut wb = self.workbook.borrow_mut();
        let ws = wb
            .worksheet_mut(self.sheet_index)
            .ok_or_else(|| JsError::new("Worksheet no longer exists"))?;
        let chart = ws
            .build_pivot_chart(
                &options.pivot_name,
                chart_type,
                duke_sheets::DrawingAnchor::default(),
            )
            .map_err(to_js_error)?;
        ws.add_chart(chart.clone());
        to_js_value(&WasmChart::from(&chart))
    }
}

#[wasm_bindgen]
pub struct Workbook {
    pub(crate) inner: Rc<RefCell<CoreWorkbook>>,
}

#[wasm_bindgen]
impl Workbook {
    #[wasm_bindgen(constructor)]
    pub fn new() -> Self {
        Self {
            inner: Rc::new(RefCell::new(CoreWorkbook::new())),
        }
    }

    #[wasm_bindgen(js_name = loadCsvString)]
    pub fn load_csv_string(csv: &str) -> Result<Workbook, JsError> {
        let reader = Cursor::new(csv.as_bytes());
        let ws =
            duke_sheets_csv::CsvReader::read(reader, &duke_sheets_csv::CsvReadOptions::default())
                .map_err(to_js_error)?;
        let mut wb = CoreWorkbook::empty();
        wb.add_existing_worksheet(ws).map_err(to_js_error)?;
        Ok(Self {
            inner: Rc::new(RefCell::new(wb)),
        })
    }

    #[wasm_bindgen(js_name = fromCsvString)]
    pub fn from_csv_string(csv: &str) -> Result<Workbook, JsError> {
        Self::load_csv_string(csv)
    }

    /// Load a workbook from bytes, auto-detecting the format (XLSX or XLS).
    #[wasm_bindgen(js_name = fromBytes)]
    pub fn from_bytes(data: &[u8]) -> Result<Workbook, JsError> {
        use duke_sheets::WorkbookExt;
        let wb = duke_sheets_core::Workbook::from_bytes(data)
            .map_err(|e| JsError::new(&format!("Failed to read file: {}", e)))?;
        Ok(Self {
            inner: Rc::new(RefCell::new(wb)),
        })
    }

    /// Load a password-protected workbook from bytes, auto-detecting
    /// the format (XLSX or XLS). `skipIntegrityCheck` skips the HMAC
    /// integrity check on Agile-encrypted files (matches Office
    /// behaviour); defaults to false.
    #[wasm_bindgen(js_name = fromBytesWithPassword)]
    pub fn from_bytes_with_password(
        data: &[u8],
        password: &str,
        skip_integrity_check: Option<bool>,
    ) -> Result<Workbook, JsError> {
        use duke_sheets::{WorkbookExt, WorkbookOpenOptions};
        let mut opts = WorkbookOpenOptions::default().password(password);
        if skip_integrity_check.unwrap_or(false) {
            opts = opts.skip_integrity_check();
        }
        let wb = duke_sheets_core::Workbook::from_bytes_with(data, &opts)
            .map_err(|e| JsError::new(&format!("Failed to read file: {}", e)))?;
        Ok(Self {
            inner: Rc::new(RefCell::new(wb)),
        })
    }

    #[wasm_bindgen(js_name = saveXlsxBytes)]
    pub fn save_xlsx_bytes(&self) -> Result<Vec<u8>, JsError> {
        let wb = self.inner.borrow();
        let mut buf = Vec::new();
        XlsxWriter::write(&wb, Cursor::new(&mut buf)).map_err(to_js_error)?;
        Ok(buf)
    }

    /// Save the workbook as encrypted XLSX bytes. `profile` selects
    /// the encryption variant; passing `null`/`undefined` uses the
    /// Agile-256 default. Valid values: `"agile"`, `"standard"`.
    /// `keyBits` and `spinCount` override the defaults where the
    /// profile supports them.
    #[wasm_bindgen(js_name = saveXlsxBytesEncrypted)]
    pub fn save_xlsx_bytes_encrypted(
        &self,
        password: &str,
        profile: Option<String>,
        key_bits: Option<u32>,
        spin_count: Option<u32>,
    ) -> Result<Vec<u8>, JsError> {
        let wb = self.inner.borrow();
        let xlsx_profile = parse_xlsx_profile(profile.as_deref(), key_bits, spin_count)
            .map_err(|e| JsError::new(&e))?;
        XlsxWriter::write_to_bytes_encrypted(&wb, password, &xlsx_profile).map_err(to_js_error)
    }

    /// Save the workbook as encrypted XLS bytes. `profile` selects
    /// the FilePass variant; `null` defaults to RC4 CryptoAPI 128.
    /// Valid values: `"rc4-cryptoapi"`, `"rc4-legacy"`, `"xor"`.
    /// `keyBits` controls RC4 CryptoAPI key size (40 or 128). XOR is
    /// not certified to interoperate with modern Excel.
    #[wasm_bindgen(js_name = saveXlsBytesEncrypted)]
    pub fn save_xls_bytes_encrypted(
        &self,
        password: &str,
        profile: Option<String>,
        key_bits: Option<u32>,
    ) -> Result<Vec<u8>, JsError> {
        let wb = self.inner.borrow();
        let variant =
            parse_xls_variant(profile.as_deref(), key_bits).map_err(|e| JsError::new(&e))?;
        duke_sheets_xls::XlsWriter::write_to_bytes_encrypted(&wb, password, variant)
            .map_err(to_js_error)
    }

    #[wasm_bindgen(js_name = saveXlsbBytes)]
    pub fn save_xlsb_bytes(&self) -> Result<Vec<u8>, JsError> {
        let wb = self.inner.borrow();
        let mut buf = Vec::new();
        XlsbWriter::write(&wb, Cursor::new(&mut buf)).map_err(to_js_error)?;
        Ok(buf)
    }

    #[wasm_bindgen(js_name = saveCsvString)]
    pub fn save_csv_string(&self) -> Result<String, JsError> {
        let wb = self.inner.borrow();
        let ws = wb
            .worksheet(0)
            .ok_or_else(|| JsError::new("No worksheets to save"))?;
        let mut buf = Vec::new();
        duke_sheets_csv::CsvWriter::write(
            ws,
            &mut buf,
            &duke_sheets_csv::CsvWriteOptions::default(),
        )
        .map_err(to_js_error)?;
        String::from_utf8(buf).map_err(to_js_error)
    }

    #[wasm_bindgen(getter, js_name = sheetCount)]
    pub fn sheet_count(&self) -> Result<usize, JsError> {
        let wb = self.inner.borrow();
        Ok(wb.sheet_count())
    }

    #[wasm_bindgen(getter, js_name = sheetNames)]
    pub fn sheet_names(&self) -> Result<Vec<String>, JsError> {
        let wb = self.inner.borrow();
        Ok((0..wb.sheet_count())
            .filter_map(|i| wb.worksheet(i).map(|ws| ws.name().to_string()))
            .collect())
    }

    #[wasm_bindgen(js_name = getSheet)]
    pub fn get_sheet(&self, index: usize) -> Result<Worksheet, JsError> {
        let wb = self.inner.borrow();
        if index >= wb.sheet_count() {
            return Err(JsError::new(&format!("Sheet index {} out of range", index)));
        }
        drop(wb);
        Ok(Worksheet {
            workbook: Rc::clone(&self.inner),
            sheet_index: index,
        })
    }

    #[wasm_bindgen(js_name = getSheetByName)]
    pub fn get_sheet_by_name(&self, name: &str) -> Result<Worksheet, JsError> {
        let wb = self.inner.borrow();
        let index = wb
            .sheet_index(name)
            .ok_or_else(|| JsError::new(&format!("Sheet '{}' not found", name)))?;
        drop(wb);
        Ok(Worksheet {
            workbook: Rc::clone(&self.inner),
            sheet_index: index,
        })
    }

    #[wasm_bindgen(js_name = addSheet)]
    pub fn add_sheet(&self, name: &str) -> Result<usize, JsError> {
        let mut wb = self.inner.borrow_mut();
        wb.add_worksheet_with_name(name).map_err(to_js_error)
    }

    #[wasm_bindgen(js_name = removeSheet)]
    pub fn remove_sheet(&self, index: usize) -> Result<(), JsError> {
        let mut wb = self.inner.borrow_mut();
        wb.remove_worksheet(index).map(|_| ()).map_err(to_js_error)
    }

    pub fn calculate(&self, options: Option<JsValue>) -> Result<JsValue, JsError> {
        let mut wb = self.inner.borrow_mut();
        let options = if let Some(options) = options {
            // Extract JS callback functions before serde deserialization (which consumes the JsValue).
            // Serde can't deserialize JS functions, so we pull them out via Reflect::get.
            let web_service_js_fn =
                js_sys::Reflect::get(&options, &JsValue::from_str("webServiceFn"))
                    .ok()
                    .and_then(|v| v.dyn_into::<js_sys::Function>().ok());
            let rtd_js_fn = js_sys::Reflect::get(&options, &JsValue::from_str("rtdFn"))
                .ok()
                .and_then(|v| v.dyn_into::<js_sys::Function>().ok());
            let external_js_fn = js_sys::Reflect::get(&options, &JsValue::from_str("externalFn"))
                .ok()
                .and_then(|v| v.dyn_into::<js_sys::Function>().ok())
                .map(|f| (f, false))
                .or_else(|| {
                    js_sys::Reflect::get(&options, &JsValue::from_str("externalFnFn"))
                        .ok()
                        .and_then(|v| v.dyn_into::<js_sys::Function>().ok())
                        .map(|f| (f, true))
                });

            let js_opts: JsCalculationOptions =
                serde_wasm_bindgen::from_value(options).map_err(to_js_error)?;

            // Build web_service_fn from callback
            let web_service_fn = web_service_js_fn.map(|js_fn| {
                let wrapper = SendSyncFunction(js_fn);
                Arc::new(move |url: &str| -> Option<String> {
                    let result = wrapper
                        .call1(&JsValue::NULL, &JsValue::from_str(url))
                        .ok()?;
                    result.as_string()
                }) as Arc<dyn Fn(&str) -> Option<String> + Send + Sync>
            });

            // Build rtd_fn from callback
            let rtd_fn = rtd_js_fn.map(|js_fn| {
                let wrapper = SendSyncFunction(js_fn);
                Arc::new(
                    move |prog_id: &str, server: &str, topics: &[String]| -> Option<String> {
                        let topics_arr = js_sys::Array::new();
                        for t in topics {
                            topics_arr.push(&JsValue::from_str(t));
                        }
                        let result = wrapper
                            .call3(
                                &JsValue::NULL,
                                &JsValue::from_str(prog_id),
                                &JsValue::from_str(server),
                                &topics_arr.into(),
                            )
                            .ok()?;
                        result.as_string()
                    },
                )
                    as Arc<dyn Fn(&str, &str, &[String]) -> Option<String> + Send + Sync>
            });

            // Build external_fn from callback: externalFn(book, name, args[]) -> number|string|bool|null
            let external_fn = external_js_fn.map(|(js_fn, legacy_two_arg)| {
                let wrapper = SendSyncFunction(js_fn);
                Arc::new(
                    move |book: &str, name: &str, args: &[String]| -> Option<FormulaValue> {
                        let args_arr = js_sys::Array::new();
                        for a in args {
                            args_arr.push(&JsValue::from_str(a));
                        }
                        let result = if legacy_two_arg {
                            wrapper.call2(
                                &JsValue::NULL,
                                &JsValue::from_str(name),
                                &args_arr.into(),
                            )
                        } else {
                            wrapper.call3(
                                &JsValue::NULL,
                                &JsValue::from_str(book),
                                &JsValue::from_str(name),
                                &args_arr.into(),
                            )
                        }
                        .ok()?;
                        js_value_to_formula_value(result)
                    },
                )
                    as Arc<dyn Fn(&str, &str, &[String]) -> Option<FormulaValue> + Send + Sync>
            });

            CalculationOptions {
                iterative: js_opts.iterative.unwrap_or(false),
                max_iterations: js_opts.max_iterations.unwrap_or(100),
                max_change: js_opts.max_change.unwrap_or(0.001),
                force_full_calculation: js_opts.force_full_calculation.unwrap_or(true),
                calculate_volatile: js_opts.calculate_volatile.unwrap_or(true),
                sheets: js_opts.sheets.unwrap_or_default(),
                max_threads: js_opts.max_threads,
                web_service_fn,
                rtd_fn,
                external_fn,
            }
        } else {
            CalculationOptions::default()
        };
        let stats = wb.calculate_with_options(&options).map_err(to_js_error)?;
        to_js_value(&WasmCalculationStats::from(&stats))
    }

    #[wasm_bindgen(js_name = refreshPivots)]
    pub fn refresh_pivots(&self) -> Result<JsValue, JsError> {
        let mut wb = self.inner.borrow_mut();
        let stats = wb.refresh_pivots().map_err(to_js_error)?;
        to_js_value(&WasmPivotRefreshStats::from(stats))
    }

    #[wasm_bindgen(js_name = addDataConnection)]
    pub fn add_data_connection(&self, options: JsValue) -> Result<(), JsError> {
        let options: WasmWorkbookConnectionOptions =
            serde_wasm_bindgen::from_value(options).map_err(to_js_error)?;
        let mut wb = self.inner.borrow_mut();
        wb.add_data_connection(build_workbook_connection_from_wasm(options)?)
            .map_err(to_js_error)
    }

    #[wasm_bindgen(getter, js_name = dataConnectionCount)]
    pub fn data_connection_count(&self) -> usize {
        let wb = self.inner.borrow();
        wb.data_connections().len()
    }

    #[wasm_bindgen(getter, js_name = dataConnectionNames)]
    pub fn data_connection_names(&self) -> Vec<String> {
        let wb = self.inner.borrow();
        wb.data_connections()
            .iter()
            .map(|connection| connection.name.clone())
            .collect()
    }

    #[wasm_bindgen(getter, js_name = dataConnections)]
    pub fn data_connections(&self) -> Result<JsValue, JsError> {
        let wb = self.inner.borrow();
        let connections = wb
            .data_connections()
            .iter()
            .map(WasmWorkbookConnectionDefinition::from)
            .collect::<Vec<_>>();
        to_js_value(&connections)
    }

    #[wasm_bindgen(js_name = getDataConnection)]
    pub fn get_data_connection(&self, name: &str) -> Result<JsValue, JsError> {
        let wb = self.inner.borrow();
        match wb.data_connection_by_name(name) {
            Some(connection) => to_js_value(&WasmWorkbookConnectionDefinition::from(connection)),
            None => Ok(JsValue::NULL),
        }
    }

    #[wasm_bindgen(js_name = getDataConnectionById)]
    pub fn get_data_connection_by_id(&self, id: u32) -> Result<JsValue, JsError> {
        let wb = self.inner.borrow();
        match wb.data_connection_by_id(id) {
            Some(connection) => to_js_value(&WasmWorkbookConnectionDefinition::from(connection)),
            None => Ok(JsValue::NULL),
        }
    }

    #[wasm_bindgen(js_name = defineName)]
    pub fn define_name(&self, name: &str, refers_to: &str) -> Result<(), JsError> {
        let mut wb = self.inner.borrow_mut();
        wb.define_name(name, refers_to).map_err(to_js_error)
    }

    #[wasm_bindgen(js_name = getNamedRange)]
    pub fn get_named_range(&self, name: &str) -> Result<Option<String>, JsError> {
        let wb = self.inner.borrow();
        Ok(wb.get_named_range(name, 0).map(|nr| nr.refers_to.clone()))
    }
}

impl Default for Workbook {
    fn default() -> Self {
        Self::new()
    }
}

/// Convert a JS value returned by an `externalFnFn` callback into a typed FormulaValue.
/// `null`/`undefined` -> `None` (yields `#N/A`). Booleans take priority over numbers.
fn js_value_to_formula_value(value: JsValue) -> Option<FormulaValue> {
    if value.is_null() || value.is_undefined() {
        None
    } else if let Some(b) = value.as_bool() {
        Some(FormulaValue::Boolean(b))
    } else if let Some(n) = value.as_f64() {
        Some(FormulaValue::Number(n))
    } else if let Some(s) = value.as_string() {
        Some(FormulaValue::String(s))
    } else {
        None
    }
}

fn js_to_cell_value(value: JsValue) -> Result<CoreCellValue, JsError> {
    if value.is_null() || value.is_undefined() {
        Ok(CoreCellValue::Empty)
    } else if let Some(b) = value.as_bool() {
        Ok(CoreCellValue::Boolean(b))
    } else if let Some(n) = value.as_f64() {
        Ok(CoreCellValue::Number(n))
    } else if let Some(s) = value.as_string() {
        Ok(CoreCellValue::string(s))
    } else {
        Err(JsError::new(
            "Cell value must be null, boolean, number, or string",
        ))
    }
}

#[wasm_bindgen(start)]
pub fn init() {}
