//! Prelude module - common imports for duke-sheets users
//!
//! ```rust
//! use duke_sheets::prelude::*;
//! ```

pub use crate::{
    // Style types
    Alignment,
    BorderEdge,
    BorderLineStyle,
    BorderStyle,
    // Calculation types
    CalculationOptions,
    CalculationStats,
    CellAddress,
    // Comments
    CellComment,
    CellError,

    CellRange,
    // Cell types
    CellValue,
    // Conditional formatting types
    CfOperator,
    CfRuleType,
    // Chart types
    Chart,
    ChartType,

    Color,
    ConditionalFormatRule,

    CsvReader,
    CsvWriter,

    // Data validation types
    DataValidation,

    // Error types
    Error,
    FillStyle,
    // Format detection
    FileFormat,
    FontStyle,
    HorizontalAlignment,
    Hyperlink,
    IconSetStyle,
    NumberFormat,
    Result,

    // Rich text types
    RichTextRun,
    RunFont,
    Style,
    ValidationErrorStyle,
    ValidationOperator,
    ValidationType,
    VerticalAlignment,
    // Main types
    Workbook,
    // Extension traits
    WorkbookCalculationExt,
    WorkbookExt,
    Worksheet,

    // I/O types
    XlsxReader,
    XlsxWriter,
};
