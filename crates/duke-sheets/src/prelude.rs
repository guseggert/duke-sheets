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
    CheckState,
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

    DrawingAnchor,
    EmbeddedImage,
    // Error types
    Error,
    // Format detection
    FileFormat,
    FillStyle,
    FontStyle,
    // Form control types
    FormControl,
    FormControlKind,
    HorizontalAlignment,
    Hyperlink,
    IconSetStyle,
    ImageFormat,

    ImageInfo,
    ImageSizing,
    ListSelection,
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
