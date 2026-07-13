//! Prelude module - common imports for duke-sheets users
//!
//! ```rust
//! use duke_sheets::prelude::*;
//! ```

pub use crate::{
    // Drawing helpers
    default_comment_anchor,
    radio_groups,
    validate_anchor,
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
    ChartEx,
    ChartType,
    CheckState,
    // Drawing types
    ChildTransform,
    Color,
    CommentRef,
    ConditionalFormatRule,

    ControlText,
    CsvReader,
    CsvWriter,

    // Data validation types
    DataValidation,

    DrawingAnchor,
    DrawingKind,
    DrawingMeta,
    DrawingNodeMut,
    DrawingNodeRef,
    DrawingObject,
    DrawingPath,
    DrawingText,
    Drawn,
    EmbeddedImage,
    // Error types
    Error,
    // Format detection
    FileFormat,
    FillStyle,
    FontStyle,
    // Form control types
    FormControl,
    FormControlInteractionResult,
    FormControlKind,
    Group,
    GroupChild,
    GroupTransform,
    HorizontalAlignment,
    Hyperlink,
    IconSetStyle,
    ImageFormat,

    ImageInfo,
    ImageSizing,
    ListSelection,
    NumberFormat,
    PlacedControl,
    RawDrawing,
    RawRel,
    RectEmu,
    Result,

    // Rich text types
    RichTextRun,
    RunFont,
    Shape,
    ShapeFill,
    ShapeGeometry,
    ShapeLine,
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
