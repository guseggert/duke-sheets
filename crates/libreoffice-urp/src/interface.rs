//! Hardcoded UNO interface definitions for the interfaces we need.
//!
//! Each interface definition includes:
//! - The fully-qualified type name
//! - The base interface(s) — for computing cumulative method indices
//! - Method signatures (name, parameter types, return type, one-way flag)
//!
//! Method indices are cumulative: XInterface has methods 0-2 (queryInterface,
//! acquire, release). Any interface extending XInterface starts its own methods
//! at index 3.
//!
//! Note: `acquire` (index 1) and `release` (index 2) are never sent over URP.
//! Reference counting is implicit in the bridge.

use crate::types::Type;

/// A method signature in a UNO interface.
#[derive(Debug, Clone)]
pub struct MethodDef {
    pub name: &'static str,
    /// Absolute method index (cumulative across inheritance chain).
    pub index: u16,
    pub params: &'static [ParamType],
    pub return_type: Type,
    /// If true, this is a one-way method (no reply expected).
    pub one_way: bool,
}

/// A parameter type specification.
#[derive(Debug, Clone, Copy)]
pub enum ParamType {
    Type,
    String,
    Long,
    Short,
    Bool,
    Double,
    Any,
    Interface(&'static str),
    Struct(&'static str),
    Enum(&'static str),
    SequenceOfPropertyValue,
    SequenceOfProtocolProperty,
}

impl ParamType {
    pub fn to_type(&self) -> Type {
        match self {
            ParamType::Type => Type::r#type(),
            ParamType::String => Type::string(),
            ParamType::Long => Type::long(),
            ParamType::Short => Type::short(),
            ParamType::Bool => Type::boolean(),
            ParamType::Double => Type::double(),
            ParamType::Any => Type::any(),
            ParamType::Interface(name) => Type::interface(*name),
            ParamType::Struct(name) => Type::r#struct(*name),
            ParamType::Enum(name) => Type::r#enum(*name),
            ParamType::SequenceOfPropertyValue => {
                Type::sequence("com.sun.star.beans.PropertyValue")
            }
            ParamType::SequenceOfProtocolProperty => {
                Type::sequence("com.sun.star.bridge.ProtocolProperty")
            }
        }
    }
}

// ============================================================================
// XInterface (base of everything)
// Methods: queryInterface(0), acquire(1), release(2)
// ============================================================================

pub fn query_interface() -> MethodDef {
    MethodDef {
        name: "queryInterface",
        index: 0,
        params: &[ParamType::Type],
        return_type: Type::any(),
        one_way: false,
    }
}

pub fn release() -> MethodDef {
    MethodDef {
        name: "release",
        index: 2,
        params: &[],
        return_type: Type::void(),
        one_way: true,
    }
}

// ============================================================================
// XProtocolProperties (special, used for negotiation)
// Extends XInterface (0-2)
// Methods: queryInterface(0), acquire(1), release(2), ???(3),
//          requestChange(4), commitChange(5)
//
// Actually, the spec says requestChange=4, commitChange=5 as absolute indices.
// XProtocolProperties extends XInterface, and XInterface has 3 methods (0,1,2).
// XProtocolProperties adds: requestChange(3), commitChange(4)
// BUT the spec uses IDs 4 and 5 — this is because there's an intermediate
// interface. Looking at the LO source, the IDs are hardcoded as 4 and 5.
// ============================================================================

pub fn request_change() -> MethodDef {
    MethodDef {
        name: "requestChange",
        index: 4,
        params: &[ParamType::Long],
        return_type: Type::long(),
        one_way: false,
    }
}

pub fn commit_change() -> MethodDef {
    MethodDef {
        name: "commitChange",
        index: 5,
        params: &[ParamType::SequenceOfProtocolProperty],
        return_type: Type::void(),
        one_way: false,
    }
}

// ============================================================================
// XComponentContext
// Extends XInterface (0-2)
// Methods: getValueByName(3), getServiceManager(4)
// ============================================================================

pub fn get_service_manager() -> MethodDef {
    MethodDef {
        name: "getServiceManager",
        index: 4,
        params: &[],
        return_type: Type::interface("com.sun.star.lang.XMultiComponentFactory"),
        one_way: false,
    }
}

// ============================================================================
// XMultiComponentFactory
// Extends XInterface (0-2)
// Methods: createInstanceWithContext(3), createInstanceWithArgumentsAndContext(4)
// ============================================================================

pub fn create_instance_with_context() -> MethodDef {
    MethodDef {
        name: "createInstanceWithContext",
        index: 3,
        params: &[
            ParamType::String,
            ParamType::Interface("com.sun.star.uno.XComponentContext"),
        ],
        return_type: Type::interface("com.sun.star.uno.XInterface"),
        one_way: false,
    }
}

// ============================================================================
// XComponentLoader
// Extends XInterface (0-2)
// Methods: loadComponentFromURL(3)
// ============================================================================

pub fn load_component_from_url() -> MethodDef {
    MethodDef {
        name: "loadComponentFromURL",
        index: 3,
        params: &[
            ParamType::String,                  // URL
            ParamType::String,                  // TargetFrameName
            ParamType::Long,                    // SearchFlags
            ParamType::SequenceOfPropertyValue, // Arguments
        ],
        return_type: Type::interface("com.sun.star.lang.XComponent"),
        one_way: false,
    }
}

// ============================================================================
// XSpreadsheetDocument
// Extends XInterface via XModel (complex chain, but getSheets is at a known index)
//
// Inheritance: XSpreadsheetDocument -> XModel -> XComponent -> XInterface
// XComponent: dispose(3), addEventListener(4), removeEventListener(5)
// XModel: attachResource(6), getURL(7), getArgs(8), connectController(9),
//         disconnectController(10), lockControllers(11), unlockControllers(12),
//         hasControllersLocked(13), getCurrentController(14),
//         getCurrentSelection(15)
// XSpreadsheetDocument: getSheets(16)
//
// However, this depends on the exact IDL. For safety, we'll use queryInterface
// to get the right interface proxy and call methods on it.
// The method index for getSheets on XSpreadsheetDocument (which directly extends
// XInterface in its IDL, NOT XModel) is:
// XInterface(0,1,2), getSheets(3)
// ============================================================================

pub fn get_sheets() -> MethodDef {
    MethodDef {
        name: "getSheets",
        index: 3,
        params: &[],
        return_type: Type::interface("com.sun.star.sheet.XSpreadsheets"),
        one_way: false,
    }
}

// ============================================================================
// XIndexAccess
// Extends XElementAccess which extends XInterface
// XElementAccess: getElementType(3), hasElements(4)
// XIndexAccess: getCount(5), getByIndex(6)
// ============================================================================

pub fn get_by_index() -> MethodDef {
    MethodDef {
        name: "getByIndex",
        index: 6,
        params: &[ParamType::Long],
        return_type: Type::any(),
        one_way: false,
    }
}

pub fn get_count() -> MethodDef {
    MethodDef {
        name: "getCount",
        index: 5,
        params: &[],
        return_type: Type::long(),
        one_way: false,
    }
}

// ============================================================================
// XSpreadsheet
// Extends XSheetCellRange -> XCellRange -> XInterface
// XCellRange: getCellByPosition(3), getCellRangeByPosition(4), getCellRangeByName(5)
// XSheetCellRange: getSpreadsheet(6)
// XSpreadsheet: createCursor(7), createCursorByRange(8)
//
// Actually XSpreadsheet extends XSheetCellRange which extends XCellRange.
// We primarily need getCellByPosition which comes from XCellRange.
// XCellRange extends XInterface:
//   getCellByPosition(3), getCellRangeByPosition(4), getCellRangeByName(5)
// ============================================================================

pub fn get_cell_by_position() -> MethodDef {
    MethodDef {
        name: "getCellByPosition",
        index: 3,
        params: &[ParamType::Long, ParamType::Long], // column, row
        return_type: Type::interface("com.sun.star.table.XCell"),
        one_way: false,
    }
}

// ============================================================================
// XCell
// Extends XInterface
// Methods: getFormula(3), setFormula(4), getValue(5), setValue(6), getType(7), getError(8)
// ============================================================================

pub fn cell_get_formula() -> MethodDef {
    MethodDef {
        name: "getFormula",
        index: 3,
        params: &[],
        return_type: Type::string(),
        one_way: false,
    }
}

pub fn cell_set_formula() -> MethodDef {
    MethodDef {
        name: "setFormula",
        index: 4,
        params: &[ParamType::String],
        return_type: Type::void(),
        one_way: false,
    }
}

pub fn cell_get_value() -> MethodDef {
    MethodDef {
        name: "getValue",
        index: 5,
        params: &[],
        return_type: Type::double(),
        one_way: false,
    }
}

pub fn cell_set_value() -> MethodDef {
    MethodDef {
        name: "setValue",
        index: 6,
        params: &[ParamType::Double], // f64
        return_type: Type::void(),
        one_way: false,
    }
}

pub fn cell_get_type() -> MethodDef {
    MethodDef {
        name: "getType",
        index: 7,
        params: &[],
        return_type: Type::r#enum("com.sun.star.table.CellContentType"),
        one_way: false,
    }
}

// ============================================================================
// XText / XSimpleText / XTextRange — for setting string cell values
// XTextRange extends XInterface: getString(3), setString(4)
// XSimpleText extends XTextRange: createTextCursor(5), ...
// XText extends XSimpleText: ...
//
// For cells, we use XTextRange::setString/getString since XCell implements it.
// Actually, XCell doesn't directly implement XTextRange. We need to
// queryInterface for XTextRange on the cell.
//
// XTextRange extends XInterface:
//   getText(3), getStart(4), getEnd(5), getString(6), setString(7)
// ============================================================================

pub fn text_range_get_string() -> MethodDef {
    MethodDef {
        name: "getString",
        index: 6,
        params: &[],
        return_type: Type::string(),
        one_way: false,
    }
}

pub fn text_range_set_string() -> MethodDef {
    MethodDef {
        name: "setString",
        index: 7,
        params: &[ParamType::String],
        return_type: Type::void(),
        one_way: false,
    }
}

// ============================================================================
// XStorable
// Extends XInterface
// Methods: hasLocation(3), getLocation(4), isReadonly(5), store(6),
//          storeAsURL(7), storeToURL(8)
// ============================================================================

pub fn store_to_url() -> MethodDef {
    MethodDef {
        name: "storeToURL",
        index: 8,
        params: &[
            ParamType::String,                  // URL
            ParamType::SequenceOfPropertyValue, // Args
        ],
        return_type: Type::void(),
        one_way: false,
    }
}

// ============================================================================
// XCloseable
// Extends XInterface (actually extends XCloseBroadcaster which extends XInterface)
// XCloseBroadcaster: addCloseListener(3), removeCloseListener(4)
// XCloseable: close(5)
// ============================================================================

pub fn closeable_close() -> MethodDef {
    MethodDef {
        name: "close",
        index: 5,
        params: &[ParamType::Bool], // bDeliverOwnership
        return_type: Type::void(),
        one_way: false,
    }
}

// ============================================================================
// XPropertySet
// Extends XInterface
// Methods: getPropertySetInfo(3), setPropertyValue(4), getPropertyValue(5),
//          addPropertyChangeListener(6), removePropertyChangeListener(7),
//          addVetoableChangeListener(8), removeVetoableChangeListener(9)
// ============================================================================

pub fn set_property_value() -> MethodDef {
    MethodDef {
        name: "setPropertyValue",
        index: 4,
        params: &[ParamType::String, ParamType::Any],
        return_type: Type::void(),
        one_way: false,
    }
}

pub fn get_property_value() -> MethodDef {
    MethodDef {
        name: "getPropertyValue",
        index: 5,
        params: &[ParamType::String],
        return_type: Type::any(),
        one_way: false,
    }
}

// ============================================================================
// XNameContainer / XNameAccess (for XSpreadsheets which extends both)
// XNameAccess extends XElementAccess -> XInterface
// XElementAccess: getElementType(3), hasElements(4)
// XNameAccess: getByName(5), getElementNames(6), hasByName(7)
// XNameReplace extends XNameAccess: replaceByName(8)
// XNameContainer extends XNameReplace: insertByName(9), removeByName(10)
// ============================================================================

pub fn insert_new_by_name() -> MethodDef {
    // This is on XSpreadsheets specifically, not XNameContainer.
    // XSpreadsheets extends XNameContainer:
    // XNameContainer methods end at 10.
    // XSpreadsheets: insertNewByName(11), moveByName(12), copyByName(13)
    //
    // But wait: XSpreadsheets also extends XIndexAccess.
    // Multi-inheritance in UNO: the method indices depend on the IDL order.
    // XSpreadsheets inherits: XNameContainer + XIndexAccess + XEnumerationAccess
    // This gets complicated. For safety, we'll use queryInterface to get
    // the correct proxy and rely on the LibreOffice-specific numbering.
    //
    // From the IDL:
    // XSpreadsheets: XNameContainer
    //   insertNewByName(string, short): first method after XNameContainer
    //   moveByName(string, short): second
    //   copyByName(string, string, short): third
    //
    // XNameContainer extends XNameReplace(8) extends XNameAccess(5,6,7)
    //   extends XElementAccess(3,4) extends XInterface(0,1,2)
    // XNameReplace: replaceByName(8)
    // XNameContainer: insertByName(9), removeByName(10)
    // XSpreadsheets: insertNewByName(11), moveByName(12), copyByName(13)
    MethodDef {
        name: "insertNewByName",
        index: 11,
        params: &[ParamType::String, ParamType::Short],
        return_type: Type::void(),
        one_way: false,
    }
}

// ============================================================================
// XCellRange — getCellRangeByName
// Extends XInterface (0-2)
// Methods: getCellByPosition(3), getCellRangeByPosition(4), getCellRangeByName(5)
// ============================================================================

pub fn get_cell_range_by_name() -> MethodDef {
    MethodDef {
        name: "getCellRangeByName",
        index: 5,
        params: &[ParamType::String],
        return_type: Type::interface("com.sun.star.table.XCellRange"),
        one_way: false,
    }
}

// ============================================================================
// XMergeable
// Extends XInterface (0-2)
// Methods: merge(3), getIsMerged(4)
// ============================================================================

pub fn merge() -> MethodDef {
    MethodDef {
        name: "merge",
        index: 3,
        params: &[ParamType::Bool],
        return_type: Type::void(),
        one_way: false,
    }
}

// ============================================================================
// XSheetAnnotationsSupplier
// Extends XInterface (0-2)
// Methods: getAnnotations(3)
// ============================================================================

pub fn get_annotations() -> MethodDef {
    MethodDef {
        name: "getAnnotations",
        index: 3,
        params: &[],
        return_type: Type::interface("com.sun.star.sheet.XSheetAnnotations"),
        one_way: false,
    }
}

// ============================================================================
// XSheetAnnotations
// Extends XInterface -> XElementAccess -> XIndexAccess -> XSheetAnnotations
// XElementAccess: getElementType(3), hasElements(4)
// XIndexAccess: getCount(5), getByIndex(6)
// XSheetAnnotations: insertNew(7), removeByIndex(8)
// ============================================================================

pub fn annotations_insert_new() -> MethodDef {
    MethodDef {
        name: "insertNew",
        index: 7,
        params: &[
            ParamType::Struct("com.sun.star.table.CellAddress"),
            ParamType::String,
        ],
        return_type: Type::void(),
        one_way: false,
    }
}

// ============================================================================
// XColumnRowRange
// Extends XInterface (0-2)
// Methods: getColumns(3), getRows(4)
// ============================================================================

pub fn get_columns() -> MethodDef {
    MethodDef {
        name: "getColumns",
        index: 3,
        params: &[],
        return_type: Type::interface("com.sun.star.table.XTableColumns"),
        one_way: false,
    }
}

pub fn get_rows() -> MethodDef {
    MethodDef {
        name: "getRows",
        index: 4,
        params: &[],
        return_type: Type::interface("com.sun.star.table.XTableRows"),
        one_way: false,
    }
}

// ============================================================================
// XNamed
// Extends XInterface (0-2)
// Methods: getName(3), setName(4)
// ============================================================================

pub fn set_name() -> MethodDef {
    MethodDef {
        name: "setName",
        index: 4,
        params: &[ParamType::String],
        return_type: Type::void(),
        one_way: false,
    }
}

// ============================================================================
// XNumberFormatsSupplier
// Extends XInterface (0-2)
// Methods: getNumberFormatSettings(3), getNumberFormats(4)
// ============================================================================

pub fn get_number_formats() -> MethodDef {
    MethodDef {
        name: "getNumberFormats",
        index: 4,
        params: &[],
        return_type: Type::interface("com.sun.star.util.XNumberFormats"),
        one_way: false,
    }
}

// ============================================================================
// XNumberFormats
// Extends XInterface (0-2)
// Methods: getByKey(3), queryKeys(4), queryKey(5), addNew(6), ...
// ============================================================================

pub fn number_formats_query_key() -> MethodDef {
    MethodDef {
        name: "queryKey",
        index: 5,
        params: &[
            ParamType::String,                             // format string
            ParamType::Struct("com.sun.star.lang.Locale"), // locale
            ParamType::Bool,                               // bScanAllFormats
        ],
        return_type: Type::long(),
        one_way: false,
    }
}

pub fn number_formats_add_new() -> MethodDef {
    MethodDef {
        name: "addNew",
        index: 6,
        params: &[
            ParamType::String,                             // format string
            ParamType::Struct("com.sun.star.lang.Locale"), // locale
        ],
        return_type: Type::long(),
        one_way: false,
    }
}

// ============================================================================
// XNameAccess
// Extends XInterface -> XElementAccess (3,4) -> XNameAccess
// Methods: getByName(5), getElementNames(6), hasByName(7)
// ============================================================================

pub fn get_by_name() -> MethodDef {
    MethodDef {
        name: "getByName",
        index: 5,
        params: &[ParamType::String],
        return_type: Type::any(),
        one_way: false,
    }
}

// ============================================================================
// XNameContainer
// Extends XNameAccess -> XNameReplace(8) -> XNameContainer
// Methods: insertByName(9), removeByName(10)
// ============================================================================

pub fn insert_by_name() -> MethodDef {
    MethodDef {
        name: "insertByName",
        index: 9,
        params: &[ParamType::String, ParamType::Any],
        return_type: Type::void(),
        one_way: false,
    }
}

// ============================================================================
// XMultiServiceFactory (on the document, NOT the global ServiceManager)
// Extends XInterface (0-2)
// Methods: createInstance(3), createInstanceWithArguments(4), getAvailableServiceNames(5)
// ============================================================================

pub fn doc_create_instance() -> MethodDef {
    MethodDef {
        name: "createInstance",
        index: 3,
        params: &[ParamType::String],
        return_type: Type::interface("com.sun.star.uno.XInterface"),
        one_way: false,
    }
}

// ============================================================================
// XSheetConditionalEntries
// Extends XInterface -> XElementAccess -> XIndexAccess -> XSheetConditionalEntries
// Methods: addNew(7), removeByIndex(8), clear(9)
// ============================================================================

pub fn conditional_entries_add_new() -> MethodDef {
    MethodDef {
        name: "addNew",
        index: 7,
        params: &[ParamType::SequenceOfPropertyValue],
        return_type: Type::void(),
        one_way: false,
    }
}

// ============================================================================
// XStyleFamiliesSupplier
// Extends XInterface (0-2)
// Methods: getStyleFamilies(3)
// ============================================================================

pub fn get_style_families() -> MethodDef {
    MethodDef {
        name: "getStyleFamilies",
        index: 3,
        params: &[],
        return_type: Type::interface("com.sun.star.container.XNameAccess"),
        one_way: false,
    }
}
