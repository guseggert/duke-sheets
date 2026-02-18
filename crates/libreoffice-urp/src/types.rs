//! UNO type system: TypeClass, Type, and UnoValue.
//!
//! This module defines the Rust representations of UNO types as they appear
//! on the URP wire protocol.

use std::fmt;

/// UNO TypeClass — identifies the kind of a UNO type.
/// Values match the wire format encoding (bits 6..0 of the type byte).
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
#[repr(u8)]
pub enum TypeClass {
    Void = 0,
    Char = 1,
    Boolean = 2,
    Byte = 3,
    Short = 4,
    UnsignedShort = 5,
    Long = 6,
    UnsignedLong = 7,
    Hyper = 8,
    UnsignedHyper = 9,
    Float = 10,
    Double = 11,
    String = 12,
    Type = 13,
    Any = 14,
    Enum = 15,
    // 16 = TYPEDEF (not used on wire)
    Struct = 17,
    // 18 = UNION (not used)
    Exception = 19,
    Sequence = 20,
    // 21 = ARRAY (not used)
    Interface = 22,
}

impl TypeClass {
    /// Parse from the wire byte (lower 7 bits).
    pub fn from_byte(b: u8) -> Option<TypeClass> {
        match b & 0x7F {
            0 => Some(TypeClass::Void),
            1 => Some(TypeClass::Char),
            2 => Some(TypeClass::Boolean),
            3 => Some(TypeClass::Byte),
            4 => Some(TypeClass::Short),
            5 => Some(TypeClass::UnsignedShort),
            6 => Some(TypeClass::Long),
            7 => Some(TypeClass::UnsignedLong),
            8 => Some(TypeClass::Hyper),
            9 => Some(TypeClass::UnsignedHyper),
            10 => Some(TypeClass::Float),
            11 => Some(TypeClass::Double),
            12 => Some(TypeClass::String),
            13 => Some(TypeClass::Type),
            14 => Some(TypeClass::Any),
            15 => Some(TypeClass::Enum),
            17 => Some(TypeClass::Struct),
            19 => Some(TypeClass::Exception),
            20 => Some(TypeClass::Sequence),
            22 => Some(TypeClass::Interface),
            _ => None,
        }
    }

    /// Whether this is a "simple" type (no type name needed on wire).
    pub fn is_simple(self) -> bool {
        (self as u8) <= 14
    }
}

/// A fully-described UNO type (type class + optional type name).
#[derive(Debug, Clone, PartialEq, Eq, Hash)]
pub struct Type {
    pub class: TypeClass,
    /// For complex types (Enum, Struct, Exception, Sequence, Interface),
    /// this is the fully-qualified UNO type name.
    /// For simple types, this is empty.
    pub name: String,
}

impl Type {
    pub fn void() -> Self {
        Self {
            class: TypeClass::Void,
            name: String::new(),
        }
    }

    pub fn boolean() -> Self {
        Self {
            class: TypeClass::Boolean,
            name: String::new(),
        }
    }

    pub fn byte() -> Self {
        Self {
            class: TypeClass::Byte,
            name: String::new(),
        }
    }

    pub fn short() -> Self {
        Self {
            class: TypeClass::Short,
            name: String::new(),
        }
    }

    pub fn long() -> Self {
        Self {
            class: TypeClass::Long,
            name: String::new(),
        }
    }

    pub fn hyper() -> Self {
        Self {
            class: TypeClass::Hyper,
            name: String::new(),
        }
    }

    pub fn float() -> Self {
        Self {
            class: TypeClass::Float,
            name: String::new(),
        }
    }

    pub fn double() -> Self {
        Self {
            class: TypeClass::Double,
            name: String::new(),
        }
    }

    pub fn string() -> Self {
        Self {
            class: TypeClass::String,
            name: String::new(),
        }
    }

    pub fn any() -> Self {
        Self {
            class: TypeClass::Any,
            name: String::new(),
        }
    }

    pub fn r#type() -> Self {
        Self {
            class: TypeClass::Type,
            name: String::new(),
        }
    }

    pub fn interface(name: impl Into<String>) -> Self {
        Self {
            class: TypeClass::Interface,
            name: name.into(),
        }
    }

    pub fn r#enum(name: impl Into<String>) -> Self {
        Self {
            class: TypeClass::Enum,
            name: name.into(),
        }
    }

    pub fn r#struct(name: impl Into<String>) -> Self {
        Self {
            class: TypeClass::Struct,
            name: name.into(),
        }
    }

    pub fn exception(name: impl Into<String>) -> Self {
        Self {
            class: TypeClass::Exception,
            name: name.into(),
        }
    }

    pub fn sequence(element_type_name: &str) -> Self {
        Self {
            class: TypeClass::Sequence,
            name: format!("[]{element_type_name}"),
        }
    }

    pub fn sequence_of_bytes() -> Self {
        Self {
            class: TypeClass::Sequence,
            name: "[]byte".to_string(),
        }
    }
}

impl fmt::Display for Type {
    fn fmt(&self, f: &mut fmt::Formatter<'_>) -> fmt::Result {
        if self.name.is_empty() {
            write!(f, "{:?}", self.class)
        } else {
            write!(f, "{}", self.name)
        }
    }
}

/// A UNO value — the Rust representation of any value that can be
/// sent/received over URP.
#[derive(Debug, Clone, PartialEq)]
pub enum UnoValue {
    Void,
    Bool(bool),
    Byte(u8),
    Short(i16),
    UnsignedShort(u16),
    Long(i32),
    UnsignedLong(u32),
    Hyper(i64),
    UnsignedHyper(u64),
    Float(f32),
    Double(f64),
    Char(u16),
    String(String),
    Type(Type),
    Any(Box<Any>),
    Enum(i32),
    Struct(Vec<UnoValue>),
    Exception(UnoException),
    Sequence(Vec<UnoValue>),
    /// An interface reference, identified by its OID.
    /// Empty string = null reference.
    Interface(String),
}

impl UnoValue {
    /// Create a null interface reference.
    pub fn null_interface() -> Self {
        UnoValue::Interface(String::new())
    }

    /// Check if this is a null interface reference.
    pub fn is_null_interface(&self) -> bool {
        matches!(self, UnoValue::Interface(oid) if oid.is_empty())
    }

    pub fn as_string(&self) -> Option<&str> {
        match self {
            UnoValue::String(s) => Some(s),
            _ => None,
        }
    }

    pub fn as_double(&self) -> Option<f64> {
        match self {
            UnoValue::Double(d) => Some(*d),
            _ => None,
        }
    }

    pub fn as_long(&self) -> Option<i32> {
        match self {
            UnoValue::Long(n) => Some(*n),
            _ => None,
        }
    }

    pub fn as_interface_oid(&self) -> Option<&str> {
        match self {
            UnoValue::Interface(oid) if !oid.is_empty() => Some(oid),
            _ => None,
        }
    }

    pub fn into_any(self) -> Option<Any> {
        match self {
            UnoValue::Any(a) => Some(*a),
            _ => None,
        }
    }

    /// Extract an interface OID from an Any value (common pattern in UNO).
    pub fn extract_interface_oid(&self) -> Option<&str> {
        match self {
            UnoValue::Any(a) => a.value.as_interface_oid(),
            UnoValue::Interface(oid) if !oid.is_empty() => Some(oid),
            _ => None,
        }
    }

    /// Infer the UNO Type from the value's variant.
    ///
    /// This is used for generic struct serialization: struct fields are written
    /// with their concrete types (no type tags), so we need to know the wire type
    /// from the value alone. For `Any` values, the type descriptor is carried
    /// inside. For sequences and interfaces, the type name may be incomplete
    /// but the TypeClass is sufficient for the marshal layer.
    pub fn infer_type(&self) -> Type {
        match self {
            UnoValue::Void => Type::void(),
            UnoValue::Bool(_) => Type::boolean(),
            UnoValue::Byte(_) => Type::byte(),
            UnoValue::Short(_) => Type::short(),
            UnoValue::UnsignedShort(_) => Type {
                class: TypeClass::UnsignedShort,
                name: String::new(),
            },
            UnoValue::Long(_) => Type::long(),
            UnoValue::UnsignedLong(_) => Type {
                class: TypeClass::UnsignedLong,
                name: String::new(),
            },
            UnoValue::Hyper(_) => Type::hyper(),
            UnoValue::UnsignedHyper(_) => Type {
                class: TypeClass::UnsignedHyper,
                name: String::new(),
            },
            UnoValue::Float(_) => Type::float(),
            UnoValue::Double(_) => Type::double(),
            UnoValue::Char(_) => Type {
                class: TypeClass::Char,
                name: String::new(),
            },
            UnoValue::String(_) => Type::string(),
            UnoValue::Type(_) => Type::r#type(),
            UnoValue::Any(_) => Type::any(),
            UnoValue::Enum(_) => Type::r#enum(""),
            UnoValue::Struct(_) => Type::r#struct(""),
            UnoValue::Exception(_) => Type::exception(""),
            UnoValue::Sequence(_) => Type::sequence(""),
            UnoValue::Interface(_) => Type::interface(""),
        }
    }
}

/// A typed UNO value — wraps a value with its type descriptor.
#[derive(Debug, Clone, PartialEq)]
pub struct Any {
    pub type_desc: Type,
    pub value: UnoValue,
}

/// A UNO exception value.
#[derive(Debug, Clone, PartialEq)]
pub struct UnoException {
    /// Fully-qualified exception type name.
    pub type_name: String,
    /// The exception message (first member of com.sun.star.uno.Exception).
    pub message: String,
    // Note: in a full implementation, we'd also carry the exception's Context
    // interface reference and any derived members. For the prototype, the message
    // is sufficient.
}

// ============================================================================
// Well-known UNO type names
// ============================================================================

pub mod type_names {
    pub const X_INTERFACE: &str = "com.sun.star.uno.XInterface";
    pub const X_COMPONENT_CONTEXT: &str = "com.sun.star.uno.XComponentContext";
    pub const X_MULTI_COMPONENT_FACTORY: &str = "com.sun.star.lang.XMultiComponentFactory";
    pub const X_COMPONENT_LOADER: &str = "com.sun.star.frame.XComponentLoader";
    pub const X_COMPONENT: &str = "com.sun.star.lang.XComponent";
    pub const X_SPREADSHEET_DOCUMENT: &str = "com.sun.star.sheet.XSpreadsheetDocument";
    pub const X_SPREADSHEETS: &str = "com.sun.star.sheet.XSpreadsheets";
    pub const X_INDEX_ACCESS: &str = "com.sun.star.container.XIndexAccess";
    pub const X_SPREADSHEET: &str = "com.sun.star.sheet.XSpreadsheet";
    pub const X_CELL: &str = "com.sun.star.table.XCell";
    pub const X_TEXT: &str = "com.sun.star.text.XText";
    pub const X_SIMPLE_TEXT: &str = "com.sun.star.text.XSimpleText";
    pub const X_TEXT_RANGE: &str = "com.sun.star.text.XTextRange";
    pub const X_STORABLE: &str = "com.sun.star.frame.XStorable";
    pub const X_CLOSEABLE: &str = "com.sun.star.util.XCloseable";
    pub const X_PROPERTY_SET: &str = "com.sun.star.beans.XPropertySet";
    pub const X_PROTOCOL_PROPERTIES: &str = "com.sun.star.bridge.XProtocolProperties";

    pub const X_CELL_RANGE: &str = "com.sun.star.table.XCellRange";
    pub const X_MERGEABLE: &str = "com.sun.star.util.XMergeable";
    pub const X_SHEET_ANNOTATIONS_SUPPLIER: &str = "com.sun.star.sheet.XSheetAnnotationsSupplier";
    pub const X_SHEET_ANNOTATIONS: &str = "com.sun.star.sheet.XSheetAnnotations";
    pub const X_COLUMN_ROW_RANGE: &str = "com.sun.star.table.XColumnRowRange";
    pub const X_TABLE_ROWS: &str = "com.sun.star.table.XTableRows";
    pub const X_TABLE_COLUMNS: &str = "com.sun.star.table.XTableColumns";
    pub const X_NAMED: &str = "com.sun.star.container.XNamed";
    pub const X_NUMBER_FORMATS_SUPPLIER: &str = "com.sun.star.util.XNumberFormatsSupplier";
    pub const X_NUMBER_FORMATS: &str = "com.sun.star.util.XNumberFormats";
    pub const X_NAME_ACCESS: &str = "com.sun.star.container.XNameAccess";
    pub const X_NAME_CONTAINER: &str = "com.sun.star.container.XNameContainer";
    pub const X_MULTI_SERVICE_FACTORY: &str = "com.sun.star.lang.XMultiServiceFactory";
    pub const X_SHEET_CONDITIONAL_ENTRIES: &str = "com.sun.star.sheet.XSheetConditionalEntries";
    pub const X_STYLE_FAMILIES_SUPPLIER: &str = "com.sun.star.style.XStyleFamiliesSupplier";

    pub const PROPERTY_VALUE: &str = "com.sun.star.beans.PropertyValue";
    pub const PROTOCOL_PROPERTY: &str = "com.sun.star.bridge.ProtocolProperty";

    pub const SERVICE_DESKTOP: &str = "com.sun.star.frame.Desktop";
}
