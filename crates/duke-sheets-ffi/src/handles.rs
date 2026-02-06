//! Handle management for FFI

use duke_sheets::Workbook;
use lazy_static::lazy_static;
use std::collections::HashMap;
use std::sync::Mutex;

/// Opaque handle type
pub type Handle = u64;

/// Null handle constant
pub const HANDLE_NULL: Handle = 0;

/// Global context for managing FFI objects
pub struct FfiContext {
    workbooks: HashMap<Handle, Workbook>,
    next_handle: Handle,
}

impl FfiContext {
    fn new() -> Self {
        Self {
            workbooks: HashMap::new(),
            next_handle: 1, // Start at 1, 0 is null
        }
    }

    pub fn create_workbook(&mut self, wb: Workbook) -> Handle {
        let handle = self.next_handle;
        self.next_handle += 1;
        self.workbooks.insert(handle, wb);
        handle
    }

    pub fn get_workbook(&self, handle: Handle) -> Option<&Workbook> {
        self.workbooks.get(&handle)
    }

    pub fn get_workbook_mut(&mut self, handle: Handle) -> Option<&mut Workbook> {
        self.workbooks.get_mut(&handle)
    }

    pub fn destroy_workbook(&mut self, handle: Handle) -> bool {
        self.workbooks.remove(&handle).is_some()
    }
}

lazy_static! {
    pub static ref CONTEXT: Mutex<FfiContext> = Mutex::new(FfiContext::new());
}

/// Helper macro for FFI functions
#[macro_export]
macro_rules! with_context {
    (|$ctx:ident| $body:expr) => {
        match $crate::handles::CONTEXT.lock() {
            Ok($ctx) => $body,
            Err(_) => $crate::error::CELLS_ERR_INTERNAL,
        }
    };
    (|mut $ctx:ident| $body:expr) => {
        match $crate::handles::CONTEXT.lock() {
            Ok(mut $ctx) => $body,
            Err(_) => $crate::error::CELLS_ERR_INTERNAL,
        }
    };
}
