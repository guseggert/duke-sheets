//! Workbook FFI functions

use std::ffi::CStr;
use std::os::raw::{c_char, c_int};

use crate::error::*;
use crate::handles::{Handle, CONTEXT};
use crate::with_context;
use duke_sheets::Workbook;

/// Create a new empty workbook
#[no_mangle]
pub extern "C" fn cells_workbook_new(out_handle: *mut Handle) -> c_int {
    if out_handle.is_null() {
        return CELLS_ERR_NULL_PTR;
    }

    with_context!(|mut ctx| {
        let wb = Workbook::new();
        let handle = ctx.create_workbook(wb);
        unsafe {
            *out_handle = handle;
        }
        CELLS_OK
    })
}

/// Free a workbook
#[no_mangle]
pub extern "C" fn cells_workbook_free(handle: Handle) -> c_int {
    with_context!(|mut ctx| {
        if ctx.destroy_workbook(handle) {
            CELLS_OK
        } else {
            CELLS_ERR_INVALID_HANDLE
        }
    })
}

/// Get the number of worksheets
#[no_mangle]
pub extern "C" fn cells_workbook_sheet_count(handle: Handle, out_count: *mut c_int) -> c_int {
    if out_count.is_null() {
        return CELLS_ERR_NULL_PTR;
    }

    with_context!(|ctx| {
        match ctx.get_workbook(handle) {
            Some(wb) => {
                unsafe {
                    *out_count = wb.sheet_count() as c_int;
                }
                CELLS_OK
            }
            None => CELLS_ERR_INVALID_HANDLE,
        }
    })
}

// TODO: Implement more workbook functions
// - cells_workbook_open
// - cells_workbook_save
// - cells_workbook_add_sheet
// etc.
