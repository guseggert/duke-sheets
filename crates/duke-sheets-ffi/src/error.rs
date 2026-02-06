//! FFI error codes

use std::ffi::CStr;
use std::os::raw::{c_char, c_int};

// Success
pub const CELLS_OK: c_int = 0;

// General errors
pub const CELLS_ERR_NULL_PTR: c_int = -1;
pub const CELLS_ERR_INVALID_HANDLE: c_int = -2;
pub const CELLS_ERR_INTERNAL: c_int = -3;

// I/O errors
pub const CELLS_ERR_FILE_NOT_FOUND: c_int = -10;
pub const CELLS_ERR_PERMISSION_DENIED: c_int = -11;
pub const CELLS_ERR_IO: c_int = -12;

// Format errors
pub const CELLS_ERR_INVALID_FORMAT: c_int = -20;
pub const CELLS_ERR_CORRUPT_FILE: c_int = -21;
pub const CELLS_ERR_UNSUPPORTED: c_int = -22;

// Data errors
pub const CELLS_ERR_OUT_OF_BOUNDS: c_int = -30;
pub const CELLS_ERR_INVALID_ARGUMENT: c_int = -31;
pub const CELLS_ERR_BUFFER_TOO_SMALL: c_int = -32;

// Formula errors
pub const CELLS_ERR_FORMULA_PARSE: c_int = -40;
pub const CELLS_ERR_CIRCULAR_REF: c_int = -41;

/// Get error message for an error code
#[no_mangle]
pub extern "C" fn cells_error_message(code: c_int) -> *const c_char {
    let msg: &'static [u8] = match code {
        CELLS_OK => b"Success\0",
        CELLS_ERR_NULL_PTR => b"Null pointer argument\0",
        CELLS_ERR_INVALID_HANDLE => b"Invalid handle\0",
        CELLS_ERR_INTERNAL => b"Internal error\0",
        CELLS_ERR_FILE_NOT_FOUND => b"File not found\0",
        CELLS_ERR_PERMISSION_DENIED => b"Permission denied\0",
        CELLS_ERR_IO => b"I/O error\0",
        CELLS_ERR_INVALID_FORMAT => b"Invalid file format\0",
        CELLS_ERR_CORRUPT_FILE => b"Corrupt file\0",
        CELLS_ERR_UNSUPPORTED => b"Unsupported feature\0",
        CELLS_ERR_OUT_OF_BOUNDS => b"Index out of bounds\0",
        CELLS_ERR_INVALID_ARGUMENT => b"Invalid argument\0",
        CELLS_ERR_BUFFER_TOO_SMALL => b"Buffer too small\0",
        CELLS_ERR_FORMULA_PARSE => b"Formula parse error\0",
        CELLS_ERR_CIRCULAR_REF => b"Circular reference\0",
        _ => b"Unknown error\0",
    };

    msg.as_ptr() as *const c_char
}
