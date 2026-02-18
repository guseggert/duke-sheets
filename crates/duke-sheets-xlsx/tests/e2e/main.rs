//! End-to-end tests for duke-sheets-xlsx.
//!
//! Each test creates its own fixture on-demand by connecting to a running
//! LibreOffice instance (via Docker), building the exact spreadsheet it needs,
//! saving to a temp file, then reading it back with `XlsxReader` and asserting.
//!
//! ## Requirements
//!
//! A LibreOffice instance listening on `localhost:2002` with a shared volume:
//!
//! ```bash
//! docker run --rm -d -p 2002:2002 \
//!   -v /tmp/duke-sheets-urp:/tmp/duke-sheets-urp \
//!   duke-sheets-pyuno \
//!   bash -c 'soffice --headless --accept="socket,host=0.0.0.0,port=2002;urp;StarOffice.ComponentContext" & sleep infinity'
//! ```
//!
//! Tests skip gracefully if LibreOffice is not available.

mod common;
mod reading;

// Re-export common utilities for submodules
pub use common::*;
