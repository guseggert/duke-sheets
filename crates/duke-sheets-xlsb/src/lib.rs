#![cfg_attr(test, allow(clippy::approx_constant))]
pub mod biff12;
pub mod error;
pub mod reader;
pub mod writer;

pub use error::{XlsbError, XlsbResult};
pub use reader::XlsbReader;
pub use writer::XlsbWriter;
