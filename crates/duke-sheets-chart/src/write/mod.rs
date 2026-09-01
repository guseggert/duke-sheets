//! Chart XML part writers (`c:chartSpace` / `cx:chartSpace`), shared
//! by the XLSX and XLSB writers (both formats store chart parts as
//! XML).

mod chart;
mod chart_ex;
mod chart_style;

pub use chart::chart_part_bytes;
pub use chart_ex::chart_ex_part_bytes;
pub use chart_style::{
    chart_color_style_bytes, chart_color_style_part_bytes, chart_style_bytes,
    chart_style_part_bytes,
};

use std::io::Cursor;

use quick_xml::Writer;

pub(crate) type XmlWriter = Writer<Cursor<Vec<u8>>>;
pub(crate) const NS_DOC_RELS: &str =
    "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
