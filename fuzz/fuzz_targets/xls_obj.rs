//! Fuzz target for BIFF8 Obj record parsing (MS-XLS 2.4.181).
//!
//! Drives `parse_obj` over an arbitrary subrecord stream (ftCmo,
//! ftCblsData, ftRboData, ftSbs, ObjLinkFmla, ftLbsData, ...) and the
//! ftLbsData structural parser directly under both host object types,
//! since the LbsDropData tail is only present for dropdowns.

#![no_main]
use duke_sheets_xls::biff::obj;
use libfuzzer_sys::fuzz_target;

fuzz_target!(|data: &[u8]| {
    let _ = obj::parse_obj(data);
    if !data.is_empty() {
        let _ = obj::LbsData::parse(&data[1..], obj::ot::LIST_BOX);
        let _ = obj::LbsData::parse(&data[1..], obj::ot::DROPDOWN);
    }
});
