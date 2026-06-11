//! Fuzz target for BIFF12 formula token parsing.
//!
//! The input is split into `(rgce, rgcb)`: the token stream and the
//! trailing extra-data stream used by PtgArray and other tokens. This
//! exercises the fast token parser directly, including the array-constant
//! count bounds and BIFF12-wide PtgName/PtgNameX indices.

#![no_main]
use libfuzzer_sys::fuzz_target;

fuzz_target!(|data: &[u8]| {
    if data.is_empty() {
        return;
    }
    let split = (data[0] as usize).min(data.len() - 1);
    let rgce = &data[1..1 + split];
    let rgcb = &data[1 + split..];
    let _ = duke_sheets_xlsb::biff12::token_parser::parse_tokens_with_extra(rgce, rgcb);
});
