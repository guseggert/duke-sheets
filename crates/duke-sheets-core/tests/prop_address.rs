//! Property tests: address and range text forms must round-trip
//! through parse ↔ Display for the entire coordinate space, including
//! absolute-reference flags.

use duke_sheets_core::{CellAddress, CellRange, MAX_COLS, MAX_ROWS};
use proptest::prelude::*;

prop_compose! {
    fn arb_address()(
        row in 0..MAX_ROWS,
        col in 0..MAX_COLS,
        row_absolute in any::<bool>(),
        col_absolute in any::<bool>(),
    ) -> CellAddress {
        CellAddress::with_absolute(row, col, row_absolute, col_absolute)
    }
}

proptest! {
    #[test]
    fn cell_address_roundtrips_through_text(addr in arb_address()) {
        let text = addr.to_string();
        let parsed = CellAddress::parse(&text).expect("printed address must parse");
        prop_assert_eq!(parsed, addr, "text form: {}", text);
    }

    #[test]
    fn cell_range_roundtrips_through_text(a in arb_address(), b in arb_address()) {
        let range = CellRange::new(a, b);
        let text = range.to_string();
        let parsed = CellRange::parse(&text).expect("printed range must parse");
        prop_assert_eq!(parsed.to_string(), text);
    }

    #[test]
    fn parse_never_panics_on_arbitrary_short_strings(s in ".{0,12}") {
        let _ = CellAddress::parse(&s);
        let _ = CellRange::parse(&s);
    }
}
