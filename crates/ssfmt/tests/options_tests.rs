use ssfmt::{DateSystem, FormatOptions};

#[test]
fn test_default_options() {
    let opts = FormatOptions::default();
    assert_eq!(opts.date_system, DateSystem::Date1900);
}

#[test]
fn test_date_system_epoch() {
    assert_eq!(DateSystem::Date1900.epoch_year(), 1900);
    assert_eq!(DateSystem::Date1904.epoch_year(), 1904);
}
