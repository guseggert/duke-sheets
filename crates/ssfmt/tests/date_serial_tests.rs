use ssfmt::date_serial::{date_to_serial, serial_to_date};
use ssfmt::DateSystem;

#[test]
fn test_serial_to_date_1900_basic() {
    // Day 1 = January 1, 1900
    let (y, m, d) = serial_to_date(1.0, DateSystem::Date1900).unwrap();
    assert_eq!((y, m, d), (1900, 1, 1));
}

#[test]
fn test_serial_to_date_1900_day_60() {
    // Day 60 = February 29, 1900 (Excel's bug - this date doesn't exist)
    let (y, m, d) = serial_to_date(60.0, DateSystem::Date1900).unwrap();
    assert_eq!((y, m, d), (1900, 2, 29));
}

#[test]
fn test_serial_to_date_1900_day_61() {
    // Day 61 = March 1, 1900
    let (y, m, d) = serial_to_date(61.0, DateSystem::Date1900).unwrap();
    assert_eq!((y, m, d), (1900, 3, 1));
}

#[test]
fn test_serial_to_date_known_date() {
    // January 9, 2026 should be serial 46031 in 1900 system
    let (y, m, d) = serial_to_date(46031.0, DateSystem::Date1900).unwrap();
    assert_eq!((y, m, d), (2026, 1, 9));
}

#[test]
fn test_serial_to_time() {
    // 0.5 = 12:00:00 (noon)
    let (h, m, s) = ssfmt::date_serial::serial_to_time(0.5);
    assert_eq!((h, m, s), (12, 0, 0));

    // 0.75 = 18:00:00 (6 PM)
    let (h, m, s) = ssfmt::date_serial::serial_to_time(0.75);
    assert_eq!((h, m, s), (18, 0, 0));
}

#[test]
fn test_date_to_serial() {
    let serial = date_to_serial(2026, 1, 9, DateSystem::Date1900);
    assert!((serial - 46031.0).abs() < 0.0001);
}
