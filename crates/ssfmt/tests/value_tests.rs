use ssfmt::Value;

#[test]
fn test_value_from_f64() {
    let v: Value = 42.5.into();
    assert!(matches!(v, Value::Number(n) if (n - 42.5).abs() < f64::EPSILON));
}

#[test]
fn test_value_from_i64() {
    let v: Value = 42i64.into();
    assert!(matches!(v, Value::Number(n) if (n - 42.0).abs() < f64::EPSILON));
}

#[test]
fn test_value_from_str() {
    let v: Value = "hello".into();
    assert!(matches!(v, Value::Text(s) if s == "hello"));
}

#[test]
fn test_value_from_bool() {
    let v: Value = true.into();
    assert!(matches!(v, Value::Bool(true)));
}
