use crate::error::FormulaResult;
use crate::evaluator::{EvaluationContext, FormulaValue};
use duke_sheets_core::CellError;

const SQRT_2PI: f64 = 2.5066282746310002;
const EPS: f64 = 1e-12;

fn scalar_number(value: &FormulaValue) -> Result<f64, CellError> {
    match value {
        FormulaValue::Number(n) => Ok(*n),
        FormulaValue::Boolean(true) => Ok(1.0),
        FormulaValue::Boolean(false) => Ok(0.0),
        FormulaValue::String(s) => s.parse::<f64>().map_err(|_| CellError::Value),
        FormulaValue::Empty => Ok(0.0),
        FormulaValue::Error(e) => Err(*e),
        FormulaValue::Array(arr) => {
            if let Some(row) = arr.first() {
                if let Some(cell) = row.first() {
                    return scalar_number(cell);
                }
            }
            Err(CellError::Value)
        }
    }
}

fn required_number(args: &[FormulaValue], idx: usize) -> Result<f64, CellError> {
    let arg = args.get(idx).ok_or(CellError::Value)?;
    scalar_number(arg)
}

fn optional_number(args: &[FormulaValue], idx: usize, default: f64) -> Result<f64, CellError> {
    match args.get(idx) {
        None => Ok(default),
        Some(FormulaValue::Empty) => Ok(default),
        Some(v) => scalar_number(v),
    }
}

fn required_bool(args: &[FormulaValue], idx: usize) -> Result<bool, CellError> {
    let arg = args.get(idx).ok_or(CellError::Value)?;
    match arg {
        FormulaValue::Boolean(b) => Ok(*b),
        FormulaValue::Number(n) => Ok(*n != 0.0),
        FormulaValue::Empty => Ok(false),
        FormulaValue::Error(e) => Err(*e),
        FormulaValue::String(s) => {
            let upper = s.to_uppercase();
            if upper == "TRUE" {
                Ok(true)
            } else if upper == "FALSE" {
                Ok(false)
            } else {
                Err(CellError::Value)
            }
        }
        FormulaValue::Array(_) => Err(CellError::Value),
    }
}

fn collect_numbers(value: &FormulaValue, numbers: &mut Vec<f64>) -> Option<CellError> {
    match value {
        FormulaValue::Number(n) => numbers.push(*n),
        FormulaValue::Error(e) => return Some(*e),
        FormulaValue::Array(arr) => {
            for row in arr {
                for cell in row {
                    match cell {
                        FormulaValue::Number(n) => numbers.push(*n),
                        FormulaValue::Error(e) => return Some(*e),
                        _ => {}
                    }
                }
            }
        }
        _ => {}
    }
    None
}

fn collect_number_pairs(a: &FormulaValue, b: &FormulaValue) -> Result<Vec<(f64, f64)>, CellError> {
    match (a, b) {
        (FormulaValue::Array(arr1), FormulaValue::Array(arr2)) => {
            if arr1.len() != arr2.len() {
                return Err(CellError::Value);
            }
            let mut out = Vec::new();
            for (r1, r2) in arr1.iter().zip(arr2.iter()) {
                if r1.len() != r2.len() {
                    return Err(CellError::Value);
                }
                for (c1, c2) in r1.iter().zip(r2.iter()) {
                    match (c1, c2) {
                        (FormulaValue::Error(e), _) | (_, FormulaValue::Error(e)) => {
                            return Err(*e)
                        }
                        (FormulaValue::Number(x), FormulaValue::Number(y)) => out.push((*x, *y)),
                        _ => {}
                    }
                }
            }
            Ok(out)
        }
        _ => Ok(vec![(scalar_number(a)?, scalar_number(b)?)]),
    }
}

fn ln_gamma(z: f64) -> f64 {
    let coeffs = [
        0.9999999999998099,
        676.5203681218851,
        -1259.1392167224028,
        771.3234287776531,
        -176.6150291621406,
        12.507343278686905,
        -0.13857109526572012,
        0.000009984369578019572,
        0.00000015056327351493116,
    ];

    if z < 0.5 {
        return std::f64::consts::PI.ln()
            - (std::f64::consts::PI * z).sin().ln()
            - ln_gamma(1.0 - z);
    }

    let z1 = z - 1.0;
    let mut x = coeffs[0];
    for (i, c) in coeffs.iter().enumerate().skip(1) {
        x += c / (z1 + i as f64);
    }
    let t = z1 + 7.5;
    0.5 * (2.0 * std::f64::consts::PI).ln() + (z1 + 0.5) * t.ln() - t + x.ln()
}

fn beta_continued_fraction(x: f64, a: f64, b: f64) -> f64 {
    let max_iter = 200;
    let tiny = 1e-30;
    let qab = a + b;
    let qap = a + 1.0;
    let qam = a - 1.0;

    let mut c = 1.0;
    let mut d = 1.0 - qab * x / qap;
    if d.abs() < tiny {
        d = tiny;
    }
    d = 1.0 / d;
    let mut h = d;

    for m in 1..=max_iter {
        let m2 = 2.0 * m as f64;
        let mut aa = m as f64 * (b - m as f64) * x / ((qam + m2) * (a + m2));
        d = 1.0 + aa * d;
        if d.abs() < tiny {
            d = tiny;
        }
        c = 1.0 + aa / c;
        if c.abs() < tiny {
            c = tiny;
        }
        d = 1.0 / d;
        h *= d * c;

        aa = -(a + m as f64) * (qab + m as f64) * x / ((a + m2) * (qap + m2));
        d = 1.0 + aa * d;
        if d.abs() < tiny {
            d = tiny;
        }
        c = 1.0 + aa / c;
        if c.abs() < tiny {
            c = tiny;
        }
        d = 1.0 / d;
        let del = d * c;
        h *= del;
        if (del - 1.0).abs() < 1e-14 {
            break;
        }
    }
    h
}

fn regularized_beta(x: f64, a: f64, b: f64) -> f64 {
    if x <= 0.0 {
        return 0.0;
    }
    if x >= 1.0 {
        return 1.0;
    }

    let ln_bt = ln_gamma(a + b) - ln_gamma(a) - ln_gamma(b) + a * x.ln() + b * (1.0 - x).ln();
    let bt = ln_bt.exp();
    let pivot = (a + 1.0) / (a + b + 2.0);

    if x < pivot {
        bt * beta_continued_fraction(x, a, b) / a
    } else {
        1.0 - bt * beta_continued_fraction(1.0 - x, b, a) / b
    }
}

fn regularized_gamma_p(a: f64, x: f64) -> f64 {
    if x <= 0.0 {
        return 0.0;
    }
    if x < a + 1.0 {
        let mut sum = 1.0 / a;
        let mut term = sum;
        let mut n = 1.0;
        while term.abs() > 1e-14 * sum.abs() && n < 10_000.0 {
            term *= x / (a + n);
            sum += term;
            n += 1.0;
        }
        (-x + a * x.ln() - ln_gamma(a)).exp() * sum
    } else {
        let mut b = x + 1.0 - a;
        let mut c = 1.0 / 1e-30;
        let mut d = 1.0 / b;
        let mut h = d;
        for i in 1..=200 {
            let an = -(i as f64) * ((i as f64) - a);
            b += 2.0;
            d = an * d + b;
            if d.abs() < 1e-30 {
                d = 1e-30;
            }
            c = b + an / c;
            if c.abs() < 1e-30 {
                c = 1e-30;
            }
            d = 1.0 / d;
            let del = d * c;
            h *= del;
            if (del - 1.0).abs() < 1e-14 {
                break;
            }
        }
        1.0 - (-x + a * x.ln() - ln_gamma(a)).exp() * h
    }
}

fn normal_cdf(x: f64) -> f64 {
    if x == 0.0 {
        return 0.5;
    }
    let p = regularized_gamma_p(0.5, x * x / 2.0);
    if x > 0.0 {
        0.5 * (1.0 + p)
    } else {
        0.5 * (1.0 - p)
    }
}

fn normal_pdf(x: f64) -> f64 {
    (-0.5 * x * x).exp() / SQRT_2PI
}

fn inverse_normal(p: f64) -> f64 {
    let a = [
        -39.69683028665376,
        220.9460984245205,
        -275.9285104469687,
        138.357751867269,
        -30.66479806614716,
        2.506628277459239,
    ];
    let b = [
        -54.47609879822406,
        161.5858368580409,
        -155.6989798598866,
        66.80131188771972,
        -13.28068155288572,
    ];
    let c = [
        -0.007784894002430293,
        -0.3223964580411365,
        -2.400758277161838,
        -2.549732539343734,
        4.374664141464968,
        2.938163982698783,
    ];
    let d = [
        0.007784695709041462,
        0.3224671290700398,
        2.445134137142996,
        3.754408661907416,
    ];

    let plow = 0.02425;
    let phigh = 1.0 - plow;

    if p < plow {
        let q = (-2.0 * p.ln()).sqrt();
        return (((((c[0] * q + c[1]) * q + c[2]) * q + c[3]) * q + c[4]) * q + c[5])
            / ((((d[0] * q + d[1]) * q + d[2]) * q + d[3]) * q + 1.0);
    }
    if p > phigh {
        let q = (-2.0 * (1.0 - p).ln()).sqrt();
        return -(((((c[0] * q + c[1]) * q + c[2]) * q + c[3]) * q + c[4]) * q + c[5])
            / ((((d[0] * q + d[1]) * q + d[2]) * q + d[3]) * q + 1.0);
    }

    let q = p - 0.5;
    let r = q * q;
    (((((a[0] * r + a[1]) * r + a[2]) * r + a[3]) * r + a[4]) * r + a[5]) * q
        / (((((b[0] * r + b[1]) * r + b[2]) * r + b[3]) * r + b[4]) * r + 1.0)
}

fn t_cdf(x: f64, df: f64) -> f64 {
    if x == 0.0 {
        return 0.5;
    }
    let t = df / (df + x * x);
    let ib = regularized_beta(t, df / 2.0, 0.5);
    if x > 0.0 {
        1.0 - 0.5 * ib
    } else {
        0.5 * ib
    }
}

fn t_pdf(x: f64, df: f64) -> f64 {
    ((ln_gamma((df + 1.0) / 2.0) - ln_gamma(df / 2.0) - 0.5 * (df * std::f64::consts::PI).ln())
        .exp())
        * (1.0 + x * x / df).powf(-(df + 1.0) / 2.0)
}

fn chi2_cdf(x: f64, df: f64) -> f64 {
    if x <= 0.0 {
        return 0.0;
    }
    regularized_gamma_p(df / 2.0, x / 2.0)
}

fn f_cdf(x: f64, d1: f64, d2: f64) -> f64 {
    if x <= 0.0 {
        return 0.0;
    }
    let y = (d1 * x) / (d1 * x + d2);
    regularized_beta(y, d1 / 2.0, d2 / 2.0)
}

fn beta_pdf(x: f64, a: f64, b: f64) -> f64 {
    if x <= 0.0 || x >= 1.0 {
        return 0.0;
    }
    ((a - 1.0) * x.ln() + (b - 1.0) * (1.0 - x).ln()
        - (ln_gamma(a) + ln_gamma(b) - ln_gamma(a + b)))
    .exp()
}

fn gamma_pdf(x: f64, a: f64) -> f64 {
    if x <= 0.0 {
        return 0.0;
    }
    ((a - 1.0) * x.ln() - x - ln_gamma(a)).exp()
}

fn f_pdf(x: f64, d1: f64, d2: f64) -> f64 {
    if x <= 0.0 {
        return 0.0;
    }
    let a = d1 / 2.0;
    let b = d2 / 2.0;
    let ln_beta = ln_gamma(a) + ln_gamma(b) - ln_gamma(a + b);
    let ln_num = a * (d1 / d2).ln() + (a - 1.0) * x.ln();
    let ln_den = ln_beta + (a + b) * (1.0 + d1 * x / d2).ln();
    (ln_num - ln_den).exp()
}

fn inverse_beta(p: f64, a: f64, b: f64) -> f64 {
    if p <= 0.0 {
        return 0.0;
    }
    if p >= 1.0 {
        return 1.0;
    }

    let mut lo = EPS;
    let mut hi = 1.0 - EPS;
    let mut x = p.clamp(lo, hi);
    for _ in 0..80 {
        let fx = regularized_beta(x, a, b) - p;
        if fx.abs() < 1e-12 {
            break;
        }
        let pdf = beta_pdf(x, a, b);
        let mut next = if pdf > 1e-18 {
            x - fx / pdf
        } else {
            (lo + hi) * 0.5
        };
        if fx > 0.0 {
            hi = x;
        } else {
            lo = x;
        }
        if !(next > lo && next < hi) {
            next = 0.5 * (lo + hi);
        }
        x = next;
    }
    x
}

fn inverse_gamma(p: f64, a: f64) -> f64 {
    if p <= 0.0 {
        return 0.0;
    }
    if p >= 1.0 {
        return f64::INFINITY;
    }

    let mut lo = 0.0;
    let mut hi = (a + 10.0).max(1.0);
    while regularized_gamma_p(a, hi) < p {
        hi *= 2.0;
        if hi > 1e12 {
            break;
        }
    }

    let z = inverse_normal(p);
    let mut x = (a * (1.0 - 1.0 / (9.0 * a) + z / (3.0 * a.sqrt())).powi(3)).clamp(EPS, hi);
    for _ in 0..100 {
        let fx = regularized_gamma_p(a, x) - p;
        if fx.abs() < 1e-12 {
            break;
        }
        let pdf = gamma_pdf(x, a);
        let mut next = if pdf > 1e-18 {
            x - fx / pdf
        } else {
            0.5 * (lo + hi)
        };
        if fx > 0.0 {
            hi = x;
        } else {
            lo = x;
        }
        if !(next > lo && next < hi) {
            next = 0.5 * (lo + hi);
        }
        x = next;
    }
    x
}

fn inverse_chi2(p: f64, df: f64) -> f64 {
    2.0 * inverse_gamma(p, df / 2.0)
}

fn inverse_t(p: f64, df: f64) -> f64 {
    if p <= 0.0 {
        return f64::NEG_INFINITY;
    }
    if p >= 1.0 {
        return f64::INFINITY;
    }
    if p == 0.5 {
        return 0.0;
    }
    if p < 0.5 {
        return -inverse_t(1.0 - p, df);
    }

    let mut lo = 0.0;
    let mut hi = 1.0;
    while t_cdf(hi, df) < p {
        hi *= 2.0;
        if hi > 1e6 {
            break;
        }
    }
    let mut x = inverse_normal(p).max(EPS);
    if x > hi {
        x = 0.5 * (lo + hi);
    }

    for _ in 0..80 {
        let fx = t_cdf(x, df) - p;
        if fx.abs() < 1e-12 {
            break;
        }
        let pdf = t_pdf(x, df);
        let mut next = if pdf > 1e-18 {
            x - fx / pdf
        } else {
            0.5 * (lo + hi)
        };
        if fx > 0.0 {
            hi = x;
        } else {
            lo = x;
        }
        if !(next > lo && next < hi) {
            next = 0.5 * (lo + hi);
        }
        x = next;
    }
    x
}

fn inverse_f(p: f64, d1: f64, d2: f64) -> f64 {
    if p <= 0.0 {
        return 0.0;
    }
    if p >= 1.0 {
        return f64::INFINITY;
    }

    let mut lo = 0.0;
    let mut hi = 1.0;
    while f_cdf(hi, d1, d2) < p {
        hi *= 2.0;
        if hi > 1e12 {
            break;
        }
    }
    let mut x = 0.5 * (lo + hi);
    for _ in 0..100 {
        let fx = f_cdf(x, d1, d2) - p;
        if fx.abs() < 1e-12 {
            break;
        }
        let pdf = f_pdf(x, d1, d2);
        let mut next = if pdf > 1e-18 {
            x - fx / pdf
        } else {
            0.5 * (lo + hi)
        };
        if fx > 0.0 {
            hi = x;
        } else {
            lo = x;
        }
        if !(next > lo && next < hi) {
            next = 0.5 * (lo + hi);
        }
        x = next;
    }
    x
}

fn int_floor(n: f64) -> i64 {
    n.floor() as i64
}

fn ln_choose(n: i64, k: i64) -> f64 {
    ln_gamma((n + 1) as f64) - ln_gamma((k + 1) as f64) - ln_gamma((n - k + 1) as f64)
}

pub fn fn_betadist(args: &[FormulaValue], _ctx: &EvaluationContext) -> FormulaResult<FormulaValue> {
    let x = match required_number(args, 0) {
        Ok(v) => v,
        Err(e) => return Ok(FormulaValue::Error(e)),
    };
    let alpha = match required_number(args, 1) {
        Ok(v) => v,
        Err(e) => return Ok(FormulaValue::Error(e)),
    };
    let beta = match required_number(args, 2) {
        Ok(v) => v,
        Err(e) => return Ok(FormulaValue::Error(e)),
    };
    let a = match optional_number(args, 3, 0.0) {
        Ok(v) => v,
        Err(e) => return Ok(FormulaValue::Error(e)),
    };
    let b = match optional_number(args, 4, 1.0) {
        Ok(v) => v,
        Err(e) => return Ok(FormulaValue::Error(e)),
    };
    if alpha <= 0.0 || beta <= 0.0 || b <= a || x < a || x > b {
        return Ok(FormulaValue::Error(CellError::Num));
    }
    let y = (x - a) / (b - a);
    Ok(FormulaValue::Number(regularized_beta(y, alpha, beta)))
}

pub fn fn_betainv(args: &[FormulaValue], _ctx: &EvaluationContext) -> FormulaResult<FormulaValue> {
    let p = match required_number(args, 0) {
        Ok(v) => v,
        Err(e) => return Ok(FormulaValue::Error(e)),
    };
    let alpha = match required_number(args, 1) {
        Ok(v) => v,
        Err(e) => return Ok(FormulaValue::Error(e)),
    };
    let beta = match required_number(args, 2) {
        Ok(v) => v,
        Err(e) => return Ok(FormulaValue::Error(e)),
    };
    let a = match optional_number(args, 3, 0.0) {
        Ok(v) => v,
        Err(e) => return Ok(FormulaValue::Error(e)),
    };
    let b = match optional_number(args, 4, 1.0) {
        Ok(v) => v,
        Err(e) => return Ok(FormulaValue::Error(e)),
    };
    if alpha <= 0.0 || beta <= 0.0 || b <= a || !(0.0..=1.0).contains(&p) {
        return Ok(FormulaValue::Error(CellError::Num));
    }
    let y = inverse_beta(p, alpha, beta);
    Ok(FormulaValue::Number(a + (b - a) * y))
}

pub fn fn_binomdist(
    args: &[FormulaValue],
    _ctx: &EvaluationContext,
) -> FormulaResult<FormulaValue> {
    let k = int_floor(match required_number(args, 0) {
        Ok(v) => v,
        Err(e) => return Ok(FormulaValue::Error(e)),
    });
    let n = int_floor(match required_number(args, 1) {
        Ok(v) => v,
        Err(e) => return Ok(FormulaValue::Error(e)),
    });
    let p = match required_number(args, 2) {
        Ok(v) => v,
        Err(e) => return Ok(FormulaValue::Error(e)),
    };
    let cumulative = match required_bool(args, 3) {
        Ok(v) => v,
        Err(e) => return Ok(FormulaValue::Error(e)),
    };
    if n < 0 || k < 0 || k > n || !(0.0..=1.0).contains(&p) {
        return Ok(FormulaValue::Error(CellError::Num));
    }
    let pmf =
        |x: i64| (ln_choose(n, x) + (x as f64) * p.ln() + ((n - x) as f64) * (1.0 - p).ln()).exp();
    let result = if cumulative {
        (0..=k).map(pmf).sum()
    } else {
        pmf(k)
    };
    Ok(FormulaValue::Number(result))
}

pub fn fn_chidist(args: &[FormulaValue], _ctx: &EvaluationContext) -> FormulaResult<FormulaValue> {
    let x = match required_number(args, 0) {
        Ok(v) => v,
        Err(e) => return Ok(FormulaValue::Error(e)),
    };
    let df = match required_number(args, 1) {
        Ok(v) => v,
        Err(e) => return Ok(FormulaValue::Error(e)),
    };
    if x < 0.0 || df <= 0.0 {
        return Ok(FormulaValue::Error(CellError::Num));
    }
    Ok(FormulaValue::Number(1.0 - chi2_cdf(x, df)))
}

pub fn fn_chiinv(args: &[FormulaValue], _ctx: &EvaluationContext) -> FormulaResult<FormulaValue> {
    let prob = match required_number(args, 0) {
        Ok(v) => v,
        Err(e) => return Ok(FormulaValue::Error(e)),
    };
    let df = match required_number(args, 1) {
        Ok(v) => v,
        Err(e) => return Ok(FormulaValue::Error(e)),
    };
    if prob <= 0.0 || prob > 1.0 || df <= 0.0 {
        return Ok(FormulaValue::Error(CellError::Num));
    }
    Ok(FormulaValue::Number(inverse_chi2(1.0 - prob, df)))
}

pub fn fn_chitest(args: &[FormulaValue], _ctx: &EvaluationContext) -> FormulaResult<FormulaValue> {
    let (actual, expected) = match (args.get(0), args.get(1)) {
        (Some(a), Some(b)) => (a, b),
        _ => return Ok(FormulaValue::Error(CellError::Value)),
    };
    let (arr_a, arr_e) = match (actual, expected) {
        (FormulaValue::Array(a), FormulaValue::Array(e)) => (a, e),
        _ => return Ok(FormulaValue::Error(CellError::Value)),
    };
    if arr_a.len() != arr_e.len() || arr_a.is_empty() {
        return Ok(FormulaValue::Error(CellError::Value));
    }
    let rows = arr_a.len();
    let cols = arr_a[0].len();
    if cols == 0 {
        return Ok(FormulaValue::Error(CellError::Value));
    }
    let mut chi = 0.0;
    for (ra, re) in arr_a.iter().zip(arr_e.iter()) {
        if ra.len() != cols || re.len() != cols {
            return Ok(FormulaValue::Error(CellError::Value));
        }
        for (a, e) in ra.iter().zip(re.iter()) {
            let oa = match scalar_number(a) {
                Ok(v) => v,
                Err(err) => return Ok(FormulaValue::Error(err)),
            };
            let oe = match scalar_number(e) {
                Ok(v) => v,
                Err(err) => return Ok(FormulaValue::Error(err)),
            };
            if oe <= 0.0 {
                return Ok(FormulaValue::Error(CellError::Num));
            }
            let d = oa - oe;
            chi += d * d / oe;
        }
    }
    let df = ((rows - 1) * (cols - 1)) as f64;
    if df <= 0.0 {
        return Ok(FormulaValue::Error(CellError::Num));
    }
    Ok(FormulaValue::Number(1.0 - chi2_cdf(chi, df)))
}

pub fn fn_confidence(
    args: &[FormulaValue],
    _ctx: &EvaluationContext,
) -> FormulaResult<FormulaValue> {
    let alpha = match required_number(args, 0) {
        Ok(v) => v,
        Err(e) => return Ok(FormulaValue::Error(e)),
    };
    let stdev = match required_number(args, 1) {
        Ok(v) => v,
        Err(e) => return Ok(FormulaValue::Error(e)),
    };
    let size = match required_number(args, 2) {
        Ok(v) => v,
        Err(e) => return Ok(FormulaValue::Error(e)),
    };
    if alpha <= 0.0 || alpha >= 1.0 || stdev <= 0.0 || size < 1.0 {
        return Ok(FormulaValue::Error(CellError::Num));
    }
    let z = inverse_normal(1.0 - alpha / 2.0);
    Ok(FormulaValue::Number(z * stdev / size.sqrt()))
}

pub fn fn_covar(args: &[FormulaValue], _ctx: &EvaluationContext) -> FormulaResult<FormulaValue> {
    let (a, b) = match (args.get(0), args.get(1)) {
        (Some(a), Some(b)) => (a, b),
        _ => return Ok(FormulaValue::Error(CellError::Value)),
    };
    let pairs = match collect_number_pairs(a, b) {
        Ok(v) => v,
        Err(e) => return Ok(FormulaValue::Error(e)),
    };
    if pairs.is_empty() {
        return Ok(FormulaValue::Error(CellError::Div0));
    }
    let n = pairs.len() as f64;
    let mean_x = pairs.iter().map(|(x, _)| x).sum::<f64>() / n;
    let mean_y = pairs.iter().map(|(_, y)| y).sum::<f64>() / n;
    let cov = pairs
        .iter()
        .map(|(x, y)| (x - mean_x) * (y - mean_y))
        .sum::<f64>()
        / n;
    Ok(FormulaValue::Number(cov))
}

pub fn fn_critbinom(
    args: &[FormulaValue],
    _ctx: &EvaluationContext,
) -> FormulaResult<FormulaValue> {
    let n = int_floor(match required_number(args, 0) {
        Ok(v) => v,
        Err(e) => return Ok(FormulaValue::Error(e)),
    });
    let p = match required_number(args, 1) {
        Ok(v) => v,
        Err(e) => return Ok(FormulaValue::Error(e)),
    };
    let alpha = match required_number(args, 2) {
        Ok(v) => v,
        Err(e) => return Ok(FormulaValue::Error(e)),
    };
    if n < 0 || !(0.0..=1.0).contains(&p) || !(0.0..=1.0).contains(&alpha) {
        return Ok(FormulaValue::Error(CellError::Num));
    }
    let pmf =
        |x: i64| (ln_choose(n, x) + (x as f64) * p.ln() + ((n - x) as f64) * (1.0 - p).ln()).exp();
    let mut cdf = 0.0;
    for x in 0..=n {
        cdf += pmf(x);
        if cdf >= alpha {
            return Ok(FormulaValue::Number(x as f64));
        }
    }
    Ok(FormulaValue::Number(n as f64))
}

pub fn fn_expondist(
    args: &[FormulaValue],
    _ctx: &EvaluationContext,
) -> FormulaResult<FormulaValue> {
    let x = match required_number(args, 0) {
        Ok(v) => v,
        Err(e) => return Ok(FormulaValue::Error(e)),
    };
    let lambda = match required_number(args, 1) {
        Ok(v) => v,
        Err(e) => return Ok(FormulaValue::Error(e)),
    };
    let cumulative = match required_bool(args, 2) {
        Ok(v) => v,
        Err(e) => return Ok(FormulaValue::Error(e)),
    };
    if x < 0.0 || lambda <= 0.0 {
        return Ok(FormulaValue::Error(CellError::Num));
    }
    let result = if cumulative {
        1.0 - (-lambda * x).exp()
    } else {
        lambda * (-lambda * x).exp()
    };
    Ok(FormulaValue::Number(result))
}

pub fn fn_fdist(args: &[FormulaValue], _ctx: &EvaluationContext) -> FormulaResult<FormulaValue> {
    let x = match required_number(args, 0) {
        Ok(v) => v,
        Err(e) => return Ok(FormulaValue::Error(e)),
    };
    let d1 = match required_number(args, 1) {
        Ok(v) => v,
        Err(e) => return Ok(FormulaValue::Error(e)),
    };
    let d2 = match required_number(args, 2) {
        Ok(v) => v,
        Err(e) => return Ok(FormulaValue::Error(e)),
    };
    if x < 0.0 || d1 <= 0.0 || d2 <= 0.0 {
        return Ok(FormulaValue::Error(CellError::Num));
    }
    Ok(FormulaValue::Number(1.0 - f_cdf(x, d1, d2)))
}

pub fn fn_finv(args: &[FormulaValue], _ctx: &EvaluationContext) -> FormulaResult<FormulaValue> {
    let p = match required_number(args, 0) {
        Ok(v) => v,
        Err(e) => return Ok(FormulaValue::Error(e)),
    };
    let d1 = match required_number(args, 1) {
        Ok(v) => v,
        Err(e) => return Ok(FormulaValue::Error(e)),
    };
    let d2 = match required_number(args, 2) {
        Ok(v) => v,
        Err(e) => return Ok(FormulaValue::Error(e)),
    };
    if p <= 0.0 || p >= 1.0 || d1 <= 0.0 || d2 <= 0.0 {
        return Ok(FormulaValue::Error(CellError::Num));
    }
    Ok(FormulaValue::Number(inverse_f(1.0 - p, d1, d2)))
}

pub fn fn_ftest(args: &[FormulaValue], _ctx: &EvaluationContext) -> FormulaResult<FormulaValue> {
    let mut a = Vec::new();
    let mut b = Vec::new();
    if let Some(err) = collect_numbers(args.get(0).unwrap_or(&FormulaValue::Empty), &mut a) {
        return Ok(FormulaValue::Error(err));
    }
    if let Some(err) = collect_numbers(args.get(1).unwrap_or(&FormulaValue::Empty), &mut b) {
        return Ok(FormulaValue::Error(err));
    }
    if a.len() < 2 || b.len() < 2 {
        return Ok(FormulaValue::Error(CellError::Div0));
    }

    let mean = |v: &[f64]| v.iter().sum::<f64>() / v.len() as f64;
    let var_s = |v: &[f64]| {
        let m = mean(v);
        v.iter().map(|x| (x - m) * (x - m)).sum::<f64>() / (v.len() as f64 - 1.0)
    };

    let v1 = var_s(&a);
    let v2 = var_s(&b);
    if v1 == 0.0 || v2 == 0.0 {
        return Ok(FormulaValue::Error(CellError::Div0));
    }
    let f = v1 / v2;
    let d1 = (a.len() - 1) as f64;
    let d2 = (b.len() - 1) as f64;
    let cdf = f_cdf(f, d1, d2);
    let p = 2.0 * cdf.min(1.0 - cdf);
    Ok(FormulaValue::Number(p.clamp(0.0, 1.0)))
}

pub fn fn_gammadist(
    args: &[FormulaValue],
    _ctx: &EvaluationContext,
) -> FormulaResult<FormulaValue> {
    let x = match required_number(args, 0) {
        Ok(v) => v,
        Err(e) => return Ok(FormulaValue::Error(e)),
    };
    let alpha = match required_number(args, 1) {
        Ok(v) => v,
        Err(e) => return Ok(FormulaValue::Error(e)),
    };
    let beta = match required_number(args, 2) {
        Ok(v) => v,
        Err(e) => return Ok(FormulaValue::Error(e)),
    };
    let cumulative = match required_bool(args, 3) {
        Ok(v) => v,
        Err(e) => return Ok(FormulaValue::Error(e)),
    };
    if x < 0.0 || alpha <= 0.0 || beta <= 0.0 {
        return Ok(FormulaValue::Error(CellError::Num));
    }
    let result = if cumulative {
        regularized_gamma_p(alpha, x / beta)
    } else if x == 0.0 {
        0.0
    } else {
        ((alpha - 1.0) * x.ln() - x / beta - ln_gamma(alpha) - alpha * beta.ln()).exp()
    };
    Ok(FormulaValue::Number(result))
}

pub fn fn_gammainv(args: &[FormulaValue], _ctx: &EvaluationContext) -> FormulaResult<FormulaValue> {
    let p = match required_number(args, 0) {
        Ok(v) => v,
        Err(e) => return Ok(FormulaValue::Error(e)),
    };
    let alpha = match required_number(args, 1) {
        Ok(v) => v,
        Err(e) => return Ok(FormulaValue::Error(e)),
    };
    let beta = match required_number(args, 2) {
        Ok(v) => v,
        Err(e) => return Ok(FormulaValue::Error(e)),
    };
    if p <= 0.0 || p >= 1.0 || alpha <= 0.0 || beta <= 0.0 {
        return Ok(FormulaValue::Error(CellError::Num));
    }
    Ok(FormulaValue::Number(inverse_gamma(p, alpha) * beta))
}

pub fn fn_hypgeomdist(
    args: &[FormulaValue],
    _ctx: &EvaluationContext,
) -> FormulaResult<FormulaValue> {
    let x = int_floor(match required_number(args, 0) {
        Ok(v) => v,
        Err(e) => return Ok(FormulaValue::Error(e)),
    });
    let n = int_floor(match required_number(args, 1) {
        Ok(v) => v,
        Err(e) => return Ok(FormulaValue::Error(e)),
    });
    let k = int_floor(match required_number(args, 2) {
        Ok(v) => v,
        Err(e) => return Ok(FormulaValue::Error(e)),
    });
    let n_pop = int_floor(match required_number(args, 3) {
        Ok(v) => v,
        Err(e) => return Ok(FormulaValue::Error(e)),
    });
    if n_pop <= 0 || k < 0 || n < 0 || x < 0 || k > n_pop || n > n_pop || x > k || x > n {
        return Ok(FormulaValue::Error(CellError::Num));
    }
    let low = (n + k - n_pop).max(0);
    let high = n.min(k);
    if x < low || x > high {
        return Ok(FormulaValue::Number(0.0));
    }
    let ln_p = ln_choose(k, x) + ln_choose(n_pop - k, n - x) - ln_choose(n_pop, n);
    Ok(FormulaValue::Number(ln_p.exp()))
}

pub fn fn_loginv(args: &[FormulaValue], _ctx: &EvaluationContext) -> FormulaResult<FormulaValue> {
    let p = match required_number(args, 0) {
        Ok(v) => v,
        Err(e) => return Ok(FormulaValue::Error(e)),
    };
    let mean = match required_number(args, 1) {
        Ok(v) => v,
        Err(e) => return Ok(FormulaValue::Error(e)),
    };
    let stdev = match required_number(args, 2) {
        Ok(v) => v,
        Err(e) => return Ok(FormulaValue::Error(e)),
    };
    if p <= 0.0 || p >= 1.0 || stdev <= 0.0 {
        return Ok(FormulaValue::Error(CellError::Num));
    }
    Ok(FormulaValue::Number(
        (mean + stdev * inverse_normal(p)).exp(),
    ))
}

pub fn fn_lognormdist(
    args: &[FormulaValue],
    _ctx: &EvaluationContext,
) -> FormulaResult<FormulaValue> {
    let x = match required_number(args, 0) {
        Ok(v) => v,
        Err(e) => return Ok(FormulaValue::Error(e)),
    };
    let mean = match required_number(args, 1) {
        Ok(v) => v,
        Err(e) => return Ok(FormulaValue::Error(e)),
    };
    let stdev = match required_number(args, 2) {
        Ok(v) => v,
        Err(e) => return Ok(FormulaValue::Error(e)),
    };
    if x <= 0.0 || stdev <= 0.0 {
        return Ok(FormulaValue::Error(CellError::Num));
    }
    Ok(FormulaValue::Number(normal_cdf((x.ln() - mean) / stdev)))
}

pub fn fn_negbinomdist(
    args: &[FormulaValue],
    _ctx: &EvaluationContext,
) -> FormulaResult<FormulaValue> {
    let f = int_floor(match required_number(args, 0) {
        Ok(v) => v,
        Err(e) => return Ok(FormulaValue::Error(e)),
    });
    let s = int_floor(match required_number(args, 1) {
        Ok(v) => v,
        Err(e) => return Ok(FormulaValue::Error(e)),
    });
    let p = match required_number(args, 2) {
        Ok(v) => v,
        Err(e) => return Ok(FormulaValue::Error(e)),
    };
    if f < 0 || s <= 0 || !(0.0..=1.0).contains(&p) {
        return Ok(FormulaValue::Error(CellError::Num));
    }
    let ln_p = ln_choose(f + s - 1, f) + (f as f64) * (1.0 - p).ln() + (s as f64) * p.ln();
    Ok(FormulaValue::Number(ln_p.exp()))
}

pub fn fn_normdist(args: &[FormulaValue], _ctx: &EvaluationContext) -> FormulaResult<FormulaValue> {
    let x = match required_number(args, 0) {
        Ok(v) => v,
        Err(e) => return Ok(FormulaValue::Error(e)),
    };
    let mean = match required_number(args, 1) {
        Ok(v) => v,
        Err(e) => return Ok(FormulaValue::Error(e)),
    };
    let stdev = match required_number(args, 2) {
        Ok(v) => v,
        Err(e) => return Ok(FormulaValue::Error(e)),
    };
    let cumulative = match required_bool(args, 3) {
        Ok(v) => v,
        Err(e) => return Ok(FormulaValue::Error(e)),
    };
    if stdev <= 0.0 {
        return Ok(FormulaValue::Error(CellError::Num));
    }
    let z = (x - mean) / stdev;
    let result = if cumulative {
        normal_cdf(z)
    } else {
        normal_pdf(z) / stdev
    };
    Ok(FormulaValue::Number(result))
}

pub fn fn_norm_inv(args: &[FormulaValue], _ctx: &EvaluationContext) -> FormulaResult<FormulaValue> {
    let p = match required_number(args, 0) {
        Ok(v) => v,
        Err(e) => return Ok(FormulaValue::Error(e)),
    };
    let mean = match required_number(args, 1) {
        Ok(v) => v,
        Err(e) => return Ok(FormulaValue::Error(e)),
    };
    let stdev = match required_number(args, 2) {
        Ok(v) => v,
        Err(e) => return Ok(FormulaValue::Error(e)),
    };
    if p <= 0.0 || p >= 1.0 || stdev <= 0.0 {
        return Ok(FormulaValue::Error(CellError::Num));
    }
    Ok(FormulaValue::Number(mean + stdev * inverse_normal(p)))
}

pub fn fn_normsdist(
    args: &[FormulaValue],
    _ctx: &EvaluationContext,
) -> FormulaResult<FormulaValue> {
    let z = match required_number(args, 0) {
        Ok(v) => v,
        Err(e) => return Ok(FormulaValue::Error(e)),
    };
    Ok(FormulaValue::Number(normal_cdf(z)))
}

pub fn fn_normsinv(args: &[FormulaValue], _ctx: &EvaluationContext) -> FormulaResult<FormulaValue> {
    let p = match required_number(args, 0) {
        Ok(v) => v,
        Err(e) => return Ok(FormulaValue::Error(e)),
    };
    if p <= 0.0 || p >= 1.0 {
        return Ok(FormulaValue::Error(CellError::Num));
    }
    Ok(FormulaValue::Number(inverse_normal(p)))
}

pub fn fn_poisson(args: &[FormulaValue], _ctx: &EvaluationContext) -> FormulaResult<FormulaValue> {
    let x = int_floor(match required_number(args, 0) {
        Ok(v) => v,
        Err(e) => return Ok(FormulaValue::Error(e)),
    });
    let mean = match required_number(args, 1) {
        Ok(v) => v,
        Err(e) => return Ok(FormulaValue::Error(e)),
    };
    let cumulative = match required_bool(args, 2) {
        Ok(v) => v,
        Err(e) => return Ok(FormulaValue::Error(e)),
    };
    if x < 0 || mean <= 0.0 {
        return Ok(FormulaValue::Error(CellError::Num));
    }
    let pmf = |k: i64| ((k as f64) * mean.ln() - mean - ln_gamma((k + 1) as f64)).exp();
    let result = if cumulative {
        (0..=x).map(pmf).sum()
    } else {
        pmf(x)
    };
    Ok(FormulaValue::Number(result))
}

pub fn fn_tdist(args: &[FormulaValue], _ctx: &EvaluationContext) -> FormulaResult<FormulaValue> {
    let x = match required_number(args, 0) {
        Ok(v) => v,
        Err(e) => return Ok(FormulaValue::Error(e)),
    };
    let df = match required_number(args, 1) {
        Ok(v) => v,
        Err(e) => return Ok(FormulaValue::Error(e)),
    };
    let tails = int_floor(match required_number(args, 2) {
        Ok(v) => v,
        Err(e) => return Ok(FormulaValue::Error(e)),
    });
    if x < 0.0 || df <= 0.0 || (tails != 1 && tails != 2) {
        return Ok(FormulaValue::Error(CellError::Num));
    }
    let right_tail = 1.0 - t_cdf(x, df);
    let result = if tails == 1 {
        right_tail
    } else {
        2.0 * right_tail
    };
    Ok(FormulaValue::Number(result))
}

pub fn fn_tinv(args: &[FormulaValue], _ctx: &EvaluationContext) -> FormulaResult<FormulaValue> {
    let p = match required_number(args, 0) {
        Ok(v) => v,
        Err(e) => return Ok(FormulaValue::Error(e)),
    };
    let df = match required_number(args, 1) {
        Ok(v) => v,
        Err(e) => return Ok(FormulaValue::Error(e)),
    };
    if p <= 0.0 || p >= 1.0 || df <= 0.0 {
        return Ok(FormulaValue::Error(CellError::Num));
    }
    Ok(FormulaValue::Number(inverse_t(1.0 - p / 2.0, df)))
}

pub fn fn_ttest(args: &[FormulaValue], _ctx: &EvaluationContext) -> FormulaResult<FormulaValue> {
    let tails = int_floor(match required_number(args, 2) {
        Ok(v) => v,
        Err(e) => return Ok(FormulaValue::Error(e)),
    });
    let ttype = int_floor(match required_number(args, 3) {
        Ok(v) => v,
        Err(e) => return Ok(FormulaValue::Error(e)),
    });
    if tails != 1 && tails != 2 {
        return Ok(FormulaValue::Error(CellError::Num));
    }

    let mut a = Vec::new();
    let mut b = Vec::new();
    if let Some(err) = collect_numbers(args.get(0).unwrap_or(&FormulaValue::Empty), &mut a) {
        return Ok(FormulaValue::Error(err));
    }
    if let Some(err) = collect_numbers(args.get(1).unwrap_or(&FormulaValue::Empty), &mut b) {
        return Ok(FormulaValue::Error(err));
    }

    let mean = |v: &[f64]| v.iter().sum::<f64>() / v.len() as f64;
    let var_s = |v: &[f64]| {
        let m = mean(v);
        v.iter().map(|x| (x - m) * (x - m)).sum::<f64>() / (v.len() as f64 - 1.0)
    };

    let (t_stat, df) = match ttype {
        1 => {
            if a.len() != b.len() || a.len() < 2 {
                return Ok(FormulaValue::Error(CellError::Num));
            }
            let diffs: Vec<f64> = a.iter().zip(b.iter()).map(|(x, y)| x - y).collect();
            let m = mean(&diffs);
            let s2 = var_s(&diffs);
            if s2 <= 0.0 {
                return Ok(FormulaValue::Error(CellError::Div0));
            }
            (
                m / (s2.sqrt() / (diffs.len() as f64).sqrt()),
                (diffs.len() - 1) as f64,
            )
        }
        2 => {
            if a.len() < 2 || b.len() < 2 {
                return Ok(FormulaValue::Error(CellError::Num));
            }
            let n1 = a.len() as f64;
            let n2 = b.len() as f64;
            let m1 = mean(&a);
            let m2 = mean(&b);
            let s1 = var_s(&a);
            let s2 = var_s(&b);
            let sp2 = ((n1 - 1.0) * s1 + (n2 - 1.0) * s2) / (n1 + n2 - 2.0);
            if sp2 <= 0.0 {
                return Ok(FormulaValue::Error(CellError::Div0));
            }
            (
                (m1 - m2) / (sp2 * (1.0 / n1 + 1.0 / n2)).sqrt(),
                n1 + n2 - 2.0,
            )
        }
        3 => {
            if a.len() < 2 || b.len() < 2 {
                return Ok(FormulaValue::Error(CellError::Num));
            }
            let n1 = a.len() as f64;
            let n2 = b.len() as f64;
            let m1 = mean(&a);
            let m2 = mean(&b);
            let s1 = var_s(&a);
            let s2 = var_s(&b);
            let v = s1 / n1 + s2 / n2;
            if v <= 0.0 {
                return Ok(FormulaValue::Error(CellError::Div0));
            }
            let df = v * v
                / (((s1 / n1) * (s1 / n1)) / (n1 - 1.0) + ((s2 / n2) * (s2 / n2)) / (n2 - 1.0));
            ((m1 - m2) / v.sqrt(), df)
        }
        _ => return Ok(FormulaValue::Error(CellError::Num)),
    };

    let one_tail = 1.0 - t_cdf(t_stat.abs(), df);
    let p = if tails == 1 { one_tail } else { 2.0 * one_tail };
    Ok(FormulaValue::Number(p.clamp(0.0, 1.0)))
}

pub fn fn_weibull(args: &[FormulaValue], _ctx: &EvaluationContext) -> FormulaResult<FormulaValue> {
    let x = match required_number(args, 0) {
        Ok(v) => v,
        Err(e) => return Ok(FormulaValue::Error(e)),
    };
    let alpha = match required_number(args, 1) {
        Ok(v) => v,
        Err(e) => return Ok(FormulaValue::Error(e)),
    };
    let beta = match required_number(args, 2) {
        Ok(v) => v,
        Err(e) => return Ok(FormulaValue::Error(e)),
    };
    let cumulative = match required_bool(args, 3) {
        Ok(v) => v,
        Err(e) => return Ok(FormulaValue::Error(e)),
    };
    if x < 0.0 || alpha <= 0.0 || beta <= 0.0 {
        return Ok(FormulaValue::Error(CellError::Num));
    }
    let xb = (x / beta).powf(alpha);
    let result = if cumulative {
        1.0 - (-xb).exp()
    } else if x == 0.0 {
        if alpha < 1.0 {
            f64::INFINITY
        } else if alpha == 1.0 {
            1.0 / beta
        } else {
            0.0
        }
    } else {
        (alpha / beta) * (x / beta).powf(alpha - 1.0) * (-xb).exp()
    };
    Ok(FormulaValue::Number(result))
}

pub fn fn_ztest(args: &[FormulaValue], _ctx: &EvaluationContext) -> FormulaResult<FormulaValue> {
    let mut values = Vec::new();
    if let Some(err) = collect_numbers(args.get(0).unwrap_or(&FormulaValue::Empty), &mut values) {
        return Ok(FormulaValue::Error(err));
    }
    if values.is_empty() {
        return Ok(FormulaValue::Error(CellError::Div0));
    }
    let x = match required_number(args, 1) {
        Ok(v) => v,
        Err(e) => return Ok(FormulaValue::Error(e)),
    };

    let n = values.len() as f64;
    let mean = values.iter().sum::<f64>() / n;
    let sigma = if args.len() >= 3 {
        match required_number(args, 2) {
            Ok(v) => v,
            Err(e) => return Ok(FormulaValue::Error(e)),
        }
    } else {
        if values.len() < 2 {
            return Ok(FormulaValue::Error(CellError::Div0));
        }
        let var = values.iter().map(|v| (v - mean) * (v - mean)).sum::<f64>() / (n - 1.0);
        var.sqrt()
    };
    if sigma <= 0.0 {
        return Ok(FormulaValue::Error(CellError::Num));
    }
    let z = (mean - x) / (sigma / n.sqrt());
    Ok(FormulaValue::Number(1.0 - normal_cdf(z)))
}

pub fn fn_ceiling(args: &[FormulaValue], _ctx: &EvaluationContext) -> FormulaResult<FormulaValue> {
    let number = match required_number(args, 0) {
        Ok(v) => v,
        Err(e) => return Ok(FormulaValue::Error(e)),
    };
    let significance = match required_number(args, 1) {
        Ok(v) => v,
        Err(e) => return Ok(FormulaValue::Error(e)),
    };
    if significance == 0.0 {
        return Ok(FormulaValue::Number(0.0));
    }
    if (number > 0.0 && significance < 0.0) || (number < 0.0 && significance > 0.0) {
        return Ok(FormulaValue::Error(CellError::Num));
    }
    Ok(FormulaValue::Number(
        (number / significance).ceil() * significance,
    ))
}

pub fn fn_floor(args: &[FormulaValue], _ctx: &EvaluationContext) -> FormulaResult<FormulaValue> {
    let number = match required_number(args, 0) {
        Ok(v) => v,
        Err(e) => return Ok(FormulaValue::Error(e)),
    };
    let significance = match required_number(args, 1) {
        Ok(v) => v,
        Err(e) => return Ok(FormulaValue::Error(e)),
    };
    if significance == 0.0 {
        return Ok(FormulaValue::Number(0.0));
    }
    if (number > 0.0 && significance < 0.0) || (number < 0.0 && significance > 0.0) {
        return Ok(FormulaValue::Error(CellError::Num));
    }
    Ok(FormulaValue::Number(
        (number / significance).floor() * significance,
    ))
}

#[cfg(test)]
mod tests {
    use super::*;

    fn n(x: f64) -> FormulaValue {
        FormulaValue::Number(x)
    }

    fn b(v: bool) -> FormulaValue {
        FormulaValue::Boolean(v)
    }

    fn arr(values: &[f64]) -> FormulaValue {
        FormulaValue::Array(vec![values
            .iter()
            .map(|v| FormulaValue::Number(*v))
            .collect()])
    }

    fn assert_close(v: FormulaValue, expected: f64, tol: f64) {
        match v {
            FormulaValue::Number(x) => {
                assert!((x - expected).abs() <= tol, "{} != {}", x, expected)
            }
            _ => panic!("expected number"),
        }
    }

    #[test]
    fn test_betadist() {
        let ctx = EvaluationContext::simple();
        assert_close(
            fn_betadist(&[n(0.3), n(1.0), n(1.0)], &ctx).unwrap(),
            0.3,
            1e-9,
        );
        assert_close(
            fn_betadist(&[n(3.0), n(1.0), n(1.0), n(2.0), n(6.0)], &ctx).unwrap(),
            0.25,
            1e-9,
        );
    }

    #[test]
    fn test_betainv() {
        let ctx = EvaluationContext::simple();
        assert_close(
            fn_betainv(&[n(0.3), n(1.0), n(1.0)], &ctx).unwrap(),
            0.3,
            1e-8,
        );
        assert_close(
            fn_betainv(&[n(0.25), n(1.0), n(1.0), n(2.0), n(6.0)], &ctx).unwrap(),
            3.0,
            1e-8,
        );
    }

    #[test]
    fn test_binomdist() {
        let ctx = EvaluationContext::simple();
        assert_close(
            fn_binomdist(&[n(6.0), n(10.0), n(0.5), b(false)], &ctx).unwrap(),
            0.205078125,
            1e-12,
        );
        assert_close(
            fn_binomdist(&[n(6.0), n(10.0), n(0.5), b(true)], &ctx).unwrap(),
            0.828125,
            1e-12,
        );
    }

    #[test]
    fn test_chidist() {
        let ctx = EvaluationContext::simple();
        assert_close(fn_chidist(&[n(18.307), n(10.0)], &ctx).unwrap(), 0.05, 1e-2);
        assert_close(fn_chidist(&[n(0.0), n(2.0)], &ctx).unwrap(), 1.0, 1e-12);
    }

    #[test]
    fn test_chiinv() {
        let ctx = EvaluationContext::simple();
        let x = fn_chiinv(&[n(0.05), n(10.0)], &ctx).unwrap();
        if let FormulaValue::Number(v) = x {
            assert!((v - 18.307).abs() < 0.2);
            assert_close(fn_chidist(&[n(v), n(10.0)], &ctx).unwrap(), 0.05, 2e-3);
        } else {
            panic!("expected number")
        }
    }

    #[test]
    fn test_chitest() {
        let ctx = EvaluationContext::simple();
        let a = FormulaValue::Array(vec![vec![n(10.0), n(20.0)], vec![n(20.0), n(40.0)]]);
        let e = FormulaValue::Array(vec![vec![n(10.0), n(20.0)], vec![n(20.0), n(40.0)]]);
        assert_close(fn_chitest(&[a.clone(), e], &ctx).unwrap(), 1.0, 1e-12);
        let e2 = FormulaValue::Array(vec![vec![n(15.0), n(15.0)], vec![n(15.0), n(45.0)]]);
        if let FormulaValue::Number(p) = fn_chitest(&[a, e2], &ctx).unwrap() {
            assert!(p < 1.0);
        } else {
            panic!("expected number")
        }
    }

    #[test]
    fn test_confidence() {
        let ctx = EvaluationContext::simple();
        assert_close(
            fn_confidence(&[n(0.05), n(2.0), n(100.0)], &ctx).unwrap(),
            0.3919927969,
            1e-6,
        );
        assert_eq!(
            fn_confidence(&[n(1.0), n(2.0), n(100.0)], &ctx).unwrap(),
            FormulaValue::Error(CellError::Num)
        );
    }

    #[test]
    fn test_covar() {
        let ctx = EvaluationContext::simple();
        assert_close(
            fn_covar(&[arr(&[1.0, 2.0, 3.0]), arr(&[1.0, 2.0, 3.0])], &ctx).unwrap(),
            2.0 / 3.0,
            1e-12,
        );
        assert_close(
            fn_covar(&[arr(&[1.0, 2.0, 3.0]), arr(&[3.0, 2.0, 1.0])], &ctx).unwrap(),
            -2.0 / 3.0,
            1e-12,
        );
    }

    #[test]
    fn test_critbinom() {
        let ctx = EvaluationContext::simple();
        assert_eq!(
            fn_critbinom(&[n(10.0), n(0.5), n(0.5)], &ctx).unwrap(),
            n(5.0)
        );
        assert_eq!(
            fn_critbinom(&[n(10.0), n(0.5), n(0.9)], &ctx).unwrap(),
            n(7.0)
        );
    }

    #[test]
    fn test_expondist() {
        let ctx = EvaluationContext::simple();
        assert_close(
            fn_expondist(&[n(1.0), n(1.0), b(true)], &ctx).unwrap(),
            1.0 - (-1.0f64).exp(),
            1e-12,
        );
        assert_close(
            fn_expondist(&[n(1.0), n(1.0), b(false)], &ctx).unwrap(),
            (-1.0f64).exp(),
            1e-12,
        );
    }

    #[test]
    fn test_fdist() {
        let ctx = EvaluationContext::simple();
        if let FormulaValue::Number(p) = fn_fdist(&[n(2.0), n(5.0), n(10.0)], &ctx).unwrap() {
            assert!(p > 0.0 && p < 1.0);
        } else {
            panic!("expected number")
        }
        assert_close(
            fn_fdist(&[n(0.0), n(5.0), n(10.0)], &ctx).unwrap(),
            1.0,
            1e-12,
        );
    }

    #[test]
    fn test_finv() {
        let ctx = EvaluationContext::simple();
        let x = fn_finv(&[n(0.05), n(5.0), n(10.0)], &ctx).unwrap();
        if let FormulaValue::Number(v) = x {
            assert_close(
                fn_fdist(&[n(v), n(5.0), n(10.0)], &ctx).unwrap(),
                0.05,
                5e-3,
            );
            assert!(v > 0.0);
        } else {
            panic!("expected number")
        }
    }

    #[test]
    fn test_ftest() {
        let ctx = EvaluationContext::simple();
        assert_close(
            fn_ftest(&[arr(&[1.0, 2.0, 3.0]), arr(&[1.0, 2.0, 3.0])], &ctx).unwrap(),
            1.0,
            1e-6,
        );
        if let FormulaValue::Number(p) =
            fn_ftest(&[arr(&[1.0, 2.0, 3.0]), arr(&[1.0, 10.0, 20.0])], &ctx).unwrap()
        {
            assert!(p < 1.0);
        } else {
            panic!("expected number")
        }
    }

    #[test]
    fn test_gammadist() {
        let ctx = EvaluationContext::simple();
        assert_close(
            fn_gammadist(&[n(2.0), n(1.0), n(2.0), b(true)], &ctx).unwrap(),
            1.0 - (-1.0f64).exp(),
            1e-8,
        );
        assert_close(
            fn_gammadist(&[n(2.0), n(1.0), n(2.0), b(false)], &ctx).unwrap(),
            (-1.0f64).exp() / 2.0,
            1e-8,
        );
    }

    #[test]
    fn test_gammainv() {
        let ctx = EvaluationContext::simple();
        assert_close(
            fn_gammainv(&[n(1.0 - (-1.0f64).exp()), n(1.0), n(2.0)], &ctx).unwrap(),
            2.0,
            1e-5,
        );
        assert_eq!(
            fn_gammainv(&[n(0.0), n(1.0), n(2.0)], &ctx).unwrap(),
            FormulaValue::Error(CellError::Num)
        );
    }

    #[test]
    fn test_hypgeomdist() {
        let ctx = EvaluationContext::simple();
        assert_close(
            fn_hypgeomdist(&[n(1.0), n(2.0), n(5.0), n(10.0)], &ctx).unwrap(),
            25.0 / 45.0,
            1e-12,
        );
        assert_close(
            fn_hypgeomdist(&[n(2.0), n(2.0), n(5.0), n(10.0)], &ctx).unwrap(),
            10.0 / 45.0,
            1e-12,
        );
    }

    #[test]
    fn test_loginv() {
        let ctx = EvaluationContext::simple();
        assert_close(
            fn_loginv(&[n(0.5), n(0.0), n(1.0)], &ctx).unwrap(),
            1.0,
            1e-12,
        );
        assert_close(
            fn_loginv(&[n(0.841344746), n(0.0), n(1.0)], &ctx).unwrap(),
            std::f64::consts::E,
            1e-5,
        );
    }

    #[test]
    fn test_lognormdist() {
        let ctx = EvaluationContext::simple();
        assert_close(
            fn_lognormdist(&[n(1.0), n(0.0), n(1.0)], &ctx).unwrap(),
            0.5,
            1e-12,
        );
        assert_close(
            fn_lognormdist(&[n(std::f64::consts::E), n(0.0), n(1.0)], &ctx).unwrap(),
            0.841344746,
            1e-5,
        );
    }

    #[test]
    fn test_negbinomdist() {
        let ctx = EvaluationContext::simple();
        assert_close(
            fn_negbinomdist(&[n(0.0), n(1.0), n(0.2)], &ctx).unwrap(),
            0.2,
            1e-12,
        );
        assert_close(
            fn_negbinomdist(&[n(1.0), n(1.0), n(0.2)], &ctx).unwrap(),
            0.16,
            1e-12,
        );
    }

    #[test]
    fn test_normdist() {
        let ctx = EvaluationContext::simple();
        assert_close(
            fn_normdist(&[n(0.0), n(0.0), n(1.0), b(true)], &ctx).unwrap(),
            0.5,
            1e-12,
        );
        assert_close(
            fn_normdist(&[n(0.0), n(0.0), n(1.0), b(false)], &ctx).unwrap(),
            0.3989422804,
            1e-9,
        );
    }

    #[test]
    fn test_norm_inv() {
        let ctx = EvaluationContext::simple();
        assert_close(
            fn_norm_inv(&[n(0.5), n(0.0), n(1.0)], &ctx).unwrap(),
            0.0,
            1e-12,
        );
        assert_close(
            fn_norm_inv(&[n(0.841344746), n(0.0), n(1.0)], &ctx).unwrap(),
            1.0,
            1e-5,
        );
    }

    #[test]
    fn test_normsdist() {
        let ctx = EvaluationContext::simple();
        assert_close(fn_normsdist(&[n(0.0)], &ctx).unwrap(), 0.5, 1e-12);
        assert_close(fn_normsdist(&[n(1.0)], &ctx).unwrap(), 0.841344746, 1e-5);
    }

    #[test]
    fn test_normsinv() {
        let ctx = EvaluationContext::simple();
        assert_close(fn_normsinv(&[n(0.5)], &ctx).unwrap(), 0.0, 1e-12);
        assert_close(fn_normsinv(&[n(0.841344746)], &ctx).unwrap(), 1.0, 1e-5);
    }

    #[test]
    fn test_poisson() {
        let ctx = EvaluationContext::simple();
        assert_close(
            fn_poisson(&[n(0.0), n(2.0), b(false)], &ctx).unwrap(),
            (-2.0f64).exp(),
            1e-12,
        );
        assert_close(
            fn_poisson(&[n(1.0), n(2.0), b(true)], &ctx).unwrap(),
            (-2.0f64).exp() * 3.0,
            1e-12,
        );
    }

    #[test]
    fn test_tdist() {
        let ctx = EvaluationContext::simple();
        assert_close(
            fn_tdist(&[n(1.96), n(100.0), n(2.0)], &ctx).unwrap(),
            0.0528,
            6e-3,
        );
        assert_close(
            fn_tdist(&[n(1.96), n(100.0), n(1.0)], &ctx).unwrap(),
            0.0264,
            3e-3,
        );
    }

    #[test]
    fn test_tinv() {
        let ctx = EvaluationContext::simple();
        let x = fn_tinv(&[n(0.05), n(100.0)], &ctx).unwrap();
        if let FormulaValue::Number(v) = x {
            assert!((v - 1.984).abs() < 0.05);
            assert_close(
                fn_tdist(&[n(v), n(100.0), n(2.0)], &ctx).unwrap(),
                0.05,
                3e-3,
            );
        } else {
            panic!("expected number")
        }
    }

    #[test]
    fn test_ttest() {
        let ctx = EvaluationContext::simple();
        assert_close(
            fn_ttest(
                &[arr(&[1.0, 2.0, 3.0]), arr(&[1.0, 2.0, 3.0]), n(2.0), n(2.0)],
                &ctx,
            )
            .unwrap(),
            1.0,
            1e-8,
        );
        if let FormulaValue::Number(p) = fn_ttest(
            &[arr(&[1.0, 2.0, 3.0]), arr(&[5.0, 6.0, 7.0]), n(2.0), n(3.0)],
            &ctx,
        )
        .unwrap()
        {
            assert!(p < 0.1);
        } else {
            panic!("expected number")
        }
    }

    #[test]
    fn test_weibull() {
        let ctx = EvaluationContext::simple();
        assert_close(
            fn_weibull(&[n(2.0), n(1.0), n(2.0), b(true)], &ctx).unwrap(),
            1.0 - (-1.0f64).exp(),
            1e-10,
        );
        assert_close(
            fn_weibull(&[n(2.0), n(1.0), n(2.0), b(false)], &ctx).unwrap(),
            (-1.0f64).exp() / 2.0,
            1e-10,
        );
    }

    #[test]
    fn test_ztest() {
        let ctx = EvaluationContext::simple();
        assert_close(
            fn_ztest(&[arr(&[1.0, 2.0, 3.0, 4.0, 5.0]), n(3.0)], &ctx).unwrap(),
            0.5,
            1e-12,
        );
        if let FormulaValue::Number(p) =
            fn_ztest(&[arr(&[1.0, 2.0, 3.0, 4.0, 5.0]), n(2.0), n(1.0)], &ctx).unwrap()
        {
            assert!(p < 0.02);
        } else {
            panic!("expected number")
        }
    }

    #[test]
    fn test_ceiling() {
        let ctx = EvaluationContext::simple();
        assert_eq!(fn_ceiling(&[n(2.5), n(1.0)], &ctx).unwrap(), n(3.0));
        assert_eq!(fn_ceiling(&[n(-2.5), n(-2.0)], &ctx).unwrap(), n(-4.0));
    }

    #[test]
    fn test_floor() {
        let ctx = EvaluationContext::simple();
        assert_eq!(fn_floor(&[n(2.5), n(1.0)], &ctx).unwrap(), n(2.0));
        assert_eq!(fn_floor(&[n(-2.5), n(-2.0)], &ctx).unwrap(), n(-2.0));
    }

    #[test]
    fn test_betadist_errors() {
        let ctx = EvaluationContext::simple();
        assert_eq!(
            fn_betadist(&[n(2.0), n(1.0), n(1.0)], &ctx).unwrap(),
            FormulaValue::Error(CellError::Num)
        );
    }

    #[test]
    fn test_betainv_errors() {
        let ctx = EvaluationContext::simple();
        assert_eq!(
            fn_betainv(&[n(-0.1), n(1.0), n(1.0)], &ctx).unwrap(),
            FormulaValue::Error(CellError::Num)
        );
    }

    #[test]
    fn test_binomdist_errors() {
        let ctx = EvaluationContext::simple();
        assert_eq!(
            fn_binomdist(&[n(6.0), n(5.0), n(0.5), b(false)], &ctx).unwrap(),
            FormulaValue::Error(CellError::Num)
        );
    }

    #[test]
    fn test_chidist_errors() {
        let ctx = EvaluationContext::simple();
        assert_eq!(
            fn_chidist(&[n(-1.0), n(10.0)], &ctx).unwrap(),
            FormulaValue::Error(CellError::Num)
        );
    }

    #[test]
    fn test_chiinv_errors() {
        let ctx = EvaluationContext::simple();
        assert_eq!(
            fn_chiinv(&[n(0.0), n(10.0)], &ctx).unwrap(),
            FormulaValue::Error(CellError::Num)
        );
    }

    #[test]
    fn test_chitest_errors() {
        let ctx = EvaluationContext::simple();
        assert_eq!(
            fn_chitest(&[arr(&[1.0, 2.0]), arr(&[1.0])], &ctx).unwrap(),
            FormulaValue::Error(CellError::Value)
        );
    }

    #[test]
    fn test_confidence_errors() {
        let ctx = EvaluationContext::simple();
        assert_eq!(
            fn_confidence(&[n(0.05), n(0.0), n(10.0)], &ctx).unwrap(),
            FormulaValue::Error(CellError::Num)
        );
    }

    #[test]
    fn test_covar_errors() {
        let ctx = EvaluationContext::simple();
        let bad = FormulaValue::Array(vec![vec![n(1.0)], vec![n(2.0)]]);
        assert_eq!(
            fn_covar(&[bad, arr(&[1.0, 2.0])], &ctx).unwrap(),
            FormulaValue::Error(CellError::Value)
        );
    }

    #[test]
    fn test_critbinom_edges() {
        let ctx = EvaluationContext::simple();
        assert_eq!(
            fn_critbinom(&[n(10.0), n(0.5), n(0.0)], &ctx).unwrap(),
            n(0.0)
        );
    }

    #[test]
    fn test_expondist_errors() {
        let ctx = EvaluationContext::simple();
        assert_eq!(
            fn_expondist(&[n(-1.0), n(1.0), b(true)], &ctx).unwrap(),
            FormulaValue::Error(CellError::Num)
        );
    }

    #[test]
    fn test_fdist_errors() {
        let ctx = EvaluationContext::simple();
        assert_eq!(
            fn_fdist(&[n(1.0), n(0.0), n(10.0)], &ctx).unwrap(),
            FormulaValue::Error(CellError::Num)
        );
    }

    #[test]
    fn test_finv_errors() {
        let ctx = EvaluationContext::simple();
        assert_eq!(
            fn_finv(&[n(1.0), n(5.0), n(10.0)], &ctx).unwrap(),
            FormulaValue::Error(CellError::Num)
        );
    }

    #[test]
    fn test_ftest_errors() {
        let ctx = EvaluationContext::simple();
        assert_eq!(
            fn_ftest(&[arr(&[1.0]), arr(&[1.0])], &ctx).unwrap(),
            FormulaValue::Error(CellError::Div0)
        );
    }

    #[test]
    fn test_gammadist_errors() {
        let ctx = EvaluationContext::simple();
        assert_eq!(
            fn_gammadist(&[n(-1.0), n(1.0), n(1.0), b(true)], &ctx).unwrap(),
            FormulaValue::Error(CellError::Num)
        );
    }

    #[test]
    fn test_gammainv_errors() {
        let ctx = EvaluationContext::simple();
        assert_eq!(
            fn_gammainv(&[n(1.0), n(1.0), n(1.0)], &ctx).unwrap(),
            FormulaValue::Error(CellError::Num)
        );
    }

    #[test]
    fn test_hypgeomdist_errors() {
        let ctx = EvaluationContext::simple();
        assert_eq!(
            fn_hypgeomdist(&[n(3.0), n(2.0), n(5.0), n(10.0)], &ctx).unwrap(),
            FormulaValue::Error(CellError::Num)
        );
    }

    #[test]
    fn test_loginv_errors() {
        let ctx = EvaluationContext::simple();
        assert_eq!(
            fn_loginv(&[n(0.0), n(0.0), n(1.0)], &ctx).unwrap(),
            FormulaValue::Error(CellError::Num)
        );
    }

    #[test]
    fn test_lognormdist_errors() {
        let ctx = EvaluationContext::simple();
        assert_eq!(
            fn_lognormdist(&[n(0.0), n(0.0), n(1.0)], &ctx).unwrap(),
            FormulaValue::Error(CellError::Num)
        );
    }

    #[test]
    fn test_negbinomdist_errors() {
        let ctx = EvaluationContext::simple();
        assert_eq!(
            fn_negbinomdist(&[n(1.0), n(0.0), n(0.2)], &ctx).unwrap(),
            FormulaValue::Error(CellError::Num)
        );
    }

    #[test]
    fn test_normdist_errors() {
        let ctx = EvaluationContext::simple();
        assert_eq!(
            fn_normdist(&[n(0.0), n(0.0), n(0.0), b(true)], &ctx).unwrap(),
            FormulaValue::Error(CellError::Num)
        );
    }

    #[test]
    fn test_norm_inv_errors() {
        let ctx = EvaluationContext::simple();
        assert_eq!(
            fn_norm_inv(&[n(0.0), n(0.0), n(1.0)], &ctx).unwrap(),
            FormulaValue::Error(CellError::Num)
        );
    }

    #[test]
    fn test_normsdist_extra() {
        let ctx = EvaluationContext::simple();
        assert_close(fn_normsdist(&[n(-1.0)], &ctx).unwrap(), 0.158655254, 1e-5);
    }

    #[test]
    fn test_normsinv_errors() {
        let ctx = EvaluationContext::simple();
        assert_eq!(
            fn_normsinv(&[n(1.0)], &ctx).unwrap(),
            FormulaValue::Error(CellError::Num)
        );
    }

    #[test]
    fn test_poisson_errors() {
        let ctx = EvaluationContext::simple();
        assert_eq!(
            fn_poisson(&[n(-1.0), n(2.0), b(false)], &ctx).unwrap(),
            FormulaValue::Error(CellError::Num)
        );
    }

    #[test]
    fn test_tdist_errors() {
        let ctx = EvaluationContext::simple();
        assert_eq!(
            fn_tdist(&[n(1.0), n(10.0), n(3.0)], &ctx).unwrap(),
            FormulaValue::Error(CellError::Num)
        );
    }

    #[test]
    fn test_tinv_errors() {
        let ctx = EvaluationContext::simple();
        assert_eq!(
            fn_tinv(&[n(1.0), n(10.0)], &ctx).unwrap(),
            FormulaValue::Error(CellError::Num)
        );
    }

    #[test]
    fn test_ttest_errors() {
        let ctx = EvaluationContext::simple();
        assert_eq!(
            fn_ttest(&[arr(&[1.0]), arr(&[1.0]), n(2.0), n(2.0)], &ctx).unwrap(),
            FormulaValue::Error(CellError::Num)
        );
    }

    #[test]
    fn test_weibull_errors() {
        let ctx = EvaluationContext::simple();
        assert_eq!(
            fn_weibull(&[n(-1.0), n(1.0), n(1.0), b(true)], &ctx).unwrap(),
            FormulaValue::Error(CellError::Num)
        );
    }

    #[test]
    fn test_ztest_errors() {
        let ctx = EvaluationContext::simple();
        assert_eq!(
            fn_ztest(&[arr(&[1.0]), n(1.0)], &ctx).unwrap(),
            FormulaValue::Error(CellError::Div0)
        );
    }

    #[test]
    fn test_ceiling_errors() {
        let ctx = EvaluationContext::simple();
        assert_eq!(
            fn_ceiling(&[n(2.0), n(-1.0)], &ctx).unwrap(),
            FormulaValue::Error(CellError::Num)
        );
    }

    #[test]
    fn test_floor_errors() {
        let ctx = EvaluationContext::simple();
        assert_eq!(
            fn_floor(&[n(2.0), n(-1.0)], &ctx).unwrap(),
            FormulaValue::Error(CellError::Num)
        );
    }
}
