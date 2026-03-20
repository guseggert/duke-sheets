//! Additional statistical functions.

use crate::error::FormulaResult;
use crate::evaluator::{EvaluationContext, FormulaValue};
use duke_sheets_core::CellError;

const SQRT_2PI: f64 = 2.506_628_274_631_000_2;

fn scalar_number(value: &FormulaValue) -> Result<f64, FormulaValue> {
    match value {
        FormulaValue::Error(e) => Err(FormulaValue::Error(*e)),
        FormulaValue::Array { .. } => Err(FormulaValue::Error(CellError::Value)),
        _ => value
            .as_number()
            .ok_or(FormulaValue::Error(CellError::Value)),
    }
}

fn required_number(args: &[FormulaValue], idx: usize) -> Result<f64, FormulaValue> {
    args.get(idx)
        .ok_or(FormulaValue::Error(CellError::Value))
        .and_then(scalar_number)
}

fn optional_number(args: &[FormulaValue], idx: usize, default: f64) -> Result<f64, FormulaValue> {
    match args.get(idx).filter(|v| !matches!(v, FormulaValue::Empty)) {
        Some(v) => scalar_number(v),
        None => Ok(default),
    }
}

fn scalar_bool(value: &FormulaValue) -> Result<bool, FormulaValue> {
    match value {
        FormulaValue::Error(e) => Err(FormulaValue::Error(*e)),
        FormulaValue::Array { .. } => Err(FormulaValue::Error(CellError::Value)),
        _ => value.as_bool().ok_or(FormulaValue::Error(CellError::Value)),
    }
}

fn required_bool(args: &[FormulaValue], idx: usize) -> Result<bool, FormulaValue> {
    args.get(idx)
        .ok_or(FormulaValue::Error(CellError::Value))
        .and_then(scalar_bool)
}

fn optional_bool(args: &[FormulaValue], idx: usize, default: bool) -> Result<bool, FormulaValue> {
    match args.get(idx).filter(|v| !matches!(v, FormulaValue::Empty)) {
        Some(v) => scalar_bool(v),
        None => Ok(default),
    }
}

fn flatten_numbers(value: &FormulaValue, out: &mut Vec<f64>) -> Result<(), FormulaValue> {
    match value {
        FormulaValue::Error(e) => Err(FormulaValue::Error(*e)),
        FormulaValue::Array { data: rows, .. } => {
            for row in rows {
                for cell in row {
                    flatten_numbers(cell, out)?;
                }
            }
            Ok(())
        }
        _ => {
            out.push(
                value
                    .as_number()
                    .ok_or(FormulaValue::Error(CellError::Value))?,
            );
            Ok(())
        }
    }
}

fn matrix_numbers(value: &FormulaValue) -> Result<Vec<Vec<f64>>, FormulaValue> {
    match value {
        FormulaValue::Error(e) => Err(FormulaValue::Error(*e)),
        FormulaValue::Array { data: rows, .. } => {
            if rows.is_empty() {
                return Err(FormulaValue::Error(CellError::Value));
            }
            let cols = rows[0].len();
            if cols == 0 {
                return Err(FormulaValue::Error(CellError::Value));
            }
            let mut out = Vec::with_capacity(rows.len());
            for row in rows {
                if row.len() != cols {
                    return Err(FormulaValue::Error(CellError::Value));
                }
                let mut out_row = Vec::with_capacity(cols);
                for cell in row {
                    out_row.push(
                        cell.as_number()
                            .ok_or(FormulaValue::Error(CellError::Value))?,
                    );
                }
                out.push(out_row);
            }
            Ok(out)
        }
        _ => Ok(vec![vec![scalar_number(value)?]]),
    }
}

fn transpose(m: &[Vec<f64>]) -> Vec<Vec<f64>> {
    if m.is_empty() || m[0].is_empty() {
        return vec![];
    }
    let rows = m.len();
    let cols = m[0].len();
    let mut t = vec![vec![0.0; rows]; cols];
    for (r, row) in m.iter().enumerate() {
        for (c, value) in row.iter().enumerate() {
            t[c][r] = *value;
        }
    }
    t
}

fn normal_pdf(x: f64) -> f64 {
    (-0.5 * x * x).exp() / SQRT_2PI
}

fn erf(x: f64) -> f64 {
    let sign = if x < 0.0 { -1.0 } else { 1.0 };
    let ax = x.abs();
    let t = 1.0 / (1.0 + 0.327_591_1 * ax);
    let a1 = 0.254_829_592;
    let a2 = -0.284_496_736;
    let a3 = 1.421_413_741;
    let a4 = -1.453_152_027;
    let a5 = 1.061_405_429;
    let y = 1.0 - (((((a5 * t + a4) * t + a3) * t + a2) * t + a1) * t) * (-(ax * ax)).exp();
    sign * y
}

fn normal_cdf(x: f64) -> f64 {
    0.5 * (1.0 + erf(x / std::f64::consts::SQRT_2))
}

fn inverse_normal(p: f64) -> f64 {
    let a = [
        -3.969_683_028_665_376e1,
        2.209_460_984_245_205e2,
        -2.759_285_104_469_687e2,
        1.383_577_518_672_69e2,
        -3.066_479_806_614_716e1,
        2.506_628_277_459_239,
    ];
    let b = [
        -5.447_609_879_822_406e1,
        1.615_858_368_580_409e2,
        -1.556_989_798_598_866e2,
        6.680_131_188_771_972e1,
        -1.328_068_155_288_572e1,
    ];
    let c = [
        -7.784_894_002_430_293e-3,
        -3.223_964_580_411_365e-1,
        -2.400_758_277_161_838,
        -2.549_732_539_343_734,
        4.374_664_141_464_968,
        2.938_163_982_698_783,
    ];
    let d = [
        7.784_695_709_041_462e-3,
        3.224_671_290_700_398e-1,
        2.445_134_137_142_996,
        3.754_408_661_907_416,
    ];

    let plow = 0.024_25;
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

#[allow(clippy::needless_range_loop)]
fn solve_gaussian(mut a: Vec<Vec<f64>>, mut b: Vec<f64>) -> Option<Vec<f64>> {
    let n = a.len();
    if n == 0 || b.len() != n {
        return None;
    }

    for i in 0..n {
        let mut pivot = i;
        let mut pivot_value = a[i][i].abs();
        for (r, row) in a.iter().enumerate().skip(i + 1) {
            if row[i].abs() > pivot_value {
                pivot = r;
                pivot_value = row[i].abs();
            }
        }
        if pivot_value < 1e-12 {
            return None;
        }

        if pivot != i {
            a.swap(i, pivot);
            b.swap(i, pivot);
        }

        let diag = a[i][i];
        for c in i..n {
            a[i][c] /= diag;
        }
        b[i] /= diag;

        for r in 0..n {
            if r == i {
                continue;
            }
            let factor = a[r][i];
            if factor.abs() < 1e-15 {
                continue;
            }
            for c in i..n {
                a[r][c] -= factor * a[i][c];
            }
            b[r] -= factor * b[i];
        }
    }

    Some(b)
}

fn invert_matrix(mut m: Vec<Vec<f64>>) -> Option<Vec<Vec<f64>>> {
    let n = m.len();
    if n == 0 || m.iter().any(|row| row.len() != n) {
        return None;
    }
    let mut inv = vec![vec![0.0; n]; n];
    for (i, row) in inv.iter_mut().enumerate().take(n) {
        row[i] = 1.0;
    }

    for i in 0..n {
        let mut pivot = i;
        let mut pivot_value = m[i][i].abs();
        for (r, row) in m.iter().enumerate().skip(i + 1) {
            if row[i].abs() > pivot_value {
                pivot = r;
                pivot_value = row[i].abs();
            }
        }
        if pivot_value < 1e-12 {
            return None;
        }

        if pivot != i {
            m.swap(i, pivot);
            inv.swap(i, pivot);
        }

        let diag = m[i][i];
        for c in 0..n {
            m[i][c] /= diag;
            inv[i][c] /= diag;
        }

        for r in 0..n {
            if r == i {
                continue;
            }
            let factor = m[r][i];
            if factor.abs() < 1e-15 {
                continue;
            }
            for c in 0..n {
                m[r][c] -= factor * m[i][c];
                inv[r][c] -= factor * inv[i][c];
            }
        }
    }

    Some(inv)
}

fn parse_known_x(value: Option<&FormulaValue>, n: usize) -> Result<Vec<Vec<f64>>, FormulaValue> {
    match value {
        None => {
            let mut x = Vec::with_capacity(n);
            for i in 0..n {
                x.push(vec![(i + 1) as f64]);
            }
            Ok(x)
        }
        Some(v) => {
            let m = matrix_numbers(v)?;
            let rows = m.len();
            let cols = m[0].len();

            if rows == n {
                return Ok(m);
            }
            if cols == n {
                return Ok(transpose(&m));
            }
            if rows * cols == n {
                let mut x = Vec::with_capacity(n);
                for row in &m {
                    for val in row {
                        x.push(vec![*val]);
                    }
                }
                return Ok(x);
            }
            Err(FormulaValue::Error(CellError::Value))
        }
    }
}

fn parse_new_x(
    value: Option<&FormulaValue>,
    default_x: &[Vec<f64>],
    predictors: usize,
) -> Result<Vec<Vec<f64>>, FormulaValue> {
    match value {
        None => Ok(default_x.to_vec()),
        Some(v) => {
            let m = matrix_numbers(v)?;
            let rows = m.len();
            let cols = m[0].len();

            if cols == predictors {
                return Ok(m);
            }
            if rows == predictors {
                return Ok(transpose(&m));
            }
            if predictors == 1 {
                let mut out = Vec::new();
                for row in &m {
                    for val in row {
                        out.push(vec![*val]);
                    }
                }
                return Ok(out);
            }
            Err(FormulaValue::Error(CellError::Value))
        }
    }
}

#[derive(Clone)]
struct LinearFit {
    slopes: Vec<f64>,
    intercept: f64,
    stderr_slopes: Vec<f64>,
    stderr_intercept: f64,
    r2: f64,
    se_y: f64,
    f_stat: f64,
    df: f64,
    ss_reg: f64,
    ss_res: f64,
}

fn fit_linear_regression(
    y: &[f64],
    x: &[Vec<f64>],
    include_const: bool,
) -> Result<LinearFit, FormulaValue> {
    let n = y.len();
    if n == 0 || x.len() != n {
        return Err(FormulaValue::Error(CellError::Value));
    }
    let p = x[0].len();
    if p == 0 || x.iter().any(|row| row.len() != p) {
        return Err(FormulaValue::Error(CellError::Value));
    }

    let k = p + if include_const { 1 } else { 0 };
    if n < k.max(1) {
        return Err(FormulaValue::Error(CellError::Num));
    }

    let mut xtx = vec![vec![0.0; k]; k];
    let mut xty = vec![0.0; k];

    for (i, row_x) in x.iter().enumerate().take(n) {
        let mut row = Vec::with_capacity(k);
        row.extend_from_slice(row_x);
        if include_const {
            row.push(1.0);
        }

        for r in 0..k {
            xty[r] += row[r] * y[i];
            for c in 0..k {
                xtx[r][c] += row[r] * row[c];
            }
        }
    }

    let beta = solve_gaussian(xtx.clone(), xty).ok_or(FormulaValue::Error(CellError::Num))?;
    let mut slopes = beta[..p].to_vec();
    let intercept = if include_const { beta[p] } else { 0.0 };

    let mut y_hat = vec![0.0; n];
    for (i, row_x) in x.iter().enumerate().take(n) {
        let mut pred = intercept;
        for j in 0..p {
            pred += slopes[j] * row_x[j];
        }
        y_hat[i] = pred;
    }

    let mean_y = y.iter().sum::<f64>() / n as f64;
    let ss_res = y
        .iter()
        .zip(y_hat.iter())
        .map(|(yy, pp)| (yy - pp) * (yy - pp))
        .sum::<f64>();
    let ss_tot = y
        .iter()
        .map(|yy| (yy - mean_y) * (yy - mean_y))
        .sum::<f64>();
    let ss_reg = (ss_tot - ss_res).max(0.0);
    let r2 = if ss_tot <= 1e-15 {
        if ss_res <= 1e-15 {
            1.0
        } else {
            0.0
        }
    } else {
        (1.0 - ss_res / ss_tot).clamp(0.0, 1.0)
    };

    let df = (n as f64 - k as f64).max(0.0);
    let mse = if df > 0.0 { ss_res / df } else { f64::NAN };
    let se_y = if mse.is_finite() {
        mse.sqrt()
    } else {
        f64::NAN
    };

    let xtx_inv = invert_matrix(xtx).ok_or(FormulaValue::Error(CellError::Num))?;
    let mut stderr_slopes = vec![f64::NAN; p];
    for (j, stderr) in stderr_slopes.iter_mut().enumerate().take(p) {
        *stderr = if mse.is_finite() {
            (mse * xtx_inv[j][j]).max(0.0).sqrt()
        } else {
            f64::NAN
        };
    }
    let stderr_intercept = if include_const {
        if mse.is_finite() {
            (mse * xtx_inv[p][p]).max(0.0).sqrt()
        } else {
            f64::NAN
        }
    } else {
        0.0
    };

    if !include_const {
        slopes.shrink_to_fit();
    }

    let dfn = p as f64;
    let f_stat = if dfn > 0.0 && df > 0.0 && ss_res > 0.0 {
        (ss_reg / dfn) / (ss_res / df)
    } else {
        f64::NAN
    };

    Ok(LinearFit {
        slopes,
        intercept,
        stderr_slopes,
        stderr_intercept,
        r2,
        se_y,
        f_stat,
        df,
        ss_reg,
        ss_res,
    })
}

fn as_formula_row(values: &[f64]) -> Vec<FormulaValue> {
    values.iter().map(|v| FormulaValue::Number(*v)).collect()
}

fn reverse_slopes_with_intercept(slopes: &[f64], intercept: f64) -> Vec<f64> {
    let mut row = slopes.to_vec();
    row.reverse();
    row.push(intercept);
    row
}

fn linest_impl(
    known_y: &FormulaValue,
    known_x: Option<&FormulaValue>,
    include_const: bool,
    with_stats: bool,
) -> Result<FormulaValue, FormulaValue> {
    let mut y = Vec::new();
    flatten_numbers(known_y, &mut y)?;
    if y.is_empty() {
        return Err(FormulaValue::Error(CellError::Num));
    }

    let x = parse_known_x(known_x, y.len())?;
    let fit = fit_linear_regression(&y, &x, include_const)?;

    let coef_row = reverse_slopes_with_intercept(&fit.slopes, fit.intercept);
    if !with_stats {
        return Ok(FormulaValue::Array { data: vec![as_formula_row(&coef_row)], source: None });
    }

    let stderr_row = reverse_slopes_with_intercept(&fit.stderr_slopes, fit.stderr_intercept);
    let cols = coef_row.len();
    let mut row3 = vec![0.0; cols];
    let mut row4 = vec![0.0; cols];
    let mut row5 = vec![0.0; cols];
    row3[0] = fit.r2;
    if cols > 1 {
        row3[1] = fit.se_y;
    }
    row4[0] = fit.f_stat;
    if cols > 1 {
        row4[1] = fit.df;
    }
    row5[0] = fit.ss_reg;
    if cols > 1 {
        row5[1] = fit.ss_res;
    }

    Ok(FormulaValue::Array { data: vec![
        as_formula_row(&coef_row),
        as_formula_row(&stderr_row),
        as_formula_row(&row3),
        as_formula_row(&row4),
        as_formula_row(&row5),
    ], source: None })
}

fn logest_impl(
    known_y: &FormulaValue,
    known_x: Option<&FormulaValue>,
    include_const: bool,
    with_stats: bool,
) -> Result<FormulaValue, FormulaValue> {
    let mut y = Vec::new();
    flatten_numbers(known_y, &mut y)?;
    if y.is_empty() || y.iter().any(|v| *v <= 0.0) {
        return Err(FormulaValue::Error(CellError::Num));
    }
    let log_y = FormulaValue::Array { data: y.iter()
        .map(|v| vec![FormulaValue::Number(v.ln())])
        .collect::<Vec<_>>(), source: None };

    let line = linest_impl(&log_y, known_x, include_const, with_stats)?;
    match line {
        FormulaValue::Array { data: mut rows, .. } => {
            if let Some(first) = rows.first_mut() {
                for cell in first {
                    if let FormulaValue::Number(n) = cell {
                        *cell = FormulaValue::Number(n.exp());
                    }
                }
            }
            Ok(FormulaValue::Array { data: rows, source: None })
        }
        _ => Err(FormulaValue::Error(CellError::Value)),
    }
}

fn predict_linear(rows: &[Vec<f64>], slopes: &[f64], intercept: f64) -> Vec<f64> {
    let mut out = Vec::with_capacity(rows.len());
    for row in rows {
        let mut y = intercept;
        for (j, slope) in slopes.iter().enumerate() {
            y += slope * row[j];
        }
        out.push(y);
    }
    out
}

fn predict_log(rows: &[Vec<f64>], slopes: &[f64], intercept: f64) -> Vec<f64> {
    let mut out = Vec::with_capacity(rows.len());
    for row in rows {
        let mut ln_y = intercept;
        for (j, slope) in slopes.iter().enumerate() {
            ln_y += slope * row[j];
        }
        out.push(ln_y.exp());
    }
    out
}

fn mean_step_size(timeline: &[f64]) -> Option<f64> {
    if timeline.len() < 2 {
        return None;
    }
    let mut diffs = Vec::new();
    for pair in timeline.windows(2) {
        let d = pair[1] - pair[0];
        if d > 0.0 {
            diffs.push(d);
        }
    }
    if diffs.is_empty() {
        None
    } else {
        Some(diffs.iter().sum::<f64>() / diffs.len() as f64)
    }
}

fn autocorrelation(values: &[f64], lag: usize) -> f64 {
    let n = values.len();
    if lag == 0 || lag >= n {
        return 0.0;
    }
    let mean = values.iter().sum::<f64>() / n as f64;
    let mut num = 0.0;
    let mut den = 0.0;
    for i in 0..(n - lag) {
        num += (values[i] - mean) * (values[i + lag] - mean);
    }
    for v in values {
        den += (v - mean) * (v - mean);
    }
    if den.abs() < 1e-15 {
        0.0
    } else {
        num / den
    }
}

fn detect_seasonality(values: &[f64], timeline: &[f64]) -> usize {
    if values.len() < 4 || values.len() != timeline.len() {
        return 1;
    }
    let max_lag = (values.len() / 2).min(24);
    let mut best_lag = 1usize;
    let mut best_corr = 0.0;
    for lag in 2..=max_lag {
        let corr = autocorrelation(values, lag);
        if corr > best_corr {
            best_corr = corr;
            best_lag = lag;
        }
    }
    if best_corr > 0.3 {
        best_lag
    } else {
        1
    }
}

#[derive(Clone)]
struct HoltResult {
    alpha: f64,
    beta: f64,
    level: f64,
    trend: f64,
    fitted: Vec<f64>,
    mae: f64,
    rmse: f64,
}

fn holt_with_params(values: &[f64], alpha: f64, beta: f64) -> HoltResult {
    let n = values.len();
    let mut fitted = vec![values[0]; n];
    let mut level = values[0];
    let mut trend = if n >= 2 { values[1] - values[0] } else { 0.0 };

    for (t, value) in values.iter().enumerate().skip(1) {
        fitted[t] = level + trend;
        let new_level = alpha * *value + (1.0 - alpha) * (level + trend);
        let new_trend = beta * (new_level - level) + (1.0 - beta) * trend;
        level = new_level;
        trend = new_trend;
    }

    let mut sae = 0.0;
    let mut sse = 0.0;
    let mut count = 0usize;
    for t in 1..n {
        let e = values[t] - fitted[t];
        sae += e.abs();
        sse += e * e;
        count += 1;
    }
    let mae = if count > 0 { sae / count as f64 } else { 0.0 };
    let rmse = if count > 0 {
        (sse / count as f64).sqrt()
    } else {
        0.0
    };

    HoltResult {
        alpha,
        beta,
        level,
        trend,
        fitted,
        mae,
        rmse,
    }
}

fn fit_holt(values: &[f64]) -> HoltResult {
    // Nested bisection optimization over (alpha, beta) to minimize MSE.
    const RESOLUTION: f64 = 0.001;

    let optimize_beta = |alpha: f64| -> f64 {
        let mut lo = 0.0_f64;
        let mut hi = 1.0_f64;
        let mut e_lo = holt_with_params(values, alpha, lo).rmse;
        let mut e_hi = holt_with_params(values, alpha, hi).rmse;
        while (hi - lo) > RESOLUTION {
            let mid = (lo + hi) / 2.0;
            if e_hi > e_lo { hi = mid; e_hi = holt_with_params(values, alpha, hi).rmse; }
            else { lo = mid; e_lo = holt_with_params(values, alpha, lo).rmse; }
        }
        (lo + hi) / 2.0
    };

    let mut lo = 0.0_f64;
    let mut hi = 1.0_f64;
    let mut e_lo = holt_with_params(values, lo, optimize_beta(lo)).rmse;
    let mut e_hi = holt_with_params(values, hi, optimize_beta(hi)).rmse;
    while (hi - lo) > RESOLUTION {
        let mid = (lo + hi) / 2.0;
        if e_hi > e_lo { hi = mid; e_hi = holt_with_params(values, hi, optimize_beta(hi)).rmse; }
        else { lo = mid; e_lo = holt_with_params(values, lo, optimize_beta(lo)).rmse; }
    }
    let best_alpha = (lo + hi) / 2.0;
    holt_with_params(values, best_alpha, optimize_beta(best_alpha))
}

fn parse_values_timeline(
    values_arg: Option<&FormulaValue>,
    timeline_arg: Option<&FormulaValue>,
) -> Result<(Vec<f64>, Vec<f64>), FormulaValue> {
    let mut values = Vec::new();
    let mut timeline = Vec::new();
    let Some(v) = values_arg else {
        return Err(FormulaValue::Error(CellError::Value));
    };
    let Some(t) = timeline_arg else {
        return Err(FormulaValue::Error(CellError::Value));
    };

    flatten_numbers(v, &mut values)?;
    flatten_numbers(t, &mut timeline)?;

    if values.len() != timeline.len() || values.len() < 2 {
        return Err(FormulaValue::Error(CellError::Value));
    }
    Ok((values, timeline))
}

pub fn fn_lognorm_dist(
    args: &[FormulaValue],
    _ctx: &EvaluationContext,
) -> FormulaResult<FormulaValue> {
    let x = match required_number(args, 0) {
        Ok(v) => v,
        Err(e) => return Ok(e),
    };
    let mean = match required_number(args, 1) {
        Ok(v) => v,
        Err(e) => return Ok(e),
    };
    let sigma = match required_number(args, 2) {
        Ok(v) => v,
        Err(e) => return Ok(e),
    };
    let cumulative = match required_bool(args, 3) {
        Ok(v) => v,
        Err(e) => return Ok(e),
    };

    if sigma <= 0.0 {
        return Ok(FormulaValue::Error(CellError::Num));
    }
    if x <= 0.0 {
        return Ok(FormulaValue::Number(0.0));
    }

    let z = (x.ln() - mean) / sigma;
    let result = if cumulative {
        normal_cdf(z)
    } else {
        normal_pdf(z) / (x * sigma)
    };
    Ok(FormulaValue::Number(result))
}

pub fn fn_lognorm_inv(
    args: &[FormulaValue],
    _ctx: &EvaluationContext,
) -> FormulaResult<FormulaValue> {
    let p = match required_number(args, 0) {
        Ok(v) => v,
        Err(e) => return Ok(e),
    };
    let mean = match required_number(args, 1) {
        Ok(v) => v,
        Err(e) => return Ok(e),
    };
    let sigma = match required_number(args, 2) {
        Ok(v) => v,
        Err(e) => return Ok(e),
    };

    if sigma <= 0.0 || !(0.0 < p && p < 1.0) {
        return Ok(FormulaValue::Error(CellError::Num));
    }
    Ok(FormulaValue::Number(
        (mean + sigma * inverse_normal(p)).exp(),
    ))
}

pub fn fn_linest(args: &[FormulaValue], _ctx: &EvaluationContext) -> FormulaResult<FormulaValue> {
    let known_y = match args.first() {
        Some(v) => v,
        None => return Ok(FormulaValue::Error(CellError::Value)),
    };
    let known_x = args.get(1);
    let include_const = match optional_bool(args, 2, true) {
        Ok(v) => v,
        Err(e) => return Ok(e),
    };
    let with_stats = match optional_bool(args, 3, false) {
        Ok(v) => v,
        Err(e) => return Ok(e),
    };

    match linest_impl(known_y, known_x, include_const, with_stats) {
        Ok(v) => Ok(v),
        Err(e) => Ok(e),
    }
}

pub fn fn_logest(args: &[FormulaValue], _ctx: &EvaluationContext) -> FormulaResult<FormulaValue> {
    let known_y = match args.first() {
        Some(v) => v,
        None => return Ok(FormulaValue::Error(CellError::Value)),
    };
    let known_x = args.get(1);
    let include_const = match optional_bool(args, 2, true) {
        Ok(v) => v,
        Err(e) => return Ok(e),
    };
    let with_stats = match optional_bool(args, 3, false) {
        Ok(v) => v,
        Err(e) => return Ok(e),
    };

    match logest_impl(known_y, known_x, include_const, with_stats) {
        Ok(v) => Ok(v),
        Err(e) => Ok(e),
    }
}

pub fn fn_growth(args: &[FormulaValue], _ctx: &EvaluationContext) -> FormulaResult<FormulaValue> {
    let known_y = match args.first() {
        Some(v) => v,
        None => return Ok(FormulaValue::Error(CellError::Value)),
    };
    let known_x_arg = args.get(1);
    let mut y = Vec::new();
    if let Err(e) = flatten_numbers(known_y, &mut y) {
        return Ok(e);
    }
    if y.is_empty() || y.iter().any(|v| *v <= 0.0) {
        return Ok(FormulaValue::Error(CellError::Num));
    }
    let known_x = match parse_known_x(known_x_arg, y.len()) {
        Ok(v) => v,
        Err(e) => return Ok(e),
    };

    let new_x = match parse_new_x(args.get(2), &known_x, known_x[0].len()) {
        Ok(v) => v,
        Err(e) => return Ok(e),
    };
    let include_const = match optional_bool(args, 3, true) {
        Ok(v) => v,
        Err(e) => return Ok(e),
    };

    let ln_y = y.iter().map(|v| v.ln()).collect::<Vec<_>>();
    let fit = match fit_linear_regression(&ln_y, &known_x, include_const) {
        Ok(v) => v,
        Err(e) => return Ok(e),
    };
    let preds = predict_log(&new_x, &fit.slopes, fit.intercept);
    let arr = preds
        .into_iter()
        .map(|v| vec![FormulaValue::Number(v)])
        .collect::<Vec<_>>();
    Ok(FormulaValue::Array { data: arr, source: None })
}

pub fn fn_trend(args: &[FormulaValue], _ctx: &EvaluationContext) -> FormulaResult<FormulaValue> {
    let known_y = match args.first() {
        Some(v) => v,
        None => return Ok(FormulaValue::Error(CellError::Value)),
    };
    let known_x_arg = args.get(1);
    let mut y = Vec::new();
    if let Err(e) = flatten_numbers(known_y, &mut y) {
        return Ok(e);
    }
    if y.is_empty() {
        return Ok(FormulaValue::Error(CellError::Num));
    }
    let known_x = match parse_known_x(known_x_arg, y.len()) {
        Ok(v) => v,
        Err(e) => return Ok(e),
    };
    let new_x = match parse_new_x(args.get(2), &known_x, known_x[0].len()) {
        Ok(v) => v,
        Err(e) => return Ok(e),
    };
    let include_const = match optional_bool(args, 3, true) {
        Ok(v) => v,
        Err(e) => return Ok(e),
    };

    let fit = match fit_linear_regression(&y, &known_x, include_const) {
        Ok(v) => v,
        Err(e) => return Ok(e),
    };
    let preds = predict_linear(&new_x, &fit.slopes, fit.intercept);
    let arr = preds
        .into_iter()
        .map(|v| vec![FormulaValue::Number(v)])
        .collect::<Vec<_>>();
    Ok(FormulaValue::Array { data: arr, source: None })
}

pub fn fn_forecast_ets(
    args: &[FormulaValue],
    _ctx: &EvaluationContext,
) -> FormulaResult<FormulaValue> {
    let target = match required_number(args, 0) {
        Ok(v) => v,
        Err(e) => return Ok(e),
    };
    let (values, timeline) = match parse_values_timeline(args.get(1), args.get(2)) {
        Ok(v) => v,
        Err(e) => return Ok(e),
    };

    let _seasonality = match optional_number(args, 3, 1.0) {
        Ok(v) => v,
        Err(e) => return Ok(e),
    };
    let _data_completion = match optional_number(args, 4, 1.0) {
        Ok(v) => v,
        Err(e) => return Ok(e),
    };
    let _aggregation = match optional_number(args, 5, 0.0) {
        Ok(v) => v,
        Err(e) => return Ok(e),
    };

    let step = mean_step_size(&timeline).unwrap_or(1.0);
    let horizon = ((target - timeline[timeline.len() - 1]) / step).max(0.0);
    let model = fit_holt(&values);
    Ok(FormulaValue::Number(model.level + horizon * model.trend))
}

pub fn fn_forecast_ets_confint(
    args: &[FormulaValue],
    _ctx: &EvaluationContext,
) -> FormulaResult<FormulaValue> {
    let target = match required_number(args, 0) {
        Ok(v) => v,
        Err(e) => return Ok(e),
    };
    let (values, timeline) = match parse_values_timeline(args.get(1), args.get(2)) {
        Ok(v) => v,
        Err(e) => return Ok(e),
    };

    let confidence = match optional_number(args, 3, 0.95) {
        Ok(v) => v,
        Err(e) => return Ok(e),
    };
    if !(0.0 < confidence && confidence < 1.0) {
        return Ok(FormulaValue::Error(CellError::Num));
    }

    let _seasonality = match optional_number(args, 4, 1.0) {
        Ok(v) => v,
        Err(e) => return Ok(e),
    };
    let _data_completion = match optional_number(args, 5, 1.0) {
        Ok(v) => v,
        Err(e) => return Ok(e),
    };
    let _aggregation = match optional_number(args, 6, 0.0) {
        Ok(v) => v,
        Err(e) => return Ok(e),
    };

    let step = mean_step_size(&timeline).unwrap_or(1.0);
    let horizon = ((target - timeline[timeline.len() - 1]) / step).max(1.0);
    let model = fit_holt(&values);
    let z = inverse_normal(0.5 + confidence / 2.0);
    let stderr = model.rmse * horizon.sqrt();
    Ok(FormulaValue::Number(z * stderr))
}

pub fn fn_forecast_ets_seasonality(
    args: &[FormulaValue],
    _ctx: &EvaluationContext,
) -> FormulaResult<FormulaValue> {
    let (values, timeline) = match parse_values_timeline(args.first(), args.get(1)) {
        Ok(v) => v,
        Err(e) => return Ok(e),
    };
    let _data_completion = match optional_number(args, 2, 1.0) {
        Ok(v) => v,
        Err(e) => return Ok(e),
    };
    let _aggregation = match optional_number(args, 3, 0.0) {
        Ok(v) => v,
        Err(e) => return Ok(e),
    };

    Ok(FormulaValue::Number(
        detect_seasonality(&values, &timeline) as f64
    ))
}

pub fn fn_forecast_ets_stat(
    args: &[FormulaValue],
    _ctx: &EvaluationContext,
) -> FormulaResult<FormulaValue> {
    let (values, timeline) = match parse_values_timeline(args.first(), args.get(1)) {
        Ok(v) => v,
        Err(e) => return Ok(e),
    };
    let _seasonality = match optional_number(args, 2, 1.0) {
        Ok(v) => v,
        Err(e) => return Ok(e),
    };
    let _data_completion = match optional_number(args, 3, 1.0) {
        Ok(v) => v,
        Err(e) => return Ok(e),
    };
    let _aggregation = match optional_number(args, 4, 0.0) {
        Ok(v) => v,
        Err(e) => return Ok(e),
    };

    let model = fit_holt(&values);
    let mae = model.mae;
    let rmse = model.rmse;
    let mut smape_sum = 0.0;
    let mut count = 0usize;
    for (actual, fitted) in values.iter().zip(model.fitted.iter()).skip(1) {
        let denom = actual.abs() + fitted.abs();
        if denom > 1e-12 {
            smape_sum += (2.0 * (actual - fitted).abs()) / denom;
            count += 1;
        }
    }
    let smape = if count > 0 {
        smape_sum / count as f64
    } else {
        0.0
    };

    let scale = if values.len() > 1 {
        let mut d = 0.0;
        for pair in values.windows(2) {
            d += (pair[1] - pair[0]).abs();
        }
        d / (values.len() - 1) as f64
    } else {
        0.0
    };
    let mase = if scale > 1e-12 { mae / scale } else { 0.0 };
    let step_size = mean_step_size(&timeline).unwrap_or(1.0);

    let row = vec![
        FormulaValue::Number(model.alpha),
        FormulaValue::Number(model.beta),
        FormulaValue::Number(0.0),
        FormulaValue::Number(mase),
        FormulaValue::Number(smape),
        FormulaValue::Number(mae),
        FormulaValue::Number(rmse),
        FormulaValue::Number(step_size),
    ];
    Ok(FormulaValue::Array { data: vec![row], source: None })
}

#[cfg(test)]
mod tests {
    use super::*;

    fn eval(formula: &str) -> FormulaResult<FormulaValue> {
        let ast = crate::parser::parse_formula(formula)?;
        crate::evaluator::evaluate(&ast, &EvaluationContext::simple())
    }

    fn as_number(v: FormulaValue) -> f64 {
        match v {
            FormulaValue::Number(n) => n,
            _ => panic!("Expected Number"),
        }
    }

    #[test]
    fn test_lognorm_dist() {
        let ctx = EvaluationContext::simple();
        let args = vec![
            eval("=1").unwrap(),
            eval("=0").unwrap(),
            eval("=1").unwrap(),
            eval("=TRUE").unwrap(),
        ];
        let result = fn_lognorm_dist(&args, &ctx).unwrap();
        let n = as_number(result);
        assert!((n - 0.5).abs() < 1e-6);
    }

    #[test]
    fn test_lognorm_inv() {
        let ctx = EvaluationContext::simple();
        let args = vec![
            eval("=0.5").unwrap(),
            eval("=0").unwrap(),
            eval("=1").unwrap(),
        ];
        let result = fn_lognorm_inv(&args, &ctx).unwrap();
        let n = as_number(result);
        assert!((n - 1.0).abs() < 1e-7);
    }

    #[test]
    fn test_linest() {
        let ctx = EvaluationContext::simple();
        let args = vec![eval("={1,2,3}").unwrap(), eval("={1,2,3}").unwrap()];
        let result = fn_linest(&args, &ctx).unwrap();
        match result {
            FormulaValue::Array { data: rows, .. } => {
                assert_eq!(rows.len(), 1);
                if let FormulaValue::Number(slope) = rows[0][0] {
                    assert!((slope - 1.0).abs() < 1e-8);
                } else {
                    panic!("Expected slope number");
                }
                if let FormulaValue::Number(intercept) = rows[0][1] {
                    assert!(intercept.abs() < 1e-8);
                } else {
                    panic!("Expected intercept number");
                }
            }
            _ => panic!("Expected Array"),
        }
    }

    #[test]
    fn test_logest() {
        let ctx = EvaluationContext::simple();
        let args = vec![eval("={2,4,8}").unwrap(), eval("={1,2,3}").unwrap()];
        let result = fn_logest(&args, &ctx).unwrap();
        match result {
            FormulaValue::Array { data: rows, .. } => {
                if let FormulaValue::Number(m) = rows[0][0] {
                    assert!((m - 2.0).abs() < 1e-6);
                } else {
                    panic!("Expected m number");
                }
                if let FormulaValue::Number(b) = rows[0][1] {
                    assert!((b - 1.0).abs() < 1e-6);
                } else {
                    panic!("Expected b number");
                }
            }
            _ => panic!("Expected Array"),
        }
    }

    #[test]
    fn test_growth() {
        let ctx = EvaluationContext::simple();
        let args = vec![
            eval("={2,4,8}").unwrap(),
            eval("={1,2,3}").unwrap(),
            eval("={4,5}").unwrap(),
        ];
        let result = fn_growth(&args, &ctx).unwrap();
        match result {
            FormulaValue::Array { data: rows, .. } => {
                assert_eq!(rows.len(), 2);
                if let FormulaValue::Number(v1) = rows[0][0] {
                    assert!((v1 - 16.0).abs() < 1e-4);
                } else {
                    panic!("Expected first prediction");
                }
                if let FormulaValue::Number(v2) = rows[1][0] {
                    assert!((v2 - 32.0).abs() < 1e-3);
                } else {
                    panic!("Expected second prediction");
                }
            }
            _ => panic!("Expected Array"),
        }
    }

    #[test]
    fn test_trend() {
        let ctx = EvaluationContext::simple();
        let args = vec![
            eval("={3,5,7}").unwrap(),
            eval("={1,2,3}").unwrap(),
            eval("={4}").unwrap(),
        ];
        let result = fn_trend(&args, &ctx).unwrap();
        match result {
            FormulaValue::Array { data: rows, .. } => {
                if let FormulaValue::Number(v) = rows[0][0] {
                    assert!((v - 9.0).abs() < 1e-8);
                } else {
                    panic!("Expected number prediction");
                }
            }
            _ => panic!("Expected Array"),
        }
    }

    #[test]
    fn test_forecast_ets() {
        let ctx = EvaluationContext::simple();
        let args = vec![
            eval("=4").unwrap(),
            eval("={10,12,14}").unwrap(),
            eval("={1,2,3}").unwrap(),
        ];
        let result = fn_forecast_ets(&args, &ctx).unwrap();
        let n = as_number(result);
        assert!(n > 15.0 && n < 17.5);
    }

    #[test]
    fn test_forecast_ets_confint() {
        let ctx = EvaluationContext::simple();
        let args = vec![
            eval("=5").unwrap(),
            eval("={10,12,14,16}").unwrap(),
            eval("={1,2,3,4}").unwrap(),
            eval("=0.95").unwrap(),
        ];
        let result = fn_forecast_ets_confint(&args, &ctx).unwrap();
        let n = as_number(result);
        assert!(n >= 0.0);
    }

    #[test]
    fn test_forecast_ets_seasonality() {
        let ctx = EvaluationContext::simple();
        let args = vec![
            eval("={1,2,1,2,1,2}").unwrap(),
            eval("={1,2,3,4,5,6}").unwrap(),
        ];
        let result = fn_forecast_ets_seasonality(&args, &ctx).unwrap();
        let n = as_number(result);
        assert!((n - 2.0).abs() < 1e-8);
    }

    #[test]
    fn test_forecast_ets_stat() {
        let ctx = EvaluationContext::simple();
        let args = vec![eval("={10,12,14,16}").unwrap(), eval("={1,2,3,4}").unwrap()];
        let result = fn_forecast_ets_stat(&args, &ctx).unwrap();
        match result {
            FormulaValue::Array { data: rows, .. } => {
                assert_eq!(rows.len(), 1);
                assert_eq!(rows[0].len(), 8);
                if let FormulaValue::Number(alpha) = rows[0][0] {
                    assert!(alpha > 0.0 && alpha <= 1.0);
                } else {
                    panic!("Expected alpha number");
                }
            }
            _ => panic!("Expected Array"),
        }
    }

    // ===== DOCS-BASED TESTS =====

    #[test]
    fn test_lognorm_dist_docs() {
        // Docs: x=4, mean=3.5, standard_dev=1.2
        // =LOGNORM.DIST(4,3.5,1.2,TRUE) = 0.0390836
        let result = eval("=LOGNORM.DIST(4,3.5,1.2,TRUE)").unwrap();
        if let FormulaValue::Number(n) = result {
            assert!((n - 0.0390836).abs() < 1e-6);
        } else {
            panic!("Expected Number");
        }
        // =LOGNORM.DIST(4,3.5,1.2,FALSE) = 0.0176176
        let result = eval("=LOGNORM.DIST(4,3.5,1.2,FALSE)").unwrap();
        if let FormulaValue::Number(n) = result {
            assert!((n - 0.0176176).abs() < 1e-6);
        } else {
            panic!("Expected Number");
        }
    }

    #[test]
    fn test_lognorm_inv_docs() {
        // Docs: probability=0.039084, mean=3.5, standard_dev=1.2
        // =LOGNORM.INV(0.039084,3.5,1.2) = 4.0000252
        let result = eval("=LOGNORM.INV(0.039084,3.5,1.2)").unwrap();
        if let FormulaValue::Number(n) = result {
            assert!((n - 4.0000252).abs() < 1e-4);
        } else {
            panic!("Expected Number");
        }
    }

    #[test]
    fn test_linest_docs() {
        // --- Example 1: slope and y-intercept ---
        // known_y={1,9,5,7}, known_x={0,4,2,3} -> slope=2, intercept=1
        let result = eval("=LINEST({1,9,5,7},{0,4,2,3})").unwrap();
        match result {
            FormulaValue::Array { data: rows, .. } => {
                assert_eq!(rows.len(), 1);
                if let FormulaValue::Number(slope) = rows[0][0] {
                    assert!((slope - 2.0).abs() < 1e-8);
                } else {
                    panic!("Expected slope number");
                }
                if let FormulaValue::Number(intercept) = rows[0][1] {
                    assert!((intercept - 1.0).abs() < 1e-8);
                } else {
                    panic!("Expected intercept number");
                }
            }
            _ => panic!("Expected Array"),
        }

        // --- Example 2: simple linear regression ---
        // months 1-6, sales {3100,4500,4400,5400,7500,8100}
        // Docs: =SUM(LINEST(B1:B6,A1:A6)*{9,1}) = $11,000
        let result = eval("=LINEST({3100,4500,4400,5400,7500,8100},{1,2,3,4,5,6})").unwrap();
        match result {
            FormulaValue::Array { data: rows, .. } => {
                if let (FormulaValue::Number(slope), FormulaValue::Number(intercept)) =
                    (&rows[0][0], &rows[0][1])
                {
                    let predicted = slope * 9.0 + intercept;
                    assert!((predicted - 11000.0).abs() < 1.0);
                } else {
                    panic!("Expected slope and intercept numbers");
                }
            }
            _ => panic!("Expected Array"),
        }

        // --- Example 3: multiple linear regression with stats ---
        // Docs first-column values: m4=-234.2371645, se4=13.26801148,
        // r^2=0.996747993, F=459.7536742, ssreg=1732393319
        let result = eval(concat!(
            "=LINEST({142000;144000;151000;150000;139000;169000;126000;142900;163000;169000;149000},",
            "{2310,2,2,20;2333,2,2,12;2356,3,1.5,33;2379,3,2,43;2402,2,3,53;",
            "2425,4,2,23;2448,2,1.5,99;2471,2,2,34;2494,3,3,23;2517,4,4,55;2540,2,3,22},",
            "TRUE,TRUE)"
        ))
        .unwrap();
        match result {
            FormulaValue::Array { data: rows, .. } => {
                assert_eq!(rows.len(), 5, "stats=TRUE should return 5 rows");
                // m4 (age coefficient)
                if let FormulaValue::Number(m4) = rows[0][0] {
                    assert!((m4 - (-234.2371645)).abs() < 0.001);
                } else {
                    panic!("Expected m4 number");
                }
                // se4
                if let FormulaValue::Number(se4) = rows[1][0] {
                    assert!((se4 - 13.26801148).abs() < 0.001);
                } else {
                    panic!("Expected se4 number");
                }
                // r^2
                if let FormulaValue::Number(r2) = rows[2][0] {
                    assert!((r2 - 0.996747993).abs() < 1e-6);
                } else {
                    panic!("Expected r^2 number");
                }
                // F statistic
                if let FormulaValue::Number(f) = rows[3][0] {
                    assert!((f - 459.7536742).abs() < 0.001);
                } else {
                    panic!("Expected F number");
                }
                // ssreg
                if let FormulaValue::Number(ssreg) = rows[4][0] {
                    assert!((ssreg - 1732393319.0).abs() < 1.0);
                } else {
                    panic!("Expected ssreg number");
                }
            }
            _ => panic!("Expected Array"),
        }
    }

    #[test]
    fn test_logest_docs() {
        // Docs: single result (m) = 1.4633
        // Data from GROWTH example: y={33100,...,220000}, x={11,...,16}
        let result =
            eval("=LOGEST({33100,47300,69000,102000,150000,220000},{11,12,13,14,15,16})").unwrap();
        match result {
            FormulaValue::Array { data: rows, .. } => {
                assert_eq!(rows.len(), 1);
                if let FormulaValue::Number(m) = rows[0][0] {
                    assert!((m - 1.4633).abs() < 0.001);
                } else {
                    panic!("Expected m number");
                }
            }
            _ => panic!("Expected Array"),
        }
    }

    #[test]
    fn test_growth_docs() {
        // --- Fitting example ---
        // Docs: GROWTH(B2:B7,A2:A7) fitted values
        // Month={11,...,16}, Units={33100,...,220000}
        // Expected: {32618, 47729, 69841, 102197, 149542, 218822}
        let result =
            eval("=GROWTH({33100,47300,69000,102000,150000,220000},{11,12,13,14,15,16})").unwrap();
        let expected_fit = [32618.0, 47729.0, 69841.0, 102197.0, 149542.0, 218822.0];
        match result {
            FormulaValue::Array { data: rows, .. } => {
                assert_eq!(rows.len(), expected_fit.len());
                for (i, &exp) in expected_fit.iter().enumerate() {
                    if let FormulaValue::Number(v) = rows[i][0] {
                        assert!((v - exp).abs() < 1.0, "fit[{i}]: got {v}, expected {exp}");
                    } else {
                        panic!("Expected number at index {i}");
                    }
                }
            }
            _ => panic!("Expected Array"),
        }

        // --- Prediction example ---
        // Docs: GROWTH(B2:B7,A2:A7,A9:A10) for months 17, 18
        // Expected: {320197, 468536}
        let result =
            eval("=GROWTH({33100,47300,69000,102000,150000,220000},{11,12,13,14,15,16},{17,18})")
                .unwrap();
        let expected_pred = [320197.0, 468536.0];
        match result {
            FormulaValue::Array { data: rows, .. } => {
                assert_eq!(rows.len(), expected_pred.len());
                for (i, &exp) in expected_pred.iter().enumerate() {
                    if let FormulaValue::Number(v) = rows[i][0] {
                        assert!((v - exp).abs() < 1.0, "pred[{i}]: got {v}, expected {exp}");
                    } else {
                        panic!("Expected number at index {i}");
                    }
                }
            }
            _ => panic!("Expected Array"),
        }
    }

    #[test]
    fn test_trend_docs() {
        // From LINEST Example 2: months 1-6, sales {3100,4500,4400,5400,7500,8100}
        // Docs: SUM(LINEST(...)*{9,1}) = $11,000 — TREND for month 9 must match
        let result = eval("=TREND({3100,4500,4400,5400,7500,8100},{1,2,3,4,5,6},{9})").unwrap();
        match result {
            FormulaValue::Array { data: rows, .. } => {
                if let FormulaValue::Number(v) = rows[0][0] {
                    assert!((v - 11000.0).abs() < 1.0);
                } else {
                    panic!("Expected number prediction");
                }
            }
            _ => panic!("Expected Array"),
        }
    }
}
