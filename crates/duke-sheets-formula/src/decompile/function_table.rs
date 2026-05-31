//! Function metadata accessors keyed by BIFF8 function table (Ftab) index.
//!
//! The actual function definitions live on [`crate::functions::FunctionDef`]
//! and are registered into the global [`crate::functions::FunctionRegistry`]
//! singleton. This module exposes the iftab-keyed lookups used by the
//! XLS (BIFF8) and XLSB (BIFF12) writer and decompiler paths.
//!
//! Argument count encoding (MS-XLS [MS-XLS] §2.5.198.63 Ftab):
//! - 0..253: fixed argument count
//! - 254: variable, minimum from this value
//! - 255: variable, 0 or more

use crate::functions::registry;

/// PTG operand class for function arguments and reference tokens.
///
/// See [MS-XLS] §2.5.198 "Tokens (Ptg)" for the V/R/A class distinction.
/// The class is encoded in the low bits of certain PTG opcodes (e.g.
/// PtgRef 0x24=R, 0x44=V, 0x64=A; PtgArea 0x25=R, 0x45=V, 0x65=A).
///
/// Only the R and V classes are used by the current writer paths.
/// Array class (A) is reserved for array-formula encoding.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum OperandClass {
    /// Reference class — operand must be a cell/range/name reference.
    R,
    /// Value class — operand is coerced to a value before use.
    V,
}

/// Look up a function's BIFF8 iftab index by name (case-insensitive).
///
/// Returns `None` for unknown names and for newer functions that lack an
/// iftab (XLOOKUP, RANDARRAY, dynamic-array functions, etc.).
pub fn function_index(name: &str) -> Option<u16> {
    registry().get(name).and_then(|def| def.iftab)
}

/// Look up a function name by BIFF8 iftab index.
///
/// Returns the function name, or an empty string for unknown/reserved indices.
pub fn function_name(idx: u16) -> &'static str {
    registry()
        .get_by_iftab(idx)
        .map(|def| def.name)
        .unwrap_or("")
}

/// Look up the declared argument count for a BIFF8 function.
///
/// Returns the argc value (see encoding notes at top).
/// Returns 255 (variable) for unknown indices.
pub fn function_argc(idx: u16) -> u16 {
    registry()
        .get_by_iftab(idx)
        .map(|def| def.declared_argc)
        .unwrap_or(255)
}

/// Minimum number of arguments function `iftab` accepts, or `None` for an
/// unknown index. Useful for synthesizing a minimal valid call.
pub fn function_min_args(idx: u16) -> Option<usize> {
    registry().get_by_iftab(idx).map(|def| def.min_args)
}

/// Maximum number of arguments function `iftab` accepts (`None` = unbounded or
/// unknown index).
pub fn function_max_args(idx: u16) -> Option<usize> {
    registry().get_by_iftab(idx).and_then(|def| def.max_args)
}

/// True if function `iftab` accepts `actual_argc` as a fixed-arity call,
/// suitable for PtgFunc (0x41) emission. Otherwise the writer must emit
/// PtgFuncVar (0x42) with an explicit argument count.
///
/// MS-XLS [MS-XLS] §2.5.198.62 / §2.5.198.63 defines PtgFunc and PtgFuncVar:
/// PtgFunc carries no argc byte and assumes the Ftab declares a fixed count
/// matching `actual_argc`. The per-function [`FunctionDef::fixed_arity`]
/// allow-list is grown empirically from Excel-authored byte-parity tests.
///
/// Analysis-ToolPak add-in functions (Ftab 384..=476) follow a uniform rule:
/// those with a single valid arity (`min_args == max_args`) are emitted by
/// Excel as PtgFunc, those with an argument range as PtgFuncVar. This was
/// confirmed across the whole range by the comprehensive XLSB ATP parity test
/// (`excel_byte_parity_for_all_xlsb_atp_functions_we_emit`), so the rule is
/// derived from min/max rather than a per-function `fixed_arity` flag.
///
/// [`FunctionDef::fixed_arity`]: crate::functions::FunctionDef::fixed_arity
pub fn function_is_fixed_arity(iftab: u16, actual_argc: usize) -> bool {
    let Some(def) = registry().get_by_iftab(iftab) else {
        return false;
    };
    if function_is_biff8_addin(iftab) {
        return def.max_args == Some(def.min_args) && actual_argc == def.min_args;
    }
    def.fixed_arity && def.declared_argc as usize == actual_argc
}

/// PTG operand class Excel expects for the `arg_idx`-th argument of the
/// function whose Ftab index is `iftab`. Defaults to V for unrecognized
/// functions.
///
/// Reads [`FunctionDef::arg_classes`] (per-position overrides) first, then
/// falls back to [`FunctionDef::default_arg_class`].
///
/// [`FunctionDef::arg_classes`]: crate::functions::FunctionDef::arg_classes
/// [`FunctionDef::default_arg_class`]: crate::functions::FunctionDef::default_arg_class
pub fn function_arg_class(iftab: u16, arg_idx: usize) -> OperandClass {
    let Some(def) = registry().get_by_iftab(iftab) else {
        return OperandClass::V;
    };
    def.arg_classes
        .get(arg_idx)
        .copied()
        .unwrap_or(def.default_arg_class)
}

/// True if BIFF8 emits function `iftab` via the Analysis-ToolPak add-in
/// mechanism — a PtgNameX referencing an EXTERNNAME in an AddIn SUPBOOK,
/// followed by PtgFuncVar with iftab=255 (UDF) — rather than a native
/// PtgFunc/PtgFuncVar carrying the Ftab index.
///
/// This is exactly the contiguous Ftab range `384..=476` (HEX2BIN through
/// FVSCHEDULE: the engineering, complex-number, financial, and date ATP
/// functions). Determined by authoring every Ftab function in Excel and
/// classifying its emitted token stream (comprehensive, not sampled): the
/// range is contiguous with no exceptions. Functions 477..=484 (the CUBE
/// functions and the Excel-2007 IFERROR/COUNTIFS/SUMIFS/AVERAGEIF/AVERAGEIFS
/// block) and everything `<=383` are native.
///
/// Cross-checks with Apache POI's `AnalysisToolPak`: POI's list equals this
/// range plus the 2007-native functions, which POI groups with ATP for
/// *evaluation* but Excel serializes natively — so the over-inclusion is
/// expected and this range is the authoritative BIFF8 *serialization* set.
pub fn function_is_biff8_addin(iftab: u16) -> bool {
    (384..=476).contains(&iftab)
}

/// Operand class for the top-level expression of a defined-name / built-in
/// name body (the `refers_to` formula).
///
/// A name body that is a bare reference or range must be reference-class: a
/// value-class range makes Excel apply implicit intersection when the name is
/// used, collapsing e.g. `Data!A1:A3` to a single cell — so `SUM(Numbers)`
/// would sum one cell instead of the range. The reference operators
/// (range/union/intersect) are likewise reference-class; constants and
/// arithmetic are value-class. Shared by the XLS (BIFF8) and XLSB (BIFF12)
/// writers so the rule can't drift between them. MS-XLS [MS-XLS] §2.5.198.103.
pub fn name_body_operand_class(expr: &crate::FormulaExpr) -> OperandClass {
    use crate::ast::BinaryOperator;
    use crate::FormulaExpr;
    match expr {
        FormulaExpr::CellRef(_)
        | FormulaExpr::RangeRef(_)
        | FormulaExpr::NameRef(_)
        | FormulaExpr::BinaryOp {
            op: BinaryOperator::Range | BinaryOperator::Union | BinaryOperator::Intersect,
            ..
        } => OperandClass::R,
        _ => OperandClass::V,
    }
}

/// True if `iftab` names a reference-class function — one that can return a
/// reference and therefore takes the operand class of the position it occupies
/// (R when used as a reference argument, V otherwise). Pure value functions
/// return `false` and are always emitted V-class.
///
/// MS-XLS [MS-XLS] §2.5.198.103. Drives the function-token class bits in the
/// XLS writer for nested calls.
pub fn function_returns_reference(iftab: u16) -> bool {
    registry()
        .get_by_iftab(iftab)
        .is_some_and(|def| def.returns_reference)
}

/// True if `iftab` names a volatile function — one whose result depends on
/// state outside its direct operands so Excel must re-evaluate the formula
/// on every workbook change.
///
/// MS-XLS [MS-XLS] §2.5.198.42 PtgAttrVolatile is emitted as the first token
/// of a formula whose AST calls a volatile function transitively. Use
/// [`expr_calls_volatile_function`] for the AST walk.
pub fn function_is_volatile(iftab: u16) -> bool {
    registry()
        .get_by_iftab(iftab)
        .is_some_and(|def| def.volatile)
}

/// True if `expr` calls a volatile function transitively. Walks the AST to
/// find any [`crate::FormulaExpr::Function`] node whose `name` resolves to a
/// volatile function in the registry.
///
/// Drives PtgAttrVolatile prefix emission in the XLS/XLSB writers.
pub fn expr_calls_volatile_function(expr: &crate::FormulaExpr) -> bool {
    use crate::FormulaExpr;

    match expr {
        FormulaExpr::Function { name, args } => {
            if let Some(def) = registry().get(name) {
                if def.volatile {
                    return true;
                }
            }
            args.iter().any(expr_calls_volatile_function)
        }
        FormulaExpr::BinaryOp { left, right, .. } => {
            expr_calls_volatile_function(left) || expr_calls_volatile_function(right)
        }
        FormulaExpr::UnaryOp { operand, .. } => expr_calls_volatile_function(operand),
        FormulaExpr::Array(rows) => rows
            .iter()
            .any(|row| row.iter().any(expr_calls_volatile_function)),
        _ => false,
    }
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn test_common_functions() {
        assert_eq!(function_name(0), "COUNT");
        assert_eq!(function_name(1), "IF");
        assert_eq!(function_name(4), "SUM");
        assert_eq!(function_name(5), "AVERAGE");
        assert_eq!(function_name(6), "MIN");
        assert_eq!(function_name(7), "MAX");
        assert_eq!(function_name(32), "LEN");
        assert_eq!(function_name(64), "MATCH");
        assert_eq!(function_name(102), "VLOOKUP");
        assert_eq!(function_name(336), "CONCATENATE");
        assert_eq!(function_name(345), "SUMIF");
        assert_eq!(function_name(480), "IFERROR");
        assert_eq!(function_name(484), "AVERAGEIFS");
    }

    #[test]
    fn test_reserved_slots() {
        assert_eq!(function_name(202), "");
        assert_eq!(function_name(203), "");
        assert_eq!(function_name(217), "");
    }

    #[test]
    fn test_argc() {
        assert_eq!(function_argc(4), 255); // SUM = variable
        assert_eq!(function_argc(1), 3); // IF = 3 (fixed in BIFF8)
        assert_eq!(function_argc(32), 1); // LEN = 1
        assert_eq!(function_argc(64), 3); // MATCH = 3
        assert_eq!(function_argc(19), 0); // PI = 0
    }

    #[test]
    fn test_out_of_range() {
        assert_eq!(function_name(999), "");
        assert_eq!(function_argc(999), 255);
    }

    #[test]
    fn test_function_index_common() {
        assert_eq!(function_index("SUM"), Some(4));
        assert_eq!(function_index("IF"), Some(1));
        assert_eq!(function_index("VLOOKUP"), Some(102));
        assert_eq!(function_index("AVERAGE"), Some(5));
        assert_eq!(function_index("COUNT"), Some(0));
        assert_eq!(function_index("CONCATENATE"), Some(336));
        assert_eq!(function_index("IFERROR"), Some(480));
    }

    #[test]
    fn test_function_index_case_insensitive() {
        assert_eq!(function_index("sum"), Some(4));
        assert_eq!(function_index("Sum"), Some(4));
        assert_eq!(function_index("vlookup"), Some(102));
    }

    #[test]
    fn test_function_index_not_found() {
        assert_eq!(function_index("NOTAFUNCTION"), None);
        assert_eq!(function_index(""), None);
    }

    #[test]
    fn fixed_arity_requires_matching_argc() {
        // PI is declared 0-arg, fixed.
        assert!(function_is_fixed_arity(19, 0));
        // ABS is declared 1-arg, fixed.
        assert!(function_is_fixed_arity(24, 1));
        // ABS with 2 args is never fixed (mismatch with FTAB).
        assert!(!function_is_fixed_arity(24, 2));
        // SUM is variable-argc (FTAB declares 255); never fixed.
        assert!(!function_is_fixed_arity(4, 1));
        assert!(!function_is_fixed_arity(4, 5));
    }

    #[test]
    fn fixed_arity_excludes_non_allowlisted() {
        // COUNTA (169) has FTAB argc=255 (variable) — never fixed.
        assert!(!function_is_fixed_arity(169, 1));
        // VLOOKUP (102) has FTAB argc=4 but is not on the allow-list —
        // Excel emits PtgFuncVar for it.
        assert!(!function_is_fixed_arity(102, 4));
    }

    #[test]
    fn arg_class_aggregators_use_r_class() {
        assert_eq!(function_arg_class(0, 0), OperandClass::R); // COUNT
        assert_eq!(function_arg_class(4, 0), OperandClass::R); // SUM
        assert_eq!(function_arg_class(5, 0), OperandClass::R); // AVERAGE
        assert_eq!(function_arg_class(6, 0), OperandClass::R); // MIN
        assert_eq!(function_arg_class(7, 0), OperandClass::R); // MAX
        assert_eq!(function_arg_class(12, 0), OperandClass::R); // STDEV
        assert_eq!(function_arg_class(46, 0), OperandClass::R); // VAR
        assert_eq!(function_arg_class(169, 0), OperandClass::R); // COUNTA
        assert_eq!(function_arg_class(183, 0), OperandClass::R); // PRODUCT
    }

    #[test]
    fn arg_class_ref_functions_use_r_class() {
        assert_eq!(function_arg_class(8, 0), OperandClass::R); // ROW
        assert_eq!(function_arg_class(9, 0), OperandClass::R); // COLUMN
        assert_eq!(function_arg_class(75, 0), OperandClass::R); // AREAS
        assert_eq!(function_arg_class(78, 0), OperandClass::R); // OFFSET arg 0
        // OFFSET arg 1..3 take values (rows, cols, height, width) — V-class.
        assert_eq!(function_arg_class(78, 1), OperandClass::V);
        assert_eq!(function_arg_class(78, 2), OperandClass::V);
    }

    #[test]
    fn arg_class_default_is_v() {
        // Unknown function falls through to V.
        assert_eq!(function_arg_class(9999, 0), OperandClass::V);
        // IF (1) is not on the R-class list — V by default.
        assert_eq!(function_arg_class(1, 0), OperandClass::V);
        // VLOOKUP arg 0 is V.
        assert_eq!(function_arg_class(102, 0), OperandClass::V);
    }

    #[test]
    fn volatile_known_functions() {
        assert!(function_is_volatile(63)); // RAND
        assert!(function_is_volatile(74)); // NOW
        assert!(function_is_volatile(78)); // OFFSET
        assert!(function_is_volatile(148)); // INDIRECT
        assert!(function_is_volatile(221)); // TODAY
        assert!(function_is_volatile(244)); // INFO
        // RANDBETWEEN (464) is volatile in the runtime registry, so the
        // writer also sees it as volatile via the unified table. Closes
        // the drift bug between runtime and writer metadata.
        assert!(function_is_volatile(464)); // RANDBETWEEN
    }

    #[test]
    fn volatile_excludes_non_volatile() {
        assert!(!function_is_volatile(4)); // SUM
        assert!(!function_is_volatile(102)); // VLOOKUP
        assert!(!function_is_volatile(219)); // ADDRESS (not volatile)
        assert!(!function_is_volatile(345)); // SUMIF
    }

    #[test]
    fn expr_volatile_detects_direct_call() {
        use crate::FormulaExpr;

        let now = FormulaExpr::Function {
            name: "NOW".to_string(),
            args: vec![],
        };
        assert!(expr_calls_volatile_function(&now));

        let sum = FormulaExpr::Function {
            name: "SUM".to_string(),
            args: vec![FormulaExpr::Number(1.0), FormulaExpr::Number(2.0)],
        };
        assert!(!expr_calls_volatile_function(&sum));
    }

    #[test]
    fn expr_volatile_detects_nested_call() {
        use crate::ast::BinaryOperator;
        use crate::FormulaExpr;

        // SUM(NOW(), 1) — NOW is nested inside SUM args
        let expr = FormulaExpr::Function {
            name: "SUM".to_string(),
            args: vec![
                FormulaExpr::Function {
                    name: "NOW".to_string(),
                    args: vec![],
                },
                FormulaExpr::Number(1.0),
            ],
        };
        assert!(expr_calls_volatile_function(&expr));

        // 1 + RAND() — volatile inside binary op
        let expr = FormulaExpr::BinaryOp {
            op: BinaryOperator::Add,
            left: Box::new(FormulaExpr::Number(1.0)),
            right: Box::new(FormulaExpr::Function {
                name: "RAND".to_string(),
                args: vec![],
            }),
        };
        assert!(expr_calls_volatile_function(&expr));
    }

    #[test]
    fn expr_volatile_case_insensitive() {
        use crate::FormulaExpr;

        let lowercase = FormulaExpr::Function {
            name: "now".to_string(),
            args: vec![],
        };
        assert!(expr_calls_volatile_function(&lowercase));
    }

    #[test]
    fn expr_volatile_catches_randbetween() {
        // Regression test for the metadata-drift bug: RANDBETWEEN was
        // volatile in the runtime registry but not in the writer's
        // iftab-based volatile list, so the writer omitted PtgAttrVolatile.
        // After the registry unification both paths agree.
        use crate::FormulaExpr;

        let expr = FormulaExpr::Function {
            name: "RANDBETWEEN".to_string(),
            args: vec![FormulaExpr::Number(1.0), FormulaExpr::Number(10.0)],
        };
        assert!(expr_calls_volatile_function(&expr));
    }

    #[test]
    fn registry_iftab_name_roundtrip_is_consistent() {
        // Every function registered with an iftab must round-trip:
        // function_index(name) == iftab AND function_name(iftab) == name.
        // Catches any by_name/by_iftab index inconsistency introduced by
        // the registration migration or future edits. (Name comparison is
        // case-insensitive because function_index uppercases.)
        use crate::functions::registry;

        for def in registry().iter() {
            let Some(iftab) = def.iftab else { continue };
            assert_eq!(
                function_name(iftab),
                def.name,
                "function_name({iftab}) should be {}",
                def.name
            );
            assert_eq!(
                function_index(def.name),
                Some(iftab),
                "function_index({}) should be {iftab}",
                def.name
            );
            assert_eq!(
                function_argc(iftab),
                def.declared_argc,
                "function_argc({iftab}) should match declared_argc for {}",
                def.name
            );
        }
    }

    #[test]
    fn obsolete_biff_functions_resolve_by_iftab() {
        // Spot-check a curated set of obsolete BIFF8 macro functions against
        // their known MS-XLS Ftab indices. These have no runtime evaluator
        // but must still decompile by name. Values cross-checked against the
        // published [MS-XLS] Ftab so the migration script's output is
        // independently verified for at least these entries.
        let known: &[(u16, &str)] = &[
            (79, "ABSREF"),
            (87, "ECHO"),
            (89, "CALLER"),
            (91, "WINDOWS"),
            (186, "GET.WORKSPACE"),
            (187, "GET.WINDOW"),
            (188, "GET.DOCUMENT"),
            (94, "ACTIVE.CELL"),
            (110, "EXEC"),
            (150, "CALL"),
            (167, "IPMT"),
        ];
        for &(iftab, name) in known {
            assert_eq!(function_name(iftab), name, "iftab {iftab}");
            assert_eq!(function_index(name), Some(iftab), "name {name}");
        }
    }

    #[test]
    fn returns_reference_set_for_ref_class_functions() {
        // Reference-class functions take the operand class of the position
        // they occupy (verified against Excel: R inside SUM(...), V at top
        // level). Pure value functions are always V.
        assert!(function_returns_reference(1)); // IF
        assert!(function_returns_reference(100)); // CHOOSE
        assert!(function_returns_reference(78)); // OFFSET
        assert!(function_returns_reference(148)); // INDIRECT
        assert!(function_returns_reference(29)); // INDEX
        assert!(!function_returns_reference(24)); // ABS
        assert!(!function_returns_reference(4)); // SUM
        assert!(!function_returns_reference(102)); // VLOOKUP (returns value)
    }

    #[test]
    fn biff8_addin_range() {
        // The ATP add-in serialization range [384, 476], determined by
        // comprehensive Excel classification.
        assert!(function_is_biff8_addin(384)); // HEX2BIN
        assert!(function_is_biff8_addin(449)); // EDATE
        assert!(function_is_biff8_addin(472)); // NETWORKDAYS
        assert!(function_is_biff8_addin(476)); // FVSCHEDULE
        // Just outside the range: native.
        assert!(!function_is_biff8_addin(383)); // CUBERANKEDMEMBER
        assert!(!function_is_biff8_addin(477)); // CUBEKPIMEMBER
        assert!(!function_is_biff8_addin(480)); // IFERROR (2007 native)
        assert!(!function_is_biff8_addin(345)); // SUMIF (classic native)
        assert!(!function_is_biff8_addin(4)); // SUM
    }

    #[test]
    fn time_is_fixed_arity() {
        // Regression: an earlier probe misread TIME=66's iftab-low-byte
        // (0x42) as a PtgFuncVar opcode. TIME is fixed 3-arg → PtgFunc.
        assert!(function_is_fixed_arity(66, 3));
    }
}
