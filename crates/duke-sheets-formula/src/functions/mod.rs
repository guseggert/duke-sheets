//! Built-in Excel functions

pub mod compatibility;
pub mod criteria;
pub mod database;
pub mod date;
pub mod engineering;
pub mod financial;
pub mod financial_extra;
pub mod info;
pub mod info_extra;
pub mod logical;
pub mod logical_extra;
pub mod lookup;
pub mod lookup_extra;
pub mod math;
pub mod math_extra;
pub mod statistical;
pub mod statistical_extra;
pub mod text;
pub mod text_extra;

use crate::decompile::function_table::OperandClass;
use crate::error::{FormulaError, FormulaResult};
use crate::evaluator::{EvaluationContext, FormulaValue};
use std::collections::HashMap;
use std::sync::OnceLock;

/// Function implementation signature
///
/// Functions can consult the evaluation context (e.g. workbook settings, date system,
/// current sheet/cell) to match Excel semantics.
pub type FunctionImpl = fn(&[FormulaValue], &EvaluationContext) -> FormulaResult<FormulaValue>;

/// Stub implementation used by [`FunctionDef::default`] and by registry entries for
/// obsolete BIFF8 macro functions (ABSREF, ECHO, GET.*, etc.) that we recognize by
/// name but don't evaluate. Returns an [`FormulaError::Evaluation`].
fn fn_not_implemented(
    _args: &[FormulaValue],
    _ctx: &EvaluationContext,
) -> FormulaResult<FormulaValue> {
    Err(FormulaError::Evaluation(
        "function not implemented".to_string(),
    ))
}

/// Function definition.
///
/// Carries everything we know about a built-in function: name, evaluator,
/// argument constraints, and the BIFF8 encoding metadata used by the XLS
/// and XLSB writers. The writer paths consume `iftab`, `declared_argc`,
/// `fixed_arity`, `default_arg_class`, `arg_classes`, and `volatile`; the
/// evaluator consumes `min_args`, `max_args`, `volatile`, and `implementation`.
pub struct FunctionDef {
    /// Function name in canonical uppercase form.
    pub name: &'static str,
    /// BIFF8 function table index ([MS-XLS] §2.5.198.63 Ftab), if defined.
    /// `None` for functions added after BIFF8 (XLOOKUP, RANDARRAY, dynamic-array
    /// functions, dot-suffixed `.MATH`/`.PRECISE` variants, etc.).
    pub iftab: Option<u16>,
    /// MS-XLS Ftab declared argument count encoding:
    /// - `0..253`: fixed count
    /// - `254`: variable, minimum from this value
    /// - `255`: variable, 0 or more
    ///
    /// Used by the writer to decide PtgFunc-vs-PtgFuncVar emission. For
    /// non-BIFF8 functions any value is safe; default is `255`.
    pub declared_argc: u16,
    /// Minimum argument count enforced by the runtime evaluator.
    pub min_args: usize,
    /// Maximum argument count enforced by the runtime evaluator (`None` = unlimited).
    pub max_args: Option<usize>,
    /// Whether Excel emits this function with PtgFunc (`true`, fixed-arity opcode)
    /// rather than PtgFuncVar (`false`, variable-arity opcode) when `actual_argc`
    /// matches `declared_argc`. Empirically grown from Excel-authored byte-parity
    /// tests; defaulting to `false` is always safe (PtgFuncVar accepts any arity).
    pub fixed_arity: bool,
    /// Whether the function is volatile (recalculated every change). Drives both
    /// runtime re-evaluation and writer-side PtgAttrVolatile prefix emission.
    pub volatile: bool,
    /// Default operand class for arguments. Most functions use `V` (value);
    /// aggregators (SUM/AVERAGE/etc.) use `R` (reference) so range operands
    /// iterate rather than collapse.
    pub default_arg_class: OperandClass,
    /// Per-position operand class overrides (indexed by argument position).
    /// Arguments beyond this slice's length fall back to `default_arg_class`.
    pub arg_classes: &'static [OperandClass],
    /// Runtime evaluator implementation. Defaults to a stub that returns
    /// an `Evaluation` error.
    pub implementation: FunctionImpl,
}

impl Default for FunctionDef {
    fn default() -> Self {
        Self {
            name: "",
            iftab: None,
            declared_argc: 255,
            min_args: 0,
            max_args: None,
            fixed_arity: false,
            volatile: false,
            default_arg_class: OperandClass::V,
            arg_classes: &[],
            implementation: fn_not_implemented,
        }
    }
}

/// Function registry.
///
/// Stores function definitions in a `Vec` for owned, address-stable storage and
/// builds two indexes for lookup: by uppercase name (for the parser/evaluator)
/// and by BIFF8 iftab index (for the XLS/XLSB writer and decompiler).
pub struct FunctionRegistry {
    functions: Vec<FunctionDef>,
    by_name: HashMap<String, usize>,
    by_iftab: Vec<Option<usize>>,
}

/// Global function registry, lazily initialized on first access. Both the
/// evaluator and the decompile metadata module consult this singleton so
/// metadata never diverges between paths.
static GLOBAL_REGISTRY: OnceLock<FunctionRegistry> = OnceLock::new();

/// Return the global function registry, building it on first call.
pub fn registry() -> &'static FunctionRegistry {
    GLOBAL_REGISTRY.get_or_init(FunctionRegistry::new)
}

impl FunctionRegistry {
    /// Create a new registry with all built-in functions
    pub fn new() -> Self {
        let mut registry = Self {
            functions: Vec::new(),
            by_name: HashMap::new(),
            by_iftab: Vec::new(),
        };

        registry.register_math_functions();
        registry.register_logical_functions();
        registry.register_text_functions();
        registry.register_info_functions();
        registry.register_date_functions();
        registry.register_lookup_functions();
        registry.register_statistical_functions();
        registry.register_financial_functions();
        registry.register_database_functions();
        registry.register_engineering_functions();
        registry.register_compatibility_functions();
        registry.register_web_functions();
        registry.register_obsolete_biff_functions();

        registry
    }

    /// Look up a function by name (case-insensitive).
    pub fn get(&self, name: &str) -> Option<&FunctionDef> {
        self.by_name
            .get(&name.to_uppercase())
            .map(|&i| &self.functions[i])
    }

    /// Look up a function by BIFF8 iftab index.
    pub fn get_by_iftab(&self, iftab: u16) -> Option<&FunctionDef> {
        self.by_iftab
            .get(iftab as usize)
            .copied()
            .flatten()
            .map(|i| &self.functions[i])
    }

    /// Iterate over all registered functions.
    pub fn iter(&self) -> impl Iterator<Item = &FunctionDef> {
        self.functions.iter()
    }

    /// Register a function. Inserts into both the name and iftab indexes
    /// (the latter only if `def.iftab` is `Some`).
    pub fn register(&mut self, def: FunctionDef) {
        let idx = self.functions.len();
        let upper = def.name.to_uppercase();
        if let Some(iftab) = def.iftab {
            let ift = iftab as usize;
            if self.by_iftab.len() <= ift {
                self.by_iftab.resize(ift + 1, None);
            }
            self.by_iftab[ift] = Some(idx);
        }
        self.by_name.insert(upper, idx);
        self.functions.push(def);
    }

    fn register_math_functions(&mut self) {
        // SUM
        self.register(FunctionDef {
            name: "SUM",
            iftab: Some(4),
            min_args: 1,
            max_args: None,
            implementation: math::fn_sum,
            volatile: false,
            default_arg_class: OperandClass::R,
            ..Default::default()
        });

        // AVERAGE
        self.register(FunctionDef {
            name: "AVERAGE",
            iftab: Some(5),
            min_args: 1,
            max_args: None,
            implementation: math::fn_average,
            volatile: false,
            default_arg_class: OperandClass::R,
            ..Default::default()
        });

        // MIN
        self.register(FunctionDef {
            name: "MIN",
            iftab: Some(6),
            min_args: 1,
            max_args: None,
            implementation: math::fn_min,
            volatile: false,
            default_arg_class: OperandClass::R,
            ..Default::default()
        });

        // MAX
        self.register(FunctionDef {
            name: "MAX",
            iftab: Some(7),
            min_args: 1,
            max_args: None,
            implementation: math::fn_max,
            volatile: false,
            default_arg_class: OperandClass::R,
            ..Default::default()
        });

        // COUNT
        self.register(FunctionDef {
            name: "COUNT",
            iftab: Some(0),
            min_args: 1,
            max_args: None,
            implementation: math::fn_count,
            volatile: false,
            default_arg_class: OperandClass::R,
            ..Default::default()
        });

        // RAND (volatile)
        self.register(FunctionDef {
            name: "RAND",
            iftab: Some(63),
            declared_argc: 0,
            min_args: 0,
            max_args: Some(0),
            implementation: math::fn_rand,
            volatile: true,
            fixed_arity: true,
            ..Default::default()
        });

        // RANDBETWEEN (volatile)
        self.register(FunctionDef {
            name: "RANDBETWEEN",
            iftab: Some(464),
            declared_argc: 2,
            min_args: 2,
            max_args: Some(2),
            implementation: math::fn_randbetween,
            volatile: true,
            ..Default::default()
        });

        // ABS
        self.register(FunctionDef {
            name: "ABS",
            iftab: Some(24),
            declared_argc: 1,
            min_args: 1,
            max_args: Some(1),
            implementation: math::fn_abs,
            volatile: false,
            fixed_arity: true,
            ..Default::default()
        });

        // ROUND
        self.register(FunctionDef {
            name: "ROUND",
            iftab: Some(27),
            declared_argc: 2,
            min_args: 1,
            max_args: Some(2),
            implementation: math::fn_round,
            volatile: false,
            fixed_arity: true,
            ..Default::default()
        });

        // MOD
        self.register(FunctionDef {
            name: "MOD",
            iftab: Some(39),
            declared_argc: 2,
            min_args: 2,
            max_args: Some(2),
            implementation: math::fn_mod,
            volatile: false,
            fixed_arity: true,
            ..Default::default()
        });

        // INT
        self.register(FunctionDef {
            name: "INT",
            iftab: Some(25),
            declared_argc: 1,
            min_args: 1,
            max_args: Some(1),
            implementation: math::fn_int,
            volatile: false,
            fixed_arity: true,
            ..Default::default()
        });

        // TRUNC
        self.register(FunctionDef {
            name: "TRUNC",
            iftab: Some(197),
            declared_argc: 2,
            min_args: 1,
            max_args: Some(2),
            implementation: math::fn_trunc,
            volatile: false,
            ..Default::default()
        });

        // SIGN
        self.register(FunctionDef {
            name: "SIGN",
            iftab: Some(26),
            declared_argc: 1,
            min_args: 1,
            max_args: Some(1),
            implementation: math::fn_sign,
            volatile: false,
            fixed_arity: true,
            ..Default::default()
        });

        // SQRT
        self.register(FunctionDef {
            name: "SQRT",
            iftab: Some(20),
            declared_argc: 1,
            min_args: 1,
            max_args: Some(1),
            implementation: math::fn_sqrt,
            volatile: false,
            fixed_arity: true,
            ..Default::default()
        });

        // POWER
        self.register(FunctionDef {
            name: "POWER",
            iftab: Some(337),
            declared_argc: 2,
            min_args: 2,
            max_args: Some(2),
            implementation: math::fn_power,
            volatile: false,
            ..Default::default()
        });

        // LOG
        self.register(FunctionDef {
            name: "LOG",
            iftab: Some(109),
            declared_argc: 2,
            min_args: 1,
            max_args: Some(2),
            implementation: math::fn_log,
            volatile: false,
            ..Default::default()
        });

        // LOG10
        self.register(FunctionDef {
            name: "LOG10",
            iftab: Some(23),
            declared_argc: 1,
            min_args: 1,
            max_args: Some(1),
            implementation: math::fn_log10,
            volatile: false,
            fixed_arity: true,
            ..Default::default()
        });

        // LN
        self.register(FunctionDef {
            name: "LN",
            iftab: Some(22),
            declared_argc: 1,
            min_args: 1,
            max_args: Some(1),
            implementation: math::fn_ln,
            volatile: false,
            fixed_arity: true,
            ..Default::default()
        });

        // EXP
        self.register(FunctionDef {
            name: "EXP",
            iftab: Some(21),
            declared_argc: 1,
            min_args: 1,
            max_args: Some(1),
            implementation: math::fn_exp,
            volatile: false,
            fixed_arity: true,
            ..Default::default()
        });

        // PI
        self.register(FunctionDef {
            name: "PI",
            iftab: Some(19),
            declared_argc: 0,
            min_args: 0,
            max_args: Some(0),
            implementation: math::fn_pi,
            volatile: false,
            fixed_arity: true,
            ..Default::default()
        });

        // SUMIF
        self.register(FunctionDef {
            name: "SUMIF",
            iftab: Some(345),
            declared_argc: 3,
            min_args: 2,
            max_args: Some(3),
            implementation: math::fn_sumif,
            volatile: false,
            ..Default::default()
        });

        // SUMIFS
        self.register(FunctionDef {
            name: "SUMIFS",
            iftab: Some(482),
            declared_argc: 129,
            min_args: 3,
            max_args: None, // sum_range + up to 127 criteria pairs
            implementation: math::fn_sumifs,
            volatile: false,
            ..Default::default()
        });

        // SUMPRODUCT
        self.register(FunctionDef {
            name: "SUMPRODUCT",
            iftab: Some(228),
            min_args: 1,
            max_args: None, // Up to 255 arrays
            implementation: math::fn_sumproduct,
            volatile: false,
            ..Default::default()
        });

        // SIN
        self.register(FunctionDef {
            name: "SIN",
            iftab: Some(15),
            declared_argc: 1,
            min_args: 1,
            max_args: Some(1),
            implementation: math::fn_sin,
            volatile: false,
            fixed_arity: true,
            ..Default::default()
        });

        // COS
        self.register(FunctionDef {
            name: "COS",
            iftab: Some(16),
            declared_argc: 1,
            min_args: 1,
            max_args: Some(1),
            implementation: math::fn_cos,
            volatile: false,
            fixed_arity: true,
            ..Default::default()
        });

        // TAN
        self.register(FunctionDef {
            name: "TAN",
            iftab: Some(17),
            declared_argc: 1,
            min_args: 1,
            max_args: Some(1),
            implementation: math::fn_tan,
            volatile: false,
            fixed_arity: true,
            ..Default::default()
        });

        // ASIN
        self.register(FunctionDef {
            name: "ASIN",
            iftab: Some(98),
            declared_argc: 1,
            min_args: 1,
            max_args: Some(1),
            implementation: math::fn_asin,
            volatile: false,
            fixed_arity: true,
            ..Default::default()
        });

        // ACOS
        self.register(FunctionDef {
            name: "ACOS",
            iftab: Some(99),
            declared_argc: 1,
            min_args: 1,
            max_args: Some(1),
            implementation: math::fn_acos,
            volatile: false,
            fixed_arity: true,
            ..Default::default()
        });

        // ATAN
        self.register(FunctionDef {
            name: "ATAN",
            iftab: Some(18),
            declared_argc: 1,
            min_args: 1,
            max_args: Some(1),
            implementation: math::fn_atan,
            volatile: false,
            fixed_arity: true,
            ..Default::default()
        });

        // ATAN2
        self.register(FunctionDef {
            name: "ATAN2",
            iftab: Some(97),
            declared_argc: 2,
            min_args: 2,
            max_args: Some(2),
            implementation: math::fn_atan2,
            volatile: false,
            fixed_arity: true,
            ..Default::default()
        });

        // DEGREES
        self.register(FunctionDef {
            name: "DEGREES",
            iftab: Some(343),
            declared_argc: 1,
            min_args: 1,
            max_args: Some(1),
            implementation: math::fn_degrees,
            volatile: false,
            ..Default::default()
        });

        // RADIANS
        self.register(FunctionDef {
            name: "RADIANS",
            iftab: Some(342),
            declared_argc: 1,
            min_args: 1,
            max_args: Some(1),
            implementation: math::fn_radians,
            volatile: false,
            ..Default::default()
        });

        // ROUNDUP
        self.register(FunctionDef {
            name: "ROUNDUP",
            iftab: Some(212),
            declared_argc: 2,
            min_args: 2,
            max_args: Some(2),
            implementation: math::fn_roundup,
            volatile: false,
            ..Default::default()
        });

        // ROUNDDOWN
        self.register(FunctionDef {
            name: "ROUNDDOWN",
            iftab: Some(213),
            declared_argc: 2,
            min_args: 2,
            max_args: Some(2),
            implementation: math::fn_rounddown,
            volatile: false,
            ..Default::default()
        });

        // CEILING.MATH
        self.register(FunctionDef {
            name: "CEILING.MATH",
            min_args: 1,
            max_args: Some(3),
            implementation: math::fn_ceiling_math,
            volatile: false,
            ..Default::default()
        });

        // FLOOR.MATH
        self.register(FunctionDef {
            name: "FLOOR.MATH",
            min_args: 1,
            max_args: Some(3),
            implementation: math::fn_floor_math,
            volatile: false,
            ..Default::default()
        });

        // ODD
        self.register(FunctionDef {
            name: "ODD",
            iftab: Some(298),
            declared_argc: 1,
            min_args: 1,
            max_args: Some(1),
            implementation: math::fn_odd,
            volatile: false,
            ..Default::default()
        });

        // EVEN
        self.register(FunctionDef {
            name: "EVEN",
            iftab: Some(279),
            declared_argc: 1,
            min_args: 1,
            max_args: Some(1),
            implementation: math::fn_even,
            volatile: false,
            ..Default::default()
        });
        // math_extra functions
        self.register(FunctionDef {
            name: "ACOSH",
            iftab: Some(233),
            declared_argc: 1,
            min_args: 1,
            max_args: Some(1),
            implementation: math_extra::fn_acosh,
            volatile: false,
            ..Default::default()
        });
        self.register(FunctionDef {
            name: "ASINH",
            iftab: Some(232),
            declared_argc: 1,
            min_args: 1,
            max_args: Some(1),
            implementation: math_extra::fn_asinh,
            volatile: false,
            ..Default::default()
        });
        self.register(FunctionDef {
            name: "ATANH",
            iftab: Some(234),
            declared_argc: 1,
            min_args: 1,
            max_args: Some(1),
            implementation: math_extra::fn_atanh,
            volatile: false,
            ..Default::default()
        });
        self.register(FunctionDef {
            name: "ACOT",
            min_args: 1,
            max_args: Some(1),
            implementation: math_extra::fn_acot,
            volatile: false,
            ..Default::default()
        });
        self.register(FunctionDef {
            name: "ACOTH",
            min_args: 1,
            max_args: Some(1),
            implementation: math_extra::fn_acoth,
            volatile: false,
            ..Default::default()
        });
        self.register(FunctionDef {
            name: "COSH",
            iftab: Some(230),
            declared_argc: 1,
            min_args: 1,
            max_args: Some(1),
            implementation: math_extra::fn_cosh,
            volatile: false,
            ..Default::default()
        });
        self.register(FunctionDef {
            name: "SINH",
            iftab: Some(229),
            declared_argc: 1,
            min_args: 1,
            max_args: Some(1),
            implementation: math_extra::fn_sinh,
            volatile: false,
            ..Default::default()
        });
        self.register(FunctionDef {
            name: "TANH",
            iftab: Some(231),
            declared_argc: 1,
            min_args: 1,
            max_args: Some(1),
            implementation: math_extra::fn_tanh,
            volatile: false,
            ..Default::default()
        });
        self.register(FunctionDef {
            name: "COT",
            min_args: 1,
            max_args: Some(1),
            implementation: math_extra::fn_cot,
            volatile: false,
            ..Default::default()
        });
        self.register(FunctionDef {
            name: "COTH",
            min_args: 1,
            max_args: Some(1),
            implementation: math_extra::fn_coth,
            volatile: false,
            ..Default::default()
        });
        self.register(FunctionDef {
            name: "CSC",
            min_args: 1,
            max_args: Some(1),
            implementation: math_extra::fn_csc,
            volatile: false,
            ..Default::default()
        });
        self.register(FunctionDef {
            name: "CSCH",
            min_args: 1,
            max_args: Some(1),
            implementation: math_extra::fn_csch,
            volatile: false,
            ..Default::default()
        });
        self.register(FunctionDef {
            name: "SEC",
            min_args: 1,
            max_args: Some(1),
            implementation: math_extra::fn_sec,
            volatile: false,
            ..Default::default()
        });
        self.register(FunctionDef {
            name: "SECH",
            min_args: 1,
            max_args: Some(1),
            implementation: math_extra::fn_sech,
            volatile: false,
            ..Default::default()
        });
        self.register(FunctionDef {
            name: "COMBIN",
            iftab: Some(276),
            declared_argc: 2,
            min_args: 2,
            max_args: Some(2),
            implementation: math_extra::fn_combin,
            volatile: false,
            ..Default::default()
        });
        self.register(FunctionDef {
            name: "COMBINA",
            min_args: 2,
            max_args: Some(2),
            implementation: math_extra::fn_combina,
            volatile: false,
            ..Default::default()
        });
        self.register(FunctionDef {
            name: "FACT",
            iftab: Some(184),
            declared_argc: 1,
            min_args: 1,
            max_args: Some(1),
            implementation: math_extra::fn_fact,
            volatile: false,
            ..Default::default()
        });
        self.register(FunctionDef {
            name: "FACTDOUBLE",
            iftab: Some(415),
            declared_argc: 1,
            min_args: 1,
            max_args: Some(1),
            implementation: math_extra::fn_factdouble,
            volatile: false,
            ..Default::default()
        });
        self.register(FunctionDef {
            name: "MULTINOMIAL",
            iftab: Some(474),
            min_args: 1,
            max_args: None,
            implementation: math_extra::fn_multinomial,
            volatile: false,
            ..Default::default()
        });
        self.register(FunctionDef {
            name: "GCD",
            iftab: Some(473),
            min_args: 1,
            max_args: None,
            implementation: math_extra::fn_gcd,
            volatile: false,
            ..Default::default()
        });
        self.register(FunctionDef {
            name: "LCM",
            iftab: Some(475),
            min_args: 1,
            max_args: None,
            implementation: math_extra::fn_lcm,
            volatile: false,
            ..Default::default()
        });
        self.register(FunctionDef {
            name: "PRODUCT",
            iftab: Some(183),
            min_args: 1,
            max_args: None,
            implementation: math_extra::fn_product,
            volatile: false,
            default_arg_class: OperandClass::R,
            ..Default::default()
        });
        self.register(FunctionDef {
            name: "QUOTIENT",
            iftab: Some(417),
            declared_argc: 2,
            min_args: 2,
            max_args: Some(2),
            implementation: math_extra::fn_quotient,
            volatile: false,
            ..Default::default()
        });
        self.register(FunctionDef {
            name: "MROUND",
            iftab: Some(422),
            declared_argc: 2,
            min_args: 2,
            max_args: Some(2),
            implementation: math_extra::fn_mround,
            volatile: false,
            ..Default::default()
        });
        self.register(FunctionDef {
            name: "SUMSQ",
            iftab: Some(321),
            min_args: 1,
            max_args: None,
            implementation: math_extra::fn_sumsq,
            volatile: false,
            ..Default::default()
        });
        self.register(FunctionDef {
            name: "SQRTPI",
            iftab: Some(416),
            declared_argc: 1,
            min_args: 1,
            max_args: Some(1),
            implementation: math_extra::fn_sqrtpi,
            volatile: false,
            ..Default::default()
        });
        self.register(FunctionDef {
            name: "BASE",
            min_args: 2,
            max_args: Some(3),
            implementation: math_extra::fn_base,
            volatile: false,
            ..Default::default()
        });
        self.register(FunctionDef {
            name: "DECIMAL",
            min_args: 2,
            max_args: Some(2),
            implementation: math_extra::fn_decimal,
            volatile: false,
            ..Default::default()
        });
        self.register(FunctionDef {
            name: "ROMAN",
            iftab: Some(354),
            declared_argc: 2,
            min_args: 1,
            max_args: Some(2),
            implementation: math_extra::fn_roman,
            volatile: false,
            ..Default::default()
        });
        self.register(FunctionDef {
            name: "ARABIC",
            min_args: 1,
            max_args: Some(1),
            implementation: math_extra::fn_arabic,
            volatile: false,
            ..Default::default()
        });
        self.register(FunctionDef {
            name: "CEILING.PRECISE",
            min_args: 1,
            max_args: Some(2),
            implementation: math_extra::fn_ceiling_precise,
            volatile: false,
            ..Default::default()
        });
        self.register(FunctionDef {
            name: "FLOOR.PRECISE",
            min_args: 1,
            max_args: Some(2),
            implementation: math_extra::fn_floor_precise,
            volatile: false,
            ..Default::default()
        });
        self.register(FunctionDef {
            name: "ISO.CEILING",
            min_args: 1,
            max_args: Some(2),
            implementation: math_extra::fn_iso_ceiling,
            volatile: false,
            ..Default::default()
        });
        self.register(FunctionDef {
            name: "MDETERM",
            iftab: Some(163),
            declared_argc: 1,
            min_args: 1,
            max_args: Some(1),
            implementation: math_extra::fn_mdeterm,
            volatile: false,
            ..Default::default()
        });
        self.register(FunctionDef {
            name: "MINVERSE",
            iftab: Some(164),
            declared_argc: 1,
            min_args: 1,
            max_args: Some(1),
            implementation: math_extra::fn_minverse,
            volatile: false,
            ..Default::default()
        });
        self.register(FunctionDef {
            name: "MMULT",
            iftab: Some(165),
            declared_argc: 1,
            min_args: 2,
            max_args: Some(2),
            implementation: math_extra::fn_mmult,
            volatile: false,
            ..Default::default()
        });
        self.register(FunctionDef {
            name: "MUNIT",
            min_args: 1,
            max_args: Some(1),
            implementation: math_extra::fn_munit,
            volatile: false,
            ..Default::default()
        });
        self.register(FunctionDef {
            name: "RANDARRAY",
            min_args: 0,
            max_args: Some(5),
            implementation: math_extra::fn_randarray,
            volatile: true,
            ..Default::default()
        });
        self.register(FunctionDef {
            name: "SERIESSUM",
            iftab: Some(414),
            declared_argc: 4,
            min_args: 4,
            max_args: Some(4),
            implementation: math_extra::fn_seriessum,
            volatile: false,
            ..Default::default()
        });
        self.register(FunctionDef {
            name: "SUMX2MY2",
            iftab: Some(304),
            declared_argc: 2,
            min_args: 2,
            max_args: Some(2),
            implementation: math_extra::fn_sumx2my2,
            volatile: false,
            ..Default::default()
        });
        self.register(FunctionDef {
            name: "SUMX2PY2",
            iftab: Some(305),
            declared_argc: 2,
            min_args: 2,
            max_args: Some(2),
            implementation: math_extra::fn_sumx2py2,
            volatile: false,
            ..Default::default()
        });
        self.register(FunctionDef {
            name: "SUMXMY2",
            iftab: Some(303),
            declared_argc: 2,
            min_args: 2,
            max_args: Some(2),
            implementation: math_extra::fn_sumxmy2,
            volatile: false,
            ..Default::default()
        });
        self.register(FunctionDef {
            name: "AGGREGATE",
            min_args: 3,
            max_args: None,
            implementation: math_extra::fn_aggregate,
            volatile: false,
            ..Default::default()
        });
        self.register(FunctionDef {
            name: "SUBTOTAL",
            iftab: Some(344),
            min_args: 2,
            max_args: None,
            implementation: math_extra::fn_subtotal,
            volatile: false,
            ..Default::default()
        });
        self.register(FunctionDef {
            name: "PERCENTOF",
            min_args: 2,
            max_args: Some(2),
            implementation: math_extra::fn_percentof,
            volatile: false,
            ..Default::default()
        });
    }

    fn register_logical_functions(&mut self) {
        // IF
        self.register(FunctionDef {
            name: "IF",
            iftab: Some(1),
            declared_argc: 3,
            min_args: 2,
            max_args: Some(3),
            implementation: logical::fn_if,
            volatile: false,
            ..Default::default()
        });

        // AND
        self.register(FunctionDef {
            name: "AND",
            iftab: Some(36),
            min_args: 1,
            max_args: None,
            implementation: logical::fn_and,
            volatile: false,
            ..Default::default()
        });

        // OR
        self.register(FunctionDef {
            name: "OR",
            iftab: Some(37),
            min_args: 1,
            max_args: None,
            implementation: logical::fn_or,
            volatile: false,
            ..Default::default()
        });

        // NOT
        self.register(FunctionDef {
            name: "NOT",
            iftab: Some(38),
            declared_argc: 1,
            min_args: 1,
            max_args: Some(1),
            implementation: logical::fn_not,
            volatile: false,
            fixed_arity: true,
            ..Default::default()
        });

        // IFERROR
        self.register(FunctionDef {
            name: "IFERROR",
            iftab: Some(480),
            declared_argc: 2,
            min_args: 2,
            max_args: Some(2),
            implementation: logical::fn_iferror,
            volatile: false,
            ..Default::default()
        });

        // IFNA
        self.register(FunctionDef {
            name: "IFNA",
            min_args: 2,
            max_args: Some(2),
            implementation: logical::fn_ifna,
            volatile: false,
            ..Default::default()
        });

        // TRUE
        self.register(FunctionDef {
            name: "TRUE",
            iftab: Some(34),
            declared_argc: 0,
            min_args: 0,
            max_args: Some(0),
            implementation: logical::fn_true,
            volatile: false,
            fixed_arity: true,
            ..Default::default()
        });

        // FALSE
        self.register(FunctionDef {
            name: "FALSE",
            iftab: Some(35),
            declared_argc: 0,
            min_args: 0,
            max_args: Some(0),
            implementation: logical::fn_false,
            volatile: false,
            fixed_arity: true,
            ..Default::default()
        });

        // XOR
        self.register(FunctionDef {
            name: "XOR",
            min_args: 1,
            max_args: None,
            implementation: logical::fn_xor,
            volatile: false,
            ..Default::default()
        });

        // IFS
        self.register(FunctionDef {
            name: "IFS",
            min_args: 2,
            max_args: None, // Up to 127 condition-value pairs
            implementation: logical::fn_ifs,
            volatile: false,
            ..Default::default()
        });

        // SWITCH
        self.register(FunctionDef {
            name: "SWITCH",
            min_args: 3,
            max_args: None, // Up to 126 value-result pairs + optional default
            implementation: logical::fn_switch,
            volatile: false,
            ..Default::default()
        });
        // logical_extra functions
        self.register(FunctionDef {
            name: "LET",
            min_args: 3,
            max_args: None,
            implementation: logical_extra::fn_let,
            volatile: false,
            ..Default::default()
        });
        self.register(FunctionDef {
            name: "LAMBDA",
            min_args: 1,
            max_args: None,
            implementation: logical_extra::fn_lambda,
            volatile: false,
            ..Default::default()
        });
        self.register(FunctionDef {
            name: "MAP",
            min_args: 2,
            max_args: None,
            implementation: logical_extra::fn_map,
            volatile: false,
            ..Default::default()
        });
        self.register(FunctionDef {
            name: "REDUCE",
            min_args: 3,
            max_args: Some(3),
            implementation: logical_extra::fn_reduce,
            volatile: false,
            ..Default::default()
        });
        self.register(FunctionDef {
            name: "SCAN",
            min_args: 3,
            max_args: Some(3),
            implementation: logical_extra::fn_scan,
            volatile: false,
            ..Default::default()
        });
        self.register(FunctionDef {
            name: "BYCOL",
            min_args: 2,
            max_args: Some(2),
            implementation: logical_extra::fn_bycol,
            volatile: false,
            ..Default::default()
        });
        self.register(FunctionDef {
            name: "BYROW",
            min_args: 2,
            max_args: Some(2),
            implementation: logical_extra::fn_byrow,
            volatile: false,
            ..Default::default()
        });
        self.register(FunctionDef {
            name: "MAKEARRAY",
            min_args: 3,
            max_args: Some(3),
            implementation: logical_extra::fn_makearray,
            volatile: false,
            ..Default::default()
        });
    }

    fn register_text_functions(&mut self) {
        // LEN
        self.register(FunctionDef {
            name: "LEN",
            iftab: Some(32),
            declared_argc: 1,
            min_args: 1,
            max_args: Some(1),
            implementation: text::fn_len,
            volatile: false,
            fixed_arity: true,
            ..Default::default()
        });

        // LEFT
        self.register(FunctionDef {
            name: "LEFT",
            iftab: Some(115),
            declared_argc: 2,
            min_args: 1,
            max_args: Some(2),
            implementation: text::fn_left,
            volatile: false,
            ..Default::default()
        });

        // RIGHT
        self.register(FunctionDef {
            name: "RIGHT",
            iftab: Some(116),
            declared_argc: 2,
            min_args: 1,
            max_args: Some(2),
            implementation: text::fn_right,
            volatile: false,
            ..Default::default()
        });

        // MID
        self.register(FunctionDef {
            name: "MID",
            iftab: Some(31),
            declared_argc: 3,
            min_args: 3,
            max_args: Some(3),
            implementation: text::fn_mid,
            volatile: false,
            fixed_arity: true,
            ..Default::default()
        });

        // LOWER
        self.register(FunctionDef {
            name: "LOWER",
            iftab: Some(112),
            declared_argc: 1,
            min_args: 1,
            max_args: Some(1),
            implementation: text::fn_lower,
            volatile: false,
            ..Default::default()
        });

        // UPPER
        self.register(FunctionDef {
            name: "UPPER",
            iftab: Some(113),
            declared_argc: 1,
            min_args: 1,
            max_args: Some(1),
            implementation: text::fn_upper,
            volatile: false,
            ..Default::default()
        });

        // TRIM
        self.register(FunctionDef {
            name: "TRIM",
            iftab: Some(118),
            declared_argc: 1,
            min_args: 1,
            max_args: Some(1),
            implementation: text::fn_trim,
            volatile: false,
            ..Default::default()
        });

        // CONCAT (newer)
        self.register(FunctionDef {
            name: "CONCAT",
            min_args: 1,
            max_args: None,
            implementation: text::fn_concat,
            volatile: false,
            ..Default::default()
        });

        // CONCATENATE (legacy)
        self.register(FunctionDef {
            name: "CONCATENATE",
            iftab: Some(336),
            min_args: 1,
            max_args: None,
            implementation: text::fn_concat,
            volatile: false,
            ..Default::default()
        });

        // FIND (case-sensitive)
        self.register(FunctionDef {
            name: "FIND",
            iftab: Some(124),
            declared_argc: 3,
            min_args: 2,
            max_args: Some(3),
            implementation: text::fn_find,
            volatile: false,
            ..Default::default()
        });

        // FINDB (same as FIND for non-DBCS)
        self.register(FunctionDef {
            name: "FINDB",
            iftab: Some(205),
            declared_argc: 3,
            min_args: 2,
            max_args: Some(3),
            implementation: text::fn_find,
            volatile: false,
            ..Default::default()
        });

        // SEARCH (case-insensitive)
        self.register(FunctionDef {
            name: "SEARCH",
            iftab: Some(82),
            declared_argc: 3,
            min_args: 2,
            max_args: Some(3),
            implementation: text::fn_search,
            volatile: false,
            ..Default::default()
        });

        // SEARCHB (same as SEARCH for non-DBCS)
        self.register(FunctionDef {
            name: "SEARCHB",
            iftab: Some(206),
            declared_argc: 3,
            min_args: 2,
            max_args: Some(3),
            implementation: text::fn_search,
            volatile: false,
            ..Default::default()
        });

        // EXACT
        self.register(FunctionDef {
            name: "EXACT",
            iftab: Some(117),
            declared_argc: 2,
            min_args: 2,
            max_args: Some(2),
            implementation: text::fn_exact,
            volatile: false,
            ..Default::default()
        });

        // REPT
        self.register(FunctionDef {
            name: "REPT",
            iftab: Some(30),
            declared_argc: 2,
            min_args: 2,
            max_args: Some(2),
            implementation: text::fn_rept,
            volatile: false,
            fixed_arity: true,
            ..Default::default()
        });

        // SUBSTITUTE
        self.register(FunctionDef {
            name: "SUBSTITUTE",
            iftab: Some(120),
            declared_argc: 4,
            min_args: 3,
            max_args: Some(4),
            implementation: text::fn_substitute,
            volatile: false,
            ..Default::default()
        });

        // PROPER
        self.register(FunctionDef {
            name: "PROPER",
            iftab: Some(114),
            declared_argc: 1,
            min_args: 1,
            max_args: Some(1),
            implementation: text::fn_proper,
            volatile: false,
            ..Default::default()
        });

        // CHAR
        self.register(FunctionDef {
            name: "CHAR",
            iftab: Some(111),
            declared_argc: 1,
            min_args: 1,
            max_args: Some(1),
            implementation: text::fn_char,
            volatile: false,
            ..Default::default()
        });

        // CODE
        self.register(FunctionDef {
            name: "CODE",
            iftab: Some(121),
            declared_argc: 1,
            min_args: 1,
            max_args: Some(1),
            implementation: text::fn_code,
            volatile: false,
            ..Default::default()
        });

        // CLEAN
        self.register(FunctionDef {
            name: "CLEAN",
            iftab: Some(162),
            declared_argc: 1,
            min_args: 1,
            max_args: Some(1),
            implementation: text::fn_clean,
            volatile: false,
            ..Default::default()
        });

        // VALUE
        self.register(FunctionDef {
            name: "VALUE",
            iftab: Some(33),
            declared_argc: 1,
            min_args: 1,
            max_args: Some(1),
            implementation: text::fn_value,
            volatile: false,
            fixed_arity: true,
            ..Default::default()
        });

        // T
        self.register(FunctionDef {
            name: "T",
            iftab: Some(130),
            declared_argc: 1,
            min_args: 1,
            max_args: Some(1),
            implementation: text::fn_t,
            volatile: false,
            ..Default::default()
        });

        // N
        self.register(FunctionDef {
            name: "N",
            iftab: Some(131),
            declared_argc: 1,
            min_args: 1,
            max_args: Some(1),
            implementation: text::fn_n,
            volatile: false,
            ..Default::default()
        });

        self.register(FunctionDef {
            name: "TEXT",
            iftab: Some(48),
            declared_argc: 2,
            min_args: 2,
            max_args: Some(2),
            implementation: text::fn_text,
            volatile: false,
            ..Default::default()
        });

        self.register(FunctionDef {
            name: "TEXTJOIN",
            min_args: 3,
            max_args: None,
            implementation: text::fn_textjoin,
            volatile: false,
            ..Default::default()
        });

        self.register(FunctionDef {
            name: "FIXED",
            iftab: Some(14),
            declared_argc: 3,
            min_args: 1,
            max_args: Some(3),
            implementation: text::fn_fixed,
            volatile: false,
            ..Default::default()
        });

        self.register(FunctionDef {
            name: "DOLLAR",
            iftab: Some(13),
            declared_argc: 2,
            min_args: 1,
            max_args: Some(2),
            implementation: text::fn_dollar,
            volatile: false,
            ..Default::default()
        });

        self.register(FunctionDef {
            name: "NUMBERVALUE",
            min_args: 1,
            max_args: Some(3),
            implementation: text::fn_numbervalue,
            volatile: false,
            ..Default::default()
        });

        // LENB (same as LEN for non-DBCS)
        self.register(FunctionDef {
            name: "LENB",
            iftab: Some(211),
            declared_argc: 1,
            min_args: 1,
            max_args: Some(1),
            implementation: text::fn_len,
            volatile: false,
            ..Default::default()
        });

        // LEFTB (same as LEFT for non-DBCS)
        self.register(FunctionDef {
            name: "LEFTB",
            iftab: Some(208),
            declared_argc: 2,
            min_args: 1,
            max_args: Some(2),
            implementation: text::fn_left,
            volatile: false,
            ..Default::default()
        });

        // RIGHTB (same as RIGHT for non-DBCS)
        self.register(FunctionDef {
            name: "RIGHTB",
            iftab: Some(209),
            declared_argc: 2,
            min_args: 1,
            max_args: Some(2),
            implementation: text::fn_right,
            volatile: false,
            ..Default::default()
        });

        // MIDB (same as MID for non-DBCS)
        self.register(FunctionDef {
            name: "MIDB",
            iftab: Some(210),
            declared_argc: 3,
            min_args: 3,
            max_args: Some(3),
            implementation: text::fn_mid,
            volatile: false,
            ..Default::default()
        });
        // text_extra functions
        self.register(FunctionDef {
            name: "REPLACE",
            iftab: Some(119),
            declared_argc: 4,
            min_args: 4,
            max_args: Some(4),
            implementation: text_extra::fn_replace,
            volatile: false,
            ..Default::default()
        });
        self.register(FunctionDef {
            name: "REPLACEB",
            iftab: Some(207),
            declared_argc: 4,
            min_args: 4,
            max_args: Some(4),
            implementation: text_extra::fn_replaceb,
            volatile: false,
            ..Default::default()
        });
        self.register(FunctionDef {
            name: "TEXTBEFORE",
            min_args: 2,
            max_args: Some(6),
            implementation: text_extra::fn_textbefore,
            volatile: false,
            ..Default::default()
        });
        self.register(FunctionDef {
            name: "TEXTAFTER",
            min_args: 2,
            max_args: Some(6),
            implementation: text_extra::fn_textafter,
            volatile: false,
            ..Default::default()
        });
        self.register(FunctionDef {
            name: "TEXTSPLIT",
            min_args: 2,
            max_args: Some(6),
            implementation: text_extra::fn_textsplit,
            volatile: false,
            ..Default::default()
        });
        self.register(FunctionDef {
            name: "UNICHAR",
            min_args: 1,
            max_args: Some(1),
            implementation: text_extra::fn_unichar,
            volatile: false,
            ..Default::default()
        });
        self.register(FunctionDef {
            name: "UNICODE",
            min_args: 1,
            max_args: Some(1),
            implementation: text_extra::fn_unicode,
            volatile: false,
            ..Default::default()
        });
        self.register(FunctionDef {
            name: "ASC",
            iftab: Some(214),
            declared_argc: 1,
            min_args: 1,
            max_args: Some(1),
            implementation: text_extra::fn_asc,
            volatile: false,
            ..Default::default()
        });
        self.register(FunctionDef {
            name: "JIS",
            min_args: 1,
            max_args: Some(1),
            implementation: text_extra::fn_jis,
            volatile: false,
            ..Default::default()
        });
        self.register(FunctionDef {
            name: "DBCS",
            iftab: Some(215),
            declared_argc: 1,
            min_args: 1,
            max_args: Some(1),
            implementation: text_extra::fn_dbcs,
            volatile: false,
            ..Default::default()
        });
        self.register(FunctionDef {
            name: "BAHTTEXT",
            iftab: Some(368),
            declared_argc: 1,
            min_args: 1,
            max_args: Some(1),
            implementation: text_extra::fn_bahttext,
            volatile: false,
            ..Default::default()
        });
        self.register(FunctionDef {
            name: "PHONETIC",
            iftab: Some(360),
            declared_argc: 1,
            min_args: 1,
            max_args: Some(1),
            implementation: text_extra::fn_phonetic,
            volatile: false,
            ..Default::default()
        });
    }

    fn register_info_functions(&mut self) {
        // ISBLANK
        self.register(FunctionDef {
            name: "ISBLANK",
            iftab: Some(129),
            declared_argc: 1,
            min_args: 1,
            max_args: Some(1),
            implementation: info::fn_isblank,
            volatile: false,
            ..Default::default()
        });

        // ISNUMBER
        self.register(FunctionDef {
            name: "ISNUMBER",
            iftab: Some(128),
            declared_argc: 1,
            min_args: 1,
            max_args: Some(1),
            implementation: info::fn_isnumber,
            volatile: false,
            ..Default::default()
        });

        // ISTEXT
        self.register(FunctionDef {
            name: "ISTEXT",
            iftab: Some(127),
            declared_argc: 1,
            min_args: 1,
            max_args: Some(1),
            implementation: info::fn_istext,
            volatile: false,
            ..Default::default()
        });

        // ISERROR
        self.register(FunctionDef {
            name: "ISERROR",
            iftab: Some(3),
            declared_argc: 1,
            min_args: 1,
            max_args: Some(1),
            implementation: info::fn_iserror,
            volatile: false,
            fixed_arity: true,
            ..Default::default()
        });

        // ISNA
        self.register(FunctionDef {
            name: "ISNA",
            iftab: Some(2),
            declared_argc: 1,
            min_args: 1,
            max_args: Some(1),
            implementation: info::fn_isna,
            volatile: false,
            fixed_arity: true,
            ..Default::default()
        });

        // NA
        self.register(FunctionDef {
            name: "NA",
            iftab: Some(10),
            declared_argc: 0,
            min_args: 0,
            max_args: Some(0),
            implementation: info::fn_na,
            volatile: false,
            ..Default::default()
        });
        // info_extra functions
        self.register(FunctionDef {
            name: "ISERR",
            iftab: Some(126),
            declared_argc: 1,
            min_args: 1,
            max_args: Some(1),
            implementation: info_extra::fn_iserr,
            volatile: false,
            ..Default::default()
        });
        self.register(FunctionDef {
            name: "ISEVEN",
            iftab: Some(420),
            declared_argc: 1,
            min_args: 1,
            max_args: Some(1),
            implementation: info_extra::fn_iseven,
            volatile: false,
            ..Default::default()
        });
        self.register(FunctionDef {
            name: "ISODD",
            iftab: Some(421),
            declared_argc: 1,
            min_args: 1,
            max_args: Some(1),
            implementation: info_extra::fn_isodd,
            volatile: false,
            ..Default::default()
        });
        self.register(FunctionDef {
            name: "ISLOGICAL",
            iftab: Some(198),
            declared_argc: 1,
            min_args: 1,
            max_args: Some(1),
            implementation: info_extra::fn_islogical,
            volatile: false,
            ..Default::default()
        });
        self.register(FunctionDef {
            name: "ISNONTEXT",
            iftab: Some(190),
            declared_argc: 1,
            min_args: 1,
            max_args: Some(1),
            implementation: info_extra::fn_isnontext,
            volatile: false,
            ..Default::default()
        });
        self.register(FunctionDef {
            name: "ISREF",
            iftab: Some(105),
            declared_argc: 1,
            min_args: 1,
            max_args: Some(1),
            implementation: info_extra::fn_isref,
            volatile: false,
            ..Default::default()
        });
        self.register(FunctionDef {
            name: "ERROR.TYPE",
            iftab: Some(261),
            declared_argc: 1,
            min_args: 1,
            max_args: Some(1),
            implementation: info_extra::fn_error_type,
            volatile: false,
            ..Default::default()
        });
        self.register(FunctionDef {
            name: "TYPE",
            iftab: Some(86),
            declared_argc: 1,
            min_args: 1,
            max_args: Some(1),
            implementation: info_extra::fn_type,
            volatile: false,
            ..Default::default()
        });
        self.register(FunctionDef {
            name: "CELL",
            iftab: Some(125),
            declared_argc: 2,
            min_args: 1,
            max_args: Some(2),
            implementation: info_extra::fn_cell,
            volatile: false,
            ..Default::default()
        });
        self.register(FunctionDef {
            name: "INFO",
            iftab: Some(244),
            declared_argc: 1,
            min_args: 1,
            max_args: Some(1),
            implementation: info_extra::fn_info,
            volatile: true,
            ..Default::default()
        });
        self.register(FunctionDef {
            name: "SHEET",
            min_args: 0,
            max_args: Some(1),
            implementation: info_extra::fn_sheet,
            volatile: false,
            ..Default::default()
        });
        self.register(FunctionDef {
            name: "SHEETS",
            min_args: 0,
            max_args: Some(1),
            implementation: info_extra::fn_sheets,
            volatile: false,
            ..Default::default()
        });
        self.register(FunctionDef {
            name: "ISFORMULA",
            min_args: 1,
            max_args: Some(1),
            implementation: info_extra::fn_isformula,
            volatile: false,
            ..Default::default()
        });
        self.register(FunctionDef {
            name: "ISOMITTED",
            min_args: 1,
            max_args: Some(1),
            implementation: info_extra::fn_isomitted,
            volatile: false,
            ..Default::default()
        });
        // Stubs: external-service-dependent functions
        self.register(FunctionDef {
            name: "STOCKHISTORY",
            min_args: 1,
            max_args: None,
            implementation: info_extra::fn_stockhistory,
            volatile: false,
            ..Default::default()
        });
        self.register(FunctionDef {
            name: "CALL",
            iftab: Some(150),
            min_args: 1,
            max_args: None,
            implementation: info_extra::fn_call,
            volatile: false,
            ..Default::default()
        });
        self.register(FunctionDef {
            name: "REGISTER.ID",
            iftab: Some(267),
            declared_argc: 3,
            min_args: 1,
            max_args: None,
            implementation: info_extra::fn_register_id,
            volatile: false,
            ..Default::default()
        });
        self.register(FunctionDef {
            name: "CUBEKPIMEMBER",
            iftab: Some(477),
            declared_argc: 4,
            min_args: 1,
            max_args: None,
            implementation: info_extra::fn_cubekpimember,
            volatile: false,
            ..Default::default()
        });
        self.register(FunctionDef {
            name: "CUBEMEMBER",
            iftab: Some(381),
            declared_argc: 3,
            min_args: 1,
            max_args: None,
            implementation: info_extra::fn_cubemember,
            volatile: false,
            ..Default::default()
        });
        self.register(FunctionDef {
            name: "CUBEMEMBERPROPERTY",
            iftab: Some(382),
            declared_argc: 3,
            min_args: 1,
            max_args: None,
            implementation: info_extra::fn_cubememberproperty,
            volatile: false,
            ..Default::default()
        });
        self.register(FunctionDef {
            name: "CUBERANKEDMEMBER",
            iftab: Some(383),
            declared_argc: 4,
            min_args: 1,
            max_args: None,
            implementation: info_extra::fn_cuberankedmember,
            volatile: false,
            ..Default::default()
        });
        self.register(FunctionDef {
            name: "CUBESET",
            iftab: Some(478),
            declared_argc: 5,
            min_args: 1,
            max_args: None,
            implementation: info_extra::fn_cubeset,
            volatile: false,
            ..Default::default()
        });
        self.register(FunctionDef {
            name: "CUBESETCOUNT",
            iftab: Some(479),
            declared_argc: 1,
            min_args: 1,
            max_args: None,
            implementation: info_extra::fn_cubesetcount,
            volatile: false,
            ..Default::default()
        });
        self.register(FunctionDef {
            name: "CUBEVALUE",
            iftab: Some(380),
            min_args: 1,
            max_args: None,
            implementation: info_extra::fn_cubevalue,
            volatile: false,
            ..Default::default()
        });
    }

    fn register_date_functions(&mut self) {
        // DATE
        self.register(FunctionDef {
            name: "DATE",
            iftab: Some(65),
            declared_argc: 3,
            min_args: 3,
            max_args: Some(3),
            implementation: date::fn_date,
            volatile: false,
            ..Default::default()
        });

        // YEAR
        self.register(FunctionDef {
            name: "YEAR",
            iftab: Some(69),
            declared_argc: 1,
            min_args: 1,
            max_args: Some(1),
            implementation: date::fn_year,
            volatile: false,
            ..Default::default()
        });

        // MONTH
        self.register(FunctionDef {
            name: "MONTH",
            iftab: Some(68),
            declared_argc: 1,
            min_args: 1,
            max_args: Some(1),
            implementation: date::fn_month,
            volatile: false,
            ..Default::default()
        });

        // DAY
        self.register(FunctionDef {
            name: "DAY",
            iftab: Some(67),
            declared_argc: 1,
            min_args: 1,
            max_args: Some(1),
            implementation: date::fn_day,
            volatile: false,
            ..Default::default()
        });

        self.register(FunctionDef {
            name: "TIME",
            iftab: Some(66),
            declared_argc: 3,
            min_args: 3,
            max_args: Some(3),
            implementation: date::fn_time,
            volatile: false,
            ..Default::default()
        });

        self.register(FunctionDef {
            name: "HOUR",
            iftab: Some(71),
            declared_argc: 1,
            min_args: 1,
            max_args: Some(1),
            implementation: date::fn_hour,
            volatile: false,
            ..Default::default()
        });

        self.register(FunctionDef {
            name: "MINUTE",
            iftab: Some(72),
            declared_argc: 1,
            min_args: 1,
            max_args: Some(1),
            implementation: date::fn_minute,
            volatile: false,
            ..Default::default()
        });

        self.register(FunctionDef {
            name: "SECOND",
            iftab: Some(73),
            declared_argc: 1,
            min_args: 1,
            max_args: Some(1),
            implementation: date::fn_second,
            volatile: false,
            ..Default::default()
        });

        self.register(FunctionDef {
            name: "WEEKDAY",
            iftab: Some(70),
            declared_argc: 2,
            min_args: 1,
            max_args: Some(2),
            implementation: date::fn_weekday,
            volatile: false,
            ..Default::default()
        });

        self.register(FunctionDef {
            name: "WEEKNUM",
            iftab: Some(465),
            declared_argc: 2,
            min_args: 1,
            max_args: Some(2),
            implementation: date::fn_weeknum,
            volatile: false,
            ..Default::default()
        });

        self.register(FunctionDef {
            name: "ISOWEEKNUM",
            min_args: 1,
            max_args: Some(1),
            implementation: date::fn_isoweeknum,
            volatile: false,
            ..Default::default()
        });

        self.register(FunctionDef {
            name: "EDATE",
            iftab: Some(449),
            declared_argc: 2,
            min_args: 2,
            max_args: Some(2),
            implementation: date::fn_edate,
            volatile: false,
            ..Default::default()
        });

        self.register(FunctionDef {
            name: "EOMONTH",
            iftab: Some(450),
            declared_argc: 2,
            min_args: 2,
            max_args: Some(2),
            implementation: date::fn_eomonth,
            volatile: false,
            ..Default::default()
        });

        self.register(FunctionDef {
            name: "DAYS",
            min_args: 2,
            max_args: Some(2),
            implementation: date::fn_days,
            volatile: false,
            ..Default::default()
        });

        self.register(FunctionDef {
            name: "DAYS360",
            iftab: Some(220),
            declared_argc: 3,
            min_args: 2,
            max_args: Some(3),
            implementation: date::fn_days360,
            volatile: false,
            ..Default::default()
        });

        self.register(FunctionDef {
            name: "DATEDIF",
            iftab: Some(351),
            declared_argc: 3,
            min_args: 3,
            max_args: Some(3),
            implementation: date::fn_datedif,
            volatile: false,
            ..Default::default()
        });

        self.register(FunctionDef {
            name: "YEARFRAC",
            iftab: Some(451),
            declared_argc: 3,
            min_args: 2,
            max_args: Some(3),
            implementation: date::fn_yearfrac,
            volatile: false,
            ..Default::default()
        });

        self.register(FunctionDef {
            name: "DATEVALUE",
            iftab: Some(140),
            declared_argc: 1,
            min_args: 1,
            max_args: Some(1),
            implementation: date::fn_datevalue,
            volatile: false,
            ..Default::default()
        });

        self.register(FunctionDef {
            name: "TIMEVALUE",
            iftab: Some(141),
            declared_argc: 1,
            min_args: 1,
            max_args: Some(1),
            implementation: date::fn_timevalue,
            volatile: false,
            ..Default::default()
        });

        self.register(FunctionDef {
            name: "NETWORKDAYS",
            iftab: Some(472),
            declared_argc: 3,
            min_args: 2,
            max_args: Some(3),
            implementation: date::fn_networkdays,
            volatile: false,
            ..Default::default()
        });

        self.register(FunctionDef {
            name: "WORKDAY",
            iftab: Some(471),
            declared_argc: 3,
            min_args: 2,
            max_args: Some(3),
            implementation: date::fn_workday,
            volatile: false,
            ..Default::default()
        });

        // NOW (volatile)
        self.register(FunctionDef {
            name: "NOW",
            iftab: Some(74),
            declared_argc: 0,
            min_args: 0,
            max_args: Some(0),
            implementation: date::fn_now,
            volatile: true,
            fixed_arity: true,
            ..Default::default()
        });

        // TODAY (volatile)
        self.register(FunctionDef {
            name: "TODAY",
            iftab: Some(221),
            declared_argc: 0,
            min_args: 0,
            max_args: Some(0),
            implementation: date::fn_today,
            volatile: true,
            fixed_arity: true,
            ..Default::default()
        });
        // date functions from lookup_extra
        self.register(FunctionDef {
            name: "NETWORKDAYS.INTL",
            min_args: 2,
            max_args: Some(4),
            implementation: lookup_extra::fn_networkdays_intl,
            volatile: false,
            ..Default::default()
        });
        self.register(FunctionDef {
            name: "WORKDAY.INTL",
            min_args: 2,
            max_args: Some(4),
            implementation: lookup_extra::fn_workday_intl,
            volatile: false,
            ..Default::default()
        });
    }

    fn register_lookup_functions(&mut self) {
        // INDEX
        self.register(FunctionDef {
            name: "INDEX",
            iftab: Some(29),
            declared_argc: 4,
            min_args: 2,
            max_args: Some(3),
            implementation: lookup::fn_index,
            volatile: false,
            ..Default::default()
        });

        // MATCH
        self.register(FunctionDef {
            name: "MATCH",
            iftab: Some(64),
            declared_argc: 3,
            min_args: 2,
            max_args: Some(3),
            implementation: lookup::fn_match,
            volatile: false,
            ..Default::default()
        });

        // VLOOKUP
        self.register(FunctionDef {
            name: "VLOOKUP",
            iftab: Some(102),
            declared_argc: 4,
            min_args: 3,
            max_args: Some(4),
            implementation: lookup::fn_vlookup,
            volatile: false,
            ..Default::default()
        });

        // ROWS
        self.register(FunctionDef {
            name: "ROWS",
            iftab: Some(76),
            declared_argc: 1,
            min_args: 1,
            max_args: Some(1),
            implementation: lookup::fn_rows,
            volatile: false,
            ..Default::default()
        });

        // COLUMNS
        self.register(FunctionDef {
            name: "COLUMNS",
            iftab: Some(77),
            declared_argc: 1,
            min_args: 1,
            max_args: Some(1),
            implementation: lookup::fn_columns,
            volatile: false,
            ..Default::default()
        });

        // CHOOSE
        self.register(FunctionDef {
            name: "CHOOSE",
            iftab: Some(100),
            min_args: 2,
            max_args: None, // Up to 254 values
            implementation: lookup::fn_choose,
            volatile: false,
            ..Default::default()
        });

        // ROW
        self.register(FunctionDef {
            name: "ROW",
            iftab: Some(8),
            declared_argc: 1,
            min_args: 0,
            max_args: Some(1),
            implementation: lookup::fn_row,
            volatile: false,
            arg_classes: &[OperandClass::R],
            ..Default::default()
        });

        // COLUMN
        self.register(FunctionDef {
            name: "COLUMN",
            iftab: Some(9),
            declared_argc: 1,
            min_args: 0,
            max_args: Some(1),
            implementation: lookup::fn_column,
            volatile: false,
            arg_classes: &[OperandClass::R],
            ..Default::default()
        });

        // SEQUENCE (dynamic array function)
        self.register(FunctionDef {
            name: "SEQUENCE",
            min_args: 1,
            max_args: Some(4),
            implementation: lookup::fn_sequence,
            volatile: false,
            ..Default::default()
        });

        self.register(FunctionDef {
            name: "HLOOKUP",
            iftab: Some(101),
            declared_argc: 4,
            min_args: 3,
            max_args: Some(4),
            implementation: lookup::fn_hlookup,
            volatile: false,
            ..Default::default()
        });

        self.register(FunctionDef {
            name: "XLOOKUP",
            min_args: 3,
            max_args: Some(6),
            implementation: lookup::fn_xlookup,
            volatile: false,
            ..Default::default()
        });

        self.register(FunctionDef {
            name: "XMATCH",
            min_args: 2,
            max_args: Some(4),
            implementation: lookup::fn_xmatch,
            volatile: false,
            ..Default::default()
        });

        self.register(FunctionDef {
            name: "INDIRECT",
            iftab: Some(148),
            declared_argc: 2,
            min_args: 1,
            max_args: Some(2),
            implementation: lookup::fn_indirect,
            volatile: true,
            ..Default::default()
        });

        self.register(FunctionDef {
            name: "OFFSET",
            iftab: Some(78),
            declared_argc: 5,
            min_args: 3,
            max_args: Some(5),
            implementation: lookup::fn_offset,
            volatile: true,
            arg_classes: &[OperandClass::R],
            ..Default::default()
        });
        // lookup_extra functions
        self.register(FunctionDef {
            name: "ADDRESS",
            iftab: Some(219),
            declared_argc: 5,
            min_args: 2,
            max_args: Some(5),
            implementation: lookup_extra::fn_address,
            volatile: false,
            ..Default::default()
        });
        self.register(FunctionDef {
            name: "AREAS",
            iftab: Some(75),
            declared_argc: 1,
            min_args: 1,
            max_args: Some(1),
            implementation: lookup_extra::fn_areas,
            volatile: false,
            arg_classes: &[OperandClass::R],
            ..Default::default()
        });
        self.register(FunctionDef {
            name: "CHOOSECOLS",
            min_args: 2,
            max_args: None,
            implementation: lookup_extra::fn_choosecols,
            volatile: false,
            ..Default::default()
        });
        self.register(FunctionDef {
            name: "CHOOSEROWS",
            min_args: 2,
            max_args: None,
            implementation: lookup_extra::fn_chooserows,
            volatile: false,
            ..Default::default()
        });
        self.register(FunctionDef {
            name: "DROP",
            min_args: 2,
            max_args: Some(3),
            implementation: lookup_extra::fn_drop,
            volatile: false,
            ..Default::default()
        });
        self.register(FunctionDef {
            name: "EXPAND",
            min_args: 2,
            max_args: Some(4),
            implementation: lookup_extra::fn_expand,
            volatile: false,
            ..Default::default()
        });
        self.register(FunctionDef {
            name: "FILTER",
            min_args: 2,
            max_args: Some(3),
            implementation: lookup_extra::fn_filter,
            volatile: false,
            ..Default::default()
        });
        self.register(FunctionDef {
            name: "FORMULATEXT",
            min_args: 1,
            max_args: Some(1),
            implementation: lookup_extra::fn_formulatext,
            volatile: false,
            ..Default::default()
        });
        self.register(FunctionDef {
            name: "HSTACK",
            min_args: 1,
            max_args: None,
            implementation: lookup_extra::fn_hstack,
            volatile: false,
            ..Default::default()
        });
        self.register(FunctionDef {
            name: "LOOKUP",
            iftab: Some(28),
            declared_argc: 3,
            min_args: 2,
            max_args: Some(3),
            implementation: lookup_extra::fn_lookup,
            volatile: false,
            ..Default::default()
        });
        self.register(FunctionDef {
            name: "SORT",
            min_args: 1,
            max_args: Some(4),
            implementation: lookup_extra::fn_sort,
            volatile: false,
            ..Default::default()
        });
        self.register(FunctionDef {
            name: "SORTBY",
            min_args: 2,
            max_args: None,
            implementation: lookup_extra::fn_sortby,
            volatile: false,
            ..Default::default()
        });
        self.register(FunctionDef {
            name: "TAKE",
            min_args: 2,
            max_args: Some(3),
            implementation: lookup_extra::fn_take,
            volatile: false,
            ..Default::default()
        });
        self.register(FunctionDef {
            name: "TRIMRANGE",
            min_args: 1,
            max_args: Some(3),
            implementation: lookup_extra::fn_trimrange,
            volatile: false,
            ..Default::default()
        });
        self.register(FunctionDef {
            name: "TOCOL",
            min_args: 1,
            max_args: Some(3),
            implementation: lookup_extra::fn_tocol,
            volatile: false,
            ..Default::default()
        });
        self.register(FunctionDef {
            name: "TOROW",
            min_args: 1,
            max_args: Some(3),
            implementation: lookup_extra::fn_torow,
            volatile: false,
            ..Default::default()
        });
        self.register(FunctionDef {
            name: "TRANSPOSE",
            iftab: Some(83),
            declared_argc: 1,
            min_args: 1,
            max_args: Some(1),
            implementation: lookup_extra::fn_transpose,
            volatile: false,
            ..Default::default()
        });
        self.register(FunctionDef {
            name: "UNIQUE",
            min_args: 1,
            max_args: Some(3),
            implementation: lookup_extra::fn_unique,
            volatile: false,
            ..Default::default()
        });
        self.register(FunctionDef {
            name: "VSTACK",
            min_args: 1,
            max_args: None,
            implementation: lookup_extra::fn_vstack,
            volatile: false,
            ..Default::default()
        });
        self.register(FunctionDef {
            name: "WRAPCOLS",
            min_args: 2,
            max_args: Some(3),
            implementation: lookup_extra::fn_wrapcols,
            volatile: false,
            ..Default::default()
        });
        self.register(FunctionDef {
            name: "WRAPROWS",
            min_args: 2,
            max_args: Some(3),
            implementation: lookup_extra::fn_wraprows,
            volatile: false,
            ..Default::default()
        });
        self.register(FunctionDef {
            name: "HYPERLINK",
            iftab: Some(359),
            declared_argc: 2,
            min_args: 1,
            max_args: Some(2),
            implementation: lookup_extra::fn_hyperlink,
            volatile: false,
            ..Default::default()
        });
        self.register(FunctionDef {
            name: "GETPIVOTDATA",
            iftab: Some(358),
            declared_argc: 128,
            min_args: 2,
            max_args: None,
            implementation: lookup_extra::fn_getpivotdata,
            volatile: false,
            ..Default::default()
        });
        self.register(FunctionDef {
            name: "RTD",
            iftab: Some(379),
            min_args: 2,
            max_args: None,
            implementation: lookup_extra::fn_rtd,
            volatile: false,
            ..Default::default()
        });
        self.register(FunctionDef {
            name: "IMAGE",
            min_args: 1,
            max_args: Some(5),
            implementation: lookup_extra::fn_image,
            volatile: false,
            ..Default::default()
        });
    }

    fn register_statistical_functions(&mut self) {
        // COUNTA
        self.register(FunctionDef {
            name: "COUNTA",
            iftab: Some(169),
            min_args: 1,
            max_args: None,
            implementation: statistical::fn_counta,
            volatile: false,
            default_arg_class: OperandClass::R,
            ..Default::default()
        });

        // COUNTBLANK
        self.register(FunctionDef {
            name: "COUNTBLANK",
            iftab: Some(347),
            declared_argc: 1,
            min_args: 1,
            max_args: None,
            implementation: statistical::fn_countblank,
            volatile: false,
            ..Default::default()
        });

        // COUNTIF
        self.register(FunctionDef {
            name: "COUNTIF",
            iftab: Some(346),
            declared_argc: 2,
            min_args: 2,
            max_args: Some(2),
            implementation: statistical::fn_countif,
            volatile: false,
            ..Default::default()
        });

        // AVERAGEIF
        self.register(FunctionDef {
            name: "AVERAGEIF",
            iftab: Some(483),
            declared_argc: 3,
            min_args: 2,
            max_args: Some(3),
            implementation: statistical::fn_averageif,
            volatile: false,
            ..Default::default()
        });

        self.register(FunctionDef {
            name: "STDEV.S",
            min_args: 1,
            max_args: None,
            implementation: statistical::fn_stdev_s,
            volatile: false,
            ..Default::default()
        });

        self.register(FunctionDef {
            name: "STDEV.P",
            min_args: 1,
            max_args: None,
            implementation: statistical::fn_stdev_p,
            volatile: false,
            ..Default::default()
        });

        self.register(FunctionDef {
            name: "VAR.S",
            min_args: 1,
            max_args: None,
            implementation: statistical::fn_var_s,
            volatile: false,
            ..Default::default()
        });

        self.register(FunctionDef {
            name: "VAR.P",
            min_args: 1,
            max_args: None,
            implementation: statistical::fn_var_p,
            volatile: false,
            ..Default::default()
        });

        self.register(FunctionDef {
            name: "MODE.SNGL",
            min_args: 1,
            max_args: None,
            implementation: statistical::fn_mode_sngl,
            volatile: false,
            ..Default::default()
        });

        self.register(FunctionDef {
            name: "MAXIFS",
            min_args: 3,
            max_args: None,
            implementation: statistical::fn_maxifs,
            volatile: false,
            ..Default::default()
        });

        self.register(FunctionDef {
            name: "MINIFS",
            min_args: 3,
            max_args: None,
            implementation: statistical::fn_minifs,
            volatile: false,
            ..Default::default()
        });

        self.register(FunctionDef {
            name: "RANK.EQ",
            min_args: 2,
            max_args: Some(3),
            implementation: statistical::fn_rank_eq,
            volatile: false,
            ..Default::default()
        });

        self.register(FunctionDef {
            name: "RANK.AVG",
            min_args: 2,
            max_args: Some(3),
            implementation: statistical::fn_rank_avg,
            volatile: false,
            ..Default::default()
        });

        self.register(FunctionDef {
            name: "PERCENTILE.INC",
            min_args: 2,
            max_args: Some(2),
            implementation: statistical::fn_percentile_inc,
            volatile: false,
            ..Default::default()
        });

        self.register(FunctionDef {
            name: "PERCENTILE.EXC",
            min_args: 2,
            max_args: Some(2),
            implementation: statistical::fn_percentile_exc,
            volatile: false,
            ..Default::default()
        });

        self.register(FunctionDef {
            name: "QUARTILE.INC",
            min_args: 2,
            max_args: Some(2),
            implementation: statistical::fn_quartile_inc,
            volatile: false,
            ..Default::default()
        });

        self.register(FunctionDef {
            name: "QUARTILE.EXC",
            min_args: 2,
            max_args: Some(2),
            implementation: statistical::fn_quartile_exc,
            volatile: false,
            ..Default::default()
        });

        self.register(FunctionDef {
            name: "PERCENTRANK.INC",
            min_args: 2,
            max_args: Some(3),
            implementation: statistical::fn_percentrank_inc,
            volatile: false,
            ..Default::default()
        });

        self.register(FunctionDef {
            name: "PERCENTRANK.EXC",
            min_args: 2,
            max_args: Some(3),
            implementation: statistical::fn_percentrank_exc,
            volatile: false,
            ..Default::default()
        });

        self.register(FunctionDef {
            name: "STDEV",
            iftab: Some(12),
            min_args: 1,
            max_args: None,
            implementation: statistical::fn_stdev,
            volatile: false,
            default_arg_class: OperandClass::R,
            ..Default::default()
        });

        self.register(FunctionDef {
            name: "STDEVP",
            iftab: Some(193),
            min_args: 1,
            max_args: None,
            implementation: statistical::fn_stdevp,
            volatile: false,
            ..Default::default()
        });

        self.register(FunctionDef {
            name: "VAR",
            iftab: Some(46),
            min_args: 1,
            max_args: None,
            implementation: statistical::fn_var,
            volatile: false,
            default_arg_class: OperandClass::R,
            ..Default::default()
        });

        self.register(FunctionDef {
            name: "VARP",
            iftab: Some(194),
            min_args: 1,
            max_args: None,
            implementation: statistical::fn_varp,
            volatile: false,
            ..Default::default()
        });

        self.register(FunctionDef {
            name: "MODE",
            iftab: Some(330),
            min_args: 1,
            max_args: None,
            implementation: statistical::fn_mode,
            volatile: false,
            ..Default::default()
        });

        self.register(FunctionDef {
            name: "PERCENTILE",
            iftab: Some(328),
            declared_argc: 2,
            min_args: 2,
            max_args: Some(2),
            implementation: statistical::fn_percentile,
            volatile: false,
            ..Default::default()
        });

        self.register(FunctionDef {
            name: "QUARTILE",
            iftab: Some(327),
            declared_argc: 2,
            min_args: 2,
            max_args: Some(2),
            implementation: statistical::fn_quartile,
            volatile: false,
            ..Default::default()
        });

        self.register(FunctionDef {
            name: "RANK",
            iftab: Some(216),
            declared_argc: 3,
            min_args: 2,
            max_args: Some(3),
            implementation: statistical::fn_rank,
            volatile: false,
            ..Default::default()
        });

        self.register(FunctionDef {
            name: "PERCENTRANK",
            iftab: Some(329),
            declared_argc: 3,
            min_args: 2,
            max_args: Some(3),
            implementation: statistical::fn_percentrank,
            volatile: false,
            ..Default::default()
        });

        // MEDIAN
        self.register(FunctionDef {
            name: "MEDIAN",
            iftab: Some(227),
            min_args: 1,
            max_args: None,
            implementation: statistical::fn_median,
            volatile: false,
            ..Default::default()
        });

        // LARGE
        self.register(FunctionDef {
            name: "LARGE",
            iftab: Some(325),
            declared_argc: 2,
            min_args: 2,
            max_args: Some(2),
            implementation: statistical::fn_large,
            volatile: false,
            ..Default::default()
        });

        // SMALL
        self.register(FunctionDef {
            name: "SMALL",
            iftab: Some(326),
            declared_argc: 2,
            min_args: 2,
            max_args: Some(2),
            implementation: statistical::fn_small,
            volatile: false,
            ..Default::default()
        });

        // COUNTIFS
        self.register(FunctionDef {
            name: "COUNTIFS",
            iftab: Some(481),
            declared_argc: 128,
            min_args: 2,
            max_args: None, // Up to 127 criteria pairs
            implementation: statistical::fn_countifs,
            volatile: false,
            ..Default::default()
        });

        // AVERAGEIFS
        self.register(FunctionDef {
            name: "AVERAGEIFS",
            iftab: Some(484),
            declared_argc: 129,
            min_args: 3,
            max_args: None, // avg_range + up to 127 criteria pairs
            implementation: statistical::fn_averageifs,
            volatile: false,
            ..Default::default()
        });

        self.register(FunctionDef {
            name: "NORM.DIST",
            min_args: 4,
            max_args: Some(4),
            implementation: statistical::fn_norm_dist,
            volatile: false,
            ..Default::default()
        });
        self.register(FunctionDef {
            name: "NORM.S.DIST",
            min_args: 2,
            max_args: Some(2),
            implementation: statistical::fn_norm_s_dist,
            volatile: false,
            ..Default::default()
        });
        self.register(FunctionDef {
            name: "NORM.INV",
            min_args: 3,
            max_args: Some(3),
            implementation: statistical::fn_norm_inv,
            volatile: false,
            ..Default::default()
        });
        self.register(FunctionDef {
            name: "NORM.S.INV",
            min_args: 1,
            max_args: Some(1),
            implementation: statistical::fn_norm_s_inv,
            volatile: false,
            ..Default::default()
        });
        self.register(FunctionDef {
            name: "PHI",
            min_args: 1,
            max_args: Some(1),
            implementation: statistical::fn_phi,
            volatile: false,
            ..Default::default()
        });
        self.register(FunctionDef {
            name: "BINOM.DIST",
            min_args: 4,
            max_args: Some(4),
            implementation: statistical::fn_binom_dist,
            volatile: false,
            ..Default::default()
        });
        self.register(FunctionDef {
            name: "BINOM.DIST.RANGE",
            min_args: 3,
            max_args: Some(4),
            implementation: statistical::fn_binom_dist_range,
            volatile: false,
            ..Default::default()
        });
        self.register(FunctionDef {
            name: "BINOM.INV",
            min_args: 3,
            max_args: Some(3),
            implementation: statistical::fn_binom_inv,
            volatile: false,
            ..Default::default()
        });
        self.register(FunctionDef {
            name: "CHISQ.DIST",
            min_args: 3,
            max_args: Some(3),
            implementation: statistical::fn_chisq_dist,
            volatile: false,
            ..Default::default()
        });
        self.register(FunctionDef {
            name: "CHISQ.DIST.RT",
            min_args: 2,
            max_args: Some(2),
            implementation: statistical::fn_chisq_dist_rt,
            volatile: false,
            ..Default::default()
        });
        self.register(FunctionDef {
            name: "CHISQ.INV",
            min_args: 2,
            max_args: Some(2),
            implementation: statistical::fn_chisq_inv,
            volatile: false,
            ..Default::default()
        });
        self.register(FunctionDef {
            name: "CHISQ.INV.RT",
            min_args: 2,
            max_args: Some(2),
            implementation: statistical::fn_chisq_inv_rt,
            volatile: false,
            ..Default::default()
        });
        self.register(FunctionDef {
            name: "CHISQ.TEST",
            min_args: 2,
            max_args: Some(2),
            implementation: statistical::fn_chisq_test,
            volatile: false,
            ..Default::default()
        });
        self.register(FunctionDef {
            name: "T.DIST",
            min_args: 3,
            max_args: Some(3),
            implementation: statistical::fn_t_dist,
            volatile: false,
            ..Default::default()
        });
        self.register(FunctionDef {
            name: "T.DIST.2T",
            min_args: 2,
            max_args: Some(2),
            implementation: statistical::fn_t_dist_2t,
            volatile: false,
            ..Default::default()
        });
        self.register(FunctionDef {
            name: "T.DIST.RT",
            min_args: 2,
            max_args: Some(2),
            implementation: statistical::fn_t_dist_rt,
            volatile: false,
            ..Default::default()
        });
        self.register(FunctionDef {
            name: "T.INV",
            min_args: 2,
            max_args: Some(2),
            implementation: statistical::fn_t_inv,
            volatile: false,
            ..Default::default()
        });
        self.register(FunctionDef {
            name: "T.INV.2T",
            min_args: 2,
            max_args: Some(2),
            implementation: statistical::fn_t_inv_2t,
            volatile: false,
            ..Default::default()
        });
        self.register(FunctionDef {
            name: "T.TEST",
            min_args: 4,
            max_args: Some(4),
            implementation: statistical::fn_t_test,
            volatile: false,
            ..Default::default()
        });
        self.register(FunctionDef {
            name: "F.DIST",
            min_args: 4,
            max_args: Some(4),
            implementation: statistical::fn_f_dist,
            volatile: false,
            ..Default::default()
        });
        self.register(FunctionDef {
            name: "F.DIST.RT",
            min_args: 3,
            max_args: Some(3),
            implementation: statistical::fn_f_dist_rt,
            volatile: false,
            ..Default::default()
        });
        self.register(FunctionDef {
            name: "F.INV",
            min_args: 3,
            max_args: Some(3),
            implementation: statistical::fn_f_inv,
            volatile: false,
            ..Default::default()
        });
        self.register(FunctionDef {
            name: "F.INV.RT",
            min_args: 3,
            max_args: Some(3),
            implementation: statistical::fn_f_inv_rt,
            volatile: false,
            ..Default::default()
        });
        self.register(FunctionDef {
            name: "F.TEST",
            min_args: 2,
            max_args: Some(2),
            implementation: statistical::fn_f_test,
            volatile: false,
            ..Default::default()
        });
        self.register(FunctionDef {
            name: "GAMMA",
            min_args: 1,
            max_args: Some(1),
            implementation: statistical::fn_gamma,
            volatile: false,
            ..Default::default()
        });
        self.register(FunctionDef {
            name: "GAMMA.DIST",
            min_args: 4,
            max_args: Some(4),
            implementation: statistical::fn_gamma_dist,
            volatile: false,
            ..Default::default()
        });
        self.register(FunctionDef {
            name: "GAMMA.INV",
            min_args: 3,
            max_args: Some(3),
            implementation: statistical::fn_gamma_inv,
            volatile: false,
            ..Default::default()
        });
        self.register(FunctionDef {
            name: "GAMMALN",
            iftab: Some(271),
            declared_argc: 1,
            min_args: 1,
            max_args: Some(1),
            implementation: statistical::fn_gammaln,
            volatile: false,
            ..Default::default()
        });
        self.register(FunctionDef {
            name: "GAMMALN.PRECISE",
            min_args: 1,
            max_args: Some(1),
            implementation: statistical::fn_gammaln_precise,
            volatile: false,
            ..Default::default()
        });
        self.register(FunctionDef {
            name: "BETA.DIST",
            min_args: 4,
            max_args: Some(6),
            implementation: statistical::fn_beta_dist,
            volatile: false,
            ..Default::default()
        });
        self.register(FunctionDef {
            name: "EXPON.DIST",
            min_args: 3,
            max_args: Some(3),
            implementation: statistical::fn_expon_dist,
            volatile: false,
            ..Default::default()
        });
        self.register(FunctionDef {
            name: "HYPGEOM.DIST",
            min_args: 4,
            max_args: Some(5),
            implementation: statistical::fn_hypgeom_dist,
            volatile: false,
            ..Default::default()
        });
        self.register(FunctionDef {
            name: "NEGBINOM.DIST",
            min_args: 3,
            max_args: Some(4),
            implementation: statistical::fn_negbinom_dist,
            volatile: false,
            ..Default::default()
        });
        self.register(FunctionDef {
            name: "POISSON.DIST",
            min_args: 3,
            max_args: Some(3),
            implementation: statistical::fn_poisson_dist,
            volatile: false,
            ..Default::default()
        });
        self.register(FunctionDef {
            name: "WEIBULL.DIST",
            min_args: 4,
            max_args: Some(4),
            implementation: statistical::fn_weibull_dist,
            volatile: false,
            ..Default::default()
        });

        self.register(FunctionDef {
            name: "AVEDEV",
            iftab: Some(269),
            min_args: 1,
            max_args: None,
            implementation: statistical::fn_avedev,
            volatile: false,
            ..Default::default()
        });
        self.register(FunctionDef {
            name: "AVERAGEA",
            iftab: Some(361),
            min_args: 1,
            max_args: None,
            implementation: statistical::fn_averagea,
            volatile: false,
            ..Default::default()
        });
        self.register(FunctionDef {
            name: "DEVSQ",
            iftab: Some(318),
            min_args: 1,
            max_args: None,
            implementation: statistical::fn_devsq,
            volatile: false,
            ..Default::default()
        });
        self.register(FunctionDef {
            name: "GEOMEAN",
            iftab: Some(319),
            min_args: 1,
            max_args: None,
            implementation: statistical::fn_geomean,
            volatile: false,
            ..Default::default()
        });
        self.register(FunctionDef {
            name: "HARMEAN",
            iftab: Some(320),
            min_args: 1,
            max_args: None,
            implementation: statistical::fn_harmean,
            volatile: false,
            ..Default::default()
        });
        self.register(FunctionDef {
            name: "KURT",
            iftab: Some(322),
            min_args: 1,
            max_args: None,
            implementation: statistical::fn_kurt,
            volatile: false,
            ..Default::default()
        });
        self.register(FunctionDef {
            name: "SKEW",
            iftab: Some(323),
            min_args: 1,
            max_args: None,
            implementation: statistical::fn_skew,
            volatile: false,
            ..Default::default()
        });
        self.register(FunctionDef {
            name: "SKEW.P",
            min_args: 1,
            max_args: None,
            implementation: statistical::fn_skew_p,
            volatile: false,
            ..Default::default()
        });
        self.register(FunctionDef {
            name: "TRIMMEAN",
            iftab: Some(331),
            declared_argc: 2,
            min_args: 2,
            max_args: Some(2),
            implementation: statistical::fn_trimmean,
            volatile: false,
            ..Default::default()
        });
        self.register(FunctionDef {
            name: "STANDARDIZE",
            iftab: Some(297),
            declared_argc: 3,
            min_args: 3,
            max_args: Some(3),
            implementation: statistical::fn_standardize,
            volatile: false,
            ..Default::default()
        });
        self.register(FunctionDef {
            name: "CORREL",
            iftab: Some(307),
            declared_argc: 2,
            min_args: 2,
            max_args: Some(2),
            implementation: statistical::fn_correl,
            volatile: false,
            ..Default::default()
        });
        self.register(FunctionDef {
            name: "COVARIANCE.P",
            min_args: 2,
            max_args: Some(2),
            implementation: statistical::fn_covariance_p,
            volatile: false,
            ..Default::default()
        });
        self.register(FunctionDef {
            name: "COVARIANCE.S",
            min_args: 2,
            max_args: Some(2),
            implementation: statistical::fn_covariance_s,
            volatile: false,
            ..Default::default()
        });
        self.register(FunctionDef {
            name: "PEARSON",
            iftab: Some(312),
            declared_argc: 2,
            min_args: 2,
            max_args: Some(2),
            implementation: statistical::fn_pearson,
            volatile: false,
            ..Default::default()
        });
        self.register(FunctionDef {
            name: "RSQ",
            iftab: Some(313),
            declared_argc: 2,
            min_args: 2,
            max_args: Some(2),
            implementation: statistical::fn_rsq,
            volatile: false,
            ..Default::default()
        });
        self.register(FunctionDef {
            name: "SLOPE",
            iftab: Some(315),
            declared_argc: 2,
            min_args: 2,
            max_args: Some(2),
            implementation: statistical::fn_slope,
            volatile: false,
            ..Default::default()
        });
        self.register(FunctionDef {
            name: "INTERCEPT",
            iftab: Some(311),
            declared_argc: 2,
            min_args: 2,
            max_args: Some(2),
            implementation: statistical::fn_intercept,
            volatile: false,
            ..Default::default()
        });
        self.register(FunctionDef {
            name: "FISHER",
            iftab: Some(283),
            declared_argc: 1,
            min_args: 1,
            max_args: Some(1),
            implementation: statistical::fn_fisher,
            volatile: false,
            ..Default::default()
        });
        self.register(FunctionDef {
            name: "FISHERINV",
            iftab: Some(284),
            declared_argc: 1,
            min_args: 1,
            max_args: Some(1),
            implementation: statistical::fn_fisherinv,
            volatile: false,
            ..Default::default()
        });
        self.register(FunctionDef {
            name: "FORECAST.LINEAR",
            min_args: 3,
            max_args: Some(3),
            implementation: statistical::fn_forecast_linear,
            volatile: false,
            ..Default::default()
        });
        self.register(FunctionDef {
            name: "FORECAST",
            iftab: Some(309),
            declared_argc: 3,
            min_args: 3,
            max_args: Some(3),
            implementation: statistical::fn_forecast,
            volatile: false,
            ..Default::default()
        });
        self.register(FunctionDef {
            name: "FREQUENCY",
            iftab: Some(252),
            declared_argc: 2,
            min_args: 2,
            max_args: Some(2),
            implementation: statistical::fn_frequency,
            volatile: false,
            ..Default::default()
        });
        self.register(FunctionDef {
            name: "MAXA",
            iftab: Some(362),
            min_args: 1,
            max_args: None,
            implementation: statistical::fn_maxa,
            volatile: false,
            ..Default::default()
        });
        self.register(FunctionDef {
            name: "MINA",
            iftab: Some(363),
            min_args: 1,
            max_args: None,
            implementation: statistical::fn_mina,
            volatile: false,
            ..Default::default()
        });
        self.register(FunctionDef {
            name: "STEYX",
            iftab: Some(314),
            declared_argc: 2,
            min_args: 2,
            max_args: Some(2),
            implementation: statistical::fn_steyx,
            volatile: false,
            ..Default::default()
        });
        self.register(FunctionDef {
            name: "PROB",
            iftab: Some(317),
            declared_argc: 4,
            min_args: 3,
            max_args: Some(4),
            implementation: statistical::fn_prob,
            volatile: false,
            ..Default::default()
        });
        self.register(FunctionDef {
            name: "PERMUT",
            iftab: Some(299),
            declared_argc: 2,
            min_args: 2,
            max_args: Some(2),
            implementation: statistical::fn_permut,
            volatile: false,
            ..Default::default()
        });
        self.register(FunctionDef {
            name: "PERMUTATIONA",
            min_args: 2,
            max_args: Some(2),
            implementation: statistical::fn_permutationa,
            volatile: false,
            ..Default::default()
        });
        self.register(FunctionDef {
            name: "CONFIDENCE.NORM",
            min_args: 3,
            max_args: Some(3),
            implementation: statistical::fn_confidence_norm,
            volatile: false,
            ..Default::default()
        });
        self.register(FunctionDef {
            name: "CONFIDENCE.T",
            min_args: 3,
            max_args: Some(3),
            implementation: statistical::fn_confidence_t,
            volatile: false,
            ..Default::default()
        });
        self.register(FunctionDef {
            name: "GAUSS",
            min_args: 1,
            max_args: Some(1),
            implementation: statistical::fn_gauss,
            volatile: false,
            ..Default::default()
        });
        self.register(FunctionDef {
            name: "MODE.MULT",
            min_args: 1,
            max_args: None,
            implementation: statistical::fn_mode_mult,
            volatile: false,
            ..Default::default()
        });
        self.register(FunctionDef {
            name: "STDEVA",
            iftab: Some(366),
            min_args: 1,
            max_args: None,
            implementation: statistical::fn_stdeva,
            volatile: false,
            ..Default::default()
        });
        self.register(FunctionDef {
            name: "STDEVPA",
            iftab: Some(364),
            min_args: 1,
            max_args: None,
            implementation: statistical::fn_stdevpa,
            volatile: false,
            ..Default::default()
        });
        self.register(FunctionDef {
            name: "VARA",
            iftab: Some(367),
            min_args: 1,
            max_args: None,
            implementation: statistical::fn_vara,
            volatile: false,
            ..Default::default()
        });
        self.register(FunctionDef {
            name: "VARPA",
            iftab: Some(365),
            min_args: 1,
            max_args: None,
            implementation: statistical::fn_varpa,
            volatile: false,
            ..Default::default()
        });
        // statistical_extra functions
        self.register(FunctionDef {
            name: "LOGNORM.DIST",
            min_args: 4,
            max_args: Some(4),
            implementation: statistical_extra::fn_lognorm_dist,
            volatile: false,
            ..Default::default()
        });
        self.register(FunctionDef {
            name: "LOGNORM.INV",
            min_args: 3,
            max_args: Some(3),
            implementation: statistical_extra::fn_lognorm_inv,
            volatile: false,
            ..Default::default()
        });
        self.register(FunctionDef {
            name: "LINEST",
            iftab: Some(49),
            declared_argc: 4,
            min_args: 1,
            max_args: Some(4),
            implementation: statistical_extra::fn_linest,
            volatile: false,
            ..Default::default()
        });
        self.register(FunctionDef {
            name: "LOGEST",
            iftab: Some(51),
            declared_argc: 4,
            min_args: 1,
            max_args: Some(4),
            implementation: statistical_extra::fn_logest,
            volatile: false,
            ..Default::default()
        });
        self.register(FunctionDef {
            name: "GROWTH",
            iftab: Some(52),
            declared_argc: 4,
            min_args: 1,
            max_args: Some(4),
            implementation: statistical_extra::fn_growth,
            volatile: false,
            ..Default::default()
        });
        self.register(FunctionDef {
            name: "TREND",
            iftab: Some(50),
            declared_argc: 4,
            min_args: 1,
            max_args: Some(4),
            implementation: statistical_extra::fn_trend,
            volatile: false,
            ..Default::default()
        });
        self.register(FunctionDef {
            name: "FORECAST.ETS",
            min_args: 3,
            max_args: Some(6),
            implementation: statistical_extra::fn_forecast_ets,
            volatile: false,
            ..Default::default()
        });
        self.register(FunctionDef {
            name: "FORECAST.ETS.CONFINT",
            min_args: 3,
            max_args: Some(7),
            implementation: statistical_extra::fn_forecast_ets_confint,
            volatile: false,
            ..Default::default()
        });
        self.register(FunctionDef {
            name: "FORECAST.ETS.SEASONALITY",
            min_args: 2,
            max_args: Some(4),
            implementation: statistical_extra::fn_forecast_ets_seasonality,
            volatile: false,
            ..Default::default()
        });
        self.register(FunctionDef {
            name: "FORECAST.ETS.STAT",
            min_args: 2,
            max_args: Some(5),
            implementation: statistical_extra::fn_forecast_ets_stat,
            volatile: false,
            ..Default::default()
        });
    }

    fn register_financial_functions(&mut self) {
        // PMT
        self.register(FunctionDef {
            name: "PMT",
            iftab: Some(59),
            declared_argc: 5,
            min_args: 3,
            max_args: Some(5),
            implementation: financial::fn_pmt,
            volatile: false,
            ..Default::default()
        });

        // FV
        self.register(FunctionDef {
            name: "FV",
            iftab: Some(57),
            declared_argc: 5,
            min_args: 3,
            max_args: Some(5),
            implementation: financial::fn_fv,
            volatile: false,
            ..Default::default()
        });

        // PV
        self.register(FunctionDef {
            name: "PV",
            iftab: Some(56),
            declared_argc: 5,
            min_args: 3,
            max_args: Some(5),
            implementation: financial::fn_pv,
            volatile: false,
            ..Default::default()
        });

        // NPER
        self.register(FunctionDef {
            name: "NPER",
            iftab: Some(58),
            declared_argc: 5,
            min_args: 3,
            max_args: Some(5),
            implementation: financial::fn_nper,
            volatile: false,
            ..Default::default()
        });

        // RATE
        self.register(FunctionDef {
            name: "RATE",
            iftab: Some(60),
            declared_argc: 6,
            min_args: 3,
            max_args: Some(6),
            implementation: financial::fn_rate,
            volatile: false,
            ..Default::default()
        });

        // IPMT
        self.register(FunctionDef {
            name: "IPMT",
            iftab: Some(167),
            declared_argc: 6,
            min_args: 4,
            max_args: Some(6),
            implementation: financial::fn_ipmt,
            volatile: false,
            ..Default::default()
        });

        // PPMT
        self.register(FunctionDef {
            name: "PPMT",
            iftab: Some(168),
            declared_argc: 6,
            min_args: 4,
            max_args: Some(6),
            implementation: financial::fn_ppmt,
            volatile: false,
            ..Default::default()
        });

        // CUMIPMT
        self.register(FunctionDef {
            name: "CUMIPMT",
            iftab: Some(448),
            declared_argc: 6,
            min_args: 6,
            max_args: Some(6),
            implementation: financial::fn_cumipmt,
            volatile: false,
            ..Default::default()
        });

        // CUMPRINC
        self.register(FunctionDef {
            name: "CUMPRINC",
            iftab: Some(447),
            declared_argc: 6,
            min_args: 6,
            max_args: Some(6),
            implementation: financial::fn_cumprinc,
            volatile: false,
            ..Default::default()
        });

        // NPV
        self.register(FunctionDef {
            name: "NPV",
            iftab: Some(11),
            declared_argc: 254,
            min_args: 2,
            max_args: None,
            implementation: financial::fn_npv,
            volatile: false,
            ..Default::default()
        });

        // IRR
        self.register(FunctionDef {
            name: "IRR",
            iftab: Some(62),
            declared_argc: 2,
            min_args: 1,
            max_args: Some(2),
            implementation: financial::fn_irr,
            volatile: false,
            ..Default::default()
        });

        // MIRR
        self.register(FunctionDef {
            name: "MIRR",
            iftab: Some(61),
            declared_argc: 3,
            min_args: 3,
            max_args: Some(3),
            implementation: financial::fn_mirr,
            volatile: false,
            ..Default::default()
        });

        // XNPV
        self.register(FunctionDef {
            name: "XNPV",
            iftab: Some(430),
            declared_argc: 3,
            min_args: 3,
            max_args: Some(3),
            implementation: financial::fn_xnpv,
            volatile: false,
            ..Default::default()
        });

        // SLN
        self.register(FunctionDef {
            name: "SLN",
            iftab: Some(142),
            declared_argc: 3,
            min_args: 3,
            max_args: Some(3),
            implementation: financial::fn_sln,
            volatile: false,
            ..Default::default()
        });

        // SYD
        self.register(FunctionDef {
            name: "SYD",
            iftab: Some(143),
            declared_argc: 4,
            min_args: 4,
            max_args: Some(4),
            implementation: financial::fn_syd,
            volatile: false,
            ..Default::default()
        });

        // DB
        self.register(FunctionDef {
            name: "DB",
            iftab: Some(247),
            declared_argc: 5,
            min_args: 4,
            max_args: Some(5),
            implementation: financial::fn_db,
            volatile: false,
            ..Default::default()
        });

        // DDB
        self.register(FunctionDef {
            name: "DDB",
            iftab: Some(144),
            declared_argc: 5,
            min_args: 4,
            max_args: Some(5),
            implementation: financial::fn_ddb,
            volatile: false,
            ..Default::default()
        });

        // EFFECT
        self.register(FunctionDef {
            name: "EFFECT",
            iftab: Some(446),
            declared_argc: 2,
            min_args: 2,
            max_args: Some(2),
            implementation: financial::fn_effect,
            volatile: false,
            ..Default::default()
        });

        // NOMINAL
        self.register(FunctionDef {
            name: "NOMINAL",
            iftab: Some(445),
            declared_argc: 2,
            min_args: 2,
            max_args: Some(2),
            implementation: financial::fn_nominal,
            volatile: false,
            ..Default::default()
        });

        // PDURATION
        self.register(FunctionDef {
            name: "PDURATION",
            min_args: 3,
            max_args: Some(3),
            implementation: financial::fn_pduration,
            volatile: false,
            ..Default::default()
        });
        // financial_extra functions
        self.register(FunctionDef {
            name: "ACCRINT",
            iftab: Some(469),
            declared_argc: 8,
            min_args: 6,
            max_args: Some(8),
            implementation: financial_extra::fn_accrint,
            volatile: false,
            ..Default::default()
        });
        self.register(FunctionDef {
            name: "ACCRINTM",
            iftab: Some(470),
            declared_argc: 5,
            min_args: 3,
            max_args: Some(5),
            implementation: financial_extra::fn_accrintm,
            volatile: false,
            ..Default::default()
        });
        self.register(FunctionDef {
            name: "AMORDEGRC",
            iftab: Some(466),
            declared_argc: 7,
            min_args: 6,
            max_args: Some(7),
            implementation: financial_extra::fn_amordegrc,
            volatile: false,
            ..Default::default()
        });
        self.register(FunctionDef {
            name: "AMORLINC",
            iftab: Some(467),
            declared_argc: 7,
            min_args: 6,
            max_args: Some(7),
            implementation: financial_extra::fn_amorlinc,
            volatile: false,
            ..Default::default()
        });
        self.register(FunctionDef {
            name: "COUPDAYBS",
            iftab: Some(452),
            declared_argc: 4,
            min_args: 3,
            max_args: Some(4),
            implementation: financial_extra::fn_coupdaybs,
            volatile: false,
            ..Default::default()
        });
        self.register(FunctionDef {
            name: "COUPDAYS",
            iftab: Some(453),
            declared_argc: 4,
            min_args: 3,
            max_args: Some(4),
            implementation: financial_extra::fn_coupdays,
            volatile: false,
            ..Default::default()
        });
        self.register(FunctionDef {
            name: "COUPDAYSNC",
            iftab: Some(454),
            declared_argc: 4,
            min_args: 3,
            max_args: Some(4),
            implementation: financial_extra::fn_coupdaysnc,
            volatile: false,
            ..Default::default()
        });
        self.register(FunctionDef {
            name: "COUPNCD",
            iftab: Some(455),
            declared_argc: 4,
            min_args: 3,
            max_args: Some(4),
            implementation: financial_extra::fn_coupncd,
            volatile: false,
            ..Default::default()
        });
        self.register(FunctionDef {
            name: "COUPNUM",
            iftab: Some(456),
            declared_argc: 4,
            min_args: 3,
            max_args: Some(4),
            implementation: financial_extra::fn_coupnum,
            volatile: false,
            ..Default::default()
        });
        self.register(FunctionDef {
            name: "COUPPCD",
            iftab: Some(457),
            declared_argc: 4,
            min_args: 3,
            max_args: Some(4),
            implementation: financial_extra::fn_couppcd,
            volatile: false,
            ..Default::default()
        });
        self.register(FunctionDef {
            name: "DISC",
            iftab: Some(435),
            declared_argc: 5,
            min_args: 4,
            max_args: Some(5),
            implementation: financial_extra::fn_disc,
            volatile: false,
            ..Default::default()
        });
        self.register(FunctionDef {
            name: "DOLLARDE",
            iftab: Some(443),
            declared_argc: 2,
            min_args: 2,
            max_args: Some(2),
            implementation: financial_extra::fn_dollarde,
            volatile: false,
            ..Default::default()
        });
        self.register(FunctionDef {
            name: "DOLLARFR",
            iftab: Some(444),
            declared_argc: 2,
            min_args: 2,
            max_args: Some(2),
            implementation: financial_extra::fn_dollarfr,
            volatile: false,
            ..Default::default()
        });
        self.register(FunctionDef {
            name: "DURATION",
            iftab: Some(458),
            declared_argc: 6,
            min_args: 5,
            max_args: Some(6),
            implementation: financial_extra::fn_duration,
            volatile: false,
            ..Default::default()
        });
        self.register(FunctionDef {
            name: "FVSCHEDULE",
            iftab: Some(476),
            declared_argc: 2,
            min_args: 2,
            max_args: Some(2),
            implementation: financial_extra::fn_fvschedule,
            volatile: false,
            ..Default::default()
        });
        self.register(FunctionDef {
            name: "INTRATE",
            iftab: Some(433),
            declared_argc: 5,
            min_args: 4,
            max_args: Some(5),
            implementation: financial_extra::fn_intrate,
            volatile: false,
            ..Default::default()
        });
        self.register(FunctionDef {
            name: "ISPMT",
            iftab: Some(350),
            declared_argc: 4,
            min_args: 4,
            max_args: Some(4),
            implementation: financial_extra::fn_ispmt,
            volatile: false,
            ..Default::default()
        });
        self.register(FunctionDef {
            name: "MDURATION",
            iftab: Some(459),
            declared_argc: 6,
            min_args: 5,
            max_args: Some(6),
            implementation: financial_extra::fn_mduration,
            volatile: false,
            ..Default::default()
        });
        self.register(FunctionDef {
            name: "ODDFPRICE",
            iftab: Some(462),
            declared_argc: 8,
            min_args: 8,
            max_args: Some(9),
            implementation: financial_extra::fn_oddfprice,
            volatile: false,
            ..Default::default()
        });
        self.register(FunctionDef {
            name: "ODDFYIELD",
            iftab: Some(463),
            declared_argc: 8,
            min_args: 8,
            max_args: Some(9),
            implementation: financial_extra::fn_oddfyield,
            volatile: false,
            ..Default::default()
        });
        self.register(FunctionDef {
            name: "ODDLPRICE",
            iftab: Some(460),
            declared_argc: 8,
            min_args: 7,
            max_args: Some(8),
            implementation: financial_extra::fn_oddlprice,
            volatile: false,
            ..Default::default()
        });
        self.register(FunctionDef {
            name: "ODDLYIELD",
            iftab: Some(461),
            declared_argc: 8,
            min_args: 7,
            max_args: Some(8),
            implementation: financial_extra::fn_oddlyield,
            volatile: false,
            ..Default::default()
        });
        self.register(FunctionDef {
            name: "PRICE",
            iftab: Some(441),
            declared_argc: 7,
            min_args: 6,
            max_args: Some(7),
            implementation: financial_extra::fn_price,
            volatile: false,
            ..Default::default()
        });
        self.register(FunctionDef {
            name: "PRICEDISC",
            iftab: Some(436),
            declared_argc: 5,
            min_args: 4,
            max_args: Some(5),
            implementation: financial_extra::fn_pricedisc,
            volatile: false,
            ..Default::default()
        });
        self.register(FunctionDef {
            name: "PRICEMAT",
            iftab: Some(431),
            declared_argc: 6,
            min_args: 5,
            max_args: Some(6),
            implementation: financial_extra::fn_pricemat,
            volatile: false,
            ..Default::default()
        });
        self.register(FunctionDef {
            name: "RECEIVED",
            iftab: Some(434),
            declared_argc: 5,
            min_args: 4,
            max_args: Some(5),
            implementation: financial_extra::fn_received,
            volatile: false,
            ..Default::default()
        });
        self.register(FunctionDef {
            name: "RRI",
            min_args: 3,
            max_args: Some(3),
            implementation: financial_extra::fn_rri,
            volatile: false,
            ..Default::default()
        });
        self.register(FunctionDef {
            name: "TBILLEQ",
            iftab: Some(438),
            declared_argc: 3,
            min_args: 3,
            max_args: Some(3),
            implementation: financial_extra::fn_tbilleq,
            volatile: false,
            ..Default::default()
        });
        self.register(FunctionDef {
            name: "TBILLPRICE",
            iftab: Some(439),
            declared_argc: 3,
            min_args: 3,
            max_args: Some(3),
            implementation: financial_extra::fn_tbillprice,
            volatile: false,
            ..Default::default()
        });
        self.register(FunctionDef {
            name: "TBILLYIELD",
            iftab: Some(440),
            declared_argc: 3,
            min_args: 3,
            max_args: Some(3),
            implementation: financial_extra::fn_tbillyield,
            volatile: false,
            ..Default::default()
        });
        self.register(FunctionDef {
            name: "VDB",
            iftab: Some(222),
            declared_argc: 7,
            min_args: 5,
            max_args: Some(7),
            implementation: financial_extra::fn_vdb,
            volatile: false,
            ..Default::default()
        });
        self.register(FunctionDef {
            name: "XIRR",
            iftab: Some(429),
            declared_argc: 3,
            min_args: 2,
            max_args: Some(3),
            implementation: financial_extra::fn_xirr,
            volatile: false,
            ..Default::default()
        });
        self.register(FunctionDef {
            name: "YIELD",
            iftab: Some(442),
            declared_argc: 7,
            min_args: 6,
            max_args: Some(7),
            implementation: financial_extra::fn_yield,
            volatile: false,
            ..Default::default()
        });
        self.register(FunctionDef {
            name: "YIELDDISC",
            iftab: Some(437),
            declared_argc: 5,
            min_args: 4,
            max_args: Some(5),
            implementation: financial_extra::fn_yielddisc,
            volatile: false,
            ..Default::default()
        });
        self.register(FunctionDef {
            name: "YIELDMAT",
            iftab: Some(432),
            declared_argc: 6,
            min_args: 5,
            max_args: Some(6),
            implementation: financial_extra::fn_yieldmat,
            volatile: false,
            ..Default::default()
        });
        self.register(FunctionDef {
            name: "EUROCONVERT",
            min_args: 3,
            max_args: Some(5),
            implementation: financial_extra::fn_euroconvert,
            volatile: false,
            ..Default::default()
        });
    }

    fn register_engineering_functions(&mut self) {
        self.register(FunctionDef {
            name: "BIN2DEC",
            iftab: Some(393),
            declared_argc: 1,
            min_args: 1,
            max_args: Some(1),
            implementation: engineering::fn_bin2dec,
            volatile: false,
            ..Default::default()
        });
        self.register(FunctionDef {
            name: "BIN2HEX",
            iftab: Some(395),
            declared_argc: 2,
            min_args: 1,
            max_args: Some(2),
            implementation: engineering::fn_bin2hex,
            volatile: false,
            ..Default::default()
        });
        self.register(FunctionDef {
            name: "BIN2OCT",
            iftab: Some(394),
            declared_argc: 2,
            min_args: 1,
            max_args: Some(2),
            implementation: engineering::fn_bin2oct,
            volatile: false,
            ..Default::default()
        });
        self.register(FunctionDef {
            name: "DEC2BIN",
            iftab: Some(387),
            declared_argc: 2,
            min_args: 1,
            max_args: Some(2),
            implementation: engineering::fn_dec2bin,
            volatile: false,
            ..Default::default()
        });
        self.register(FunctionDef {
            name: "DEC2HEX",
            iftab: Some(388),
            declared_argc: 2,
            min_args: 1,
            max_args: Some(2),
            implementation: engineering::fn_dec2hex,
            volatile: false,
            ..Default::default()
        });
        self.register(FunctionDef {
            name: "DEC2OCT",
            iftab: Some(389),
            declared_argc: 2,
            min_args: 1,
            max_args: Some(2),
            implementation: engineering::fn_dec2oct,
            volatile: false,
            ..Default::default()
        });
        self.register(FunctionDef {
            name: "HEX2BIN",
            iftab: Some(384),
            declared_argc: 2,
            min_args: 1,
            max_args: Some(2),
            implementation: engineering::fn_hex2bin,
            volatile: false,
            ..Default::default()
        });
        self.register(FunctionDef {
            name: "HEX2DEC",
            iftab: Some(385),
            declared_argc: 1,
            min_args: 1,
            max_args: Some(1),
            implementation: engineering::fn_hex2dec,
            volatile: false,
            ..Default::default()
        });
        self.register(FunctionDef {
            name: "HEX2OCT",
            iftab: Some(386),
            declared_argc: 2,
            min_args: 1,
            max_args: Some(2),
            implementation: engineering::fn_hex2oct,
            volatile: false,
            ..Default::default()
        });
        self.register(FunctionDef {
            name: "OCT2BIN",
            iftab: Some(390),
            declared_argc: 2,
            min_args: 1,
            max_args: Some(2),
            implementation: engineering::fn_oct2bin,
            volatile: false,
            ..Default::default()
        });
        self.register(FunctionDef {
            name: "OCT2DEC",
            iftab: Some(392),
            declared_argc: 1,
            min_args: 1,
            max_args: Some(1),
            implementation: engineering::fn_oct2dec,
            volatile: false,
            ..Default::default()
        });
        self.register(FunctionDef {
            name: "OCT2HEX",
            iftab: Some(391),
            declared_argc: 2,
            min_args: 1,
            max_args: Some(2),
            implementation: engineering::fn_oct2hex,
            volatile: false,
            ..Default::default()
        });
        self.register(FunctionDef {
            name: "BITAND",
            min_args: 2,
            max_args: Some(2),
            implementation: engineering::fn_bitand,
            volatile: false,
            ..Default::default()
        });
        self.register(FunctionDef {
            name: "BITOR",
            min_args: 2,
            max_args: Some(2),
            implementation: engineering::fn_bitor,
            volatile: false,
            ..Default::default()
        });
        self.register(FunctionDef {
            name: "BITXOR",
            min_args: 2,
            max_args: Some(2),
            implementation: engineering::fn_bitxor,
            volatile: false,
            ..Default::default()
        });
        self.register(FunctionDef {
            name: "BITLSHIFT",
            min_args: 2,
            max_args: Some(2),
            implementation: engineering::fn_bitlshift,
            volatile: false,
            ..Default::default()
        });
        self.register(FunctionDef {
            name: "BITRSHIFT",
            min_args: 2,
            max_args: Some(2),
            implementation: engineering::fn_bitrshift,
            volatile: false,
            ..Default::default()
        });
        self.register(FunctionDef {
            name: "DELTA",
            iftab: Some(418),
            declared_argc: 2,
            min_args: 1,
            max_args: Some(2),
            implementation: engineering::fn_delta,
            volatile: false,
            ..Default::default()
        });
        self.register(FunctionDef {
            name: "GESTEP",
            iftab: Some(419),
            declared_argc: 2,
            min_args: 1,
            max_args: Some(2),
            implementation: engineering::fn_gestep,
            volatile: false,
            ..Default::default()
        });
        self.register(FunctionDef {
            name: "ERF",
            iftab: Some(423),
            declared_argc: 2,
            min_args: 1,
            max_args: Some(2),
            implementation: engineering::fn_erf,
            volatile: false,
            ..Default::default()
        });
        self.register(FunctionDef {
            name: "ERF.PRECISE",
            min_args: 1,
            max_args: Some(1),
            implementation: engineering::fn_erf_precise,
            volatile: false,
            ..Default::default()
        });
        self.register(FunctionDef {
            name: "ERFC",
            iftab: Some(424),
            declared_argc: 1,
            min_args: 1,
            max_args: Some(1),
            implementation: engineering::fn_erfc,
            volatile: false,
            ..Default::default()
        });
        self.register(FunctionDef {
            name: "ERFC.PRECISE",
            min_args: 1,
            max_args: Some(1),
            implementation: engineering::fn_erfc_precise,
            volatile: false,
            ..Default::default()
        });
        self.register(FunctionDef {
            name: "COMPLEX",
            iftab: Some(411),
            declared_argc: 3,
            min_args: 2,
            max_args: Some(3),
            implementation: engineering::fn_complex,
            volatile: false,
            ..Default::default()
        });
        self.register(FunctionDef {
            name: "IMABS",
            iftab: Some(399),
            declared_argc: 1,
            min_args: 1,
            max_args: Some(1),
            implementation: engineering::fn_imabs,
            volatile: false,
            ..Default::default()
        });
        self.register(FunctionDef {
            name: "IMAGINARY",
            iftab: Some(409),
            declared_argc: 1,
            min_args: 1,
            max_args: Some(1),
            implementation: engineering::fn_imaginary,
            volatile: false,
            ..Default::default()
        });
        self.register(FunctionDef {
            name: "IMARGUMENT",
            iftab: Some(407),
            declared_argc: 1,
            min_args: 1,
            max_args: Some(1),
            implementation: engineering::fn_imargument,
            volatile: false,
            ..Default::default()
        });
        self.register(FunctionDef {
            name: "IMCONJUGATE",
            iftab: Some(408),
            declared_argc: 1,
            min_args: 1,
            max_args: Some(1),
            implementation: engineering::fn_imconjugate,
            volatile: false,
            ..Default::default()
        });
        self.register(FunctionDef {
            name: "IMCOS",
            iftab: Some(405),
            declared_argc: 1,
            min_args: 1,
            max_args: Some(1),
            implementation: engineering::fn_imcos,
            volatile: false,
            ..Default::default()
        });
        self.register(FunctionDef {
            name: "IMCOSH",
            min_args: 1,
            max_args: Some(1),
            implementation: engineering::fn_imcosh,
            volatile: false,
            ..Default::default()
        });
        self.register(FunctionDef {
            name: "IMCOT",
            min_args: 1,
            max_args: Some(1),
            implementation: engineering::fn_imcot,
            volatile: false,
            ..Default::default()
        });
        self.register(FunctionDef {
            name: "IMCSC",
            min_args: 1,
            max_args: Some(1),
            implementation: engineering::fn_imcsc,
            volatile: false,
            ..Default::default()
        });
        self.register(FunctionDef {
            name: "IMCSCH",
            min_args: 1,
            max_args: Some(1),
            implementation: engineering::fn_imcsch,
            volatile: false,
            ..Default::default()
        });
        self.register(FunctionDef {
            name: "IMDIV",
            iftab: Some(397),
            declared_argc: 2,
            min_args: 2,
            max_args: Some(2),
            implementation: engineering::fn_imdiv,
            volatile: false,
            ..Default::default()
        });
        self.register(FunctionDef {
            name: "IMEXP",
            iftab: Some(406),
            declared_argc: 1,
            min_args: 1,
            max_args: Some(1),
            implementation: engineering::fn_imexp,
            volatile: false,
            ..Default::default()
        });
        self.register(FunctionDef {
            name: "IMLN",
            iftab: Some(401),
            declared_argc: 1,
            min_args: 1,
            max_args: Some(1),
            implementation: engineering::fn_imln,
            volatile: false,
            ..Default::default()
        });
        self.register(FunctionDef {
            name: "IMLOG10",
            iftab: Some(403),
            declared_argc: 1,
            min_args: 1,
            max_args: Some(1),
            implementation: engineering::fn_imlog10,
            volatile: false,
            ..Default::default()
        });
        self.register(FunctionDef {
            name: "IMLOG2",
            iftab: Some(402),
            declared_argc: 1,
            min_args: 1,
            max_args: Some(1),
            implementation: engineering::fn_imlog2,
            volatile: false,
            ..Default::default()
        });
        self.register(FunctionDef {
            name: "IMPOWER",
            iftab: Some(398),
            declared_argc: 2,
            min_args: 2,
            max_args: Some(2),
            implementation: engineering::fn_impower,
            volatile: false,
            ..Default::default()
        });
        self.register(FunctionDef {
            name: "IMPRODUCT",
            iftab: Some(413),
            min_args: 2,
            max_args: None,
            implementation: engineering::fn_improduct,
            volatile: false,
            ..Default::default()
        });
        self.register(FunctionDef {
            name: "IMREAL",
            iftab: Some(410),
            declared_argc: 1,
            min_args: 1,
            max_args: Some(1),
            implementation: engineering::fn_imreal,
            volatile: false,
            ..Default::default()
        });
        self.register(FunctionDef {
            name: "IMSEC",
            min_args: 1,
            max_args: Some(1),
            implementation: engineering::fn_imsec,
            volatile: false,
            ..Default::default()
        });
        self.register(FunctionDef {
            name: "IMSECH",
            min_args: 1,
            max_args: Some(1),
            implementation: engineering::fn_imsech,
            volatile: false,
            ..Default::default()
        });
        self.register(FunctionDef {
            name: "IMSIN",
            iftab: Some(404),
            declared_argc: 1,
            min_args: 1,
            max_args: Some(1),
            implementation: engineering::fn_imsin,
            volatile: false,
            ..Default::default()
        });
        self.register(FunctionDef {
            name: "IMSINH",
            min_args: 1,
            max_args: Some(1),
            implementation: engineering::fn_imsinh,
            volatile: false,
            ..Default::default()
        });
        self.register(FunctionDef {
            name: "IMSQRT",
            iftab: Some(400),
            declared_argc: 1,
            min_args: 1,
            max_args: Some(1),
            implementation: engineering::fn_imsqrt,
            volatile: false,
            ..Default::default()
        });
        self.register(FunctionDef {
            name: "IMSUB",
            iftab: Some(396),
            declared_argc: 2,
            min_args: 2,
            max_args: Some(2),
            implementation: engineering::fn_imsub,
            volatile: false,
            ..Default::default()
        });
        self.register(FunctionDef {
            name: "IMSUM",
            iftab: Some(412),
            min_args: 2,
            max_args: None,
            implementation: engineering::fn_imsum,
            volatile: false,
            ..Default::default()
        });
        self.register(FunctionDef {
            name: "IMTAN",
            min_args: 1,
            max_args: Some(1),
            implementation: engineering::fn_imtan,
            volatile: false,
            ..Default::default()
        });
        self.register(FunctionDef {
            name: "BESSELI",
            iftab: Some(428),
            declared_argc: 2,
            min_args: 2,
            max_args: Some(2),
            implementation: engineering::fn_besseli,
            volatile: false,
            ..Default::default()
        });
        self.register(FunctionDef {
            name: "BESSELJ",
            iftab: Some(425),
            declared_argc: 2,
            min_args: 2,
            max_args: Some(2),
            implementation: engineering::fn_besselj,
            volatile: false,
            ..Default::default()
        });
        self.register(FunctionDef {
            name: "BESSELK",
            iftab: Some(426),
            declared_argc: 2,
            min_args: 2,
            max_args: Some(2),
            implementation: engineering::fn_besselk,
            volatile: false,
            ..Default::default()
        });
        self.register(FunctionDef {
            name: "BESSELY",
            iftab: Some(427),
            declared_argc: 2,
            min_args: 2,
            max_args: Some(2),
            implementation: engineering::fn_bessely,
            volatile: false,
            ..Default::default()
        });
        self.register(FunctionDef {
            name: "CONVERT",
            iftab: Some(468),
            declared_argc: 3,
            min_args: 3,
            max_args: Some(3),
            implementation: engineering::fn_convert,
            volatile: false,
            ..Default::default()
        });
    }

    fn register_database_functions(&mut self) {
        self.register(FunctionDef {
            name: "DAVERAGE",
            iftab: Some(42),
            declared_argc: 3,
            min_args: 3,
            max_args: Some(3),
            implementation: database::fn_daverage,
            volatile: false,
            ..Default::default()
        });
        self.register(FunctionDef {
            name: "DCOUNT",
            iftab: Some(40),
            declared_argc: 3,
            min_args: 3,
            max_args: Some(3),
            implementation: database::fn_dcount,
            volatile: false,
            ..Default::default()
        });
        self.register(FunctionDef {
            name: "DCOUNTA",
            iftab: Some(199),
            declared_argc: 3,
            min_args: 3,
            max_args: Some(3),
            implementation: database::fn_dcounta,
            volatile: false,
            ..Default::default()
        });
        self.register(FunctionDef {
            name: "DGET",
            iftab: Some(235),
            declared_argc: 3,
            min_args: 3,
            max_args: Some(3),
            implementation: database::fn_dget,
            volatile: false,
            ..Default::default()
        });
        self.register(FunctionDef {
            name: "DMAX",
            iftab: Some(44),
            declared_argc: 3,
            min_args: 3,
            max_args: Some(3),
            implementation: database::fn_dmax,
            volatile: false,
            ..Default::default()
        });
        self.register(FunctionDef {
            name: "DMIN",
            iftab: Some(43),
            declared_argc: 3,
            min_args: 3,
            max_args: Some(3),
            implementation: database::fn_dmin,
            volatile: false,
            ..Default::default()
        });
        self.register(FunctionDef {
            name: "DPRODUCT",
            iftab: Some(189),
            declared_argc: 3,
            min_args: 3,
            max_args: Some(3),
            implementation: database::fn_dproduct,
            volatile: false,
            ..Default::default()
        });
        self.register(FunctionDef {
            name: "DSTDEV",
            iftab: Some(45),
            declared_argc: 3,
            min_args: 3,
            max_args: Some(3),
            implementation: database::fn_dstdev,
            volatile: false,
            ..Default::default()
        });
        self.register(FunctionDef {
            name: "DSTDEVP",
            iftab: Some(195),
            declared_argc: 3,
            min_args: 3,
            max_args: Some(3),
            implementation: database::fn_dstdevp,
            volatile: false,
            ..Default::default()
        });
        self.register(FunctionDef {
            name: "DSUM",
            iftab: Some(41),
            declared_argc: 3,
            min_args: 3,
            max_args: Some(3),
            implementation: database::fn_dsum,
            volatile: false,
            ..Default::default()
        });
        self.register(FunctionDef {
            name: "DVAR",
            iftab: Some(47),
            declared_argc: 3,
            min_args: 3,
            max_args: Some(3),
            implementation: database::fn_dvar,
            volatile: false,
            ..Default::default()
        });
        self.register(FunctionDef {
            name: "DVARP",
            iftab: Some(196),
            declared_argc: 3,
            min_args: 3,
            max_args: Some(3),
            implementation: database::fn_dvarp,
            volatile: false,
            ..Default::default()
        });
    }

    fn register_compatibility_functions(&mut self) {
        self.register(FunctionDef {
            name: "BETADIST",
            iftab: Some(270),
            declared_argc: 5,
            min_args: 3,
            max_args: Some(5),
            implementation: compatibility::fn_betadist,
            volatile: false,
            ..Default::default()
        });
        self.register(FunctionDef {
            name: "BETAINV",
            iftab: Some(272),
            declared_argc: 5,
            min_args: 3,
            max_args: Some(5),
            implementation: compatibility::fn_betainv,
            volatile: false,
            ..Default::default()
        });
        self.register(FunctionDef {
            name: "BINOMDIST",
            iftab: Some(273),
            declared_argc: 4,
            min_args: 4,
            max_args: Some(4),
            implementation: compatibility::fn_binomdist,
            volatile: false,
            ..Default::default()
        });
        self.register(FunctionDef {
            name: "CEILING",
            iftab: Some(288),
            declared_argc: 2,
            min_args: 2,
            max_args: Some(2),
            implementation: compatibility::fn_ceiling,
            volatile: false,
            ..Default::default()
        });
        self.register(FunctionDef {
            name: "CHIDIST",
            iftab: Some(274),
            declared_argc: 2,
            min_args: 2,
            max_args: Some(2),
            implementation: compatibility::fn_chidist,
            volatile: false,
            ..Default::default()
        });
        self.register(FunctionDef {
            name: "CHIINV",
            iftab: Some(275),
            declared_argc: 2,
            min_args: 2,
            max_args: Some(2),
            implementation: compatibility::fn_chiinv,
            volatile: false,
            ..Default::default()
        });
        self.register(FunctionDef {
            name: "CHITEST",
            iftab: Some(306),
            declared_argc: 2,
            min_args: 2,
            max_args: Some(2),
            implementation: compatibility::fn_chitest,
            volatile: false,
            ..Default::default()
        });
        self.register(FunctionDef {
            name: "CONFIDENCE",
            iftab: Some(277),
            declared_argc: 3,
            min_args: 3,
            max_args: Some(3),
            implementation: compatibility::fn_confidence,
            volatile: false,
            ..Default::default()
        });
        self.register(FunctionDef {
            name: "COVAR",
            iftab: Some(308),
            declared_argc: 2,
            min_args: 2,
            max_args: Some(2),
            implementation: compatibility::fn_covar,
            volatile: false,
            ..Default::default()
        });
        self.register(FunctionDef {
            name: "CRITBINOM",
            iftab: Some(278),
            declared_argc: 3,
            min_args: 3,
            max_args: Some(3),
            implementation: compatibility::fn_critbinom,
            volatile: false,
            ..Default::default()
        });
        self.register(FunctionDef {
            name: "EXPONDIST",
            iftab: Some(280),
            declared_argc: 3,
            min_args: 3,
            max_args: Some(3),
            implementation: compatibility::fn_expondist,
            volatile: false,
            ..Default::default()
        });
        self.register(FunctionDef {
            name: "FDIST",
            iftab: Some(281),
            declared_argc: 3,
            min_args: 3,
            max_args: Some(3),
            implementation: compatibility::fn_fdist,
            volatile: false,
            ..Default::default()
        });
        self.register(FunctionDef {
            name: "FINV",
            iftab: Some(282),
            declared_argc: 3,
            min_args: 3,
            max_args: Some(3),
            implementation: compatibility::fn_finv,
            volatile: false,
            ..Default::default()
        });
        self.register(FunctionDef {
            name: "FLOOR",
            iftab: Some(285),
            declared_argc: 2,
            min_args: 2,
            max_args: Some(2),
            implementation: compatibility::fn_floor,
            volatile: false,
            ..Default::default()
        });
        self.register(FunctionDef {
            name: "FTEST",
            iftab: Some(310),
            declared_argc: 2,
            min_args: 2,
            max_args: Some(2),
            implementation: compatibility::fn_ftest,
            volatile: false,
            ..Default::default()
        });
        self.register(FunctionDef {
            name: "GAMMADIST",
            iftab: Some(286),
            declared_argc: 4,
            min_args: 4,
            max_args: Some(4),
            implementation: compatibility::fn_gammadist,
            volatile: false,
            ..Default::default()
        });
        self.register(FunctionDef {
            name: "GAMMAINV",
            iftab: Some(287),
            declared_argc: 3,
            min_args: 3,
            max_args: Some(3),
            implementation: compatibility::fn_gammainv,
            volatile: false,
            ..Default::default()
        });
        self.register(FunctionDef {
            name: "HYPGEOMDIST",
            iftab: Some(289),
            declared_argc: 4,
            min_args: 4,
            max_args: Some(4),
            implementation: compatibility::fn_hypgeomdist,
            volatile: false,
            ..Default::default()
        });
        self.register(FunctionDef {
            name: "LOGINV",
            iftab: Some(291),
            declared_argc: 3,
            min_args: 3,
            max_args: Some(3),
            implementation: compatibility::fn_loginv,
            volatile: false,
            ..Default::default()
        });
        self.register(FunctionDef {
            name: "LOGNORMDIST",
            iftab: Some(290),
            declared_argc: 3,
            min_args: 3,
            max_args: Some(3),
            implementation: compatibility::fn_lognormdist,
            volatile: false,
            ..Default::default()
        });
        self.register(FunctionDef {
            name: "NEGBINOMDIST",
            iftab: Some(292),
            declared_argc: 3,
            min_args: 3,
            max_args: Some(3),
            implementation: compatibility::fn_negbinomdist,
            volatile: false,
            ..Default::default()
        });
        self.register(FunctionDef {
            name: "NORMDIST",
            iftab: Some(293),
            declared_argc: 4,
            min_args: 4,
            max_args: Some(4),
            implementation: compatibility::fn_normdist,
            volatile: false,
            ..Default::default()
        });
        self.register(FunctionDef {
            name: "NORMSDIST",
            iftab: Some(294),
            declared_argc: 1,
            min_args: 1,
            max_args: Some(1),
            implementation: compatibility::fn_normsdist,
            volatile: false,
            ..Default::default()
        });
        self.register(FunctionDef {
            name: "NORMSINV",
            iftab: Some(296),
            declared_argc: 1,
            min_args: 1,
            max_args: Some(1),
            implementation: compatibility::fn_normsinv,
            volatile: false,
            ..Default::default()
        });
        self.register(FunctionDef {
            name: "NORM.INV",
            min_args: 3,
            max_args: Some(3),
            implementation: compatibility::fn_norm_inv,
            volatile: false,
            ..Default::default()
        });
        self.register(FunctionDef {
            name: "POISSON",
            iftab: Some(300),
            declared_argc: 3,
            min_args: 3,
            max_args: Some(3),
            implementation: compatibility::fn_poisson,
            volatile: false,
            ..Default::default()
        });
        self.register(FunctionDef {
            name: "TDIST",
            iftab: Some(301),
            declared_argc: 3,
            min_args: 3,
            max_args: Some(3),
            implementation: compatibility::fn_tdist,
            volatile: false,
            ..Default::default()
        });
        self.register(FunctionDef {
            name: "TINV",
            iftab: Some(332),
            declared_argc: 2,
            min_args: 2,
            max_args: Some(2),
            implementation: compatibility::fn_tinv,
            volatile: false,
            ..Default::default()
        });
        self.register(FunctionDef {
            name: "TTEST",
            iftab: Some(316),
            declared_argc: 4,
            min_args: 4,
            max_args: Some(4),
            implementation: compatibility::fn_ttest,
            volatile: false,
            ..Default::default()
        });
        self.register(FunctionDef {
            name: "WEIBULL",
            iftab: Some(302),
            declared_argc: 4,
            min_args: 4,
            max_args: Some(4),
            implementation: compatibility::fn_weibull,
            volatile: false,
            ..Default::default()
        });
        self.register(FunctionDef {
            name: "ZTEST",
            iftab: Some(324),
            declared_argc: 3,
            min_args: 2,
            max_args: Some(3),
            implementation: compatibility::fn_ztest,
            volatile: false,
            ..Default::default()
        });
    }

    fn register_web_functions(&mut self) {
        self.register(FunctionDef {
            name: "ENCODEURL",
            min_args: 1,
            max_args: Some(1),
            implementation: lookup_extra::fn_encodeurl,
            volatile: false,
            ..Default::default()
        });
        self.register(FunctionDef {
            name: "FILTERXML",
            min_args: 2,
            max_args: Some(2),
            implementation: lookup_extra::fn_filterxml,
            volatile: false,
            ..Default::default()
        });
        self.register(FunctionDef {
            name: "WEBSERVICE",
            min_args: 1,
            max_args: Some(1),
            implementation: lookup_extra::fn_webservice,
            volatile: false,
            ..Default::default()
        });
    }

    /// Register obsolete BIFF8 macro functions (Lotus 1-2-3 carryovers,
    /// Excel 4 macro functions, etc.) as name+iftab stubs so the XLS/XLSB
    /// decompiler can print them when reading old workbooks. Evaluation
    /// returns an error.
    fn register_obsolete_biff_functions(&mut self) {
        self.register(FunctionDef {
            name: "ABSREF",
            iftab: Some(79),
            declared_argc: 2,
            ..Default::default()
        });
        self.register(FunctionDef {
            name: "ACTIVE.CELL",
            iftab: Some(94),
            declared_argc: 0,
            ..Default::default()
        });
        self.register(FunctionDef {
            name: "ADD.BAR",
            iftab: Some(151),
            declared_argc: 1,
            ..Default::default()
        });
        self.register(FunctionDef {
            name: "ADD.COMMAND",
            iftab: Some(153),
            declared_argc: 5,
            ..Default::default()
        });
        self.register(FunctionDef {
            name: "ADD.MENU",
            iftab: Some(152),
            declared_argc: 4,
            ..Default::default()
        });
        self.register(FunctionDef {
            name: "ADD.TOOLBAR",
            iftab: Some(253),
            declared_argc: 2,
            ..Default::default()
        });
        self.register(FunctionDef {
            name: "APP.TITLE",
            iftab: Some(262),
            declared_argc: 1,
            ..Default::default()
        });
        self.register(FunctionDef {
            name: "ARGUMENT",
            iftab: Some(81),
            declared_argc: 3,
            ..Default::default()
        });
        self.register(FunctionDef {
            name: "BREAK",
            iftab: Some(173),
            declared_argc: 0,
            ..Default::default()
        });
        self.register(FunctionDef {
            name: "CALLER",
            iftab: Some(89),
            declared_argc: 0,
            ..Default::default()
        });
        self.register(FunctionDef {
            name: "CANCEL.KEY",
            iftab: Some(170),
            declared_argc: 2,
            ..Default::default()
        });
        self.register(FunctionDef {
            name: "CHECK.COMMAND",
            iftab: Some(155),
            declared_argc: 5,
            ..Default::default()
        });
        self.register(FunctionDef {
            name: "CREATE.OBJECT",
            iftab: Some(236),
            declared_argc: 11,
            ..Default::default()
        });
        self.register(FunctionDef {
            name: "CUSTOM.REPEAT",
            iftab: Some(240),
            declared_argc: 3,
            ..Default::default()
        });
        self.register(FunctionDef {
            name: "CUSTOM.UNDO",
            iftab: Some(239),
            declared_argc: 2,
            ..Default::default()
        });
        self.register(FunctionDef {
            name: "DATESTRING",
            iftab: Some(352),
            declared_argc: 1,
            ..Default::default()
        });
        self.register(FunctionDef {
            name: "DELETE.BAR",
            iftab: Some(200),
            declared_argc: 1,
            ..Default::default()
        });
        self.register(FunctionDef {
            name: "DELETE.COMMAND",
            iftab: Some(159),
            declared_argc: 4,
            ..Default::default()
        });
        self.register(FunctionDef {
            name: "DELETE.MENU",
            iftab: Some(158),
            declared_argc: 3,
            ..Default::default()
        });
        self.register(FunctionDef {
            name: "DELETE.TOOLBAR",
            iftab: Some(254),
            declared_argc: 1,
            ..Default::default()
        });
        self.register(FunctionDef {
            name: "DEREF",
            iftab: Some(90),
            declared_argc: 1,
            ..Default::default()
        });
        self.register(FunctionDef {
            name: "DIALOG.BOX",
            iftab: Some(161),
            declared_argc: 1,
            ..Default::default()
        });
        self.register(FunctionDef {
            name: "DIRECTORY",
            iftab: Some(123),
            declared_argc: 1,
            ..Default::default()
        });
        self.register(FunctionDef {
            name: "DOCUMENTS",
            iftab: Some(93),
            declared_argc: 2,
            ..Default::default()
        });
        self.register(FunctionDef {
            name: "ECHO",
            iftab: Some(87),
            declared_argc: 1,
            ..Default::default()
        });
        self.register(FunctionDef {
            name: "ELSE",
            iftab: Some(223),
            declared_argc: 0,
            ..Default::default()
        });
        self.register(FunctionDef {
            name: "ELSE.IF",
            iftab: Some(224),
            declared_argc: 1,
            ..Default::default()
        });
        self.register(FunctionDef {
            name: "ENABLE.COMMAND",
            iftab: Some(154),
            declared_argc: 5,
            ..Default::default()
        });
        self.register(FunctionDef {
            name: "ENABLE.TOOL",
            iftab: Some(265),
            declared_argc: 3,
            ..Default::default()
        });
        self.register(FunctionDef {
            name: "END.IF",
            iftab: Some(225),
            declared_argc: 0,
            ..Default::default()
        });
        self.register(FunctionDef {
            name: "ERROR",
            iftab: Some(84),
            declared_argc: 2,
            ..Default::default()
        });
        self.register(FunctionDef {
            name: "EVALUATE",
            iftab: Some(257),
            declared_argc: 1,
            ..Default::default()
        });
        self.register(FunctionDef {
            name: "EXEC",
            iftab: Some(110),
            declared_argc: 4,
            ..Default::default()
        });
        self.register(FunctionDef {
            name: "EXECUTE",
            iftab: Some(178),
            declared_argc: 2,
            ..Default::default()
        });
        self.register(FunctionDef {
            name: "FCLOSE",
            iftab: Some(133),
            declared_argc: 1,
            ..Default::default()
        });
        self.register(FunctionDef {
            name: "FILES",
            iftab: Some(166),
            declared_argc: 2,
            ..Default::default()
        });
        self.register(FunctionDef {
            name: "FOPEN",
            iftab: Some(132),
            declared_argc: 2,
            ..Default::default()
        });
        self.register(FunctionDef {
            name: "FOR",
            iftab: Some(171),
            declared_argc: 4,
            ..Default::default()
        });
        self.register(FunctionDef {
            name: "FOR.CELL",
            iftab: Some(226),
            declared_argc: 3,
            ..Default::default()
        });
        self.register(FunctionDef {
            name: "FORMULA.CONVERT",
            iftab: Some(241),
            declared_argc: 5,
            ..Default::default()
        });
        self.register(FunctionDef {
            name: "FPOS",
            iftab: Some(139),
            declared_argc: 2,
            ..Default::default()
        });
        self.register(FunctionDef {
            name: "FREAD",
            iftab: Some(136),
            declared_argc: 2,
            ..Default::default()
        });
        self.register(FunctionDef {
            name: "FREADLN",
            iftab: Some(135),
            declared_argc: 1,
            ..Default::default()
        });
        self.register(FunctionDef {
            name: "FSIZE",
            iftab: Some(134),
            declared_argc: 1,
            ..Default::default()
        });
        self.register(FunctionDef {
            name: "FWRITE",
            iftab: Some(138),
            declared_argc: 2,
            ..Default::default()
        });
        self.register(FunctionDef {
            name: "FWRITELN",
            iftab: Some(137),
            declared_argc: 2,
            ..Default::default()
        });
        self.register(FunctionDef {
            name: "GET.BAR",
            iftab: Some(182),
            declared_argc: 4,
            ..Default::default()
        });
        self.register(FunctionDef {
            name: "GET.CELL",
            iftab: Some(185),
            declared_argc: 2,
            ..Default::default()
        });
        self.register(FunctionDef {
            name: "GET.CHART.ITEM",
            iftab: Some(160),
            declared_argc: 3,
            ..Default::default()
        });
        self.register(FunctionDef {
            name: "GET.DEF",
            iftab: Some(145),
            declared_argc: 3,
            ..Default::default()
        });
        self.register(FunctionDef {
            name: "GET.DOCUMENT",
            iftab: Some(188),
            declared_argc: 2,
            ..Default::default()
        });
        self.register(FunctionDef {
            name: "GET.FORMULA",
            iftab: Some(106),
            declared_argc: 1,
            ..Default::default()
        });
        self.register(FunctionDef {
            name: "GET.LINK.INFO",
            iftab: Some(242),
            declared_argc: 4,
            ..Default::default()
        });
        self.register(FunctionDef {
            name: "GET.MOVIE",
            iftab: Some(335),
            declared_argc: 3,
            ..Default::default()
        });
        self.register(FunctionDef {
            name: "GET.NAME",
            iftab: Some(107),
            declared_argc: 2,
            ..Default::default()
        });
        self.register(FunctionDef {
            name: "GET.NOTE",
            iftab: Some(191),
            declared_argc: 3,
            ..Default::default()
        });
        self.register(FunctionDef {
            name: "GET.OBJECT",
            iftab: Some(246),
            declared_argc: 5,
            ..Default::default()
        });
        self.register(FunctionDef {
            name: "GET.PIVOT.FIELD",
            iftab: Some(340),
            declared_argc: 3,
            ..Default::default()
        });
        self.register(FunctionDef {
            name: "GET.PIVOT.ITEM",
            iftab: Some(341),
            declared_argc: 4,
            ..Default::default()
        });
        self.register(FunctionDef {
            name: "GET.PIVOT.TABLE",
            iftab: Some(339),
            declared_argc: 2,
            ..Default::default()
        });
        self.register(FunctionDef {
            name: "GET.TOOL",
            iftab: Some(259),
            declared_argc: 3,
            ..Default::default()
        });
        self.register(FunctionDef {
            name: "GET.TOOLBAR",
            iftab: Some(258),
            declared_argc: 2,
            ..Default::default()
        });
        self.register(FunctionDef {
            name: "GET.WINDOW",
            iftab: Some(187),
            declared_argc: 2,
            ..Default::default()
        });
        self.register(FunctionDef {
            name: "GET.WORKBOOK",
            iftab: Some(268),
            declared_argc: 2,
            ..Default::default()
        });
        self.register(FunctionDef {
            name: "GET.WORKSPACE",
            iftab: Some(186),
            declared_argc: 1,
            ..Default::default()
        });
        self.register(FunctionDef {
            name: "GOTO",
            iftab: Some(53),
            declared_argc: 1,
            ..Default::default()
        });
        self.register(FunctionDef {
            name: "GROUP",
            iftab: Some(245),
            declared_argc: 0,
            ..Default::default()
        });
        self.register(FunctionDef {
            name: "HALT",
            iftab: Some(54),
            declared_argc: 1,
            ..Default::default()
        });
        self.register(FunctionDef {
            name: "HELP",
            iftab: Some(181),
            declared_argc: 1,
            ..Default::default()
        });
        self.register(FunctionDef {
            name: "INITIATE",
            iftab: Some(175),
            declared_argc: 2,
            ..Default::default()
        });
        self.register(FunctionDef {
            name: "INPUT",
            iftab: Some(104),
            declared_argc: 7,
            ..Default::default()
        });
        self.register(FunctionDef {
            name: "ISTHAIDIGIT",
            iftab: Some(375),
            declared_argc: 1,
            ..Default::default()
        });
        self.register(FunctionDef {
            name: "LAST.ERROR",
            iftab: Some(238),
            declared_argc: 0,
            ..Default::default()
        });
        self.register(FunctionDef {
            name: "LINKS",
            iftab: Some(103),
            declared_argc: 2,
            ..Default::default()
        });
        self.register(FunctionDef {
            name: "MOVIE.COMMAND",
            iftab: Some(334),
            declared_argc: 4,
            ..Default::default()
        });
        self.register(FunctionDef {
            name: "NAMES",
            iftab: Some(122),
            declared_argc: 3,
            ..Default::default()
        });
        self.register(FunctionDef {
            name: "NEXT",
            iftab: Some(174),
            declared_argc: 0,
            ..Default::default()
        });
        self.register(FunctionDef {
            name: "NORMINV",
            iftab: Some(295),
            declared_argc: 3,
            ..Default::default()
        });
        self.register(FunctionDef {
            name: "NOTE",
            iftab: Some(192),
            declared_argc: 4,
            ..Default::default()
        });
        self.register(FunctionDef {
            name: "NUMBERSTRING",
            iftab: Some(353),
            declared_argc: 2,
            ..Default::default()
        });
        self.register(FunctionDef {
            name: "OPEN.DIALOG",
            iftab: Some(355),
            declared_argc: 4,
            ..Default::default()
        });
        self.register(FunctionDef {
            name: "OPTIONS.LISTS.GET",
            iftab: Some(349),
            declared_argc: 1,
            ..Default::default()
        });
        self.register(FunctionDef {
            name: "PAUSE",
            iftab: Some(248),
            declared_argc: 1,
            ..Default::default()
        });
        self.register(FunctionDef {
            name: "PIVOT.ADD.DATA",
            iftab: Some(338),
            declared_argc: 9,
            ..Default::default()
        });
        self.register(FunctionDef {
            name: "POKE",
            iftab: Some(177),
            declared_argc: 3,
            ..Default::default()
        });
        self.register(FunctionDef {
            name: "PRESS.TOOL",
            iftab: Some(266),
            declared_argc: 3,
            ..Default::default()
        });
        self.register(FunctionDef {
            name: "REFTEXT",
            iftab: Some(146),
            declared_argc: 2,
            ..Default::default()
        });
        self.register(FunctionDef {
            name: "REGISTER",
            iftab: Some(149),
            ..Default::default()
        });
        self.register(FunctionDef {
            name: "RELREF",
            iftab: Some(80),
            declared_argc: 2,
            ..Default::default()
        });
        self.register(FunctionDef {
            name: "RENAME.COMMAND",
            iftab: Some(156),
            declared_argc: 5,
            ..Default::default()
        });
        self.register(FunctionDef {
            name: "REQUEST",
            iftab: Some(176),
            declared_argc: 2,
            ..Default::default()
        });
        self.register(FunctionDef {
            name: "RESET.TOOLBAR",
            iftab: Some(256),
            declared_argc: 1,
            ..Default::default()
        });
        self.register(FunctionDef {
            name: "RESTART",
            iftab: Some(180),
            declared_argc: 1,
            ..Default::default()
        });
        self.register(FunctionDef {
            name: "RESULT",
            iftab: Some(96),
            declared_argc: 1,
            ..Default::default()
        });
        self.register(FunctionDef {
            name: "RESUME",
            iftab: Some(251),
            declared_argc: 1,
            ..Default::default()
        });
        self.register(FunctionDef {
            name: "RETURN",
            iftab: Some(55),
            declared_argc: 1,
            ..Default::default()
        });
        self.register(FunctionDef {
            name: "ROUNDBAHTDOWN",
            iftab: Some(376),
            declared_argc: 1,
            ..Default::default()
        });
        self.register(FunctionDef {
            name: "ROUNDBAHTUP",
            iftab: Some(377),
            declared_argc: 1,
            ..Default::default()
        });
        self.register(FunctionDef {
            name: "SAVE.DIALOG",
            iftab: Some(356),
            declared_argc: 5,
            ..Default::default()
        });
        self.register(FunctionDef {
            name: "SAVE.TOOLBAR",
            iftab: Some(264),
            declared_argc: 2,
            ..Default::default()
        });
        self.register(FunctionDef {
            name: "SCENARIO.GET",
            iftab: Some(348),
            declared_argc: 2,
            ..Default::default()
        });
        self.register(FunctionDef {
            name: "SELECTION",
            iftab: Some(95),
            declared_argc: 0,
            ..Default::default()
        });
        self.register(FunctionDef {
            name: "SERIES",
            iftab: Some(92),
            declared_argc: 2,
            ..Default::default()
        });
        self.register(FunctionDef {
            name: "SET.NAME",
            iftab: Some(88),
            declared_argc: 2,
            ..Default::default()
        });
        self.register(FunctionDef {
            name: "SET.VALUE",
            iftab: Some(108),
            declared_argc: 2,
            ..Default::default()
        });
        self.register(FunctionDef {
            name: "SHOW.BAR",
            iftab: Some(157),
            declared_argc: 1,
            ..Default::default()
        });
        self.register(FunctionDef {
            name: "SPELLING.CHECK",
            iftab: Some(260),
            declared_argc: 3,
            ..Default::default()
        });
        self.register(FunctionDef {
            name: "STEP",
            iftab: Some(85),
            declared_argc: 0,
            ..Default::default()
        });
        self.register(FunctionDef {
            name: "TERMINATE",
            iftab: Some(179),
            declared_argc: 1,
            ..Default::default()
        });
        self.register(FunctionDef {
            name: "TEXT.BOX",
            iftab: Some(243),
            declared_argc: 4,
            ..Default::default()
        });
        self.register(FunctionDef {
            name: "TEXTREF",
            iftab: Some(147),
            declared_argc: 2,
            ..Default::default()
        });
        self.register(FunctionDef {
            name: "THAIDAYOFWEEK",
            iftab: Some(369),
            declared_argc: 1,
            ..Default::default()
        });
        self.register(FunctionDef {
            name: "THAIDIGIT",
            iftab: Some(370),
            declared_argc: 1,
            ..Default::default()
        });
        self.register(FunctionDef {
            name: "THAIMONTHOFYEAR",
            iftab: Some(371),
            declared_argc: 1,
            ..Default::default()
        });
        self.register(FunctionDef {
            name: "THAINUMSOUND",
            iftab: Some(372),
            declared_argc: 1,
            ..Default::default()
        });
        self.register(FunctionDef {
            name: "THAINUMSTRING",
            iftab: Some(373),
            declared_argc: 1,
            ..Default::default()
        });
        self.register(FunctionDef {
            name: "THAISTRINGLENGTH",
            iftab: Some(374),
            declared_argc: 1,
            ..Default::default()
        });
        self.register(FunctionDef {
            name: "THAIYEAR",
            iftab: Some(378),
            declared_argc: 1,
            ..Default::default()
        });
        self.register(FunctionDef {
            name: "UNREGISTER",
            iftab: Some(201),
            declared_argc: 1,
            ..Default::default()
        });
        self.register(FunctionDef {
            name: "USDOLLAR",
            iftab: Some(204),
            declared_argc: 2,
            ..Default::default()
        });
        self.register(FunctionDef {
            name: "VIEW.GET",
            iftab: Some(357),
            declared_argc: 2,
            ..Default::default()
        });
        self.register(FunctionDef {
            name: "VOLATILE",
            iftab: Some(237),
            declared_argc: 1,
            ..Default::default()
        });
        self.register(FunctionDef {
            name: "WHILE",
            iftab: Some(172),
            declared_argc: 1,
            ..Default::default()
        });
        self.register(FunctionDef {
            name: "WINDOW.TITLE",
            iftab: Some(263),
            declared_argc: 1,
            ..Default::default()
        });
        self.register(FunctionDef {
            name: "WINDOWS",
            iftab: Some(91),
            declared_argc: 2,
            ..Default::default()
        });
    }
}
impl Default for FunctionRegistry {
    fn default() -> Self {
        Self::new()
    }
}
