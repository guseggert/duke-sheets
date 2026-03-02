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

use crate::error::FormulaResult;
use crate::evaluator::{EvaluationContext, FormulaValue};
use std::collections::HashMap;

/// Function implementation signature
///
/// Functions can consult the evaluation context (e.g. workbook settings, date system,
/// current sheet/cell) to match Excel semantics.
pub type FunctionImpl = fn(&[FormulaValue], &EvaluationContext) -> FormulaResult<FormulaValue>;

/// Function definition
pub struct FunctionDef {
    /// Function name (uppercase)
    pub name: &'static str,
    /// Minimum arguments
    pub min_args: usize,
    /// Maximum arguments (None = unlimited)
    pub max_args: Option<usize>,
    /// Implementation
    pub implementation: FunctionImpl,
    /// Is volatile (recalculates every time)
    pub volatile: bool,
}

/// Function registry
pub struct FunctionRegistry {
    functions: HashMap<String, FunctionDef>,
}

impl FunctionRegistry {
    /// Create a new registry with all built-in functions
    pub fn new() -> Self {
        let mut registry = Self {
            functions: HashMap::new(),
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

        registry
    }

    /// Look up a function by name
    pub fn get(&self, name: &str) -> Option<&FunctionDef> {
        self.functions.get(&name.to_uppercase())
    }

    /// Register a function
    pub fn register(&mut self, def: FunctionDef) {
        self.functions.insert(def.name.to_uppercase(), def);
    }

    fn register_math_functions(&mut self) {
        // SUM
        self.register(FunctionDef {
            name: "SUM",
            min_args: 1,
            max_args: None,
            implementation: math::fn_sum,
            volatile: false,
        });

        // AVERAGE
        self.register(FunctionDef {
            name: "AVERAGE",
            min_args: 1,
            max_args: None,
            implementation: math::fn_average,
            volatile: false,
        });

        // MIN
        self.register(FunctionDef {
            name: "MIN",
            min_args: 1,
            max_args: None,
            implementation: math::fn_min,
            volatile: false,
        });

        // MAX
        self.register(FunctionDef {
            name: "MAX",
            min_args: 1,
            max_args: None,
            implementation: math::fn_max,
            volatile: false,
        });

        // COUNT
        self.register(FunctionDef {
            name: "COUNT",
            min_args: 1,
            max_args: None,
            implementation: math::fn_count,
            volatile: false,
        });

        // RAND (volatile)
        self.register(FunctionDef {
            name: "RAND",
            min_args: 0,
            max_args: Some(0),
            implementation: math::fn_rand,
            volatile: true,
        });

        // RANDBETWEEN (volatile)
        self.register(FunctionDef {
            name: "RANDBETWEEN",
            min_args: 2,
            max_args: Some(2),
            implementation: math::fn_randbetween,
            volatile: true,
        });

        // ABS
        self.register(FunctionDef {
            name: "ABS",
            min_args: 1,
            max_args: Some(1),
            implementation: math::fn_abs,
            volatile: false,
        });

        // ROUND
        self.register(FunctionDef {
            name: "ROUND",
            min_args: 1,
            max_args: Some(2),
            implementation: math::fn_round,
            volatile: false,
        });

        // MOD
        self.register(FunctionDef {
            name: "MOD",
            min_args: 2,
            max_args: Some(2),
            implementation: math::fn_mod,
            volatile: false,
        });

        // INT
        self.register(FunctionDef {
            name: "INT",
            min_args: 1,
            max_args: Some(1),
            implementation: math::fn_int,
            volatile: false,
        });

        // TRUNC
        self.register(FunctionDef {
            name: "TRUNC",
            min_args: 1,
            max_args: Some(2),
            implementation: math::fn_trunc,
            volatile: false,
        });

        // SIGN
        self.register(FunctionDef {
            name: "SIGN",
            min_args: 1,
            max_args: Some(1),
            implementation: math::fn_sign,
            volatile: false,
        });

        // SQRT
        self.register(FunctionDef {
            name: "SQRT",
            min_args: 1,
            max_args: Some(1),
            implementation: math::fn_sqrt,
            volatile: false,
        });

        // POWER
        self.register(FunctionDef {
            name: "POWER",
            min_args: 2,
            max_args: Some(2),
            implementation: math::fn_power,
            volatile: false,
        });

        // LOG
        self.register(FunctionDef {
            name: "LOG",
            min_args: 1,
            max_args: Some(2),
            implementation: math::fn_log,
            volatile: false,
        });

        // LOG10
        self.register(FunctionDef {
            name: "LOG10",
            min_args: 1,
            max_args: Some(1),
            implementation: math::fn_log10,
            volatile: false,
        });

        // LN
        self.register(FunctionDef {
            name: "LN",
            min_args: 1,
            max_args: Some(1),
            implementation: math::fn_ln,
            volatile: false,
        });

        // EXP
        self.register(FunctionDef {
            name: "EXP",
            min_args: 1,
            max_args: Some(1),
            implementation: math::fn_exp,
            volatile: false,
        });

        // PI
        self.register(FunctionDef {
            name: "PI",
            min_args: 0,
            max_args: Some(0),
            implementation: math::fn_pi,
            volatile: false,
        });

        // SUMIF
        self.register(FunctionDef {
            name: "SUMIF",
            min_args: 2,
            max_args: Some(3),
            implementation: math::fn_sumif,
            volatile: false,
        });

        // SUMIFS
        self.register(FunctionDef {
            name: "SUMIFS",
            min_args: 3,
            max_args: None, // sum_range + up to 127 criteria pairs
            implementation: math::fn_sumifs,
            volatile: false,
        });

        // SUMPRODUCT
        self.register(FunctionDef {
            name: "SUMPRODUCT",
            min_args: 1,
            max_args: None, // Up to 255 arrays
            implementation: math::fn_sumproduct,
            volatile: false,
        });

        // SIN
        self.register(FunctionDef {
            name: "SIN",
            min_args: 1,
            max_args: Some(1),
            implementation: math::fn_sin,
            volatile: false,
        });

        // COS
        self.register(FunctionDef {
            name: "COS",
            min_args: 1,
            max_args: Some(1),
            implementation: math::fn_cos,
            volatile: false,
        });

        // TAN
        self.register(FunctionDef {
            name: "TAN",
            min_args: 1,
            max_args: Some(1),
            implementation: math::fn_tan,
            volatile: false,
        });

        // ASIN
        self.register(FunctionDef {
            name: "ASIN",
            min_args: 1,
            max_args: Some(1),
            implementation: math::fn_asin,
            volatile: false,
        });

        // ACOS
        self.register(FunctionDef {
            name: "ACOS",
            min_args: 1,
            max_args: Some(1),
            implementation: math::fn_acos,
            volatile: false,
        });

        // ATAN
        self.register(FunctionDef {
            name: "ATAN",
            min_args: 1,
            max_args: Some(1),
            implementation: math::fn_atan,
            volatile: false,
        });

        // ATAN2
        self.register(FunctionDef {
            name: "ATAN2",
            min_args: 2,
            max_args: Some(2),
            implementation: math::fn_atan2,
            volatile: false,
        });

        // DEGREES
        self.register(FunctionDef {
            name: "DEGREES",
            min_args: 1,
            max_args: Some(1),
            implementation: math::fn_degrees,
            volatile: false,
        });

        // RADIANS
        self.register(FunctionDef {
            name: "RADIANS",
            min_args: 1,
            max_args: Some(1),
            implementation: math::fn_radians,
            volatile: false,
        });

        // ROUNDUP
        self.register(FunctionDef {
            name: "ROUNDUP",
            min_args: 2,
            max_args: Some(2),
            implementation: math::fn_roundup,
            volatile: false,
        });

        // ROUNDDOWN
        self.register(FunctionDef {
            name: "ROUNDDOWN",
            min_args: 2,
            max_args: Some(2),
            implementation: math::fn_rounddown,
            volatile: false,
        });

        // CEILING.MATH
        self.register(FunctionDef {
            name: "CEILING.MATH",
            min_args: 1,
            max_args: Some(3),
            implementation: math::fn_ceiling_math,
            volatile: false,
        });

        // FLOOR.MATH
        self.register(FunctionDef {
            name: "FLOOR.MATH",
            min_args: 1,
            max_args: Some(3),
            implementation: math::fn_floor_math,
            volatile: false,
        });

        // ODD
        self.register(FunctionDef {
            name: "ODD",
            min_args: 1,
            max_args: Some(1),
            implementation: math::fn_odd,
            volatile: false,
        });

        // EVEN
        self.register(FunctionDef {
            name: "EVEN",
            min_args: 1,
            max_args: Some(1),
            implementation: math::fn_even,
            volatile: false,
        });
        // --- math_extra functions ---
        self.register(FunctionDef {
            name: "ACOSH",
            min_args: 1,
            max_args: Some(1),
            implementation: math_extra::fn_acosh,
            volatile: false,
        });
        self.register(FunctionDef {
            name: "ASINH",
            min_args: 1,
            max_args: Some(1),
            implementation: math_extra::fn_asinh,
            volatile: false,
        });
        self.register(FunctionDef {
            name: "ATANH",
            min_args: 1,
            max_args: Some(1),
            implementation: math_extra::fn_atanh,
            volatile: false,
        });
        self.register(FunctionDef {
            name: "ACOTH",
            min_args: 1,
            max_args: Some(1),
            implementation: math_extra::fn_acoth,
            volatile: false,
        });
        self.register(FunctionDef {
            name: "COSH",
            min_args: 1,
            max_args: Some(1),
            implementation: math_extra::fn_cosh,
            volatile: false,
        });
        self.register(FunctionDef {
            name: "SINH",
            min_args: 1,
            max_args: Some(1),
            implementation: math_extra::fn_sinh,
            volatile: false,
        });
        self.register(FunctionDef {
            name: "TANH",
            min_args: 1,
            max_args: Some(1),
            implementation: math_extra::fn_tanh,
            volatile: false,
        });
        self.register(FunctionDef {
            name: "COT",
            min_args: 1,
            max_args: Some(1),
            implementation: math_extra::fn_cot,
            volatile: false,
        });
        self.register(FunctionDef {
            name: "COTH",
            min_args: 1,
            max_args: Some(1),
            implementation: math_extra::fn_coth,
            volatile: false,
        });
        self.register(FunctionDef {
            name: "CSC",
            min_args: 1,
            max_args: Some(1),
            implementation: math_extra::fn_csc,
            volatile: false,
        });
        self.register(FunctionDef {
            name: "CSCH",
            min_args: 1,
            max_args: Some(1),
            implementation: math_extra::fn_csch,
            volatile: false,
        });
        self.register(FunctionDef {
            name: "SEC",
            min_args: 1,
            max_args: Some(1),
            implementation: math_extra::fn_sec,
            volatile: false,
        });
        self.register(FunctionDef {
            name: "SECH",
            min_args: 1,
            max_args: Some(1),
            implementation: math_extra::fn_sech,
            volatile: false,
        });
        self.register(FunctionDef {
            name: "COMBIN",
            min_args: 2,
            max_args: Some(2),
            implementation: math_extra::fn_combin,
            volatile: false,
        });
        self.register(FunctionDef {
            name: "COMBINA",
            min_args: 2,
            max_args: Some(2),
            implementation: math_extra::fn_combina,
            volatile: false,
        });
        self.register(FunctionDef {
            name: "FACT",
            min_args: 1,
            max_args: Some(1),
            implementation: math_extra::fn_fact,
            volatile: false,
        });
        self.register(FunctionDef {
            name: "FACTDOUBLE",
            min_args: 1,
            max_args: Some(1),
            implementation: math_extra::fn_factdouble,
            volatile: false,
        });
        self.register(FunctionDef {
            name: "MULTINOMIAL",
            min_args: 1,
            max_args: None,
            implementation: math_extra::fn_multinomial,
            volatile: false,
        });
        self.register(FunctionDef {
            name: "GCD",
            min_args: 1,
            max_args: None,
            implementation: math_extra::fn_gcd,
            volatile: false,
        });
        self.register(FunctionDef {
            name: "LCM",
            min_args: 1,
            max_args: None,
            implementation: math_extra::fn_lcm,
            volatile: false,
        });
        self.register(FunctionDef {
            name: "PRODUCT",
            min_args: 1,
            max_args: None,
            implementation: math_extra::fn_product,
            volatile: false,
        });
        self.register(FunctionDef {
            name: "QUOTIENT",
            min_args: 2,
            max_args: Some(2),
            implementation: math_extra::fn_quotient,
            volatile: false,
        });
        self.register(FunctionDef {
            name: "MROUND",
            min_args: 2,
            max_args: Some(2),
            implementation: math_extra::fn_mround,
            volatile: false,
        });
        self.register(FunctionDef {
            name: "SUMSQ",
            min_args: 1,
            max_args: None,
            implementation: math_extra::fn_sumsq,
            volatile: false,
        });
        self.register(FunctionDef {
            name: "SQRTPI",
            min_args: 1,
            max_args: Some(1),
            implementation: math_extra::fn_sqrtpi,
            volatile: false,
        });
        self.register(FunctionDef {
            name: "BASE",
            min_args: 2,
            max_args: Some(3),
            implementation: math_extra::fn_base,
            volatile: false,
        });
        self.register(FunctionDef {
            name: "DECIMAL",
            min_args: 2,
            max_args: Some(2),
            implementation: math_extra::fn_decimal,
            volatile: false,
        });
        self.register(FunctionDef {
            name: "ROMAN",
            min_args: 1,
            max_args: Some(2),
            implementation: math_extra::fn_roman,
            volatile: false,
        });
        self.register(FunctionDef {
            name: "ARABIC",
            min_args: 1,
            max_args: Some(1),
            implementation: math_extra::fn_arabic,
            volatile: false,
        });
        self.register(FunctionDef {
            name: "CEILING.PRECISE",
            min_args: 1,
            max_args: Some(2),
            implementation: math_extra::fn_ceiling_precise,
            volatile: false,
        });
        self.register(FunctionDef {
            name: "FLOOR.PRECISE",
            min_args: 1,
            max_args: Some(2),
            implementation: math_extra::fn_floor_precise,
            volatile: false,
        });
        self.register(FunctionDef {
            name: "ISO.CEILING",
            min_args: 1,
            max_args: Some(2),
            implementation: math_extra::fn_iso_ceiling,
            volatile: false,
        });
        self.register(FunctionDef {
            name: "MDETERM",
            min_args: 1,
            max_args: Some(1),
            implementation: math_extra::fn_mdeterm,
            volatile: false,
        });
        self.register(FunctionDef {
            name: "MINVERSE",
            min_args: 1,
            max_args: Some(1),
            implementation: math_extra::fn_minverse,
            volatile: false,
        });
        self.register(FunctionDef {
            name: "MMULT",
            min_args: 2,
            max_args: Some(2),
            implementation: math_extra::fn_mmult,
            volatile: false,
        });
        self.register(FunctionDef {
            name: "MUNIT",
            min_args: 1,
            max_args: Some(1),
            implementation: math_extra::fn_munit,
            volatile: false,
        });
        self.register(FunctionDef {
            name: "RANDARRAY",
            min_args: 0,
            max_args: Some(5),
            implementation: math_extra::fn_randarray,
            volatile: true,
        });
        self.register(FunctionDef {
            name: "SERIESSUM",
            min_args: 4,
            max_args: Some(4),
            implementation: math_extra::fn_seriessum,
            volatile: false,
        });
        self.register(FunctionDef {
            name: "SUMX2MY2",
            min_args: 2,
            max_args: Some(2),
            implementation: math_extra::fn_sumx2my2,
            volatile: false,
        });
        self.register(FunctionDef {
            name: "SUMX2PY2",
            min_args: 2,
            max_args: Some(2),
            implementation: math_extra::fn_sumx2py2,
            volatile: false,
        });
        self.register(FunctionDef {
            name: "SUMXMY2",
            min_args: 2,
            max_args: Some(2),
            implementation: math_extra::fn_sumxmy2,
            volatile: false,
        });
        self.register(FunctionDef {
            name: "AGGREGATE",
            min_args: 3,
            max_args: None,
            implementation: math_extra::fn_aggregate,
            volatile: false,
        });
        self.register(FunctionDef {
            name: "SUBTOTAL",
            min_args: 2,
            max_args: None,
            implementation: math_extra::fn_subtotal,
            volatile: false,
        });
    }

    fn register_logical_functions(&mut self) {
        // IF
        self.register(FunctionDef {
            name: "IF",
            min_args: 2,
            max_args: Some(3),
            implementation: logical::fn_if,
            volatile: false,
        });

        // AND
        self.register(FunctionDef {
            name: "AND",
            min_args: 1,
            max_args: None,
            implementation: logical::fn_and,
            volatile: false,
        });

        // OR
        self.register(FunctionDef {
            name: "OR",
            min_args: 1,
            max_args: None,
            implementation: logical::fn_or,
            volatile: false,
        });

        // NOT
        self.register(FunctionDef {
            name: "NOT",
            min_args: 1,
            max_args: Some(1),
            implementation: logical::fn_not,
            volatile: false,
        });

        // IFERROR
        self.register(FunctionDef {
            name: "IFERROR",
            min_args: 2,
            max_args: Some(2),
            implementation: logical::fn_iferror,
            volatile: false,
        });

        // IFNA
        self.register(FunctionDef {
            name: "IFNA",
            min_args: 2,
            max_args: Some(2),
            implementation: logical::fn_ifna,
            volatile: false,
        });

        // TRUE
        self.register(FunctionDef {
            name: "TRUE",
            min_args: 0,
            max_args: Some(0),
            implementation: logical::fn_true,
            volatile: false,
        });

        // FALSE
        self.register(FunctionDef {
            name: "FALSE",
            min_args: 0,
            max_args: Some(0),
            implementation: logical::fn_false,
            volatile: false,
        });

        // XOR
        self.register(FunctionDef {
            name: "XOR",
            min_args: 1,
            max_args: None,
            implementation: logical::fn_xor,
            volatile: false,
        });

        // IFS
        self.register(FunctionDef {
            name: "IFS",
            min_args: 2,
            max_args: None, // Up to 127 condition-value pairs
            implementation: logical::fn_ifs,
            volatile: false,
        });

        // SWITCH
        self.register(FunctionDef {
            name: "SWITCH",
            min_args: 3,
            max_args: None, // Up to 126 value-result pairs + optional default
            implementation: logical::fn_switch,
            volatile: false,
        });
        // --- logical_extra functions ---
        self.register(FunctionDef {
            name: "LET",
            min_args: 3,
            max_args: None,
            implementation: logical_extra::fn_let,
            volatile: false,
        });
        self.register(FunctionDef {
            name: "LAMBDA",
            min_args: 1,
            max_args: None,
            implementation: logical_extra::fn_lambda,
            volatile: false,
        });
        self.register(FunctionDef {
            name: "MAP",
            min_args: 2,
            max_args: None,
            implementation: logical_extra::fn_map,
            volatile: false,
        });
        self.register(FunctionDef {
            name: "REDUCE",
            min_args: 3,
            max_args: Some(3),
            implementation: logical_extra::fn_reduce,
            volatile: false,
        });
        self.register(FunctionDef {
            name: "SCAN",
            min_args: 3,
            max_args: Some(3),
            implementation: logical_extra::fn_scan,
            volatile: false,
        });
        self.register(FunctionDef {
            name: "BYCOL",
            min_args: 2,
            max_args: Some(2),
            implementation: logical_extra::fn_bycol,
            volatile: false,
        });
        self.register(FunctionDef {
            name: "BYROW",
            min_args: 2,
            max_args: Some(2),
            implementation: logical_extra::fn_byrow,
            volatile: false,
        });
        self.register(FunctionDef {
            name: "MAKEARRAY",
            min_args: 3,
            max_args: Some(3),
            implementation: logical_extra::fn_makearray,
            volatile: false,
        });
    }

    fn register_text_functions(&mut self) {
        // LEN
        self.register(FunctionDef {
            name: "LEN",
            min_args: 1,
            max_args: Some(1),
            implementation: text::fn_len,
            volatile: false,
        });

        // LEFT
        self.register(FunctionDef {
            name: "LEFT",
            min_args: 1,
            max_args: Some(2),
            implementation: text::fn_left,
            volatile: false,
        });

        // RIGHT
        self.register(FunctionDef {
            name: "RIGHT",
            min_args: 1,
            max_args: Some(2),
            implementation: text::fn_right,
            volatile: false,
        });

        // MID
        self.register(FunctionDef {
            name: "MID",
            min_args: 3,
            max_args: Some(3),
            implementation: text::fn_mid,
            volatile: false,
        });

        // LOWER
        self.register(FunctionDef {
            name: "LOWER",
            min_args: 1,
            max_args: Some(1),
            implementation: text::fn_lower,
            volatile: false,
        });

        // UPPER
        self.register(FunctionDef {
            name: "UPPER",
            min_args: 1,
            max_args: Some(1),
            implementation: text::fn_upper,
            volatile: false,
        });

        // TRIM
        self.register(FunctionDef {
            name: "TRIM",
            min_args: 1,
            max_args: Some(1),
            implementation: text::fn_trim,
            volatile: false,
        });

        // CONCAT (newer)
        self.register(FunctionDef {
            name: "CONCAT",
            min_args: 1,
            max_args: None,
            implementation: text::fn_concat,
            volatile: false,
        });

        // CONCATENATE (legacy)
        self.register(FunctionDef {
            name: "CONCATENATE",
            min_args: 1,
            max_args: None,
            implementation: text::fn_concat,
            volatile: false,
        });

        // FIND (case-sensitive)
        self.register(FunctionDef {
            name: "FIND",
            min_args: 2,
            max_args: Some(3),
            implementation: text::fn_find,
            volatile: false,
        });

        // FINDB (same as FIND for non-DBCS)
        self.register(FunctionDef {
            name: "FINDB",
            min_args: 2,
            max_args: Some(3),
            implementation: text::fn_find,
            volatile: false,
        });

        // SEARCH (case-insensitive)
        self.register(FunctionDef {
            name: "SEARCH",
            min_args: 2,
            max_args: Some(3),
            implementation: text::fn_search,
            volatile: false,
        });

        // SEARCHB (same as SEARCH for non-DBCS)
        self.register(FunctionDef {
            name: "SEARCHB",
            min_args: 2,
            max_args: Some(3),
            implementation: text::fn_search,
            volatile: false,
        });

        // EXACT
        self.register(FunctionDef {
            name: "EXACT",
            min_args: 2,
            max_args: Some(2),
            implementation: text::fn_exact,
            volatile: false,
        });

        // REPT
        self.register(FunctionDef {
            name: "REPT",
            min_args: 2,
            max_args: Some(2),
            implementation: text::fn_rept,
            volatile: false,
        });

        // SUBSTITUTE
        self.register(FunctionDef {
            name: "SUBSTITUTE",
            min_args: 3,
            max_args: Some(4),
            implementation: text::fn_substitute,
            volatile: false,
        });

        // PROPER
        self.register(FunctionDef {
            name: "PROPER",
            min_args: 1,
            max_args: Some(1),
            implementation: text::fn_proper,
            volatile: false,
        });

        // CHAR
        self.register(FunctionDef {
            name: "CHAR",
            min_args: 1,
            max_args: Some(1),
            implementation: text::fn_char,
            volatile: false,
        });

        // CODE
        self.register(FunctionDef {
            name: "CODE",
            min_args: 1,
            max_args: Some(1),
            implementation: text::fn_code,
            volatile: false,
        });

        // CLEAN
        self.register(FunctionDef {
            name: "CLEAN",
            min_args: 1,
            max_args: Some(1),
            implementation: text::fn_clean,
            volatile: false,
        });

        // VALUE
        self.register(FunctionDef {
            name: "VALUE",
            min_args: 1,
            max_args: Some(1),
            implementation: text::fn_value,
            volatile: false,
        });

        // T
        self.register(FunctionDef {
            name: "T",
            min_args: 1,
            max_args: Some(1),
            implementation: text::fn_t,
            volatile: false,
        });

        // N
        self.register(FunctionDef {
            name: "N",
            min_args: 1,
            max_args: Some(1),
            implementation: text::fn_n,
            volatile: false,
        });

        self.register(FunctionDef {
            name: "TEXT",
            min_args: 2,
            max_args: Some(2),
            implementation: text::fn_text,
            volatile: false,
        });

        self.register(FunctionDef {
            name: "TEXTJOIN",
            min_args: 3,
            max_args: None,
            implementation: text::fn_textjoin,
            volatile: false,
        });

        self.register(FunctionDef {
            name: "FIXED",
            min_args: 1,
            max_args: Some(3),
            implementation: text::fn_fixed,
            volatile: false,
        });

        self.register(FunctionDef {
            name: "DOLLAR",
            min_args: 1,
            max_args: Some(2),
            implementation: text::fn_dollar,
            volatile: false,
        });

        self.register(FunctionDef {
            name: "NUMBERVALUE",
            min_args: 1,
            max_args: Some(3),
            implementation: text::fn_numbervalue,
            volatile: false,
        });

        // LENB (same as LEN for non-DBCS)
        self.register(FunctionDef {
            name: "LENB",
            min_args: 1,
            max_args: Some(1),
            implementation: text::fn_len,
            volatile: false,
        });

        // LEFTB (same as LEFT for non-DBCS)
        self.register(FunctionDef {
            name: "LEFTB",
            min_args: 1,
            max_args: Some(2),
            implementation: text::fn_left,
            volatile: false,
        });

        // RIGHTB (same as RIGHT for non-DBCS)
        self.register(FunctionDef {
            name: "RIGHTB",
            min_args: 1,
            max_args: Some(2),
            implementation: text::fn_right,
            volatile: false,
        });

        // MIDB (same as MID for non-DBCS)
        self.register(FunctionDef {
            name: "MIDB",
            min_args: 3,
            max_args: Some(3),
            implementation: text::fn_mid,
            volatile: false,
        });
        // --- text_extra functions ---
        self.register(FunctionDef {
            name: "REPLACE",
            min_args: 4,
            max_args: Some(4),
            implementation: text_extra::fn_replace,
            volatile: false,
        });
        self.register(FunctionDef {
            name: "REPLACEB",
            min_args: 4,
            max_args: Some(4),
            implementation: text_extra::fn_replaceb,
            volatile: false,
        });
        self.register(FunctionDef {
            name: "TEXTBEFORE",
            min_args: 2,
            max_args: Some(6),
            implementation: text_extra::fn_textbefore,
            volatile: false,
        });
        self.register(FunctionDef {
            name: "TEXTAFTER",
            min_args: 2,
            max_args: Some(6),
            implementation: text_extra::fn_textafter,
            volatile: false,
        });
        self.register(FunctionDef {
            name: "TEXTSPLIT",
            min_args: 2,
            max_args: Some(6),
            implementation: text_extra::fn_textsplit,
            volatile: false,
        });
        self.register(FunctionDef {
            name: "UNICHAR",
            min_args: 1,
            max_args: Some(1),
            implementation: text_extra::fn_unichar,
            volatile: false,
        });
        self.register(FunctionDef {
            name: "UNICODE",
            min_args: 1,
            max_args: Some(1),
            implementation: text_extra::fn_unicode,
            volatile: false,
        });
    }

    fn register_info_functions(&mut self) {
        // ISBLANK
        self.register(FunctionDef {
            name: "ISBLANK",
            min_args: 1,
            max_args: Some(1),
            implementation: info::fn_isblank,
            volatile: false,
        });

        // ISNUMBER
        self.register(FunctionDef {
            name: "ISNUMBER",
            min_args: 1,
            max_args: Some(1),
            implementation: info::fn_isnumber,
            volatile: false,
        });

        // ISTEXT
        self.register(FunctionDef {
            name: "ISTEXT",
            min_args: 1,
            max_args: Some(1),
            implementation: info::fn_istext,
            volatile: false,
        });

        // ISERROR
        self.register(FunctionDef {
            name: "ISERROR",
            min_args: 1,
            max_args: Some(1),
            implementation: info::fn_iserror,
            volatile: false,
        });

        // ISNA
        self.register(FunctionDef {
            name: "ISNA",
            min_args: 1,
            max_args: Some(1),
            implementation: info::fn_isna,
            volatile: false,
        });

        // NA
        self.register(FunctionDef {
            name: "NA",
            min_args: 0,
            max_args: Some(0),
            implementation: info::fn_na,
            volatile: false,
        });
        // --- info_extra functions ---
        self.register(FunctionDef {
            name: "ISERR",
            min_args: 1,
            max_args: Some(1),
            implementation: info_extra::fn_iserr,
            volatile: false,
        });
        self.register(FunctionDef {
            name: "ISEVEN",
            min_args: 1,
            max_args: Some(1),
            implementation: info_extra::fn_iseven,
            volatile: false,
        });
        self.register(FunctionDef {
            name: "ISODD",
            min_args: 1,
            max_args: Some(1),
            implementation: info_extra::fn_isodd,
            volatile: false,
        });
        self.register(FunctionDef {
            name: "ISLOGICAL",
            min_args: 1,
            max_args: Some(1),
            implementation: info_extra::fn_islogical,
            volatile: false,
        });
        self.register(FunctionDef {
            name: "ISNONTEXT",
            min_args: 1,
            max_args: Some(1),
            implementation: info_extra::fn_isnontext,
            volatile: false,
        });
        self.register(FunctionDef {
            name: "ISREF",
            min_args: 1,
            max_args: Some(1),
            implementation: info_extra::fn_isref,
            volatile: false,
        });
        self.register(FunctionDef {
            name: "ERROR.TYPE",
            min_args: 1,
            max_args: Some(1),
            implementation: info_extra::fn_error_type,
            volatile: false,
        });
        self.register(FunctionDef {
            name: "TYPE",
            min_args: 1,
            max_args: Some(1),
            implementation: info_extra::fn_type,
            volatile: false,
        });
        self.register(FunctionDef {
            name: "CELL",
            min_args: 1,
            max_args: Some(2),
            implementation: info_extra::fn_cell,
            volatile: false,
        });
        self.register(FunctionDef {
            name: "INFO",
            min_args: 1,
            max_args: Some(1),
            implementation: info_extra::fn_info,
            volatile: false,
        });
        self.register(FunctionDef {
            name: "SHEET",
            min_args: 0,
            max_args: Some(1),
            implementation: info_extra::fn_sheet,
            volatile: false,
        });
        self.register(FunctionDef {
            name: "SHEETS",
            min_args: 0,
            max_args: Some(1),
            implementation: info_extra::fn_sheets,
            volatile: false,
        });
        self.register(FunctionDef {
            name: "ISFORMULA",
            min_args: 1,
            max_args: Some(1),
            implementation: info_extra::fn_isformula,
            volatile: false,
        });
        self.register(FunctionDef {
            name: "ISOMITTED",
            min_args: 1,
            max_args: Some(1),
            implementation: info_extra::fn_isomitted,
            volatile: false,
        });
    }

    fn register_date_functions(&mut self) {
        // DATE
        self.register(FunctionDef {
            name: "DATE",
            min_args: 3,
            max_args: Some(3),
            implementation: date::fn_date,
            volatile: false,
        });

        // YEAR
        self.register(FunctionDef {
            name: "YEAR",
            min_args: 1,
            max_args: Some(1),
            implementation: date::fn_year,
            volatile: false,
        });

        // MONTH
        self.register(FunctionDef {
            name: "MONTH",
            min_args: 1,
            max_args: Some(1),
            implementation: date::fn_month,
            volatile: false,
        });

        // DAY
        self.register(FunctionDef {
            name: "DAY",
            min_args: 1,
            max_args: Some(1),
            implementation: date::fn_day,
            volatile: false,
        });

        self.register(FunctionDef {
            name: "TIME",
            min_args: 3,
            max_args: Some(3),
            implementation: date::fn_time,
            volatile: false,
        });

        self.register(FunctionDef {
            name: "HOUR",
            min_args: 1,
            max_args: Some(1),
            implementation: date::fn_hour,
            volatile: false,
        });

        self.register(FunctionDef {
            name: "MINUTE",
            min_args: 1,
            max_args: Some(1),
            implementation: date::fn_minute,
            volatile: false,
        });

        self.register(FunctionDef {
            name: "SECOND",
            min_args: 1,
            max_args: Some(1),
            implementation: date::fn_second,
            volatile: false,
        });

        self.register(FunctionDef {
            name: "WEEKDAY",
            min_args: 1,
            max_args: Some(2),
            implementation: date::fn_weekday,
            volatile: false,
        });

        self.register(FunctionDef {
            name: "WEEKNUM",
            min_args: 1,
            max_args: Some(2),
            implementation: date::fn_weeknum,
            volatile: false,
        });

        self.register(FunctionDef {
            name: "ISOWEEKNUM",
            min_args: 1,
            max_args: Some(1),
            implementation: date::fn_isoweeknum,
            volatile: false,
        });

        self.register(FunctionDef {
            name: "EDATE",
            min_args: 2,
            max_args: Some(2),
            implementation: date::fn_edate,
            volatile: false,
        });

        self.register(FunctionDef {
            name: "EOMONTH",
            min_args: 2,
            max_args: Some(2),
            implementation: date::fn_eomonth,
            volatile: false,
        });

        self.register(FunctionDef {
            name: "DAYS",
            min_args: 2,
            max_args: Some(2),
            implementation: date::fn_days,
            volatile: false,
        });

        self.register(FunctionDef {
            name: "DAYS360",
            min_args: 2,
            max_args: Some(3),
            implementation: date::fn_days360,
            volatile: false,
        });

        self.register(FunctionDef {
            name: "DATEDIF",
            min_args: 3,
            max_args: Some(3),
            implementation: date::fn_datedif,
            volatile: false,
        });

        self.register(FunctionDef {
            name: "YEARFRAC",
            min_args: 2,
            max_args: Some(3),
            implementation: date::fn_yearfrac,
            volatile: false,
        });

        self.register(FunctionDef {
            name: "DATEVALUE",
            min_args: 1,
            max_args: Some(1),
            implementation: date::fn_datevalue,
            volatile: false,
        });

        self.register(FunctionDef {
            name: "TIMEVALUE",
            min_args: 1,
            max_args: Some(1),
            implementation: date::fn_timevalue,
            volatile: false,
        });

        self.register(FunctionDef {
            name: "NETWORKDAYS",
            min_args: 2,
            max_args: Some(3),
            implementation: date::fn_networkdays,
            volatile: false,
        });

        self.register(FunctionDef {
            name: "WORKDAY",
            min_args: 2,
            max_args: Some(3),
            implementation: date::fn_workday,
            volatile: false,
        });

        // NOW (volatile)
        self.register(FunctionDef {
            name: "NOW",
            min_args: 0,
            max_args: Some(0),
            implementation: date::fn_now,
            volatile: true,
        });

        // TODAY (volatile)
        self.register(FunctionDef {
            name: "TODAY",
            min_args: 0,
            max_args: Some(0),
            implementation: date::fn_today,
            volatile: true,
        });
        // --- date functions from lookup_extra ---
        self.register(FunctionDef {
            name: "NETWORKDAYS.INTL",
            min_args: 2,
            max_args: Some(4),
            implementation: lookup_extra::fn_networkdays_intl,
            volatile: false,
        });
        self.register(FunctionDef {
            name: "WORKDAY.INTL",
            min_args: 2,
            max_args: Some(4),
            implementation: lookup_extra::fn_workday_intl,
            volatile: false,
        });
    }

    fn register_lookup_functions(&mut self) {
        // INDEX
        self.register(FunctionDef {
            name: "INDEX",
            min_args: 2,
            max_args: Some(3),
            implementation: lookup::fn_index,
            volatile: false,
        });

        // MATCH
        self.register(FunctionDef {
            name: "MATCH",
            min_args: 2,
            max_args: Some(3),
            implementation: lookup::fn_match,
            volatile: false,
        });

        // VLOOKUP
        self.register(FunctionDef {
            name: "VLOOKUP",
            min_args: 3,
            max_args: Some(4),
            implementation: lookup::fn_vlookup,
            volatile: false,
        });

        // ROWS
        self.register(FunctionDef {
            name: "ROWS",
            min_args: 1,
            max_args: Some(1),
            implementation: lookup::fn_rows,
            volatile: false,
        });

        // COLUMNS
        self.register(FunctionDef {
            name: "COLUMNS",
            min_args: 1,
            max_args: Some(1),
            implementation: lookup::fn_columns,
            volatile: false,
        });

        // CHOOSE
        self.register(FunctionDef {
            name: "CHOOSE",
            min_args: 2,
            max_args: None, // Up to 254 values
            implementation: lookup::fn_choose,
            volatile: false,
        });

        // ROW
        self.register(FunctionDef {
            name: "ROW",
            min_args: 0,
            max_args: Some(1),
            implementation: lookup::fn_row,
            volatile: false,
        });

        // COLUMN
        self.register(FunctionDef {
            name: "COLUMN",
            min_args: 0,
            max_args: Some(1),
            implementation: lookup::fn_column,
            volatile: false,
        });

        // SEQUENCE (dynamic array function)
        self.register(FunctionDef {
            name: "SEQUENCE",
            min_args: 1,
            max_args: Some(4),
            implementation: lookup::fn_sequence,
            volatile: false,
        });

        self.register(FunctionDef {
            name: "HLOOKUP",
            min_args: 3,
            max_args: Some(4),
            implementation: lookup::fn_hlookup,
            volatile: false,
        });

        self.register(FunctionDef {
            name: "XLOOKUP",
            min_args: 3,
            max_args: Some(6),
            implementation: lookup::fn_xlookup,
            volatile: false,
        });

        self.register(FunctionDef {
            name: "XMATCH",
            min_args: 2,
            max_args: Some(4),
            implementation: lookup::fn_xmatch,
            volatile: false,
        });

        self.register(FunctionDef {
            name: "INDIRECT",
            min_args: 1,
            max_args: Some(2),
            implementation: lookup::fn_indirect,
            volatile: true,
        });

        self.register(FunctionDef {
            name: "OFFSET",
            min_args: 3,
            max_args: Some(5),
            implementation: lookup::fn_offset,
            volatile: true,
        });
        // --- lookup_extra functions ---
        self.register(FunctionDef {
            name: "ADDRESS",
            min_args: 2,
            max_args: Some(5),
            implementation: lookup_extra::fn_address,
            volatile: false,
        });
        self.register(FunctionDef {
            name: "AREAS",
            min_args: 1,
            max_args: Some(1),
            implementation: lookup_extra::fn_areas,
            volatile: false,
        });
        self.register(FunctionDef {
            name: "CHOOSECOLS",
            min_args: 2,
            max_args: None,
            implementation: lookup_extra::fn_choosecols,
            volatile: false,
        });
        self.register(FunctionDef {
            name: "CHOOSEROWS",
            min_args: 2,
            max_args: None,
            implementation: lookup_extra::fn_chooserows,
            volatile: false,
        });
        self.register(FunctionDef {
            name: "DROP",
            min_args: 2,
            max_args: Some(3),
            implementation: lookup_extra::fn_drop,
            volatile: false,
        });
        self.register(FunctionDef {
            name: "EXPAND",
            min_args: 2,
            max_args: Some(4),
            implementation: lookup_extra::fn_expand,
            volatile: false,
        });
        self.register(FunctionDef {
            name: "FILTER",
            min_args: 2,
            max_args: Some(3),
            implementation: lookup_extra::fn_filter,
            volatile: false,
        });
        self.register(FunctionDef {
            name: "FORMULATEXT",
            min_args: 1,
            max_args: Some(1),
            implementation: lookup_extra::fn_formulatext,
            volatile: false,
        });
        self.register(FunctionDef {
            name: "HSTACK",
            min_args: 1,
            max_args: None,
            implementation: lookup_extra::fn_hstack,
            volatile: false,
        });
        self.register(FunctionDef {
            name: "LOOKUP",
            min_args: 2,
            max_args: Some(3),
            implementation: lookup_extra::fn_lookup,
            volatile: false,
        });
        self.register(FunctionDef {
            name: "SORT",
            min_args: 1,
            max_args: Some(4),
            implementation: lookup_extra::fn_sort,
            volatile: false,
        });
        self.register(FunctionDef {
            name: "SORTBY",
            min_args: 2,
            max_args: None,
            implementation: lookup_extra::fn_sortby,
            volatile: false,
        });
        self.register(FunctionDef {
            name: "TAKE",
            min_args: 2,
            max_args: Some(3),
            implementation: lookup_extra::fn_take,
            volatile: false,
        });
        self.register(FunctionDef {
            name: "TOCOL",
            min_args: 1,
            max_args: Some(3),
            implementation: lookup_extra::fn_tocol,
            volatile: false,
        });
        self.register(FunctionDef {
            name: "TOROW",
            min_args: 1,
            max_args: Some(3),
            implementation: lookup_extra::fn_torow,
            volatile: false,
        });
        self.register(FunctionDef {
            name: "TRANSPOSE",
            min_args: 1,
            max_args: Some(1),
            implementation: lookup_extra::fn_transpose,
            volatile: false,
        });
        self.register(FunctionDef {
            name: "UNIQUE",
            min_args: 1,
            max_args: Some(3),
            implementation: lookup_extra::fn_unique,
            volatile: false,
        });
        self.register(FunctionDef {
            name: "VSTACK",
            min_args: 1,
            max_args: None,
            implementation: lookup_extra::fn_vstack,
            volatile: false,
        });
        self.register(FunctionDef {
            name: "WRAPCOLS",
            min_args: 2,
            max_args: Some(3),
            implementation: lookup_extra::fn_wrapcols,
            volatile: false,
        });
        self.register(FunctionDef {
            name: "WRAPROWS",
            min_args: 2,
            max_args: Some(3),
            implementation: lookup_extra::fn_wraprows,
            volatile: false,
        });
        self.register(FunctionDef {
            name: "HYPERLINK",
            min_args: 1,
            max_args: Some(2),
            implementation: lookup_extra::fn_hyperlink,
            volatile: false,
        });
        self.register(FunctionDef {
            name: "GETPIVOTDATA",
            min_args: 2,
            max_args: None,
            implementation: lookup_extra::fn_getpivotdata,
            volatile: false,
        });
        self.register(FunctionDef {
            name: "RTD",
            min_args: 2,
            max_args: None,
            implementation: lookup_extra::fn_rtd,
            volatile: false,
        });
        self.register(FunctionDef {
            name: "IMAGE",
            min_args: 1,
            max_args: Some(4),
            implementation: lookup_extra::fn_image,
            volatile: false,
        });
    }

    fn register_statistical_functions(&mut self) {
        // COUNTA
        self.register(FunctionDef {
            name: "COUNTA",
            min_args: 1,
            max_args: None,
            implementation: statistical::fn_counta,
            volatile: false,
        });

        // COUNTBLANK
        self.register(FunctionDef {
            name: "COUNTBLANK",
            min_args: 1,
            max_args: None,
            implementation: statistical::fn_countblank,
            volatile: false,
        });

        // COUNTIF
        self.register(FunctionDef {
            name: "COUNTIF",
            min_args: 2,
            max_args: Some(2),
            implementation: statistical::fn_countif,
            volatile: false,
        });

        // AVERAGEIF
        self.register(FunctionDef {
            name: "AVERAGEIF",
            min_args: 2,
            max_args: Some(3),
            implementation: statistical::fn_averageif,
            volatile: false,
        });

        self.register(FunctionDef {
            name: "STDEV.S",
            min_args: 1,
            max_args: None,
            implementation: statistical::fn_stdev_s,
            volatile: false,
        });

        self.register(FunctionDef {
            name: "STDEV.P",
            min_args: 1,
            max_args: None,
            implementation: statistical::fn_stdev_p,
            volatile: false,
        });

        self.register(FunctionDef {
            name: "VAR.S",
            min_args: 1,
            max_args: None,
            implementation: statistical::fn_var_s,
            volatile: false,
        });

        self.register(FunctionDef {
            name: "VAR.P",
            min_args: 1,
            max_args: None,
            implementation: statistical::fn_var_p,
            volatile: false,
        });

        self.register(FunctionDef {
            name: "MODE.SNGL",
            min_args: 1,
            max_args: None,
            implementation: statistical::fn_mode_sngl,
            volatile: false,
        });

        self.register(FunctionDef {
            name: "MAXIFS",
            min_args: 3,
            max_args: None,
            implementation: statistical::fn_maxifs,
            volatile: false,
        });

        self.register(FunctionDef {
            name: "MINIFS",
            min_args: 3,
            max_args: None,
            implementation: statistical::fn_minifs,
            volatile: false,
        });

        self.register(FunctionDef {
            name: "RANK.EQ",
            min_args: 2,
            max_args: Some(3),
            implementation: statistical::fn_rank_eq,
            volatile: false,
        });

        self.register(FunctionDef {
            name: "RANK.AVG",
            min_args: 2,
            max_args: Some(3),
            implementation: statistical::fn_rank_avg,
            volatile: false,
        });

        self.register(FunctionDef {
            name: "PERCENTILE.INC",
            min_args: 2,
            max_args: Some(2),
            implementation: statistical::fn_percentile_inc,
            volatile: false,
        });

        self.register(FunctionDef {
            name: "PERCENTILE.EXC",
            min_args: 2,
            max_args: Some(2),
            implementation: statistical::fn_percentile_exc,
            volatile: false,
        });

        self.register(FunctionDef {
            name: "QUARTILE.INC",
            min_args: 2,
            max_args: Some(2),
            implementation: statistical::fn_quartile_inc,
            volatile: false,
        });

        self.register(FunctionDef {
            name: "QUARTILE.EXC",
            min_args: 2,
            max_args: Some(2),
            implementation: statistical::fn_quartile_exc,
            volatile: false,
        });

        self.register(FunctionDef {
            name: "PERCENTRANK.INC",
            min_args: 2,
            max_args: Some(3),
            implementation: statistical::fn_percentrank_inc,
            volatile: false,
        });

        self.register(FunctionDef {
            name: "PERCENTRANK.EXC",
            min_args: 2,
            max_args: Some(3),
            implementation: statistical::fn_percentrank_exc,
            volatile: false,
        });

        self.register(FunctionDef {
            name: "STDEV",
            min_args: 1,
            max_args: None,
            implementation: statistical::fn_stdev,
            volatile: false,
        });

        self.register(FunctionDef {
            name: "STDEVP",
            min_args: 1,
            max_args: None,
            implementation: statistical::fn_stdevp,
            volatile: false,
        });

        self.register(FunctionDef {
            name: "VAR",
            min_args: 1,
            max_args: None,
            implementation: statistical::fn_var,
            volatile: false,
        });

        self.register(FunctionDef {
            name: "VARP",
            min_args: 1,
            max_args: None,
            implementation: statistical::fn_varp,
            volatile: false,
        });

        self.register(FunctionDef {
            name: "MODE",
            min_args: 1,
            max_args: None,
            implementation: statistical::fn_mode,
            volatile: false,
        });

        self.register(FunctionDef {
            name: "PERCENTILE",
            min_args: 2,
            max_args: Some(2),
            implementation: statistical::fn_percentile,
            volatile: false,
        });

        self.register(FunctionDef {
            name: "QUARTILE",
            min_args: 2,
            max_args: Some(2),
            implementation: statistical::fn_quartile,
            volatile: false,
        });

        self.register(FunctionDef {
            name: "RANK",
            min_args: 2,
            max_args: Some(3),
            implementation: statistical::fn_rank,
            volatile: false,
        });

        self.register(FunctionDef {
            name: "PERCENTRANK",
            min_args: 2,
            max_args: Some(3),
            implementation: statistical::fn_percentrank,
            volatile: false,
        });

        // MEDIAN
        self.register(FunctionDef {
            name: "MEDIAN",
            min_args: 1,
            max_args: None,
            implementation: statistical::fn_median,
            volatile: false,
        });

        // LARGE
        self.register(FunctionDef {
            name: "LARGE",
            min_args: 2,
            max_args: Some(2),
            implementation: statistical::fn_large,
            volatile: false,
        });

        // SMALL
        self.register(FunctionDef {
            name: "SMALL",
            min_args: 2,
            max_args: Some(2),
            implementation: statistical::fn_small,
            volatile: false,
        });

        // COUNTIFS
        self.register(FunctionDef {
            name: "COUNTIFS",
            min_args: 2,
            max_args: None, // Up to 127 criteria pairs
            implementation: statistical::fn_countifs,
            volatile: false,
        });

        // AVERAGEIFS
        self.register(FunctionDef {
            name: "AVERAGEIFS",
            min_args: 3,
            max_args: None, // avg_range + up to 127 criteria pairs
            implementation: statistical::fn_averageifs,
            volatile: false,
        });

        self.register(FunctionDef {
            name: "NORM.DIST",
            min_args: 4,
            max_args: Some(4),
            implementation: statistical::fn_norm_dist,
            volatile: false,
        });
        self.register(FunctionDef {
            name: "NORM.S.DIST",
            min_args: 2,
            max_args: Some(2),
            implementation: statistical::fn_norm_s_dist,
            volatile: false,
        });
        self.register(FunctionDef {
            name: "NORM.INV",
            min_args: 3,
            max_args: Some(3),
            implementation: statistical::fn_norm_inv,
            volatile: false,
        });
        self.register(FunctionDef {
            name: "NORM.S.INV",
            min_args: 1,
            max_args: Some(1),
            implementation: statistical::fn_norm_s_inv,
            volatile: false,
        });
        self.register(FunctionDef {
            name: "PHI",
            min_args: 1,
            max_args: Some(1),
            implementation: statistical::fn_phi,
            volatile: false,
        });
        self.register(FunctionDef {
            name: "BINOM.DIST",
            min_args: 4,
            max_args: Some(4),
            implementation: statistical::fn_binom_dist,
            volatile: false,
        });
        self.register(FunctionDef {
            name: "BINOM.DIST.RANGE",
            min_args: 3,
            max_args: Some(4),
            implementation: statistical::fn_binom_dist_range,
            volatile: false,
        });
        self.register(FunctionDef {
            name: "BINOM.INV",
            min_args: 3,
            max_args: Some(3),
            implementation: statistical::fn_binom_inv,
            volatile: false,
        });
        self.register(FunctionDef {
            name: "CHISQ.DIST",
            min_args: 3,
            max_args: Some(3),
            implementation: statistical::fn_chisq_dist,
            volatile: false,
        });
        self.register(FunctionDef {
            name: "CHISQ.DIST.RT",
            min_args: 2,
            max_args: Some(2),
            implementation: statistical::fn_chisq_dist_rt,
            volatile: false,
        });
        self.register(FunctionDef {
            name: "CHISQ.INV",
            min_args: 2,
            max_args: Some(2),
            implementation: statistical::fn_chisq_inv,
            volatile: false,
        });
        self.register(FunctionDef {
            name: "CHISQ.INV.RT",
            min_args: 2,
            max_args: Some(2),
            implementation: statistical::fn_chisq_inv_rt,
            volatile: false,
        });
        self.register(FunctionDef {
            name: "CHISQ.TEST",
            min_args: 2,
            max_args: Some(2),
            implementation: statistical::fn_chisq_test,
            volatile: false,
        });
        self.register(FunctionDef {
            name: "T.DIST",
            min_args: 3,
            max_args: Some(3),
            implementation: statistical::fn_t_dist,
            volatile: false,
        });
        self.register(FunctionDef {
            name: "T.DIST.2T",
            min_args: 2,
            max_args: Some(2),
            implementation: statistical::fn_t_dist_2t,
            volatile: false,
        });
        self.register(FunctionDef {
            name: "T.DIST.RT",
            min_args: 2,
            max_args: Some(2),
            implementation: statistical::fn_t_dist_rt,
            volatile: false,
        });
        self.register(FunctionDef {
            name: "T.INV",
            min_args: 2,
            max_args: Some(2),
            implementation: statistical::fn_t_inv,
            volatile: false,
        });
        self.register(FunctionDef {
            name: "T.INV.2T",
            min_args: 2,
            max_args: Some(2),
            implementation: statistical::fn_t_inv_2t,
            volatile: false,
        });
        self.register(FunctionDef {
            name: "T.TEST",
            min_args: 4,
            max_args: Some(4),
            implementation: statistical::fn_t_test,
            volatile: false,
        });
        self.register(FunctionDef {
            name: "F.DIST",
            min_args: 4,
            max_args: Some(4),
            implementation: statistical::fn_f_dist,
            volatile: false,
        });
        self.register(FunctionDef {
            name: "F.DIST.RT",
            min_args: 3,
            max_args: Some(3),
            implementation: statistical::fn_f_dist_rt,
            volatile: false,
        });
        self.register(FunctionDef {
            name: "F.INV",
            min_args: 3,
            max_args: Some(3),
            implementation: statistical::fn_f_inv,
            volatile: false,
        });
        self.register(FunctionDef {
            name: "F.INV.RT",
            min_args: 3,
            max_args: Some(3),
            implementation: statistical::fn_f_inv_rt,
            volatile: false,
        });
        self.register(FunctionDef {
            name: "F.TEST",
            min_args: 2,
            max_args: Some(2),
            implementation: statistical::fn_f_test,
            volatile: false,
        });
        self.register(FunctionDef {
            name: "GAMMA",
            min_args: 1,
            max_args: Some(1),
            implementation: statistical::fn_gamma,
            volatile: false,
        });
        self.register(FunctionDef {
            name: "GAMMA.DIST",
            min_args: 4,
            max_args: Some(4),
            implementation: statistical::fn_gamma_dist,
            volatile: false,
        });
        self.register(FunctionDef {
            name: "GAMMA.INV",
            min_args: 3,
            max_args: Some(3),
            implementation: statistical::fn_gamma_inv,
            volatile: false,
        });
        self.register(FunctionDef {
            name: "GAMMALN",
            min_args: 1,
            max_args: Some(1),
            implementation: statistical::fn_gammaln,
            volatile: false,
        });
        self.register(FunctionDef {
            name: "GAMMALN.PRECISE",
            min_args: 1,
            max_args: Some(1),
            implementation: statistical::fn_gammaln_precise,
            volatile: false,
        });
        self.register(FunctionDef {
            name: "BETA.DIST",
            min_args: 4,
            max_args: Some(6),
            implementation: statistical::fn_beta_dist,
            volatile: false,
        });
        self.register(FunctionDef {
            name: "EXPON.DIST",
            min_args: 3,
            max_args: Some(3),
            implementation: statistical::fn_expon_dist,
            volatile: false,
        });
        self.register(FunctionDef {
            name: "HYPGEOM.DIST",
            min_args: 4,
            max_args: Some(5),
            implementation: statistical::fn_hypgeom_dist,
            volatile: false,
        });
        self.register(FunctionDef {
            name: "NEGBINOM.DIST",
            min_args: 3,
            max_args: Some(4),
            implementation: statistical::fn_negbinom_dist,
            volatile: false,
        });
        self.register(FunctionDef {
            name: "POISSON.DIST",
            min_args: 3,
            max_args: Some(3),
            implementation: statistical::fn_poisson_dist,
            volatile: false,
        });
        self.register(FunctionDef {
            name: "WEIBULL.DIST",
            min_args: 4,
            max_args: Some(4),
            implementation: statistical::fn_weibull_dist,
            volatile: false,
        });

        self.register(FunctionDef {
            name: "AVEDEV",
            min_args: 1,
            max_args: None,
            implementation: statistical::fn_avedev,
            volatile: false,
        });
        self.register(FunctionDef {
            name: "AVERAGEA",
            min_args: 1,
            max_args: None,
            implementation: statistical::fn_averagea,
            volatile: false,
        });
        self.register(FunctionDef {
            name: "DEVSQ",
            min_args: 1,
            max_args: None,
            implementation: statistical::fn_devsq,
            volatile: false,
        });
        self.register(FunctionDef {
            name: "GEOMEAN",
            min_args: 1,
            max_args: None,
            implementation: statistical::fn_geomean,
            volatile: false,
        });
        self.register(FunctionDef {
            name: "HARMEAN",
            min_args: 1,
            max_args: None,
            implementation: statistical::fn_harmean,
            volatile: false,
        });
        self.register(FunctionDef {
            name: "KURT",
            min_args: 4,
            max_args: None,
            implementation: statistical::fn_kurt,
            volatile: false,
        });
        self.register(FunctionDef {
            name: "SKEW",
            min_args: 3,
            max_args: None,
            implementation: statistical::fn_skew,
            volatile: false,
        });
        self.register(FunctionDef {
            name: "SKEW.P",
            min_args: 3,
            max_args: None,
            implementation: statistical::fn_skew_p,
            volatile: false,
        });
        self.register(FunctionDef {
            name: "TRIMMEAN",
            min_args: 2,
            max_args: Some(2),
            implementation: statistical::fn_trimmean,
            volatile: false,
        });
        self.register(FunctionDef {
            name: "STANDARDIZE",
            min_args: 3,
            max_args: Some(3),
            implementation: statistical::fn_standardize,
            volatile: false,
        });
        self.register(FunctionDef {
            name: "CORREL",
            min_args: 2,
            max_args: Some(2),
            implementation: statistical::fn_correl,
            volatile: false,
        });
        self.register(FunctionDef {
            name: "COVARIANCE.P",
            min_args: 2,
            max_args: Some(2),
            implementation: statistical::fn_covariance_p,
            volatile: false,
        });
        self.register(FunctionDef {
            name: "COVARIANCE.S",
            min_args: 2,
            max_args: Some(2),
            implementation: statistical::fn_covariance_s,
            volatile: false,
        });
        self.register(FunctionDef {
            name: "PEARSON",
            min_args: 2,
            max_args: Some(2),
            implementation: statistical::fn_pearson,
            volatile: false,
        });
        self.register(FunctionDef {
            name: "RSQ",
            min_args: 2,
            max_args: Some(2),
            implementation: statistical::fn_rsq,
            volatile: false,
        });
        self.register(FunctionDef {
            name: "SLOPE",
            min_args: 2,
            max_args: Some(2),
            implementation: statistical::fn_slope,
            volatile: false,
        });
        self.register(FunctionDef {
            name: "INTERCEPT",
            min_args: 2,
            max_args: Some(2),
            implementation: statistical::fn_intercept,
            volatile: false,
        });
        self.register(FunctionDef {
            name: "FISHER",
            min_args: 1,
            max_args: Some(1),
            implementation: statistical::fn_fisher,
            volatile: false,
        });
        self.register(FunctionDef {
            name: "FISHERINV",
            min_args: 1,
            max_args: Some(1),
            implementation: statistical::fn_fisherinv,
            volatile: false,
        });
        self.register(FunctionDef {
            name: "FORECAST.LINEAR",
            min_args: 3,
            max_args: Some(3),
            implementation: statistical::fn_forecast_linear,
            volatile: false,
        });
        self.register(FunctionDef {
            name: "FORECAST",
            min_args: 3,
            max_args: Some(3),
            implementation: statistical::fn_forecast,
            volatile: false,
        });
        self.register(FunctionDef {
            name: "FREQUENCY",
            min_args: 2,
            max_args: Some(2),
            implementation: statistical::fn_frequency,
            volatile: false,
        });
        self.register(FunctionDef {
            name: "MAXA",
            min_args: 1,
            max_args: None,
            implementation: statistical::fn_maxa,
            volatile: false,
        });
        self.register(FunctionDef {
            name: "MINA",
            min_args: 1,
            max_args: None,
            implementation: statistical::fn_mina,
            volatile: false,
        });
        self.register(FunctionDef {
            name: "STEYX",
            min_args: 2,
            max_args: Some(2),
            implementation: statistical::fn_steyx,
            volatile: false,
        });
        self.register(FunctionDef {
            name: "PROB",
            min_args: 3,
            max_args: Some(4),
            implementation: statistical::fn_prob,
            volatile: false,
        });
        self.register(FunctionDef {
            name: "PERMUT",
            min_args: 2,
            max_args: Some(2),
            implementation: statistical::fn_permut,
            volatile: false,
        });
        self.register(FunctionDef {
            name: "PERMUTATIONA",
            min_args: 2,
            max_args: Some(2),
            implementation: statistical::fn_permutationa,
            volatile: false,
        });
        self.register(FunctionDef {
            name: "CONFIDENCE.NORM",
            min_args: 3,
            max_args: Some(3),
            implementation: statistical::fn_confidence_norm,
            volatile: false,
        });
        self.register(FunctionDef {
            name: "CONFIDENCE.T",
            min_args: 3,
            max_args: Some(3),
            implementation: statistical::fn_confidence_t,
            volatile: false,
        });
        self.register(FunctionDef {
            name: "GAUSS",
            min_args: 1,
            max_args: Some(1),
            implementation: statistical::fn_gauss,
            volatile: false,
        });
        self.register(FunctionDef {
            name: "MODE.MULT",
            min_args: 1,
            max_args: None,
            implementation: statistical::fn_mode_mult,
            volatile: false,
        });
        self.register(FunctionDef {
            name: "STDEVA",
            min_args: 1,
            max_args: None,
            implementation: statistical::fn_stdeva,
            volatile: false,
        });
        self.register(FunctionDef {
            name: "STDEVPA",
            min_args: 1,
            max_args: None,
            implementation: statistical::fn_stdevpa,
            volatile: false,
        });
        self.register(FunctionDef {
            name: "VARA",
            min_args: 1,
            max_args: None,
            implementation: statistical::fn_vara,
            volatile: false,
        });
        self.register(FunctionDef {
            name: "VARPA",
            min_args: 1,
            max_args: None,
            implementation: statistical::fn_varpa,
            volatile: false,
        });
        // --- statistical_extra functions ---
        self.register(FunctionDef {
            name: "LOGNORM.DIST",
            min_args: 4,
            max_args: Some(4),
            implementation: statistical_extra::fn_lognorm_dist,
            volatile: false,
        });
        self.register(FunctionDef {
            name: "LOGNORM.INV",
            min_args: 3,
            max_args: Some(3),
            implementation: statistical_extra::fn_lognorm_inv,
            volatile: false,
        });
        self.register(FunctionDef {
            name: "LINEST",
            min_args: 1,
            max_args: Some(4),
            implementation: statistical_extra::fn_linest,
            volatile: false,
        });
        self.register(FunctionDef {
            name: "LOGEST",
            min_args: 1,
            max_args: Some(4),
            implementation: statistical_extra::fn_logest,
            volatile: false,
        });
        self.register(FunctionDef {
            name: "GROWTH",
            min_args: 1,
            max_args: Some(4),
            implementation: statistical_extra::fn_growth,
            volatile: false,
        });
        self.register(FunctionDef {
            name: "TREND",
            min_args: 1,
            max_args: Some(4),
            implementation: statistical_extra::fn_trend,
            volatile: false,
        });
        self.register(FunctionDef {
            name: "FORECAST.ETS",
            min_args: 3,
            max_args: Some(6),
            implementation: statistical_extra::fn_forecast_ets,
            volatile: false,
        });
        self.register(FunctionDef {
            name: "FORECAST.ETS.CONFINT",
            min_args: 3,
            max_args: Some(7),
            implementation: statistical_extra::fn_forecast_ets_confint,
            volatile: false,
        });
        self.register(FunctionDef {
            name: "FORECAST.ETS.SEASONALITY",
            min_args: 2,
            max_args: Some(4),
            implementation: statistical_extra::fn_forecast_ets_seasonality,
            volatile: false,
        });
        self.register(FunctionDef {
            name: "FORECAST.ETS.STAT",
            min_args: 2,
            max_args: Some(5),
            implementation: statistical_extra::fn_forecast_ets_stat,
            volatile: false,
        });
    }

    fn register_financial_functions(&mut self) {
        // PMT
        self.register(FunctionDef {
            name: "PMT",
            min_args: 3,
            max_args: Some(5),
            implementation: financial::fn_pmt,
            volatile: false,
        });

        // FV
        self.register(FunctionDef {
            name: "FV",
            min_args: 3,
            max_args: Some(5),
            implementation: financial::fn_fv,
            volatile: false,
        });

        // PV
        self.register(FunctionDef {
            name: "PV",
            min_args: 3,
            max_args: Some(5),
            implementation: financial::fn_pv,
            volatile: false,
        });

        // NPER
        self.register(FunctionDef {
            name: "NPER",
            min_args: 3,
            max_args: Some(5),
            implementation: financial::fn_nper,
            volatile: false,
        });

        // RATE
        self.register(FunctionDef {
            name: "RATE",
            min_args: 3,
            max_args: Some(6),
            implementation: financial::fn_rate,
            volatile: false,
        });

        // IPMT
        self.register(FunctionDef {
            name: "IPMT",
            min_args: 4,
            max_args: Some(6),
            implementation: financial::fn_ipmt,
            volatile: false,
        });

        // PPMT
        self.register(FunctionDef {
            name: "PPMT",
            min_args: 4,
            max_args: Some(6),
            implementation: financial::fn_ppmt,
            volatile: false,
        });

        // CUMIPMT
        self.register(FunctionDef {
            name: "CUMIPMT",
            min_args: 6,
            max_args: Some(6),
            implementation: financial::fn_cumipmt,
            volatile: false,
        });

        // CUMPRINC
        self.register(FunctionDef {
            name: "CUMPRINC",
            min_args: 6,
            max_args: Some(6),
            implementation: financial::fn_cumprinc,
            volatile: false,
        });

        // NPV
        self.register(FunctionDef {
            name: "NPV",
            min_args: 2,
            max_args: None,
            implementation: financial::fn_npv,
            volatile: false,
        });

        // IRR
        self.register(FunctionDef {
            name: "IRR",
            min_args: 1,
            max_args: Some(2),
            implementation: financial::fn_irr,
            volatile: false,
        });

        // MIRR
        self.register(FunctionDef {
            name: "MIRR",
            min_args: 3,
            max_args: Some(3),
            implementation: financial::fn_mirr,
            volatile: false,
        });

        // XNPV
        self.register(FunctionDef {
            name: "XNPV",
            min_args: 3,
            max_args: Some(3),
            implementation: financial::fn_xnpv,
            volatile: false,
        });

        // SLN
        self.register(FunctionDef {
            name: "SLN",
            min_args: 3,
            max_args: Some(3),
            implementation: financial::fn_sln,
            volatile: false,
        });

        // SYD
        self.register(FunctionDef {
            name: "SYD",
            min_args: 4,
            max_args: Some(4),
            implementation: financial::fn_syd,
            volatile: false,
        });

        // DB
        self.register(FunctionDef {
            name: "DB",
            min_args: 4,
            max_args: Some(5),
            implementation: financial::fn_db,
            volatile: false,
        });

        // DDB
        self.register(FunctionDef {
            name: "DDB",
            min_args: 4,
            max_args: Some(5),
            implementation: financial::fn_ddb,
            volatile: false,
        });

        // EFFECT
        self.register(FunctionDef {
            name: "EFFECT",
            min_args: 2,
            max_args: Some(2),
            implementation: financial::fn_effect,
            volatile: false,
        });

        // NOMINAL
        self.register(FunctionDef {
            name: "NOMINAL",
            min_args: 2,
            max_args: Some(2),
            implementation: financial::fn_nominal,
            volatile: false,
        });

        // PDURATION
        self.register(FunctionDef {
            name: "PDURATION",
            min_args: 3,
            max_args: Some(3),
            implementation: financial::fn_pduration,
            volatile: false,
        });
        // --- financial_extra functions ---
        self.register(FunctionDef {
            name: "ACCRINT",
            min_args: 6,
            max_args: Some(8),
            implementation: financial_extra::fn_accrint,
            volatile: false,
        });
        self.register(FunctionDef {
            name: "ACCRINTM",
            min_args: 3,
            max_args: Some(5),
            implementation: financial_extra::fn_accrintm,
            volatile: false,
        });
        self.register(FunctionDef {
            name: "AMORDEGRC",
            min_args: 6,
            max_args: Some(7),
            implementation: financial_extra::fn_amordegrc,
            volatile: false,
        });
        self.register(FunctionDef {
            name: "AMORLINC",
            min_args: 6,
            max_args: Some(7),
            implementation: financial_extra::fn_amorlinc,
            volatile: false,
        });
        self.register(FunctionDef {
            name: "COUPDAYBS",
            min_args: 3,
            max_args: Some(4),
            implementation: financial_extra::fn_coupdaybs,
            volatile: false,
        });
        self.register(FunctionDef {
            name: "COUPDAYS",
            min_args: 3,
            max_args: Some(4),
            implementation: financial_extra::fn_coupdays,
            volatile: false,
        });
        self.register(FunctionDef {
            name: "COUPDAYSNC",
            min_args: 3,
            max_args: Some(4),
            implementation: financial_extra::fn_coupdaysnc,
            volatile: false,
        });
        self.register(FunctionDef {
            name: "COUPNCD",
            min_args: 3,
            max_args: Some(4),
            implementation: financial_extra::fn_coupncd,
            volatile: false,
        });
        self.register(FunctionDef {
            name: "COUPNUM",
            min_args: 3,
            max_args: Some(4),
            implementation: financial_extra::fn_coupnum,
            volatile: false,
        });
        self.register(FunctionDef {
            name: "COUPPCD",
            min_args: 3,
            max_args: Some(4),
            implementation: financial_extra::fn_couppcd,
            volatile: false,
        });
        self.register(FunctionDef {
            name: "DISC",
            min_args: 4,
            max_args: Some(5),
            implementation: financial_extra::fn_disc,
            volatile: false,
        });
        self.register(FunctionDef {
            name: "DOLLARDE",
            min_args: 2,
            max_args: Some(2),
            implementation: financial_extra::fn_dollarde,
            volatile: false,
        });
        self.register(FunctionDef {
            name: "DOLLARFR",
            min_args: 2,
            max_args: Some(2),
            implementation: financial_extra::fn_dollarfr,
            volatile: false,
        });
        self.register(FunctionDef {
            name: "DURATION",
            min_args: 5,
            max_args: Some(6),
            implementation: financial_extra::fn_duration,
            volatile: false,
        });
        self.register(FunctionDef {
            name: "FVSCHEDULE",
            min_args: 2,
            max_args: Some(2),
            implementation: financial_extra::fn_fvschedule,
            volatile: false,
        });
        self.register(FunctionDef {
            name: "INTRATE",
            min_args: 4,
            max_args: Some(5),
            implementation: financial_extra::fn_intrate,
            volatile: false,
        });
        self.register(FunctionDef {
            name: "ISPMT",
            min_args: 4,
            max_args: Some(4),
            implementation: financial_extra::fn_ispmt,
            volatile: false,
        });
        self.register(FunctionDef {
            name: "MDURATION",
            min_args: 5,
            max_args: Some(6),
            implementation: financial_extra::fn_mduration,
            volatile: false,
        });
        self.register(FunctionDef {
            name: "ODDFPRICE",
            min_args: 8,
            max_args: Some(9),
            implementation: financial_extra::fn_oddfprice,
            volatile: false,
        });
        self.register(FunctionDef {
            name: "ODDFYIELD",
            min_args: 8,
            max_args: Some(9),
            implementation: financial_extra::fn_oddfyield,
            volatile: false,
        });
        self.register(FunctionDef {
            name: "ODDLPRICE",
            min_args: 7,
            max_args: Some(8),
            implementation: financial_extra::fn_oddlprice,
            volatile: false,
        });
        self.register(FunctionDef {
            name: "ODDLYIELD",
            min_args: 7,
            max_args: Some(8),
            implementation: financial_extra::fn_oddlyield,
            volatile: false,
        });
        self.register(FunctionDef {
            name: "PRICE",
            min_args: 6,
            max_args: Some(7),
            implementation: financial_extra::fn_price,
            volatile: false,
        });
        self.register(FunctionDef {
            name: "PRICEDISC",
            min_args: 4,
            max_args: Some(5),
            implementation: financial_extra::fn_pricedisc,
            volatile: false,
        });
        self.register(FunctionDef {
            name: "PRICEMAT",
            min_args: 5,
            max_args: Some(6),
            implementation: financial_extra::fn_pricemat,
            volatile: false,
        });
        self.register(FunctionDef {
            name: "RECEIVED",
            min_args: 4,
            max_args: Some(5),
            implementation: financial_extra::fn_received,
            volatile: false,
        });
        self.register(FunctionDef {
            name: "RRI",
            min_args: 3,
            max_args: Some(3),
            implementation: financial_extra::fn_rri,
            volatile: false,
        });
        self.register(FunctionDef {
            name: "TBILLEQ",
            min_args: 3,
            max_args: Some(3),
            implementation: financial_extra::fn_tbilleq,
            volatile: false,
        });
        self.register(FunctionDef {
            name: "TBILLPRICE",
            min_args: 3,
            max_args: Some(3),
            implementation: financial_extra::fn_tbillprice,
            volatile: false,
        });
        self.register(FunctionDef {
            name: "TBILLYIELD",
            min_args: 3,
            max_args: Some(3),
            implementation: financial_extra::fn_tbillyield,
            volatile: false,
        });
        self.register(FunctionDef {
            name: "VDB",
            min_args: 5,
            max_args: Some(7),
            implementation: financial_extra::fn_vdb,
            volatile: false,
        });
        self.register(FunctionDef {
            name: "XIRR",
            min_args: 2,
            max_args: Some(3),
            implementation: financial_extra::fn_xirr,
            volatile: false,
        });
        self.register(FunctionDef {
            name: "YIELD",
            min_args: 6,
            max_args: Some(7),
            implementation: financial_extra::fn_yield,
            volatile: false,
        });
        self.register(FunctionDef {
            name: "YIELDDISC",
            min_args: 4,
            max_args: Some(5),
            implementation: financial_extra::fn_yielddisc,
            volatile: false,
        });
        self.register(FunctionDef {
            name: "YIELDMAT",
            min_args: 5,
            max_args: Some(6),
            implementation: financial_extra::fn_yieldmat,
            volatile: false,
        });
    }

    fn register_engineering_functions(&mut self) {
        self.register(FunctionDef {
            name: "BIN2DEC",
            min_args: 1,
            max_args: Some(1),
            implementation: engineering::fn_bin2dec,
            volatile: false,
        });
        self.register(FunctionDef {
            name: "BIN2HEX",
            min_args: 1,
            max_args: Some(2),
            implementation: engineering::fn_bin2hex,
            volatile: false,
        });
        self.register(FunctionDef {
            name: "BIN2OCT",
            min_args: 1,
            max_args: Some(2),
            implementation: engineering::fn_bin2oct,
            volatile: false,
        });
        self.register(FunctionDef {
            name: "DEC2BIN",
            min_args: 1,
            max_args: Some(2),
            implementation: engineering::fn_dec2bin,
            volatile: false,
        });
        self.register(FunctionDef {
            name: "DEC2HEX",
            min_args: 1,
            max_args: Some(2),
            implementation: engineering::fn_dec2hex,
            volatile: false,
        });
        self.register(FunctionDef {
            name: "DEC2OCT",
            min_args: 1,
            max_args: Some(2),
            implementation: engineering::fn_dec2oct,
            volatile: false,
        });
        self.register(FunctionDef {
            name: "HEX2BIN",
            min_args: 1,
            max_args: Some(2),
            implementation: engineering::fn_hex2bin,
            volatile: false,
        });
        self.register(FunctionDef {
            name: "HEX2DEC",
            min_args: 1,
            max_args: Some(1),
            implementation: engineering::fn_hex2dec,
            volatile: false,
        });
        self.register(FunctionDef {
            name: "HEX2OCT",
            min_args: 1,
            max_args: Some(2),
            implementation: engineering::fn_hex2oct,
            volatile: false,
        });
        self.register(FunctionDef {
            name: "OCT2BIN",
            min_args: 1,
            max_args: Some(2),
            implementation: engineering::fn_oct2bin,
            volatile: false,
        });
        self.register(FunctionDef {
            name: "OCT2DEC",
            min_args: 1,
            max_args: Some(1),
            implementation: engineering::fn_oct2dec,
            volatile: false,
        });
        self.register(FunctionDef {
            name: "OCT2HEX",
            min_args: 1,
            max_args: Some(2),
            implementation: engineering::fn_oct2hex,
            volatile: false,
        });
        self.register(FunctionDef {
            name: "BITAND",
            min_args: 2,
            max_args: Some(2),
            implementation: engineering::fn_bitand,
            volatile: false,
        });
        self.register(FunctionDef {
            name: "BITOR",
            min_args: 2,
            max_args: Some(2),
            implementation: engineering::fn_bitor,
            volatile: false,
        });
        self.register(FunctionDef {
            name: "BITXOR",
            min_args: 2,
            max_args: Some(2),
            implementation: engineering::fn_bitxor,
            volatile: false,
        });
        self.register(FunctionDef {
            name: "BITLSHIFT",
            min_args: 2,
            max_args: Some(2),
            implementation: engineering::fn_bitlshift,
            volatile: false,
        });
        self.register(FunctionDef {
            name: "BITRSHIFT",
            min_args: 2,
            max_args: Some(2),
            implementation: engineering::fn_bitrshift,
            volatile: false,
        });
        self.register(FunctionDef {
            name: "DELTA",
            min_args: 1,
            max_args: Some(2),
            implementation: engineering::fn_delta,
            volatile: false,
        });
        self.register(FunctionDef {
            name: "GESTEP",
            min_args: 1,
            max_args: Some(2),
            implementation: engineering::fn_gestep,
            volatile: false,
        });
        self.register(FunctionDef {
            name: "ERF",
            min_args: 1,
            max_args: Some(2),
            implementation: engineering::fn_erf,
            volatile: false,
        });
        self.register(FunctionDef {
            name: "ERF.PRECISE",
            min_args: 1,
            max_args: Some(1),
            implementation: engineering::fn_erf_precise,
            volatile: false,
        });
        self.register(FunctionDef {
            name: "ERFC",
            min_args: 1,
            max_args: Some(1),
            implementation: engineering::fn_erfc,
            volatile: false,
        });
        self.register(FunctionDef {
            name: "ERFC.PRECISE",
            min_args: 1,
            max_args: Some(1),
            implementation: engineering::fn_erfc_precise,
            volatile: false,
        });
        self.register(FunctionDef {
            name: "COMPLEX",
            min_args: 2,
            max_args: Some(3),
            implementation: engineering::fn_complex,
            volatile: false,
        });
        self.register(FunctionDef {
            name: "IMABS",
            min_args: 1,
            max_args: Some(1),
            implementation: engineering::fn_imabs,
            volatile: false,
        });
        self.register(FunctionDef {
            name: "IMAGINARY",
            min_args: 1,
            max_args: Some(1),
            implementation: engineering::fn_imaginary,
            volatile: false,
        });
        self.register(FunctionDef {
            name: "IMARGUMENT",
            min_args: 1,
            max_args: Some(1),
            implementation: engineering::fn_imargument,
            volatile: false,
        });
        self.register(FunctionDef {
            name: "IMCONJUGATE",
            min_args: 1,
            max_args: Some(1),
            implementation: engineering::fn_imconjugate,
            volatile: false,
        });
        self.register(FunctionDef {
            name: "IMCOS",
            min_args: 1,
            max_args: Some(1),
            implementation: engineering::fn_imcos,
            volatile: false,
        });
        self.register(FunctionDef {
            name: "IMCOSH",
            min_args: 1,
            max_args: Some(1),
            implementation: engineering::fn_imcosh,
            volatile: false,
        });
        self.register(FunctionDef {
            name: "IMCOT",
            min_args: 1,
            max_args: Some(1),
            implementation: engineering::fn_imcot,
            volatile: false,
        });
        self.register(FunctionDef {
            name: "IMCSC",
            min_args: 1,
            max_args: Some(1),
            implementation: engineering::fn_imcsc,
            volatile: false,
        });
        self.register(FunctionDef {
            name: "IMCSCH",
            min_args: 1,
            max_args: Some(1),
            implementation: engineering::fn_imcsch,
            volatile: false,
        });
        self.register(FunctionDef {
            name: "IMDIV",
            min_args: 2,
            max_args: Some(2),
            implementation: engineering::fn_imdiv,
            volatile: false,
        });
        self.register(FunctionDef {
            name: "IMEXP",
            min_args: 1,
            max_args: Some(1),
            implementation: engineering::fn_imexp,
            volatile: false,
        });
        self.register(FunctionDef {
            name: "IMLN",
            min_args: 1,
            max_args: Some(1),
            implementation: engineering::fn_imln,
            volatile: false,
        });
        self.register(FunctionDef {
            name: "IMLOG10",
            min_args: 1,
            max_args: Some(1),
            implementation: engineering::fn_imlog10,
            volatile: false,
        });
        self.register(FunctionDef {
            name: "IMLOG2",
            min_args: 1,
            max_args: Some(1),
            implementation: engineering::fn_imlog2,
            volatile: false,
        });
        self.register(FunctionDef {
            name: "IMPOWER",
            min_args: 2,
            max_args: Some(2),
            implementation: engineering::fn_impower,
            volatile: false,
        });
        self.register(FunctionDef {
            name: "IMPRODUCT",
            min_args: 2,
            max_args: None,
            implementation: engineering::fn_improduct,
            volatile: false,
        });
        self.register(FunctionDef {
            name: "IMREAL",
            min_args: 1,
            max_args: Some(1),
            implementation: engineering::fn_imreal,
            volatile: false,
        });
        self.register(FunctionDef {
            name: "IMSEC",
            min_args: 1,
            max_args: Some(1),
            implementation: engineering::fn_imsec,
            volatile: false,
        });
        self.register(FunctionDef {
            name: "IMSECH",
            min_args: 1,
            max_args: Some(1),
            implementation: engineering::fn_imsech,
            volatile: false,
        });
        self.register(FunctionDef {
            name: "IMSIN",
            min_args: 1,
            max_args: Some(1),
            implementation: engineering::fn_imsin,
            volatile: false,
        });
        self.register(FunctionDef {
            name: "IMSINH",
            min_args: 1,
            max_args: Some(1),
            implementation: engineering::fn_imsinh,
            volatile: false,
        });
        self.register(FunctionDef {
            name: "IMSQRT",
            min_args: 1,
            max_args: Some(1),
            implementation: engineering::fn_imsqrt,
            volatile: false,
        });
        self.register(FunctionDef {
            name: "IMSUB",
            min_args: 2,
            max_args: Some(2),
            implementation: engineering::fn_imsub,
            volatile: false,
        });
        self.register(FunctionDef {
            name: "IMSUM",
            min_args: 2,
            max_args: None,
            implementation: engineering::fn_imsum,
            volatile: false,
        });
        self.register(FunctionDef {
            name: "IMTAN",
            min_args: 1,
            max_args: Some(1),
            implementation: engineering::fn_imtan,
            volatile: false,
        });
        self.register(FunctionDef {
            name: "BESSELI",
            min_args: 2,
            max_args: Some(2),
            implementation: engineering::fn_besseli,
            volatile: false,
        });
        self.register(FunctionDef {
            name: "BESSELJ",
            min_args: 2,
            max_args: Some(2),
            implementation: engineering::fn_besselj,
            volatile: false,
        });
        self.register(FunctionDef {
            name: "BESSELK",
            min_args: 2,
            max_args: Some(2),
            implementation: engineering::fn_besselk,
            volatile: false,
        });
        self.register(FunctionDef {
            name: "BESSELY",
            min_args: 2,
            max_args: Some(2),
            implementation: engineering::fn_bessely,
            volatile: false,
        });
        self.register(FunctionDef {
            name: "CONVERT",
            min_args: 3,
            max_args: Some(3),
            implementation: engineering::fn_convert,
            volatile: false,
        });
    }

    fn register_database_functions(&mut self) {
        self.register(FunctionDef {
            name: "DAVERAGE",
            min_args: 3,
            max_args: Some(3),
            implementation: database::fn_daverage,
            volatile: false,
        });
        self.register(FunctionDef {
            name: "DCOUNT",
            min_args: 3,
            max_args: Some(3),
            implementation: database::fn_dcount,
            volatile: false,
        });
        self.register(FunctionDef {
            name: "DCOUNTA",
            min_args: 3,
            max_args: Some(3),
            implementation: database::fn_dcounta,
            volatile: false,
        });
        self.register(FunctionDef {
            name: "DGET",
            min_args: 3,
            max_args: Some(3),
            implementation: database::fn_dget,
            volatile: false,
        });
        self.register(FunctionDef {
            name: "DMAX",
            min_args: 3,
            max_args: Some(3),
            implementation: database::fn_dmax,
            volatile: false,
        });
        self.register(FunctionDef {
            name: "DMIN",
            min_args: 3,
            max_args: Some(3),
            implementation: database::fn_dmin,
            volatile: false,
        });
        self.register(FunctionDef {
            name: "DPRODUCT",
            min_args: 3,
            max_args: Some(3),
            implementation: database::fn_dproduct,
            volatile: false,
        });
        self.register(FunctionDef {
            name: "DSTDEV",
            min_args: 3,
            max_args: Some(3),
            implementation: database::fn_dstdev,
            volatile: false,
        });
        self.register(FunctionDef {
            name: "DSTDEVP",
            min_args: 3,
            max_args: Some(3),
            implementation: database::fn_dstdevp,
            volatile: false,
        });
        self.register(FunctionDef {
            name: "DSUM",
            min_args: 3,
            max_args: Some(3),
            implementation: database::fn_dsum,
            volatile: false,
        });
        self.register(FunctionDef {
            name: "DVAR",
            min_args: 3,
            max_args: Some(3),
            implementation: database::fn_dvar,
            volatile: false,
        });
        self.register(FunctionDef {
            name: "DVARP",
            min_args: 3,
            max_args: Some(3),
            implementation: database::fn_dvarp,
            volatile: false,
        });
    }

    fn register_compatibility_functions(&mut self) {
        self.register(FunctionDef {
            name: "BETADIST",
            min_args: 3,
            max_args: Some(5),
            implementation: compatibility::fn_betadist,
            volatile: false,
        });
        self.register(FunctionDef {
            name: "BETAINV",
            min_args: 3,
            max_args: Some(5),
            implementation: compatibility::fn_betainv,
            volatile: false,
        });
        self.register(FunctionDef {
            name: "BINOMDIST",
            min_args: 4,
            max_args: Some(4),
            implementation: compatibility::fn_binomdist,
            volatile: false,
        });
        self.register(FunctionDef {
            name: "CEILING",
            min_args: 2,
            max_args: Some(2),
            implementation: compatibility::fn_ceiling,
            volatile: false,
        });
        self.register(FunctionDef {
            name: "CHIDIST",
            min_args: 2,
            max_args: Some(2),
            implementation: compatibility::fn_chidist,
            volatile: false,
        });
        self.register(FunctionDef {
            name: "CHIINV",
            min_args: 2,
            max_args: Some(2),
            implementation: compatibility::fn_chiinv,
            volatile: false,
        });
        self.register(FunctionDef {
            name: "CHITEST",
            min_args: 2,
            max_args: Some(2),
            implementation: compatibility::fn_chitest,
            volatile: false,
        });
        self.register(FunctionDef {
            name: "CONFIDENCE",
            min_args: 3,
            max_args: Some(3),
            implementation: compatibility::fn_confidence,
            volatile: false,
        });
        self.register(FunctionDef {
            name: "COVAR",
            min_args: 2,
            max_args: Some(2),
            implementation: compatibility::fn_covar,
            volatile: false,
        });
        self.register(FunctionDef {
            name: "CRITBINOM",
            min_args: 3,
            max_args: Some(3),
            implementation: compatibility::fn_critbinom,
            volatile: false,
        });
        self.register(FunctionDef {
            name: "EXPONDIST",
            min_args: 3,
            max_args: Some(3),
            implementation: compatibility::fn_expondist,
            volatile: false,
        });
        self.register(FunctionDef {
            name: "FDIST",
            min_args: 3,
            max_args: Some(3),
            implementation: compatibility::fn_fdist,
            volatile: false,
        });
        self.register(FunctionDef {
            name: "FINV",
            min_args: 3,
            max_args: Some(3),
            implementation: compatibility::fn_finv,
            volatile: false,
        });
        self.register(FunctionDef {
            name: "FLOOR",
            min_args: 2,
            max_args: Some(2),
            implementation: compatibility::fn_floor,
            volatile: false,
        });
        self.register(FunctionDef {
            name: "FTEST",
            min_args: 2,
            max_args: Some(2),
            implementation: compatibility::fn_ftest,
            volatile: false,
        });
        self.register(FunctionDef {
            name: "GAMMADIST",
            min_args: 4,
            max_args: Some(4),
            implementation: compatibility::fn_gammadist,
            volatile: false,
        });
        self.register(FunctionDef {
            name: "GAMMAINV",
            min_args: 3,
            max_args: Some(3),
            implementation: compatibility::fn_gammainv,
            volatile: false,
        });
        self.register(FunctionDef {
            name: "HYPGEOMDIST",
            min_args: 4,
            max_args: Some(4),
            implementation: compatibility::fn_hypgeomdist,
            volatile: false,
        });
        self.register(FunctionDef {
            name: "LOGINV",
            min_args: 3,
            max_args: Some(3),
            implementation: compatibility::fn_loginv,
            volatile: false,
        });
        self.register(FunctionDef {
            name: "LOGNORMDIST",
            min_args: 3,
            max_args: Some(3),
            implementation: compatibility::fn_lognormdist,
            volatile: false,
        });
        self.register(FunctionDef {
            name: "NEGBINOMDIST",
            min_args: 3,
            max_args: Some(3),
            implementation: compatibility::fn_negbinomdist,
            volatile: false,
        });
        self.register(FunctionDef {
            name: "NORMDIST",
            min_args: 4,
            max_args: Some(4),
            implementation: compatibility::fn_normdist,
            volatile: false,
        });
        self.register(FunctionDef {
            name: "NORMSDIST",
            min_args: 1,
            max_args: Some(1),
            implementation: compatibility::fn_normsdist,
            volatile: false,
        });
        self.register(FunctionDef {
            name: "NORMSINV",
            min_args: 1,
            max_args: Some(1),
            implementation: compatibility::fn_normsinv,
            volatile: false,
        });
        self.register(FunctionDef {
            name: "NORM.INV",
            min_args: 3,
            max_args: Some(3),
            implementation: compatibility::fn_norm_inv,
            volatile: false,
        });
        self.register(FunctionDef {
            name: "POISSON",
            min_args: 3,
            max_args: Some(3),
            implementation: compatibility::fn_poisson,
            volatile: false,
        });
        self.register(FunctionDef {
            name: "TDIST",
            min_args: 3,
            max_args: Some(3),
            implementation: compatibility::fn_tdist,
            volatile: false,
        });
        self.register(FunctionDef {
            name: "TINV",
            min_args: 2,
            max_args: Some(2),
            implementation: compatibility::fn_tinv,
            volatile: false,
        });
        self.register(FunctionDef {
            name: "TTEST",
            min_args: 4,
            max_args: Some(4),
            implementation: compatibility::fn_ttest,
            volatile: false,
        });
        self.register(FunctionDef {
            name: "WEIBULL",
            min_args: 4,
            max_args: Some(4),
            implementation: compatibility::fn_weibull,
            volatile: false,
        });
        self.register(FunctionDef {
            name: "ZTEST",
            min_args: 2,
            max_args: Some(3),
            implementation: compatibility::fn_ztest,
            volatile: false,
        });
    }

    fn register_web_functions(&mut self) {
        self.register(FunctionDef {
            name: "ENCODEURL",
            min_args: 1,
            max_args: Some(1),
            implementation: lookup_extra::fn_encodeurl,
            volatile: false,
        });
        self.register(FunctionDef {
            name: "FILTERXML",
            min_args: 2,
            max_args: Some(2),
            implementation: lookup_extra::fn_filterxml,
            volatile: false,
        });
        self.register(FunctionDef {
            name: "WEBSERVICE",
            min_args: 1,
            max_args: Some(1),
            implementation: lookup_extra::fn_webservice,
            volatile: false,
        });
    }
}
impl Default for FunctionRegistry {
    fn default() -> Self {
        Self::new()
    }
}
