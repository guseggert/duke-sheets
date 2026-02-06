//! Built-in Excel functions

pub mod criteria;
pub mod date;
pub mod info;
pub mod logical;
pub mod lookup;
pub mod math;
pub mod statistical;
pub mod text;

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
    }
}

impl Default for FunctionRegistry {
    fn default() -> Self {
        Self::new()
    }
}
