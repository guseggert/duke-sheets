//! Built-in Excel functions

pub mod criteria;
pub mod database;
pub mod date;
pub mod engineering;
pub mod financial;
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
        registry.register_financial_functions();
        registry.register_database_functions();
        registry.register_engineering_functions();

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
}

impl Default for FunctionRegistry {
    fn default() -> Self {
        Self::new()
    }
}
