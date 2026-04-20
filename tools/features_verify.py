#!/usr/bin/env python3
from __future__ import annotations

import re
import subprocess
import sys
from pathlib import Path
from typing import Iterator

REPO = Path(__file__).resolve().parent.parent
FEATURES = REPO / "FEATURES.md"

CLAIM_CHARS = {"\u2714", "\u25cf"}
SKIP_CHARS = {"\u2716", "\u2014"}

CRATE_TEST_MAP: dict[str, tuple[str, str | None]] = {
    "xlsx_roundtrip": ("duke-sheets", "xlsx_roundtrip"),
    "xlsx_formatting_roundtrip": ("duke-sheets", "xlsx_formatting_roundtrip"),
    "xlsx_style_roundtrip": ("duke-sheets", "xlsx_style_roundtrip"),
    "formula_evaluation": ("duke-sheets", "formula_evaluation"),
    "formula_parity": ("duke-sheets", "formula_parity"),
    "chart_corpus": ("duke-sheets", "chart_corpus"),
    "chart_parity": ("duke-sheets", "chart_parity"),
    "xlsb_parity": ("duke-sheets", "xlsb_parity"),
    "read_corpus": ("duke-sheets", "read_corpus"),
    "calculation": ("duke-sheets", None),
    "xlsx_e2e": ("duke-sheets-xlsx", None),
    "xls_e2e": ("duke-sheets-xls", None),
    "com_e2e": ("duke-sheets-excel-com", None),
    "duke-sheets-formula": ("duke-sheets-formula", None),
    "duke-sheets-xlsx": ("duke-sheets-xlsx", None),
    "duke-sheets-xls": ("duke-sheets-xls", None),
    "duke-sheets-xlsb": ("duke-sheets-xlsb", None),
    "duke-sheets-core": ("duke-sheets-core", None),
    "duke-sheets-chart": ("duke-sheets-chart", None),
}

CORPUS_GATED = {
    "chart_corpus",
    "read_corpus",
    "formula_parity",
    "chart_parity",
    "xlsb_parity",
}
ENV_GATED = {"com_e2e"}


class Row:
    __slots__ = ("line", "feature", "claims", "tests", "notes")

    def __init__(
        self, line: int, feature: str, claims: bool, tests: str, notes: str
    ) -> None:
        self.line = line
        self.feature = feature
        self.claims = claims
        self.tests = tests
        self.notes = notes


def iter_rows(src: str) -> Iterator[Row]:
    in_table = False
    awaiting_header = False
    awaiting_sep = False
    test_idx: int | None = None
    has_test_col = False
    for lineno, raw in enumerate(src.splitlines(), start=1):
        line = raw.rstrip()
        if not line.startswith("|"):
            in_table = False
            awaiting_header = False
            awaiting_sep = False
            test_idx = None
            has_test_col = False
            continue
        cells = [c.strip() for c in line.strip("|").split("|")]
        if not in_table:
            in_table = True
            awaiting_sep = True
            headers = [c.strip().lower() for c in cells]
            test_idx = None
            for i, h in enumerate(headers):
                if h in ("test", "tests"):
                    test_idx = i
                    break
            has_test_col = test_idx is not None
            continue
        if awaiting_sep:
            awaiting_sep = False
            continue
        if len(cells) < 2:
            continue
        feature = cells[0]
        claims = any(ch in CLAIM_CHARS for cell in cells for ch in cell)
        tests = (
            cells[test_idx]
            if has_test_col and test_idx is not None and test_idx < len(cells)
            else ""
        )
        notes = cells[-1] if has_test_col and cells[-1] and cells[-1] != tests else ""
        if not has_test_col:
            yield Row(lineno, feature, False, "", "")
        else:
            yield Row(lineno, feature, claims, tests, notes)


def extract_test_refs(tests_cell: str) -> list[str]:
    if not tests_cell or tests_cell.strip() in {"-", "\u2014", ""}:
        return []
    refs: list[str] = []
    for m in re.finditer(r"`([^`]+)`", tests_cell):
        token = m.group(1).strip()
        if token and token != "\u2014":
            refs.append(token)
    return refs


def resolve_ref(ref: str) -> tuple[str, str, str | None] | None:
    ref = ref.split(" ", 1)[0]
    parts = ref.split("::")
    if not parts:
        return None
    head = parts[0]
    mapping = CRATE_TEST_MAP.get(head)
    if not mapping:
        return None
    crate, test_file = mapping
    if len(parts) == 1:
        return (crate, test_file or "", "")
    filter_parts = parts[1:]
    name_filter = filter_parts[-1].rstrip("*")
    return (crate, test_file or "", name_filter)


def plan_invocations(
    refs: list[tuple[str, str, str]],
) -> list[tuple[list[str], list[str]]]:
    by_key: dict[tuple[str, str], set[str]] = {}
    for crate, test_file, name in refs:
        by_key.setdefault((crate, test_file), set()).add(name)
    plans: list[tuple[list[str], list[str]]] = []
    for (crate, test_file), names in sorted(by_key.items()):
        cmd = ["cargo", "test", "-p", crate, "--no-fail-fast"]
        if test_file:
            cmd.extend(["--test", test_file])
        filters = sorted(n for n in names if n)
        plans.append((cmd, filters))
    return plans


def run_plan(cmd: list[str], filters: list[str]) -> int:
    full = cmd + ["--"] + filters if filters else cmd
    print("    $", " ".join(full))
    try:
        rc = subprocess.call(full, cwd=str(REPO))
    except FileNotFoundError as e:
        print(f"    ERROR: {e}", file=sys.stderr)
        return 127
    return rc


def main() -> int:
    list_only = "--list" in sys.argv
    if not FEATURES.exists():
        print(f"ERROR: {FEATURES} not found", file=sys.stderr)
        return 2
    src = FEATURES.read_text(encoding="utf-8")
    missing: list[Row] = []
    unresolved: list[tuple[Row, str]] = []
    env_skipped: list[tuple[Row, str]] = []
    corpus_skipped: list[tuple[Row, str]] = []
    collected: list[tuple[str, str, str]] = []
    total_claim_rows = 0
    for row in iter_rows(src):
        if not row.claims:
            continue
        total_claim_rows += 1
        refs = extract_test_refs(row.tests)
        if not refs:
            missing.append(row)
            continue
        for ref in refs:
            head = ref.split("::", 1)[0].split(" ", 1)[0]
            if head in ENV_GATED:
                env_skipped.append((row, ref))
                continue
            if head in CORPUS_GATED:
                corpus_skipped.append((row, ref))
                continue
            resolved = resolve_ref(ref)
            if resolved is None:
                unresolved.append((row, ref))
                continue
            collected.append(resolved)

    print(f"FEATURES.md claim rows (\u2714 or \u25cf): {total_claim_rows}")
    print(f"  test refs collected:    {len(collected)}")
    print(
        f"  env-gated skipped:      {len(env_skipped)} (com_e2e requires Windows+Excel)"
    )
    print(
        f"  corpus-gated skipped:   {len(corpus_skipped)} (require data/ corpus files)"
    )
    print(f"  unresolved test paths:  {len(unresolved)}")
    print(f"  rows missing a test:    {len(missing)}")

    if missing or unresolved:
        print()
        for row in missing[:20]:
            print(f"  MISSING TEST: line {row.line}: {row.feature}")
        for row, ref in unresolved[:20]:
            print(f"  UNRESOLVED:   line {row.line}: {row.feature} -> `{ref}`")
        for row in missing[20:]:
            print(f"  MISSING TEST: line {row.line}: {row.feature}")
        for row, ref in unresolved[20:]:
            print(f"  UNRESOLVED:   line {row.line}: {row.feature} -> `{ref}`")

    if not collected:
        print("\nNo resolvable claims to verify.")
        return 1 if (missing or unresolved) else 0

    plans = plan_invocations(collected)
    if list_only:
        print(f"\n{len(plans)} cargo-test invocation(s) would run:")
        for cmd, filters in plans:
            suffix = (" -- " + " ".join(filters)) if filters else ""
            print(f"  {' '.join(cmd)}{suffix}")
        return 1 if (missing or unresolved) else 0

    print(f"\nRunning {len(plans)} cargo-test invocation(s):")
    exit_code = 0
    for cmd, filters in plans:
        rc = run_plan(cmd, filters)
        if rc != 0:
            exit_code = rc
    if missing or unresolved:
        exit_code = max(exit_code, 1)
    return exit_code


if __name__ == "__main__":
    sys.exit(main())
