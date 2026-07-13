#!/usr/bin/env python3
"""Validate format-specific evidence in FEATURES.md."""

from __future__ import annotations

import re
import subprocess
import sys
from collections import Counter, defaultdict
from dataclasses import dataclass, field
from pathlib import Path


ROOT = Path(__file__).resolve().parent.parent
FEATURES = ROOT / "FEATURES.md"
FORMATS = ("XLSX", "XLSB", "XLS")
SUPPORTED = {"\u2714", "\u25cf"}
STATUS_RE = re.compile(r"([RW])([\u2714\u25cf\u2716-])")
REF_RE = re.compile(r"`([^`]+)`")
FN_RE = re.compile(
    r"^\s*(?:pub(?:\([^)]*\))?\s+)?(?:const\s+)?(?:async\s+)?"
    r"(?:unsafe\s+)?fn\s+([A-Za-z_][A-Za-z0-9_]*)"
)
MOD_RE = re.compile(r"^(\s*)(?:pub(?:\([^)]*\))?\s+)?mod\s+([A-Za-z_][A-Za-z0-9_]*)\s*\{")

CANONICAL_PARITY = {
    "crates/duke-sheets-excel-com/tests/e2e/writing.rs": "XLSX",
    "crates/duke-sheets-excel-com/tests/e2e/writing_xlsb.rs": "XLSB",
    "crates/duke-sheets-excel-com/tests/e2e/writing_xls.rs": "XLS",
}

HARNESS_ALIASES = {
    "xlsx-e2e": ("duke-sheets-xlsx", "tests/e2e"),
    "xls-e2e": ("duke-sheets-xls", "tests/e2e"),
    "excel-com-e2e": ("duke-sheets-excel-com", "tests/e2e"),
    "com-e2e": ("duke-sheets-excel-com", "tests/e2e"),
}

CRATE_ALIASES = {
    "duke-sheets": "duke-sheets",
    "duke-sheets-chart": "duke-sheets-chart",
    "duke-sheets-core": "duke-sheets-core",
    "duke-sheets-crypto": "duke-sheets-crypto",
    "duke-sheets-excel-com": "duke-sheets-excel-com",
    "duke-sheets-formula": "duke-sheets-formula",
    "duke-sheets-xls": "duke-sheets-xls",
    "duke-sheets-xlsb": "duke-sheets-xlsb",
    "duke-sheets-xlsx": "duke-sheets-xlsx",
    "xlsb": "duke-sheets-xlsb",
}


def normalize(value: str) -> str:
    return value.strip().replace("_", "-").lower()


@dataclass
class FeatureRow:
    section: str
    line: int
    feature: str
    cells: dict[str, str]
    refs: list[str]
    notes: str

    @property
    def key(self) -> tuple[str, str]:
        return self.section, self.feature


@dataclass
class TestFunction:
    crate: str
    path: str
    name: str
    line: int
    modules: tuple[str, ...]
    body: str
    formats: dict[str, set[str]] = field(default_factory=dict)
    parity_format: str | None = None

    @property
    def in_process(self) -> bool:
        return not self.path.startswith("crates/duke-sheets-excel-com/")

    @property
    def aliases(self) -> set[str]:
        parts = Path(self.path).parts
        aliases = {normalize(self.crate), normalize(Path(self.path).stem)}
        aliases.update(normalize(part) for part in parts)
        aliases.update(normalize(module) for module in self.modules)
        if self.crate == "duke-sheets-xlsb":
            aliases.add("xlsb")
        return aliases


def split_markdown_row(line: str) -> list[str]:
    return [cell.strip() for cell in line.strip().strip("|").split("|")]


def parse_features(text: str) -> tuple[list[FeatureRow], list[str]]:
    lines = text.splitlines()
    rows: list[FeatureRow] = []
    structure_errors: list[str] = []
    section = ""
    index = 0
    while index < len(lines):
        line = lines[index]
        if line.startswith("## "):
            section = line[3:].strip()
            index += 1
            continue
        if not line.startswith("|") or index + 1 >= len(lines):
            index += 1
            continue
        headers = split_markdown_row(line)
        separator = split_markdown_row(lines[index + 1]) if lines[index + 1].startswith("|") else []
        if not separator or not all(re.fullmatch(r":?-{3,}:?", cell) for cell in separator):
            index += 1
            continue
        table_start = index
        index += 2
        if not any(fmt in headers for fmt in FORMATS):
            while index < len(lines) and lines[index].startswith("|"):
                index += 1
            continue
        test_index = next(
            (i for i, header in enumerate(headers) if header.lower() in {"test", "tests"}),
            None,
        )
        notes_index = next(
            (i for i, header in enumerate(headers) if header.lower() == "notes"),
            None,
        )
        while index < len(lines) and lines[index].startswith("|"):
            cells = split_markdown_row(lines[index])
            feature = cells[0] if cells else f"table at line {table_start + 1}"
            # Convention: exactly one unheaded trailing cell carries
            # Notes. Fewer cells silently drop claims; more than one
            # extra means a mangled row.
            if len(cells) < len(headers):
                structure_errors.append(
                    f"line {index + 1} [{feature}]: row has {len(cells)} cells "
                    f"but the table header has {len(headers)} columns"
                )
            elif len(cells) > len(headers) + 1:
                structure_errors.append(
                    f"line {index + 1} [{feature}]: row has {len(cells) - len(headers)} "
                    "unheaded trailing cells (at most one Notes overflow is allowed)"
                )
            named = {header: cells[i] if i < len(cells) else "" for i, header in enumerate(headers)}
            extras = cells[len(headers) :]
            notes = named.get(headers[notes_index], "") if notes_index is not None else ""
            if extras:
                notes = " | ".join(part for part in [notes, *extras] if part)
            tests = cells[test_index] if test_index is not None and test_index < len(cells) else ""
            for fmt in FORMATS:
                cell = named.get(fmt, "")
                statuses = STATUS_RE.findall(cell)
                directions = [direction for direction, _ in statuses]
                if len(directions) != len(set(directions)):
                    structure_errors.append(
                        f"line {index + 1} [{feature}] {fmt}: duplicate direction in cell {cell!r}"
                    )
            rows.append(
                FeatureRow(
                    section=section,
                    line=index + 1,
                    feature=feature,
                    cells={fmt: named.get(fmt, "") for fmt in FORMATS if fmt in headers},
                    refs=[match.group(1).strip() for match in REF_RE.finditer(tests)],
                    notes=notes.strip(),
                )
            )
            index += 1
    return rows, structure_errors


def function_body(lines: list[str], start: int) -> str:
    output: list[str] = []
    depth = 0
    started = False
    for line in lines[start:]:
        output.append(line)
        depth += line.count("{") - line.count("}")
        started = started or "{" in line
        if started and depth <= 0:
            break
    return "".join(output)


def is_test_attribute(attributes: list[str]) -> bool:
    joined = " ".join(attributes)
    return bool(
        re.search(r"#\[(?:[A-Za-z0-9_]+::)?test(?:\([^]]*\))?\]", joined)
        or re.search(r"#\[(?:rstest|proptest)(?:\([^]]*\))?\]", joined)
    )


def inventory_tests() -> list[TestFunction]:
    tests: list[TestFunction] = []
    crates = ROOT / "crates"
    for path in sorted(crates.rglob("*.rs")):
        if any(part in {"target", "node_modules", ".git"} for part in path.parts):
            continue
        relative = path.relative_to(ROOT).as_posix()
        crate = path.relative_to(crates).parts[0]
        lines = path.read_text(encoding="utf-8", errors="replace").splitlines(keepends=True)
        pending_attributes: list[str] = []
        modules: list[tuple[str, int]] = []
        depth = 0
        for line_number, line in enumerate(lines, start=1):
            stripped = line.strip()
            module_match = MOD_RE.match(line)
            if module_match:
                modules.append((module_match.group(2), depth))
            if stripped.startswith("#["):
                pending_attributes.append(stripped)
            else:
                function_match = FN_RE.match(line)
                if function_match and is_test_attribute(pending_attributes):
                    tests.append(
                        TestFunction(
                            crate=crate,
                            path=relative,
                            name=function_match.group(1),
                            line=line_number,
                            modules=tuple(module for module, _ in modules),
                            body=function_body(lines, line_number - 1),
                        )
                    )
                if stripped and not stripped.startswith("//"):
                    pending_attributes = []
            depth += line.count("{") - line.count("}")
            while modules and depth <= modules[-1][1]:
                modules.pop()
    for test in tests:
        classify_test(test)
    return tests


def body_formats(test: TestFunction) -> set[str]:
    text = f"{test.name}\n{test.body}"
    matches: set[str] = set()
    patterns = {
        "XLSX": r"Xlsx(?:Reader|Writer)|FileFormat::Xlsx|\.xlsx\b|[\"']xlsx[\"']|(?:^|_)xlsx(?:_|$)",
        "XLSB": r"Xlsb(?:Reader|Writer)|FileFormat::Xlsb|\.xlsb\b|[\"']xlsb[\"']|(?:^|_)xlsb(?:_|$)",
        "XLS": r"Xls(?:Reader|Writer)|FileFormat::Xls\b|\.xls\b|[\"']xls[\"']|(?:^|_)xls(?:_|$)",
    }
    for fmt, pattern in patterns.items():
        if re.search(pattern, text, re.IGNORECASE | re.MULTILINE):
            matches.add(fmt)
    return matches


def body_directions(test: TestFunction, fmt: str) -> set[str]:
    body = test.body
    prefix = {"XLSX": "Xlsx", "XLSB": "Xlsb", "XLS": "Xls"}[fmt]
    directions: set[str] = set()
    if re.search(rf"{prefix}Writer|write_{fmt.lower()}|FileFormat::{prefix}", body):
        directions.add("W")
    if re.search(rf"{prefix}Reader|read_{fmt.lower()}", body):
        directions.add("R")
    if re.search(r"\.(?:save|save_with)\s*\(", body):
        directions.add("W")
    if re.search(r"\bWorkbook::(?:open|open_with|from_bytes)\s*\(", body):
        directions.add("R")
    # Round-trip helper calls (round_trip, round_trip_xlsx, ...) are
    # body evidence; test names alone grant nothing.
    if re.search(rf"\bround_?trip(?:_{fmt.lower()})?\s*\(", body):
        directions.update(("R", "W"))
    if re.search(r"\b(?:write_xlsb_bytes|sheet1_records|first_dval_header_and_formula1)\s*\(", body):
        directions.add("W")
    return directions


def set_formats(test: TestFunction, formats: set[str], directions: set[str]) -> None:
    for fmt in formats:
        test.formats.setdefault(fmt, set()).update(directions)


def classify_test(test: TestFunction) -> None:
    path = test.path
    if path in CANONICAL_PARITY:
        fmt = CANONICAL_PARITY[path]
        writer = {
            "XLSX": r"roundtrip_through_excel\s*\(|XlsxWriter",
            "XLSB": r"roundtrip_through_excel_xlsb(?:_bytes)?\s*\(|XlsbWriter",
            "XLS": r"roundtrip_through_excel_xls(?:_bytes)?\s*\(|XlsWriter",
        }[fmt]
        if re.search(writer, test.body):
            test.parity_format = fmt
            directions = {"W"}
            # The roundtrip helpers re-read Excel's output with our
            # reader; a bare Excel writer call alone is write-only.
            if re.search(
                r"roundtrip_through_excel(?:_xlsb?|_xls)?(?:_bytes)?\s*\(",
                test.body,
            ) or re.search(
                {"XLSX": r"XlsxReader", "XLSB": r"XlsbReader", "XLS": r"XlsReader"}[fmt],
                test.body,
            ):
                directions.add("R")
            set_formats(test, {fmt}, directions)
        else:
            directions = body_directions(test, fmt)
            if directions:
                set_formats(test, {fmt}, directions)
        return

    if path.startswith("crates/duke-sheets-xlsx/"):
        if "/tests/e2e/reading/" in path:
            set_formats(test, {"XLSX"}, {"R"})
        elif path.endswith("/tests/e2e/writing.rs"):
            set_formats(test, {"XLSX"}, {"W"})
        elif "/src/reader/" in path or path.endswith("/src/reader.rs"):
            set_formats(test, {"XLSX"}, {"R"})
        elif "/src/writer/" in path or path.endswith("/src/writer.rs"):
            set_formats(test, {"XLSX"}, {"W"})
        elif "/tests/" in path:
            directions = body_directions(test, "XLSX")
            if "compat" in path and not directions:
                directions = {"R"}
            set_formats(test, {"XLSX"}, directions or {"R", "W"})
        return

    if path.startswith("crates/duke-sheets-xlsb/"):
        if "/src/reader/" in path or path.endswith("/src/reader.rs"):
            set_formats(test, {"XLSB"}, {"R"})
        elif "/src/writer/" in path or path.endswith("/src/writer.rs"):
            set_formats(test, {"XLSB"}, body_directions(test, "XLSB") or {"W"})
        elif "/biff12/compiler" in path:
            set_formats(test, {"XLSB"}, {"W"})
        elif "/biff12/" in path:
            set_formats(test, {"XLSB"}, {"R"})
        else:
            set_formats(test, {"XLSB"}, body_directions(test, "XLSB") or {"R", "W"})
        return

    if path.startswith("crates/duke-sheets-xls/"):
        if "/tests/e2e/reading/" in path:
            set_formats(test, {"XLS"}, {"R"})
        elif "/src/reader/" in path or path.endswith("/src/reader.rs"):
            set_formats(test, {"XLS"}, {"R"})
        elif "/src/writer/" in path or path.endswith("/src/writer.rs"):
            set_formats(test, {"XLS"}, {"W"})
        elif "/tests/" in path:
            directions = body_directions(test, "XLS")
            if "cryptoapi" in path and "write" not in path and not directions:
                directions = {"R"}
            set_formats(test, {"XLS"}, directions or {"R", "W"})
        return

    if path.startswith("crates/duke-sheets-excel-com/tests/e2e/"):
        formats = body_formats(test)
        for fmt in formats:
            set_formats(test, {fmt}, body_directions(test, fmt) or {"R"})
        return

    static_files = {
        "xlsx_roundtrip.rs": ("XLSX", {"R", "W"}),
        "xlsx_formatting_roundtrip.rs": ("XLSX", {"R", "W"}),
        "xlsx_style_roundtrip.rs": ("XLSX", {"R", "W"}),
        "xlsx_drawing_order.rs": ("XLSX", {"R", "W"}),
        "chart_parity.rs": ("XLSX", {"R", "W"}),
        "chart_corpus.rs": ("XLSX", {"R"}),
        "formula_parity.rs": ("XLSX", {"R"}),
        "xls_drawing_order.rs": ("XLS", {"R", "W"}),
        "xls_save_with.rs": ("XLS", {"R", "W"}),
        "xlsb_drawing_order.rs": ("XLSB", {"R", "W"}),
    }
    if path.startswith("crates/duke-sheets/tests/") and Path(path).name in static_files:
        fmt, directions = static_files[Path(path).name]
        set_formats(test, {fmt}, directions)
        return

    formats = body_formats(test)
    for fmt in formats:
        directions = body_directions(test, fmt)
        if directions:
            set_formats(test, {fmt}, directions)


def prefix_matches(test: TestFunction, prefix: list[str]) -> bool:
    if not prefix:
        return True
    normalized = [normalize(part) for part in prefix]
    head = normalized[0]
    remaining = normalized[1:]
    harness = HARNESS_ALIASES.get(head)
    if harness:
        crate, path_fragment = harness
        if test.crate != crate or path_fragment not in test.path:
            return False
    elif head in CRATE_ALIASES:
        if test.crate != CRATE_ALIASES[head]:
            return False
    else:
        remaining = normalized
    aliases = test.aliases
    return all(part in aliases for part in remaining if part not in {"mod"})


def resolve_ref(ref: str, by_name: dict[str, list[TestFunction]]) -> list[TestFunction]:
    if "*" in ref or any(character.isspace() for character in ref):
        return []
    parts = ref.split("::")
    if not parts or not parts[-1]:
        return []
    return [test for test in by_name.get(parts[-1], []) if prefix_matches(test, parts[:-1])]


def parse_status(cell: str) -> dict[str, str]:
    return dict(STATUS_RE.findall(cell))


def parity_impossible(notes: str) -> bool:
    return "parity impossible" in notes.lower()


def validate(rows: list[FeatureRow], tests: list[TestFunction]) -> tuple[list[str], Counter[str]]:
    by_name: dict[str, list[TestFunction]] = defaultdict(list)
    for test in tests:
        by_name[test.name].append(test)
    errors: list[str] = []
    totals: Counter[str] = Counter()
    for row in rows:
        resolved: dict[str, list[TestFunction]] = {}
        for ref in row.refs:
            matches = resolve_ref(ref, by_name)
            resolved[ref] = matches
            if "*" in ref:
                errors.append(
                    f"line {row.line} [{row.feature}] Test: wildcard reference `{ref}` is not allowed"
                )
            elif not matches:
                errors.append(
                    f"line {row.line} [{row.feature}] Test: exact test function `{ref}` does not exist at that path"
                )
            elif len({(test.path, test.line) for test in matches}) > 1:
                locations = ", ".join(
                    f"{test.path}:{test.line}" for test in sorted(matches, key=lambda item: item.path)
                )
                errors.append(
                    f"line {row.line} [{row.feature}] Test: `{ref}` is ambiguous ({locations})"
                )
        linked_tests = [test for matches in resolved.values() for test in matches]
        for fmt, cell in row.cells.items():
            statuses = parse_status(cell)
            if not statuses:
                continue
            totals[f"{fmt}:cells"] += 1
            for direction, status in statuses.items():
                totals[f"{fmt}:{direction}{status}"] += 1
                if status not in SUPPORTED:
                    continue
                totals[f"claims"] += 1
                matching = [
                    test
                    for test in linked_tests
                    if (direction != "W" or (test.parity_format is None and test.in_process))
                    and direction in test.formats.get(fmt, set())
                ]
                if not matching:
                    evidence = (
                        f"in-process {fmt} writer"
                        if direction == "W"
                        else f"format-specific {fmt} reader"
                    )
                    errors.append(
                        f"line {row.line} [{row.feature}] {fmt} {direction}{status}: "
                        f"no linked {evidence} test"
                    )
                if status == "\u25cf" and (not row.notes or row.notes == "-"):
                    errors.append(
                        f"line {row.line} [{row.feature}] {fmt} {direction}\u25cf: "
                        "Notes must explain the limitation"
                    )
            if statuses.get("W") == "\u2714":
                parity = [test for test in linked_tests if test.parity_format == fmt]
                if not parity and not parity_impossible(row.notes):
                    expected = {
                        "XLSX": "writing.rs",
                        "XLSB": "writing_xlsb.rs",
                        "XLS": "writing_xls.rs",
                    }[fmt]
                    errors.append(
                        f"line {row.line} [{row.feature}] {fmt} W\u2714: "
                        f"no canonical Excel parity test from `{expected}`"
                    )
                totals[f"{fmt}:parity"] += bool(parity)
    totals["rows"] = len(rows)
    totals["refs"] = sum(len(row.refs) for row in rows)
    totals["tests"] = len(tests)
    return errors, totals


def status_counts(rows: list[FeatureRow]) -> dict[str, Counter[str]]:
    counts = {fmt: Counter() for fmt in FORMATS}
    for row in rows:
        for fmt, cell in row.cells.items():
            for direction, status in parse_status(cell).items():
                counts[fmt][f"{direction}{status}"] += 1
    return counts


def git_head_features() -> str | None:
    try:
        result = subprocess.run(
            ["git", "show", "HEAD:FEATURES.md"],
            cwd=ROOT,
            check=True,
            capture_output=True,
            text=True,
        )
    except (OSError, subprocess.CalledProcessError):
        return None
    return result.stdout


def change_summary(
    current_rows: list[FeatureRow], baseline_text: str | None, tests: list[TestFunction]
) -> list[str]:
    if baseline_text is None:
        return []
    baseline_rows, _ = parse_features(baseline_text)
    baseline_by_key = {row.key: row for row in baseline_rows}
    current_by_key = {row.key: row for row in current_rows}
    before = status_counts(baseline_rows)
    after = status_counts(current_rows)
    output = ["Changes from HEAD:"]
    for fmt in FORMATS:
        labels = ("R\u2714", "R\u25cf", "R\u2716", "W\u2714", "W\u25cf", "W\u2716")
        old = " ".join(f"{label}={before[fmt][label]}" for label in labels)
        new = " ".join(f"{label}={after[fmt][label]}" for label in labels)
        output.append(f"  {fmt}: {old} -> {new}")

    links_added = 0
    links_removed = 0
    downgraded = Counter()
    rank = {"\u2714": 2, "\u25cf": 1, "\u2716": 0, "-": -1}
    for key, current in current_by_key.items():
        baseline = baseline_by_key.get(key)
        if baseline is None:
            links_added += len(set(current.refs))
            continue
        links_added += len(set(current.refs) - set(baseline.refs))
        links_removed += len(set(baseline.refs) - set(current.refs))
        for fmt in FORMATS:
            old_status = parse_status(baseline.cells.get(fmt, ""))
            new_status = parse_status(current.cells.get(fmt, ""))
            for direction in ("R", "W"):
                old = old_status.get(direction)
                new = new_status.get(direction)
                if old in rank and new in rank and rank[new] < rank[old]:
                    downgraded[direction] += 1
                    downgraded[f"{fmt}:{direction}"] += 1

    by_name: dict[str, list[TestFunction]] = defaultdict(list)
    for test in tests:
        by_name[test.name].append(test)
    invalid_before: set[tuple[tuple[str, str], str]] = set()
    invalid_after: set[tuple[tuple[str, str], str]] = set()
    for row, target in ((row, invalid_before) for row in baseline_rows):
        for ref in row.refs:
            if "*" in ref or not resolve_ref(ref, by_name):
                target.add((row.key, ref))
    for row, target in ((row, invalid_after) for row in current_rows):
        for ref in row.refs:
            if "*" in ref or not resolve_ref(ref, by_name):
                target.add((row.key, ref))
    output.append(
        f"  links added={links_added}, removed={links_removed}, net={links_added - links_removed:+d}; "
        f"non-exact/wildcard links fixed={len(invalid_before - invalid_after)}"
    )
    output.append(
        f"  downgraded R={downgraded['R']} "
        f"(XLSX={downgraded['XLSX:R']}, XLSB={downgraded['XLSB:R']}, XLS={downgraded['XLS:R']}), "
        f"W={downgraded['W']} "
        f"(XLSX={downgraded['XLSX:W']}, XLSB={downgraded['XLSB:W']}, XLS={downgraded['XLS:W']})"
    )
    return output


def print_totals(totals: Counter[str]) -> None:
    print(
        f"Checked {totals['rows']} format rows, {totals['claims']} supported directions, "
        f"{totals['refs']} links, and {totals['tests']} Rust tests."
    )
    for fmt in FORMATS:
        print(
            f"  {fmt}: R\u2714={totals[f'{fmt}:R\u2714']} R\u25cf={totals[f'{fmt}:R\u25cf']} "
            f"W\u2714={totals[f'{fmt}:W\u2714']} W\u25cf={totals[f'{fmt}:W\u25cf']} "
            f"canonical parity={totals[f'{fmt}:parity']}"
        )


def main() -> int:
    if not FEATURES.is_file():
        print(f"ERROR: {FEATURES} does not exist", file=sys.stderr)
        return 2
    rows, structure_errors = parse_features(FEATURES.read_text(encoding="utf-8"))
    tests = inventory_tests()
    errors, totals = validate(rows, tests)
    errors = structure_errors + errors
    print_totals(totals)
    for line in change_summary(rows, git_head_features(), tests):
        print(line)
    if errors:
        print(f"FAIL: {len(errors)} feature evidence error(s):", file=sys.stderr)
        for error in errors:
            print(f"  - {error}", file=sys.stderr)
        return 1
    print("PASS: FEATURES.md evidence is format-specific and complete.")
    return 0


if __name__ == "__main__":
    sys.exit(main())
