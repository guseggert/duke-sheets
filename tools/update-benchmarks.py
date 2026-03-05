#!/usr/bin/env python3
"""Read Criterion JSON results and update README benchmark section.

Usage:
    cargo bench --features full -p duke-sheets
    python3 tools/update-benchmarks.py

Reads structured data from target/criterion/*/new/{benchmark,estimates}.json
instead of parsing text output.
"""

from __future__ import annotations

import json
import re
import subprocess
import sys
from dataclasses import dataclass
from pathlib import Path

START_MARKER = "<!-- BENCHMARKS:START -->"
END_MARKER = "<!-- BENCHMARKS:END -->"

CRITERION_DIR = Path("target/criterion")

GROUP_ORDER = [
    "xlsx_read",
    "xlsx_write_serialize",
    "xlsx_write_full",
    "csv_read",
    "csv_write",
    "formula_parse",
    "calculation/linear_chain",
    "calculation/fan_out",
    "calculation/cross_sheet",
    "calculation/mixed",
]


@dataclass
class BenchResult:
    group: str
    case: str
    library: str
    time: str


def format_ns(ns: float) -> str:
    """Format nanoseconds into human-readable time."""
    if ns < 1_000:
        return f"{ns:.4g} ns"
    elif ns < 1_000_000:
        return f"{ns / 1_000:.4g} µs"
    elif ns < 1_000_000_000:
        return f"{ns / 1_000_000:.4g} ms"
    else:
        return f"{ns / 1_000_000_000:.4g} s"


def collect_results(criterion_dir: Path) -> list[BenchResult]:
    """Walk target/criterion/ and read benchmark.json + estimates.json pairs."""
    results: list[BenchResult] = []

    for estimates_path in sorted(criterion_dir.rglob("new/estimates.json")):
        bench_path = estimates_path.with_name("benchmark.json")
        if not bench_path.exists():
            continue

        bench = json.loads(bench_path.read_text())
        estimates = json.loads(estimates_path.read_text())

        group_id: str = bench["group_id"]
        function_id: str | None = bench.get("function_id")
        value_str: str | None = bench.get("value_str")
        median_ns: float = estimates["median"]["point_estimate"]

        case = value_str or "—"
        library = function_id or "—"

        results.append(
            BenchResult(
                group=group_id,
                case=case,
                library=library,
                time=format_ns(median_ns),
            )
        )

    return results


def natural_sort_key(s: str) -> list[int | str]:
    """Sort key that handles embedded numbers naturally (100 < 500 < 1000)."""
    parts: list[int | str] = []
    for tok in re.split(r'(\d+)', s):
        parts.append(int(tok) if tok.isdigit() else tok)
    return parts


def group_rank(group: str) -> tuple[int, str]:
    """Rank by longest matching GROUP_ORDER prefix."""
    best_rank = 10_000
    for i, candidate in enumerate(GROUP_ORDER):
        if group == candidate or group.startswith(candidate + "/"):
            if len(candidate) > len(GROUP_ORDER[best_rank]) if best_rank < len(GROUP_ORDER) else True:
                best_rank = i
    return (best_rank, group)


def render_table(results: list[BenchResult]) -> str:
    results.sort(key=lambda r: (group_rank(r.group), natural_sort_key(r.case), r.library))

    out = [
        "| Group | Case | Library | Time |",
        "|-------|------|---------|------|",
    ]
    for r in results:
        out.append(f"| {r.group} | {r.case} | {r.library} | {r.time} |")
    return "\n".join(out)


def git_short_rev() -> str:
    try:
        return (
            subprocess.check_output(
                ["git", "rev-parse", "--short", "HEAD"], stderr=subprocess.DEVNULL
            )
            .decode()
            .strip()
        )
    except Exception:
        return "unknown"


def update_readme(table_md: str, result_count: int, readme_path: Path) -> None:
    commit = git_short_rev()
    date_cmd = (
        subprocess.check_output(
            ["git", "log", "-1", "--format=%ci"], stderr=subprocess.DEVNULL
        )
        .decode()
        .strip()[:10]
    )

    section = (
        f"{START_MARKER}\n"
        f"### Benchmarks\n\n"
        f"> Last updated: {date_cmd} &middot; commit [`{commit}`](../../commit/{commit})\n"
        f">\n"
        f"> `cargo bench --features full -p duke-sheets`\n\n"
        f"{table_md}\n"
        f"{END_MARKER}"
    )

    readme = readme_path.read_text()
    if START_MARKER in readme and END_MARKER in readme:
        pattern = re.compile(
            re.escape(START_MARKER) + r".*?" + re.escape(END_MARKER),
            re.DOTALL,
        )
        updated = pattern.sub(section, readme)
    elif "\n## License" in readme:
        updated = readme.replace("\n## License", f"\n{section}\n\n## License")
    else:
        updated = readme.rstrip() + "\n\n" + section + "\n"

    readme_path.write_text(updated)
    print(f"README.md updated ({result_count} results)", file=sys.stderr)


def main() -> int:
    if not CRITERION_DIR.is_dir():
        print(
            f"No criterion results at {CRITERION_DIR}/. "
            "Run `cargo bench --features full -p duke-sheets` first.",
            file=sys.stderr,
        )
        return 1

    results = collect_results(CRITERION_DIR)
    if not results:
        print("No benchmark results found.", file=sys.stderr)
        return 1

    table_md = render_table(results)
    update_readme(table_md, len(results), Path("README.md"))
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
