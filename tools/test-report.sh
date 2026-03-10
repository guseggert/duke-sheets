#!/bin/bash
# Run unit/lib/doc Rust tests and print a per-crate summary report.
# Skips e2e tests (LO/COM) and corpus tests by default.
# Use 'mise run test:all' or 'mise run test:e2e' for e2e coverage.
set -euo pipefail

mkdir -p /tmp/duke-sheets-urp /tmp/duke-sheets-excel

LOGFILE=$(mktemp /tmp/duke-sheets-test-report.XXXXXX)
trap 'rm -f "$LOGFILE"' EXIT

# Helper: run a command, tee live to stdout, append to log.
run() { "$@" 2>&1 | tee -a "$LOGFILE"; }

# Phase 1: Unit tests, doc-tests, and non-e2e integration tests.
# Exclude crates whose test targets need external services or corpus
# data — those deadlock or take too long for a quick report.
#   - duke-sheets-xlsx/xls/excel-com/libreoffice: e2e needs LO/COM
#   - duke-sheets: corpus tests are slow and need excel_samples/
run cargo test --workspace \
    --exclude duke-sheets \
    --exclude duke-sheets-xlsx \
    --exclude duke-sheets-xls \
    --exclude duke-sheets-excel-com \
    --exclude duke-sheets-libreoffice || true

# Still run the unit tests (--lib) for the excluded crates.
for pkg in duke-sheets duke-sheets-xlsx duke-sheets-xls duke-sheets-excel-com duke-sheets-libreoffice; do
    run cargo test -p "$pkg" --lib || true
done

echo ""
echo "═══════════════════════════════════════════════════════════"
echo "  TEST REPORT"
echo "═══════════════════════════════════════════════════════════"
echo ""

TOTAL_PASS=0
TOTAL_FAIL=0
TOTAL_IGNORE=0
ANY_FAIL=false
crate=""
src=""

while IFS= read -r line; do
    # Match: Running unittests src/lib.rs (target/debug/deps/duke_sheets_core-abc123)
    #    or: Running tests/e2e/main.rs (target/debug/deps/e2e-abc123)
    if [[ "$line" =~ Running\ (.+)\ \(target/debug/deps/([a-zA-Z0-9_]+)- ]]; then
        src="${BASH_REMATCH[1]}"
        crate="${BASH_REMATCH[2]}"
    fi
    # Match: Doc-tests crate_name
    if [[ "$line" =~ Doc-tests\ ([a-zA-Z0-9_]+) ]]; then
        src="doc-tests"
        crate="${BASH_REMATCH[1]}"
    fi
    # Match: test result: ok. N passed; N failed; N ignored; ...
    if [[ "$line" =~ test\ result:\ (ok|FAILED)\.\ ([0-9]+)\ passed\;\ ([0-9]+)\ failed\;\ ([0-9]+)\ ignored ]]; then
        pass="${BASH_REMATCH[2]}"
        fail="${BASH_REMATCH[3]}"
        ign="${BASH_REMATCH[4]}"
        TOTAL_PASS=$((TOTAL_PASS + pass))
        TOTAL_FAIL=$((TOTAL_FAIL + fail))
        TOTAL_IGNORE=$((TOTAL_IGNORE + ign))
        if [ "$fail" -gt 0 ]; then
            ANY_FAIL=true
            printf "  ❌  %-45s %4d passed, %d failed, %d ignored\n" "$crate ($src)" "$pass" "$fail" "$ign"
        elif [ "$pass" -gt 0 ]; then
            printf "  ✅  %-45s %4d passed" "$crate ($src)" "$pass"
            if [ "$ign" -gt 0 ]; then printf ", %d ignored" "$ign"; fi
            echo ""
        fi
    fi
done < "$LOGFILE"

echo ""
echo "───────────────────────────────────────────────────────────"
if [ "$ANY_FAIL" = true ]; then
    printf "  TOTAL: %d passed, %d FAILED, %d ignored\n" "$TOTAL_PASS" "$TOTAL_FAIL" "$TOTAL_IGNORE"
    echo "═══════════════════════════════════════════════════════════"
    exit 1
else
    printf "  TOTAL: %d passed, %d ignored — all ok ✅\n" "$TOTAL_PASS" "$TOTAL_IGNORE"
    echo "═══════════════════════════════════════════════════════════"
fi
