#!/usr/bin/env python3
"""Group-aware parallel test runner.

`cargo test` runs test binaries one at a time, and the repo's mise tasks
additionally forced `RUST_TEST_THREADS=1` so the LibreOffice and Excel
bridge singletons wouldn't be entered concurrently. That serialized the
entire suite, including the ~2900 tests that never touch a backend.

This runner splits the workspace's test binaries into three groups and
runs the groups concurrently:

  pure   in-process only; a worker pool runs several binaries at once,
         each with its own libtest threads
  lo     touches the LibreOffice URP container; one binary at a time,
         --test-threads=1
  excel  touches the Excel COM bridge VM; one binary at a time,
         --test-threads=1

`lo` and `excel` stay serial internally because a single soffice daemon
and a single `[STAThread]` COM server can only service one request
stream, but there is no reason for them to block `pure` or each other.
The `lo` group is serial across binaries too, not just within one:
`ensure_lo()` replaces the container on first call in each test process,
which would kill a concurrent binary's URP connection.

Group membership is derived from each target's source files (following
`mod` declarations) rather than a hand-maintained list. Misclassifying a
pure target as `lo`/`excel` only costs a little parallelism, while the
reverse produces a loud test failure, so the marker set below
deliberately errs toward over-classification. That is why the
`duke-sheets-xls` `*_round_trip.rs` targets land in `lo`: each carries
one `#[ignore]`d `lo_can_*` test, and their non-ignored tests cost well
under a second in total.

Usage:
    tools/run-tests.py [runner opts] [cargo args...] [-- libtest args...]

    tools/run-tests.py                          # whole workspace
    tools/run-tests.py --groups pure            # skip both backends
    tools/run-tests.py -p duke-sheets-xls
    tools/run-tests.py -- --ignored
"""

import argparse
import json
import os
import queue
import re
import shutil
import subprocess
import sys
import threading
import time
from pathlib import Path

REPO_ROOT = Path(__file__).resolve().parent.parent

# Substrings that mean a target drives a backend. Matched against every
# source file linked into the target.
EXCEL_MARKERS = (
    "ensure_excel_bridge",
    "excel_bridge",
    "roundtrip_through_excel",
    "push_file_to_vm",
    "pull_file_from_vm",
)
LO_MARKERS = (
    "ensure_lo",
    "lo_bridge",
    "LibreOfficeBridge",
    "skip_if_no_lo",
)

MOD_RE = re.compile(r"(?m)^[^\S\n]*(?:pub(?:\([^)]*\))?[^\S\n]+)?mod[^\S\n]+([A-Za-z_]\w*)[^\S\n]*;")
PATH_ATTR_RE = re.compile(r'#\[path\s*=\s*"([^"]+)"\]')
SHARED_DIRS = ("/tmp/duke-sheets-urp", "/tmp/duke-sheets-excel")


def target_sources(src_path: Path) -> set[Path]:
    """Every .rs file linked into a target, resolved by walking `mod` decls."""
    seen: set[Path] = set()
    pending = [src_path]
    while pending:
        current = pending.pop()
        try:
            current = current.resolve()
        except OSError:
            continue
        if current in seen or not current.is_file():
            continue
        seen.add(current)
        try:
            text = current.read_text(errors="replace")
        except OSError:
            continue

        parent = current.parent
        # A crate/target root (main.rs, lib.rs, mod.rs) resolves `mod x` to
        # ./x.rs; any other file resolves it to ./<stem>/x.rs. Try both --
        # a spurious extra file can only over-classify, which is safe.
        roots = [parent, parent / current.stem]
        for name in MOD_RE.findall(text):
            for root in roots:
                for candidate in (root / f"{name}.rs", root / name / "mod.rs"):
                    if candidate.is_file():
                        pending.append(candidate)
        for rel in PATH_ATTR_RE.findall(text):
            for root in roots:
                candidate = root / rel
                if candidate.is_file():
                    pending.append(candidate)
    return seen


def classify(src_path: Path) -> str:
    blob = ""
    for path in target_sources(src_path):
        try:
            blob += path.read_text(errors="replace")
        except OSError:
            pass
    if any(m in blob for m in EXCEL_MARKERS):
        return "excel"
    if any(m in blob for m in LO_MARKERS):
        return "lo"
    return "pure"


class Target:
    __slots__ = ("pkg", "name", "kind", "exe", "group", "secs", "code", "output", "summary")

    def __init__(self, pkg, name, kind, exe, group):
        self.pkg, self.name, self.kind, self.exe, self.group = pkg, name, kind, exe, group
        self.secs = 0.0
        self.code = None
        self.output = ""
        self.summary = ""

    @property
    def label(self):
        suffix = " (lib)" if self.kind == "lib" else ""
        return f"{self.pkg}/{self.name}{suffix}"


def enumerate_targets(cargo_args: list[str]) -> list[Target]:
    """Build the test binaries and return them classified by group."""
    cmd = ["cargo", "test", "--no-run", "--message-format=json-render-diagnostics"]
    if not any(a in ("-p", "--package", "--workspace", "--all") for a in cargo_args):
        cmd.append("--workspace")
    cmd += cargo_args

    print(f"[run-tests] building: {' '.join(cmd)}", file=sys.stderr, flush=True)
    proc = subprocess.Popen(cmd, stdout=subprocess.PIPE, text=True, cwd=REPO_ROOT)
    targets = []
    assert proc.stdout is not None
    for line in proc.stdout:
        line = line.strip()
        if not line.startswith("{"):
            if line:
                print(line, file=sys.stderr, flush=True)
            continue
        try:
            msg = json.loads(line)
        except json.JSONDecodeError:
            continue
        reason = msg.get("reason")
        if reason == "compiler-message":
            rendered = msg.get("message", {}).get("rendered")
            if rendered:
                print(rendered, file=sys.stderr, end="", flush=True)
            continue
        if reason != "compiler-artifact":
            continue
        if not msg.get("profile", {}).get("test"):
            continue
        exe = msg.get("executable")
        if not exe:
            continue
        target = msg["target"]
        pkg = msg["package_id"].split("#")[0].rstrip("/").split("/")[-1].split("@")[0]
        group = classify(Path(target["src_path"]))
        targets.append(Target(pkg, target["name"], target["kind"][0], exe, group))
    if proc.wait() != 0:
        print("[run-tests] build failed", file=sys.stderr)
        sys.exit(proc.returncode or 1)
    return targets


def run_target(target: Target, threads: int, libtest_args: list[str]) -> None:
    cmd = [target.exe]
    # An explicit --test-threads in the passthrough args wins; libtest
    # would otherwise see the flag twice.
    if not any(a == "--test-threads" or a.startswith("--test-threads=") for a in libtest_args):
        cmd += ["--test-threads", str(threads)]
    cmd += libtest_args
    started = time.monotonic()
    try:
        proc = subprocess.run(cmd, capture_output=True, text=True, cwd=REPO_ROOT)
        target.code = proc.returncode
        target.output = proc.stdout + proc.stderr
    except Exception as exc:  # noqa: BLE001 - must never kill the worker thread
        target.code = 127
        target.output = f"failed to run {target.exe}: {exc!r}"
    target.secs = time.monotonic() - started
    for line in reversed(target.output.splitlines()):
        if line.startswith("test result:"):
            target.summary = line
            break


def main() -> int:
    parser = argparse.ArgumentParser(add_help=False)
    parser.add_argument("-j", "--jobs", type=int, default=os.cpu_count() or 4,
                        help="pure-group binaries to run at once (default: nproc)")
    parser.add_argument("--inner-threads", type=int, default=2,
                        help="libtest threads per pure-group binary (default: 2)")
    parser.add_argument("--groups", default="pure,lo,excel",
                        help="comma-separated subset of pure,lo,excel")
    parser.add_argument("-h", "--help", action="store_true")
    args, rest = parser.parse_known_args()

    if args.help:
        print(__doc__)
        return 0

    if "--" in rest:
        split = rest.index("--")
        cargo_args, libtest_args = rest[:split], rest[split + 1:]
    else:
        cargo_args, libtest_args = rest, []

    # One libtest thread per soffice instance: the container runs
    # DUKE_LO_INSTANCES of them (default 8) and each concurrent test needs
    # its own, since a single instance serializes UNO calls internally.
    lo_threads = max(1, int(os.environ.get("DUKE_LO_INSTANCES", "8")))

    wanted = {g.strip() for g in args.groups.split(",") if g.strip()}
    unknown = wanted - {"pure", "lo", "excel"}
    if unknown:
        print(f"[run-tests] unknown group(s): {', '.join(sorted(unknown))}", file=sys.stderr)
        return 2

    for d in SHARED_DIRS:
        Path(d).mkdir(parents=True, exist_ok=True)

    targets = enumerate_targets(cargo_args)
    groups = {g: [t for t in targets if t.group == g] for g in ("pure", "lo", "excel")}
    selected = [t for t in targets if t.group in wanted]

    print(
        f"[run-tests] {len(selected)} test binaries: "
        + ", ".join(f"{g}={len(groups[g])}" for g in ("pure", "lo", "excel"))
        + f"  (running: {','.join(sorted(wanted))})",
        file=sys.stderr, flush=True,
    )

    done = queue.Queue()
    total = len(selected)
    started_at = time.monotonic()

    def report(target, exc=None):
        # Every scheduled target must land on the queue exactly once, or the
        # collector below waits forever.
        if exc is not None:
            target.code = 127
            target.output = f"runner error: {exc!r}"
        done.put(target)

    def run_one(target, threads):
        try:
            run_target(target, threads, libtest_args)
        except BaseException as exc:  # noqa: BLE001
            report(target, exc)
        else:
            report(target)

    def excel_stream():
        # Strictly serial: the bridge server serves one client at a time and
        # Excel's COM apartment is single-threaded.
        for target in groups["excel"]:
            run_one(target, 1)

    def lo_stream():
        # A worker pool, like `pure`. The container's leastconn proxy
        # spreads connections across soffice backends globally, and tests
        # hold one connection per test body, so fairness lives server-side
        # and mild over-subscription just queues gracefully.
        work = queue.Queue()
        for target in groups["lo"]:
            work.put(target)

        def worker():
            while True:
                try:
                    target = work.get_nowait()
                except queue.Empty:
                    return
                run_one(target, lo_threads)

        pool = [threading.Thread(target=worker, daemon=True)
                for _ in range(max(1, min(lo_threads, len(groups["lo"]))))]
        for t in pool:
            t.start()
        for t in pool:
            t.join()

    def pure_stream():
        work = queue.Queue()
        for target in groups["pure"]:
            work.put(target)

        def worker():
            while True:
                try:
                    target = work.get_nowait()
                except queue.Empty:
                    return
                try:
                    run_target(target, args.inner_threads, libtest_args)
                except BaseException as exc:  # noqa: BLE001
                    report(target, exc)
                else:
                    report(target)

        pool = [threading.Thread(target=worker, daemon=True)
                for _ in range(max(1, min(args.jobs, len(groups["pure"]))))]
        for t in pool:
            t.start()
        for t in pool:
            t.join()

    streams = []
    if "pure" in wanted and groups["pure"]:
        streams.append(threading.Thread(target=pure_stream, daemon=True))
    if "lo" in wanted and groups["lo"]:
        streams.append(threading.Thread(target=lo_stream, daemon=True))
    if "excel" in wanted and groups["excel"]:
        streams.append(threading.Thread(target=excel_stream, daemon=True))
    for s in streams:
        s.start()

    completed = []
    failures = []
    tty = sys.stderr.isatty()
    while len(completed) < total:
        try:
            target = done.get(timeout=1.0)
        except queue.Empty:
            # Belt and braces: if every stream thread has exited without
            # reporting all its targets, stop rather than block forever.
            if any(s.is_alive() for s in streams):
                continue
            missing = [t for t in selected if t not in completed]
            for t in missing:
                t.code = t.code if t.code is not None else 127
                if not t.output:
                    t.output = "runner error: stream exited without reporting this target"
                completed.append(t)
                failures.append(t)
                print(f"FAIL  {t.label}  (not reported)", file=sys.stderr)
            break
        completed.append(target)
        if target.code != 0:
            failures.append(target)
            if tty:
                print("\r\033[K", end="", file=sys.stderr)
            print(f"FAIL  {target.label}  ({target.secs:.1f}s)", file=sys.stderr, flush=True)
        elif tty:
            print(
                f"\r\033[K[{len(completed)}/{total}] {time.monotonic() - started_at:6.1f}s  "
                f"{target.label}",
                end="", file=sys.stderr, flush=True,
            )
    if tty:
        print("\r\033[K", end="", file=sys.stderr)
    for s in streams:
        s.join()

    wall = time.monotonic() - started_at
    log_path = Path("/tmp/duke-sheets-test-run.log")
    with log_path.open("w") as fh:
        for target in sorted(completed, key=lambda t: -t.secs):
            fh.write(f"===== {target.label} [{target.group}] {target.secs:.2f}s "
                     f"rc={target.code}\n{target.output}\n")

    for target in failures:
        print(f"\n{'=' * 70}\n{target.label} [{target.group}] failed (rc={target.code})\n"
              f"{'=' * 70}\n{target.output}", file=sys.stderr)

    passed = failed = ignored = 0
    for target in completed:
        m = re.search(r"(\d+) passed; (\d+) failed; (\d+) ignored", target.summary)
        if m:
            passed += int(m.group(1))
            failed += int(m.group(2))
            ignored += int(m.group(3))

    print(f"\n{'=' * 70}", file=sys.stderr)
    print("slowest binaries:", file=sys.stderr)
    for target in sorted(completed, key=lambda t: -t.secs)[:10]:
        print(f"  {target.secs:7.1f}s  [{target.group:5}] {target.label}", file=sys.stderr)
    for group in ("pure", "lo", "excel"):
        members = [t for t in completed if t.group == group]
        if members:
            print(f"  {group:5} group: {len(members)} binaries, "
                  f"{sum(t.secs for t in members):.1f}s of work", file=sys.stderr)
    print(f"\n{passed} passed; {failed} failed; {ignored} ignored "
          f"across {len(completed)} binaries in {wall:.1f}s wall", file=sys.stderr)
    print(f"full log: {log_path}", file=sys.stderr)
    if failures:
        print(f"FAILED: {len(failures)} binaries: "
              f"{', '.join(t.label for t in failures)}", file=sys.stderr)
        return 1
    print("ok", file=sys.stderr)
    return 0


if __name__ == "__main__":
    if shutil.which("cargo") is None:
        print("[run-tests] cargo not found on PATH", file=sys.stderr)
        sys.exit(127)
    try:
        sys.exit(main())
    except KeyboardInterrupt:
        print("\n[run-tests] interrupted", file=sys.stderr)
        sys.exit(130)
