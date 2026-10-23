#!/usr/bin/env python3
"""Repair pip/certifi in the interpreter running this script; --check only audits."""

import argparse
import hashlib
import importlib.metadata
import re
import subprocess
import sys
import tempfile
import urllib.request
from pathlib import Path

# Independently published hashes from the Python incident advisory and PyPI.
# Downloads may traverse an old trust store; verify bytes before executing them.
PACKAGES = {
    "pip": (
        "26.2.2",
        "https://files.pythonhosted.org/packages/3f/06/3e4420302363fbb71be3c7f35ce042f54bfca330e3e66b3a2e3c4cca579f/pip-26.2.2-py3-none-any.whl",
        "fd7f4dfa8190e014013b9f08f3e973f4e4c26ffd179a204895533a7254813a70",
    ),
    "certifi": (
        "2026.10.19",
        "https://files.pythonhosted.org/packages/c4/51/29046e48f693c54e5b7412216d447a1d8362c9653148d98c02cdf65fa733/certifi-2026.10.19-py3-none-any.whl",
        "42d35340cc5b9ba78e818ddbff0c5b1c2b779de448a340fd1b989f5fa8670e1b",
    ),
}


MAX_WHEEL_BYTES = 8 * 1024 * 1024


def outdated_packages():
    outdated = []
    for name, (minimum, _, _) in PACKAGES.items():
        try:
            installed = importlib.metadata.version(name)
        except importlib.metadata.PackageNotFoundError:
            installed = None
        # Do not silently replace vendor builds or prereleases we cannot compare.
        if installed and not re.fullmatch(r"\d+(?:\.\d+)*", installed):
            raise ValueError(f"Review nonstandard {name} version manually: {installed}")
        if installed is None or tuple(map(int, installed.split("."))) < tuple(
            map(int, minimum.split("."))
        ):
            outdated.append(name)
        print(f"{name}: {installed or 'missing'} (minimum {minimum})", flush=True)
    return outdated


def download(name, directory):
    _, url, expected_hash = PACKAGES[name]
    with urllib.request.urlopen(url, timeout=60) as response:
        data = response.read(MAX_WHEEL_BYTES + 1)
    if len(data) > MAX_WHEEL_BYTES:
        raise ValueError(f"Wheel for {name} exceeds {MAX_WHEEL_BYTES} bytes; refusing to install")
    if hashlib.sha256(data).hexdigest() != expected_hash:
        raise ValueError(f"SHA-256 mismatch for {name}; refusing to install")
    wheel = directory / url.rsplit("/", 1)[1]
    wheel.write_bytes(data)
    return wheel


def main():
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--check", action="store_true")
    args = parser.parse_args()
    if sys.version_info < (3, 10):
        parser.error("Upgrade Python to a supported version (>=3.10) first")
    print(f"Interpreter: {sys.executable}", flush=True)
    outdated = outdated_packages()
    if args.check:
        return int(bool(outdated))
    if not outdated:
        return 0
    with tempfile.TemporaryDirectory(prefix="python-trust-") as directory:
        # Execute a verified pip wheel directly: works even in pip-less uv venvs.
        # --no-index/--no-deps prevent any unverified bootstrap downloads.
        wheels = {name: download(name, Path(directory)) for name in {"pip", *outdated}}
        subprocess.run(
            [
                sys.executable,
                "-I",
                str(wheels["pip"]) + "/pip",
                "--isolated",
                "--disable-pip-version-check",
                "install",
                "--no-index",
                "--no-deps",
                *[str(wheels[name]) for name in outdated],
            ],
            check=True,
        )
    importlib.invalidate_caches()
    return int(bool(outdated_packages()))


if __name__ == "__main__":
    try:
        sys.exit(main())
    except (OSError, ValueError, subprocess.CalledProcessError) as error:
        print(f"Python trust repair failed: {error}", file=sys.stderr)
        sys.exit(1)
