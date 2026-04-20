#!/usr/bin/env python3
"""Fetch [MS-ODRAWXML] Appendix A schema sections from learn.microsoft.com.

Microsoft hosts the authoritative XSDs for DrawingML extensions (ChartEx,
etc.) inside the [MS-ODRAWXML] specification pages. Same rendered-HTML
pattern as [MS-XLSX] Appendix A - the XSD lives in a `<pre>...</pre>` block
with entity-encoded markup.

Currently we fetch section 5.22 (ChartEx, namespace
`http://schemas.microsoft.com/office/drawing/2014/chartex`). Other
[MS-ODRAWXML] sections can be added to SCHEMAS as needed.

One-shot fetch - after running once, the `.xsd` files are committed and
regeneration reads from disk. Rerun only to pick up spec updates from
Microsoft.
"""

from __future__ import annotations

import html
import re
import sys
import urllib.request
from pathlib import Path

OUT_DIR = (
    Path(__file__).resolve().parent.parent
    / "spec"
    / "sources"
    / "ms-odrawxml"
    / "schemas"
)

# (section, namespace_url, guid) - from [MS-ODRAWXML] Appendix A index.
SCHEMAS = [
    (
        "5.22",
        "http://schemas.microsoft.com/office/drawing/2014/chartex",
        "e2723b0a-9120-42a5-bd11-c252ccb13c1e",
    ),
]

URL_TMPL = (
    "https://learn.microsoft.com/en-us/openspecs/office_standards/ms-odrawxml/{guid}"
)


def slugify(namespace: str) -> str:
    path = namespace.rsplit("//", 1)[-1]
    path = path.replace("http:/", "")
    path = path.replace("schemas.microsoft.com/office/", "")
    path = path.replace("/", "-").replace(" ", "-").lower()
    path = re.sub(r"[^a-z0-9-]", "", path)
    return path


def extract_xsd(html_text: str) -> str:
    match = re.search(r"<pre>(.+?)</pre>", html_text, re.DOTALL)
    if not match:
        raise RuntimeError("No <pre> block found")
    raw = match.group(1)
    stripped = re.sub(r"<[^>]+>", "", raw)
    xsd = html.unescape(stripped).strip()
    if not xsd.startswith("<xsd:schema") and "<xsd:schema" not in xsd[:200]:
        raise RuntimeError(
            f"Content doesn't look like XSD; first 200 chars: {xsd[:200]!r}"
        )
    return xsd


def fetch(guid: str) -> str:
    url = URL_TMPL.format(guid=guid)
    req = urllib.request.Request(url, headers={"User-Agent": "Mozilla/5.0"})
    with urllib.request.urlopen(req, timeout=60) as resp:
        return resp.read().decode("utf-8", errors="replace")


def main() -> int:
    force = "--force" in sys.argv
    OUT_DIR.mkdir(parents=True, exist_ok=True)
    errors: list[str] = []
    for section, namespace, guid in SCHEMAS:
        slug = slugify(namespace)
        out_path = OUT_DIR / f"{section}_{slug}.xsd"
        if out_path.exists() and not force:
            print(f"  skip (exists): {out_path.name}")
            continue
        try:
            page = fetch(guid)
            xsd = extract_xsd(page)
        except Exception as exc:
            errors.append(f"{section} {namespace}: {exc}")
            print(f"  ERR  {out_path.name}: {exc}", file=sys.stderr)
            continue
        out_path.write_text(xsd + "\n", encoding="utf-8")
        print(f"  wrote ({len(xsd)} chars): {out_path.name}")
    if errors:
        print(f"\n{len(errors)} error(s):", file=sys.stderr)
        for e in errors:
            print(f"  {e}", file=sys.stderr)
        return 1
    print(f"\nDone. {len(SCHEMAS)} schema file(s) in {OUT_DIR.relative_to(Path.cwd())}")
    return 0


if __name__ == "__main__":
    sys.exit(main())
