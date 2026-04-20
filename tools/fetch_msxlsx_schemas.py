#!/usr/bin/env python3
"""Fetch [MS-XLSX] Appendix A schema sections from learn.microsoft.com.

Microsoft hosts the authoritative XSDs for all Excel proprietary extensions
(x14, x15, x16, xr, xr2, etc.) as inline content in the [MS-XLSX] specification
pages. This tool fetches each of the 45 sections listed in Appendix A §5 and
writes each one to `spec/sources/ms-xlsx/schemas/<section>_<slug>.xsd`.

The content lives inside a `<pre>...</pre>` block in the rendered HTML with
XSD markup HTML-entity-encoded. We strip real HTML tags first, then decode
entities, which leaves plain XSD text.

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
    Path(__file__).resolve().parent.parent / "spec" / "sources" / "ms-xlsx" / "schemas"
)

# (section, namespace_url, guid) - from [MS-XLSX] Appendix A §5 index,
# fetched from https://learn.microsoft.com/en-us/openspecs/office_standards/ms-xlsx/6624db33-496c-47f7-a562-a54cb01b133f
SCHEMAS = [
    (
        "5.1",
        "http://schemas.microsoft.com/office/excel/2006/main",
        "b0ddba06-ac73-4d3d-a913-8fbe4a56cf28",
    ),
    (
        "5.2",
        "http://schemas.microsoft.com/office/drawing/2010/slicer",
        "8ccfc6ad-cdae-4ac6-9643-17d999d1a7e6",
    ),
    (
        "5.3",
        "http://schemas.microsoft.com/office/spreadsheetml/2010/11/main",
        "e42bbfd7-2a3d-4308-a4f3-30313fc506b9",
    ),
    (
        "5.4",
        "http://schemas.microsoft.com/office/spreadsheetml/2009/9/main",
        "0a377581-c743-4ace-bfcd-22f359e70165",
    ),
    (
        "5.5",
        "http://schemas.microsoft.com/office/spreadsheetml/2009/9/ac",
        "7148c81d-4d6e-4204-be07-a7c09f4277c7",
    ),
    (
        "5.6",
        "http://schemas.microsoft.com/office/spreadsheetml/2011/1/ac",
        "aa12452a-467d-4192-8ebf-0a90d86dd64b",
    ),
    (
        "5.7",
        "http://schemas.microsoft.com/office/drawing/2012/timeslicer",
        "c4febed7-8f37-443b-b83d-668ffd09d082",
    ),
    (
        "5.8",
        "http://schemas.microsoft.com/office/excel/2010/spreadsheetDrawing",
        "b54c3971-c168-4f9c-8791-c07f8d8d8002",
    ),
    (
        "5.9",
        "http://schemas.microsoft.com/office/spreadsheetml/2010/11/ac",
        "ce659475-d580-4104-895f-b72947105f24",
    ),
    (
        "5.10",
        "http://schemas.microsoft.com/office/spreadsheetml/2014/11/main",
        "4fe24a21-6f69-4680-882f-f86b16a69119",
    ),
    (
        "5.11",
        "http://schemas.microsoft.com/office/spreadsheetml/2015/02/main",
        "e29f5a54-ff71-47e9-9183-440b65eb27ad",
    ),
    (
        "5.12",
        "http://schemas.microsoft.com/office/spreadsheetml/2016/revision10",
        "8a818a7f-92d0-4390-9855-1a484d8f9347",
    ),
    (
        "5.13",
        "http://schemas.microsoft.com/office/spreadsheetml/2016/revision9",
        "1df7cd78-d8bc-4ec4-8952-1b61beeaee76",
    ),
    (
        "5.14",
        "http://schemas.microsoft.com/office/spreadsheetml/2016/revision6",
        "e47e052c-40d0-43cd-add4-73a220966b69",
    ),
    (
        "5.15",
        "http://schemas.microsoft.com/office/spreadsheetml/2014/revision",
        "5ba37a83-255b-4e2d-ae7b-b4d14d44ad3a",
    ),
    (
        "5.16",
        "http://schemas.microsoft.com/office/spreadsheetml/2015/revision2",
        "7b729443-d0cd-434e-8d8c-b228f701a55b",
    ),
    (
        "5.17",
        "http://schemas.microsoft.com/office/spreadsheetml/2016/revision3",
        "1c6fdaed-8a40-4086-be07-0135b10b8f90",
    ),
    (
        "5.18",
        "http://schemas.microsoft.com/office/spreadsheetml/2016/revision5",
        "83ec1155-c05a-4c5e-84b5-844b18358690",
    ),
    (
        "5.19",
        "http://schemas.microsoft.com/office/spreadsheetml/2016/pivotdefaultlayout",
        "9b107b7d-1fe8-476f-9ba5-f237ceaf6a0e",
    ),
    (
        "5.20",
        "http://schemas.microsoft.com/office/spreadsheetml/2017/richdata2",
        "71491664-45c9-427a-b1d8-69dc70a1dd07",
    ),
    (
        "5.21",
        "http://schemas.microsoft.com/office/spreadsheetml/2017/richdata",
        "9058338f-cf3b-4a7b-8f84-622e180eacc5",
    ),
    (
        "5.22",
        "http://schemas.microsoft.com/office/spreadsheetml/2018/calcfeatures",
        "6c4054c8-05a9-49dc-baa9-13e5a6282989",
    ),
    (
        "5.23",
        "http://schemas.microsoft.com/office/spreadsheetml/2018/threadedcomments",
        "adb84732-9fc8-48b6-bddc-6b0bcdaad940",
    ),
    (
        "5.24",
        "http://schemas.microsoft.com/office/spreadsheetml/2018/08/main",
        "b38f8644-d3e6-4d20-889d-56a9df09cad8",
    ),
    (
        "5.25",
        "http://schemas.microsoft.com/office/spreadsheetml/2017/dynamicarray",
        "878c8ac4-8548-4bfb-9ad4-b1378e0efed1",
    ),
    (
        "5.26",
        "http://schemas.microsoft.com/office/spreadsheetml/2019/namedsheetviews",
        "55396aef-0c07-4ffb-9ffb-e48b6d339abe",
    ),
    (
        "5.27",
        "http://schemas.microsoft.com/office/spreadsheetml/2019/extlinksprops",
        "b2939e11-39c1-4756-9909-7a2dc79e8c2e",
    ),
    (
        "5.28",
        "http://schemas.microsoft.com/office/spreadsheetml/2020/richdatawebimage",
        "1dc262ba-8b45-4611-b6fb-ce1b04a561e0",
    ),
    (
        "5.29",
        "http://schemas.microsoft.com/office/spreadsheetml/2020/pivotNov2020",
        "a9f2fa9d-070f-4491-80e8-832783f0ada2",
    ),
    (
        "5.30",
        "http://schemas.microsoft.com/office/spreadsheetml/2020/threadedcomments2",
        "716e392d-d858-443b-a393-141a983260b0",
    ),
    (
        "5.31",
        "http://schemas.microsoft.com/office/spreadsheetml/2020/richvaluerefresh",
        "7c97123f-a895-4094-afe6-b661d7f3e634",
    ),
    (
        "5.32",
        "http://schemas.microsoft.com/office/spreadsheetml/2022/pivotVersionInfo",
        "ce4604c1-dd30-452c-9fc1-630dc4012961",
    ),
    (
        "5.33",
        "http://schemas.microsoft.com/office/spreadsheetml/2022/pivotRichData",
        "0f1d2cd1-bba1-4ad3-b607-ffa7aa1e296f",
    ),
    (
        "5.34",
        "http://schemas.microsoft.com/office/spreadsheetml/2021/extlinks2021",
        "ffd38ada-cfa5-4242-976b-198f6fd508d9",
    ),
    (
        "5.35",
        "http://schemas.microsoft.com/office/spreadsheetml/2022/richvaluerel",
        "56bafcc4-1ae8-4894-a86c-1e92f165a3c5",
    ),
    (
        "5.36",
        "http://schemas.microsoft.com/office/spreadsheetml/2022/featurepropertybag",
        "9e45d09d-206b-4134-a90e-b5be9c08b2fa",
    ),
    (
        "5.37",
        "http://schemas.microsoft.com/office/spreadsheetml/2023/msForms",
        "e38e8867-d5f0-492a-b64b-b57779e18216",
    ),
    (
        "5.38",
        "http://schemas.microsoft.com/office/spreadsheetml/2023/externalCodeService",
        "b3c832bd-13c7-4c2f-869e-ac31bf13ea7c",
    ),
    (
        "5.39",
        "http://schemas.microsoft.com/office/spreadsheetml/2023/python",
        "ef230d23-8d80-473c-925e-3e5b98abdbba",
    ),
    (
        "5.40",
        "http://schemas.microsoft.com/office/spreadsheetml/2023/pivot2023Calculation",
        "bc0e64e7-19ba-49a8-8526-22960eacc4ba",
    ),
    (
        "5.41",
        "http://schemas.microsoft.com/office/spreadsheetml/2024/pivotAutoRefresh",
        "6a531596-7c64-455b-93fa-7a8d7e3fca15",
    ),
    (
        "5.42",
        "http://schemas.microsoft.com/office/spreadsheetml/2024/workbookCompatibilityVersion",
        "a5954858-15b3-4cde-98b3-670ea2710489",
    ),
    (
        "5.43",
        "http://schemas.microsoft.com/office/spreadsheetml/2023/showDataTypeIcons",
        "8d375463-bcab-4946-80c7-67c59b9b40c0",
    ),
    (
        "5.44",
        "http://schemas.microsoft.com/office/spreadsheetml/2025/externalCodeService2",
        "4747e200-ff5c-428f-ba09-0d2586d791ed",
    ),
    (
        "5.45",
        "http://schemas.microsoft.com/office/spreadsheetml/2025/pivotDataSource",
        "04ae0b01-21e2-49a1-8f77-46d3b9180217",
    ),
]

URL_TMPL = "https://learn.microsoft.com/en-us/openspecs/office_standards/ms-xlsx/{guid}"


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
