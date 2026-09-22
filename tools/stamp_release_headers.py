"""Stamp the release version, date, and MIT license into every shipped module.

ROneCOne.cls downloaded on its own, or a module copied out of a demo workbook,
travels without the repository's LICENSE file, so each one carries the full
notice and the release it belongs to. The version and date come from the newest
dated heading in CHANGELOG.md. Run this after dating a release there, then
repackage the demo workbooks; the source contract fails until every header
matches.
"""

from __future__ import annotations

import re
from pathlib import Path


ROOT = Path(__file__).resolve().parents[1]
CHANGELOG = ROOT / "CHANGELOG.md"
LICENSE = ROOT / "LICENSE"
RELEASE_HEADING = re.compile(r"^## (\d+\.\d+\.\d+) - (\d{4}-\d{2}-\d{2})$", re.MULTILINE)
RELEASE_LINE = re.compile(r"^' ROneCOne \d+\.\d+\.\d+, released \d{4}-\d{2}-\d{2}$")


def release_modules() -> list[Path]:
    """The runtime and every module the demo workbooks package."""
    demo = ROOT / "demo" / "vba"
    return [
        ROOT / "src" / "ROneCOne.cls",
        *sorted(demo.glob("*.bas")),
        *sorted(demo.glob("*.cls")),
    ]


def latest_release() -> tuple[str, str]:
    match = RELEASE_HEADING.search(CHANGELOG.read_text(encoding="utf-8"))
    if match is None:
        raise ValueError("CHANGELOG.md has no dated release heading")
    return match.group(1), match.group(2)


def release_header(version: str, date: str) -> list[str]:
    notice = LICENSE.read_text(encoding="utf-8").strip().splitlines()
    return [
        f"' ROneCOne {version}, released {date}",
        "'",
        *(f"' {line}".rstrip() for line in notice),
    ]


def stamp(text: str, header: list[str]) -> str:
    """Put the header directly below Option Explicit, replacing an older one."""
    newline = "\r\n" if "\r\n" in text else "\n"
    lines = text.split(newline)
    if "Option Explicit" not in lines:
        raise ValueError("module has no Option Explicit line")
    start = lines.index("Option Explicit") + 2
    if lines[start - 1] != "":
        raise ValueError("Option Explicit must be followed by a blank line")
    if RELEASE_LINE.match(lines[start]):
        end = lines.index("", start)
    else:
        end = start
        header = [*header, ""]
    return newline.join([*lines[:start], *header, *lines[end:]])


def main() -> None:
    header = release_header(*latest_release())
    for path in release_modules():
        text = path.read_bytes().decode("utf-8")
        stamped = stamp(text, header)
        if stamped != text:
            path.write_bytes(stamped.encode("utf-8"))
            print(path.relative_to(ROOT))


if __name__ == "__main__":
    main()
