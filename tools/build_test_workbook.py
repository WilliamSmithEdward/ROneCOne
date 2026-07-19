from __future__ import annotations

import argparse
from pathlib import Path

from pyopenvba import ExcelFile, VBAModuleKind


ROOT = Path(__file__).resolve().parents[1]
DEFAULT_OUTPUT = ROOT / "tests" / "output" / "ROneCOne_DelegateTests.xlsm"


def read_vba(path: Path) -> str:
    text = path.read_text(encoding="utf-8")
    return text.replace("\r\n", "\n").replace("\r", "\n").replace("\n", "\r\n")


def prepare_class_source(path: Path) -> str:
    """Remove the VBE export-only VERSION/BEGIN preamble for binary injection."""
    source = read_vba(path)
    lines = source.split("\r\n")
    if not lines or lines[0].strip().upper() != "VERSION 1.0 CLASS":
        return source

    for index, line in enumerate(lines[1:], start=1):
        if line.strip().upper() == "END":
            return "\r\n".join(lines[index + 1 :])
    raise ValueError(f"Class export preamble has no END marker: {path}")


def build(output: Path) -> None:
    output.parent.mkdir(parents=True, exist_ok=True)

    runtime_source = prepare_class_source(ROOT / "src" / "ROneCOne.cls")
    fixture_source = prepare_class_source(ROOT / "tests" / "vba" / "DelegateFixture.cls")
    test_source = read_vba(ROOT / "tests" / "vba" / "TestDelegates.bas")

    with ExcelFile.create_new(output) as workbook:
        project = workbook.vba_project()
        if "Module1" in workbook.module_names():
            project.delete_module("Module1")
        project.add_module("ROneCOne", runtime_source, kind=VBAModuleKind.other)
        project.add_module("DelegateFixture", fixture_source, kind=VBAModuleKind.other)
        project.add_module("TestDelegates", test_source, kind=VBAModuleKind.standard)
        workbook.save()

    with ExcelFile(output) as verification:
        expected = {"ROneCOne", "DelegateFixture", "TestDelegates"}
        actual = set(verification.module_names())
        missing = expected - actual
        if missing:
            raise RuntimeError(f"Workbook is missing modules: {sorted(missing)}")
        if verification.get_module("ROneCOne") != runtime_source:
            raise RuntimeError("ROneCOne source did not round-trip byte-for-byte")


def main() -> None:
    parser = argparse.ArgumentParser(description="Build the ROneCOne live Excel test workbook")
    parser.add_argument("--output", type=Path, default=DEFAULT_OUTPUT)
    args = parser.parse_args()
    build(args.output.resolve())
    print(args.output.resolve())


if __name__ == "__main__":
    main()
