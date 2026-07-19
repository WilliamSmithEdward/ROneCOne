from __future__ import annotations

from pathlib import Path

from pyopenvba import ExcelFile, VBAModuleKind

from build_test_workbook import prepare_class_source, read_vba


ROOT = Path(__file__).resolve().parents[1]
WORKBOOK = ROOT / "demo" / "ROneCOne_Demo.xlsm"


def package(workbook_path: Path = WORKBOOK) -> None:
    runtime_source = prepare_class_source(ROOT / "src" / "ROneCOne.cls")
    demo_source = read_vba(ROOT / "demo" / "vba" / "DemoUsage.bas")

    with ExcelFile(workbook_path) as workbook:
        project = workbook.vba_project()
        existing = set(workbook.module_names())
        for module_name in ("Module1", "ROneCOne", "DemoUsage"):
            if module_name in existing:
                project.delete_module(module_name)
        project.add_module("ROneCOne", runtime_source, kind=VBAModuleKind.other)
        project.add_module("DemoUsage", demo_source, kind=VBAModuleKind.standard)
        workbook.save()

    with ExcelFile(workbook_path) as verification:
        expected = {"ROneCOne", "DemoUsage"}
        actual = set(verification.module_names())
        missing = expected - actual
        if missing:
            raise RuntimeError(f"Demo workbook is missing modules: {sorted(missing)}")
        if verification.get_module("ROneCOne") != runtime_source:
            raise RuntimeError("Demo ROneCOne source did not round-trip byte-for-byte")


if __name__ == "__main__":
    package()
    print(WORKBOOK)
