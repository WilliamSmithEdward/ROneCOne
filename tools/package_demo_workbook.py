from __future__ import annotations

import argparse
from pathlib import Path

from pyopenvba import ExcelFile, VBAModuleKind

from build_test_workbook import prepare_class_source, read_vba


ROOT = Path(__file__).resolve().parents[1]
DELEGATES_WORKBOOK = ROOT / "demo" / "ROneCOne_Delegates_Demo.xlsm"
COLLECTIONS_WORKBOOK = ROOT / "demo" / "ROneCOne_Collections_Demo.xlsm"


def package_delegates(workbook_path: Path = DELEGATES_WORKBOOK) -> None:
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
        if verification.get_module("DemoUsage") != demo_source:
            raise RuntimeError("Demo usage source did not round-trip byte-for-byte")


def package_collections(workbook_path: Path = COLLECTIONS_WORKBOOK) -> None:
    runtime_source = prepare_class_source(ROOT / "src" / "ROneCOne.cls")
    customer_source = prepare_class_source(ROOT / "demo" / "vba" / "DemoCustomer.cls")
    demo_source = read_vba(ROOT / "demo" / "vba" / "CollectionsDemoUsage.bas")

    with ExcelFile(workbook_path) as workbook:
        project = workbook.vba_project()
        existing = set(workbook.module_names())
        names = (
            "Module1",
            "ROneCOne",
            "DemoCustomer",
            "DemoCustomerQuery",
            "CollectionsDemoUsage",
        )
        for module_name in names:
            if module_name in existing:
                project.delete_module(module_name)
        project.add_module("ROneCOne", runtime_source, kind=VBAModuleKind.other)
        project.add_module("DemoCustomer", customer_source, kind=VBAModuleKind.other)
        project.add_module(
            "CollectionsDemoUsage", demo_source, kind=VBAModuleKind.standard
        )
        workbook.save()

    with ExcelFile(workbook_path) as verification:
        expected = {
            "ROneCOne",
            "DemoCustomer",
            "CollectionsDemoUsage",
        }
        actual = set(verification.module_names())
        missing = expected - actual
        if missing:
            raise RuntimeError(f"Collections demo is missing modules: {sorted(missing)}")
        if verification.get_module("ROneCOne") != runtime_source:
            raise RuntimeError("Collections ROneCOne source did not round-trip")
        if verification.get_module("DemoCustomer") != customer_source:
            raise RuntimeError("Collections DemoCustomer source did not round-trip")
        if verification.get_module("CollectionsDemoUsage") != demo_source:
            raise RuntimeError("Collections demo usage source did not round-trip")


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser()
    parser.add_argument(
        "--kind",
        choices=("delegates", "collections", "all"),
        default="all",
    )
    parser.add_argument("--delegates-workbook", type=Path, default=DELEGATES_WORKBOOK)
    parser.add_argument("--collections-workbook", type=Path, default=COLLECTIONS_WORKBOOK)
    return parser.parse_args()


if __name__ == "__main__":
    arguments = parse_args()
    if arguments.kind in ("delegates", "all"):
        package_delegates(arguments.delegates_workbook)
        print(arguments.delegates_workbook)
    if arguments.kind in ("collections", "all"):
        package_collections(arguments.collections_workbook)
        print(arguments.collections_workbook)
