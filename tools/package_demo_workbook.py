from __future__ import annotations

import argparse
from pathlib import Path

from pyopenvba import ExcelFile, VBAModuleKind

from build_test_workbook import prepare_class_source, read_vba


ROOT = Path(__file__).resolve().parents[1]
DELEGATES_WORKBOOK = ROOT / "demo" / "ROneCOne_Delegates_Demo.xlsm"
COLLECTIONS_WORKBOOK = ROOT / "demo" / "ROneCOne_Collections_Demo.xlsm"
EVENTS_WORKBOOK = ROOT / "demo" / "ROneCOne_Events_Demo.xlsm"
EXCEPTIONS_WORKBOOK = ROOT / "demo" / "ROneCOne_Exceptions_Demo.xlsm"
TASKS_WORKBOOK = ROOT / "demo" / "ROneCOne_Tasks_Demo.xlsm"
DATA_WORKBOOK = ROOT / "demo" / "ROneCOne_Data_Demo.xlsm"
HTTP_WORKBOOK = ROOT / "demo" / "ROneCOne_Http_Demo.xlsm"
JSON_WORKBOOK = ROOT / "demo" / "ROneCOne_Json_Demo.xlsm"
FILES_WORKBOOK = ROOT / "demo" / "ROneCOne_Files_Demo.xlsm"
PROCESS_WORKBOOK = ROOT / "demo" / "ROneCOne_Process_Demo.xlsm"
TEXT_WORKBOOK = ROOT / "demo" / "ROneCOne_Text_Demo.xlsm"


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


def package_capability(
    workbook_path: Path,
    module_name: str,
    source_path: Path,
    include_customer: bool = False,
) -> None:
    runtime_source = prepare_class_source(ROOT / "src" / "ROneCOne.cls")
    demo_source = read_vba(source_path)
    customer_source = (
        prepare_class_source(ROOT / "demo" / "vba" / "DemoCustomer.cls")
        if include_customer
        else None
    )

    with ExcelFile(workbook_path) as workbook:
        project = workbook.vba_project()
        existing = set(workbook.module_names())
        for name in ("Module1", "ROneCOne", "DemoCustomer", module_name):
            if name in existing:
                project.delete_module(name)
        project.add_module("ROneCOne", runtime_source, kind=VBAModuleKind.other)
        if customer_source is not None:
            project.add_module(
                "DemoCustomer", customer_source, kind=VBAModuleKind.other
            )
        project.add_module(module_name, demo_source, kind=VBAModuleKind.standard)
        workbook.save()

    with ExcelFile(workbook_path) as verification:
        expected = {"ROneCOne", module_name}
        if customer_source is not None:
            expected.add("DemoCustomer")
        missing = expected - set(verification.module_names())
        if missing:
            raise RuntimeError(f"Demo workbook is missing modules: {sorted(missing)}")
        if verification.get_module("ROneCOne") != runtime_source:
            raise RuntimeError("Capability demo runtime did not round-trip")
        if customer_source is not None:
            if verification.get_module("DemoCustomer") != customer_source:
                raise RuntimeError("Capability DemoCustomer did not round-trip")
        if verification.get_module(module_name) != demo_source:
            raise RuntimeError(f"{module_name} did not round-trip")


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser()
    parser.add_argument(
        "--kind",
        choices=(
            "delegates",
            "collections",
            "events",
            "exceptions",
            "tasks",
            "data",
            "http",
            "json",
            "files",
            "process",
            "text",
            "all",
        ),
        default="all",
    )
    parser.add_argument("--delegates-workbook", type=Path, default=DELEGATES_WORKBOOK)
    parser.add_argument("--collections-workbook", type=Path, default=COLLECTIONS_WORKBOOK)
    parser.add_argument("--events-workbook", type=Path, default=EVENTS_WORKBOOK)
    parser.add_argument("--exceptions-workbook", type=Path, default=EXCEPTIONS_WORKBOOK)
    parser.add_argument("--tasks-workbook", type=Path, default=TASKS_WORKBOOK)
    parser.add_argument("--data-workbook", type=Path, default=DATA_WORKBOOK)
    parser.add_argument("--http-workbook", type=Path, default=HTTP_WORKBOOK)
    parser.add_argument("--json-workbook", type=Path, default=JSON_WORKBOOK)
    parser.add_argument("--files-workbook", type=Path, default=FILES_WORKBOOK)
    parser.add_argument("--process-workbook", type=Path, default=PROCESS_WORKBOOK)
    parser.add_argument("--text-workbook", type=Path, default=TEXT_WORKBOOK)
    return parser.parse_args()


if __name__ == "__main__":
    arguments = parse_args()
    if arguments.kind in ("delegates", "all"):
        package_delegates(arguments.delegates_workbook)
        print(arguments.delegates_workbook)
    if arguments.kind in ("collections", "all"):
        package_collections(arguments.collections_workbook)
        print(arguments.collections_workbook)
    if arguments.kind in ("events", "all"):
        package_capability(
            arguments.events_workbook,
            "EventsDemoUsage",
            ROOT / "demo" / "vba" / "EventsDemoUsage.bas",
        )
        print(arguments.events_workbook)
    if arguments.kind in ("exceptions", "all"):
        package_capability(
            arguments.exceptions_workbook,
            "ExceptionsDemoUsage",
            ROOT / "demo" / "vba" / "ExceptionsDemoUsage.bas",
        )
        print(arguments.exceptions_workbook)
    if arguments.kind in ("tasks", "all"):
        package_capability(
            arguments.tasks_workbook,
            "TasksDemoUsage",
            ROOT / "demo" / "vba" / "TasksDemoUsage.bas",
        )
        print(arguments.tasks_workbook)
    if arguments.kind in ("data", "all"):
        package_capability(
            arguments.data_workbook,
            "DataDemoUsage",
            ROOT / "demo" / "vba" / "DataDemoUsage.bas",
        )
        print(arguments.data_workbook)
    if arguments.kind in ("http", "all"):
        package_capability(
            arguments.http_workbook,
            "HttpDemoUsage",
            ROOT / "demo" / "vba" / "HttpDemoUsage.bas",
        )
        print(arguments.http_workbook)
    if arguments.kind in ("json", "all"):
        package_capability(
            arguments.json_workbook,
            "JsonDemoUsage",
            ROOT / "demo" / "vba" / "JsonDemoUsage.bas",
            include_customer=True,
        )
        print(arguments.json_workbook)
    if arguments.kind in ("files", "all"):
        package_capability(
            arguments.files_workbook,
            "FilesDemoUsage",
            ROOT / "demo" / "vba" / "FilesDemoUsage.bas",
        )
        print(arguments.files_workbook)
    if arguments.kind in ("process", "all"):
        package_capability(
            arguments.process_workbook,
            "ProcessDemoUsage",
            ROOT / "demo" / "vba" / "ProcessDemoUsage.bas",
        )
        print(arguments.process_workbook)
    if arguments.kind in ("text", "all"):
        package_capability(
            arguments.text_workbook,
            "TextDemoUsage",
            ROOT / "demo" / "vba" / "TextDemoUsage.bas",
        )
        print(arguments.text_workbook)
