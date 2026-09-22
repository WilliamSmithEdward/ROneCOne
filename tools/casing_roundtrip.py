"""Build the casing round-trip workbook, and compare a VBE export of it with
the sources that went in (issue #6).

The workbook holds the runtime and a host module written the way a user
writes one: Excel members, named arguments, an event-style Target parameter,
and ROneCOne's own members, with nothing declared that shares their names.
tools/run_casing_roundtrip.ps1 opens it in Excel, exports every module through
the VBE, and then runs `compare`, which requires every code token back
exactly as written.

Usage:
    casing_roundtrip.py build <workbook> [--runtime <ROneCOne.cls>]
    casing_roundtrip.py compare <export-dir> [--runtime <ROneCOne.cls>]
"""

from __future__ import annotations

import argparse
import json
import sys
from pathlib import Path

from pyopenvba import ExcelFile, VBAModuleKind

from build_test_workbook import prepare_class_source

ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(ROOT / "tests" / "python"))
from test_casing import code_tokens  # noqa: E402  one token definition for both checks

HOST_PROBE = """\
Attribute VB_Name = "HostProbe"
Option Explicit

' Host code as a ROneCOne user writes it. It declares nothing that shares a
' name with Excel or ROneCOne, so every token must come back as written.
Public Sub BuildHostProbe()
    Dim ws As Worksheet
    Dim squares As ROneCOne
    Dim ages As ROneCOne
    Dim pending As ROneCOne

    Set ws = ActiveSheet
    ws.Range("A1").Value = 1
    ws.Range("A2").Formula = "=A1*2"
    Debug.Print ws.Range("A1").Text, ws.Cells(1, 1).Value, ws.Name, ws.Index
    Debug.Print ThisWorkbook.Names.Count, ws.Columns.Count, ws.Rows.Count
    Debug.Print ws.Cells.Item(RowIndex:=1, ColumnIndex:=2).Row, ws.Range("B2").Column
    ws.Hyperlinks.Add Anchor:=ws.Range("B1"), Address:=""
    ws.Sort.SortFields.Add Key:=ws.Range("A1")
    Debug.Print Application.Cursor, ws.Range("A1").Width, ws.Range("A1").Font.Size
    Debug.Print Err.Source, Err.Description, Err.Number, ws.Previous Is Nothing
    Debug.Print ws.Shapes.Item(1).Line.Visible, ws.Shapes.Count

    Set squares = ROneCOne.ListOf(vbLong, 1, 4, 9)
    squares.Add 16
    Debug.Print squares.Count, squares.Item(0), squares.First, squares.Contains(4)
    Debug.Print squares.IndexOf(9), squares.ToList.Count
    Set ages = ROneCOne.DictionaryOf(vbString, vbLong)
    ages.Add "Ada", 36
    Debug.Print ages.ContainsKey("Ada"), ages.Keys.Count, ages.Values.Count
    Debug.Print ROneCOne.Json.Serialize(squares), ROneCOne.Json.Deserialize("[1]").Count
    Set pending = ROneCOne.Task.FromResult(42)
    Debug.Print pending.Result, pending.Await
End Sub

Private Sub HandleChange(ByVal Target As Range, Cancel As Boolean)
    Debug.Print Target.Address, Target.Row, Target.Column, Target.Value, Cancel
End Sub
"""


def modules(runtime: Path) -> dict[str, tuple[str, VBAModuleKind]]:
    return {
        "ROneCOne": (prepare_class_source(runtime), VBAModuleKind.other),
        "HostProbe": (HOST_PROBE.replace("\n", "\r\n"), VBAModuleKind.standard),
    }


def build(workbook: Path, runtime: Path) -> None:
    workbook.parent.mkdir(parents=True, exist_ok=True)
    workbook.unlink(missing_ok=True)
    with ExcelFile.create_new(workbook) as book:
        project = book.vba_project()
        if "Module1" in book.module_names():
            project.delete_module("Module1")
        for name, (source, kind) in modules(runtime).items():
            project.add_module(name, source, kind=kind)
        book.save()
    with ExcelFile(workbook) as check:
        for name, (source, _) in modules(runtime).items():
            if check.get_module(name) != source:
                raise RuntimeError(f"{name} did not round-trip byte-for-byte")
    print(workbook)


def compare(export_dir: Path, runtime: Path) -> int:
    changed = []
    for name, (source, _) in modules(runtime).items():
        exported = next(export_dir.glob(f"{name}.*"), None)
        if exported is None:
            changed.append(f"{name}: not exported")
            continue
        before = code_tokens(source)
        after = code_tokens(exported.read_text(encoding="utf-8", errors="replace"))
        if len(before) != len(after):
            changed.append(f"{name}: token count {len(before)} -> {len(after)}")
            continue
        pairs = sorted({(a, b) for a, b in zip(before, after) if a != b})
        changed.extend(f"{name}: {a} -> {b}" for a, b in pairs)
    print(json.dumps({"status": "FAIL" if changed else "PASS", "recased": changed}))
    return 1 if changed else 0


def main() -> int:
    parser = argparse.ArgumentParser()
    parser.add_argument("action", choices=("build", "compare"))
    parser.add_argument("path", type=Path)
    parser.add_argument("--runtime", type=Path, default=ROOT / "src" / "ROneCOne.cls")
    arguments = parser.parse_args()
    if arguments.action == "build":
        build(arguments.path, arguments.runtime)
        return 0
    return compare(arguments.path, arguments.runtime)


if __name__ == "__main__":
    sys.exit(main())
