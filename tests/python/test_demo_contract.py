from __future__ import annotations

import unittest
from pathlib import Path


ROOT = Path(__file__).resolve().parents[2]
DEMO_VBA = ROOT / "demo" / "vba"
COLLECTIONS_DEMO = DEMO_VBA / "CollectionsDemoUsage.bas"
DELEGATES_DEMO = DEMO_VBA / "DemoUsage.bas"
EVENTS_DEMO = DEMO_VBA / "EventsDemoUsage.bas"
EXCEPTIONS_DEMO = DEMO_VBA / "ExceptionsDemoUsage.bas"
CUSTOMER = DEMO_VBA / "DemoCustomer.cls"
COLLECTIONS_BUILDER = ROOT / "tools" / "build_collections_demo_workbook.cjs"
CAPABILITY_BUILDER = ROOT / "tools" / "build_capability_demo_workbooks.cjs"
PACKAGER = ROOT / "tools" / "package_demo_workbook.py"
RENDERER = ROOT / "tools" / "render_demo_workbook.ps1"
README = ROOT / "README.md"


class DemoContractTests(unittest.TestCase):
    def test_collection_demo_separates_primitive_and_user_class_examples(self) -> None:
        source = COLLECTIONS_DEMO.read_text(encoding="utf-8")

        self.assertIn("Private Sub WritePrimitiveCollectionExamples()", source)
        self.assertIn("Private Sub WriteUserClassLinqExamples()", source)
        self.assertIn("Private Sub RunCollectionBenchmark()", source)
        self.assertIn('Private Const USER_CLASS_SHEET As String = "User Class LINQ"', source)

    def test_delegate_demo_leads_with_inferred_func_syntax(self) -> None:
        source = DELEGATES_DEMO.read_text(encoding="utf-8")

        self.assertIn("Set x = ROneCOne.Var(vbLong)", source)
        self.assertIn("Set square = x.Multiply(x).AsFunc", source)
        self.assertIn("Set addValues = x.Add(y).AsFunc", source)

    def test_delegate_demo_exercises_the_universal_surface(self) -> None:
        source = DELEGATES_DEMO.read_text(encoding="utf-8")
        builder = (ROOT / "tools" / "build_demo_workbook.cjs").read_text(
            encoding="utf-8"
        )

        required_syntax = (
            'ROneCOne.Func(worksheetFunctions, "Max")',
            'ROneCOne.Func("DemoUsage.DemoAddValues")',
            "workbookAdd.DynamicInvoke",
            "ROneCOne.Combine(firstAction, secondAction)",
            "ROneCOne.NativeAction",
            "ROneCOne.RefLong(value)",
            "workbookAdd.Signature",
        )
        for syntax in required_syntax:
            self.assertIn(syntax, source)
        self.assertNotIn("FromMethod", source)
        self.assertIn("DynamicInvoke", builder)
        self.assertIn("Multicast Action", builder)
        self.assertIn("True ByRef", builder)
        self.assertNotIn("FromMethod", builder)

    def test_delegate_demo_uses_execute_and_inline_byref_sugar(self) -> None:
        source = DELEGATES_DEMO.read_text(encoding="utf-8")

        self.assertIn('combined.Execute "value"', source)
        self.assertIn("increment.Execute ROneCOne.RefLong(value)", source)
        self.assertNotIn("Dim ignored As Variant", source)

    def test_user_class_model_exposes_demo_fields(self) -> None:
        source = CUSTOMER.read_text(encoding="utf-8")

        for property_name in ("CustomerName", "Age", "City"):
            self.assertIn(f"Public Property Get {property_name}", source)
            self.assertIn(f"Public Property Let {property_name}", source)

    def test_user_class_demo_uses_runtime_syntax_sugar(self) -> None:
        source = COLLECTIONS_DEMO.read_text(encoding="utf-8")

        required_syntax = (
            "Set customer = customers.Element",
            'customers.Where(customer("Age").AtLeast(CLng(40)))',
            '.Map(customer("CustomerName"), vbString)',
            ".Sorted",
            'customers.Exists(customer("City").EqualTo("London"))',
        )
        for syntax in required_syntax:
            self.assertIn(syntax, source)
        self.assertNotIn("DemoCustomerQuery", source)
        self.assertNotIn("FromMethod", source)
        self.assertIn("ROneCOne.ListFrom", source)
        self.assertIn(".ForEach ROneCOne.Action", source)
        self.assertIn(".JoinText", source)

    def test_event_demo_leads_with_typed_fluent_events(self) -> None:
        source = EVENTS_DEMO.read_text(encoding="utf-8")

        self.assertIn("ROneCOne.EventOf(vbString)", source)
        self.assertIn(".Subscribe(firstHandler)", source)
        self.assertIn('changed.Emit "ready"', source)
        self.assertIn("changed.Unsubscribe(secondHandler)", source)

    def test_exception_demo_leads_with_structured_flow(self) -> None:
        source = EXCEPTIONS_DEMO.read_text(encoding="utf-8")

        self.assertIn("ROneCOne.Try(failingWork)", source)
        self.assertIn(".Catch(errorHandler)", source)
        self.assertIn(".Finally(cleanup)", source)
        self.assertIn("attempt.Execute", source)

    def test_capability_builder_and_packager_ship_separate_workbooks(self) -> None:
        builder = CAPABILITY_BUILDER.read_text(encoding="utf-8")
        packager = PACKAGER.read_text(encoding="utf-8")

        for name in ("ROneCOne_Events_Demo", "ROneCOne_Exceptions_Demo"):
            self.assertIn(name, builder)
            self.assertIn(name, packager)

    def test_collections_workbook_includes_user_class_linq_sheet(self) -> None:
        source = COLLECTIONS_BUILDER.read_text(encoding="utf-8")

        self.assertIn('workbook.worksheets.add("User Class LINQ")', source)
        self.assertIn("Deferred class Where", source)
        self.assertIn("Object ordering", source)
        self.assertIn("Quantifiers", source)
        self.assertIn("Aggregate projection", source)
        self.assertIn("customers.Element", source)
        self.assertIn('.Where(customer("Age").AtLeast(40))', source)
        self.assertIn('.Map(customer("CustomerName"), vbString)', source)

    def test_collections_packager_removes_obsolete_query_helper(self) -> None:
        source = PACKAGER.read_text(encoding="utf-8")

        self.assertIn('"DemoCustomerQuery"', source)
        self.assertNotIn("query_source", source)

    def test_renderer_has_a_non_cim_excel_ownership_fallback(self) -> None:
        source = RENDERER.read_text(encoding="utf-8")

        self.assertIn("Get-Process EXCEL", source)
        self.assertIn("CreationTimeUtc", source)
        self.assertIn("lock-matched Excel process", source)

    def test_demo_modules_are_ascii_explicit_and_readable(self) -> None:
        for path in (
            COLLECTIONS_DEMO,
            DELEGATES_DEMO,
            EVENTS_DEMO,
            EXCEPTIONS_DEMO,
            CUSTOMER,
        ):
            with self.subTest(path=path.name):
                source = path.read_text(encoding="utf-8")
                source.encode("ascii")
                self.assertIn("Option Explicit", source)
                long_lines = [
                    number
                    for number, line in enumerate(source.splitlines(), start=1)
                    if len(line) > 100
                ]
                self.assertEqual([], long_lines)

    def test_public_materials_do_not_describe_the_runtime_as_experimental(self) -> None:
        public_materials = [
            README,
            *sorted((ROOT / "docs").glob("*.md")),
            ROOT / "tools" / "build_demo_workbook.cjs",
            ROOT / "tools" / "build_collections_demo_workbook.cjs",
            CAPABILITY_BUILDER,
        ]
        for path in public_materials:
            with self.subTest(path=path.name):
                source = path.read_text(encoding="utf-8").lower()
                self.assertNotIn("experiment", source)
                self.assertNotIn("future runtime", source)
                self.assertNotIn("future scheduler", source)
                self.assertNotIn("long-term goal", source)


if __name__ == "__main__":
    unittest.main()
