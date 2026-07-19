from __future__ import annotations

import unittest
from pathlib import Path


ROOT = Path(__file__).resolve().parents[2]
DEMO_VBA = ROOT / "demo" / "vba"
COLLECTIONS_DEMO = DEMO_VBA / "CollectionsDemoUsage.bas"
DELEGATES_DEMO = DEMO_VBA / "DemoUsage.bas"
CUSTOMER = DEMO_VBA / "DemoCustomer.cls"
CUSTOMER_QUERY = DEMO_VBA / "DemoCustomerQuery.cls"
COLLECTIONS_BUILDER = ROOT / "tools" / "build_collections_demo_workbook.cjs"
PACKAGER = ROOT / "tools" / "package_demo_workbook.py"
README = ROOT / "README.md"


class DemoContractTests(unittest.TestCase):
    def test_collection_demo_separates_primitive_and_user_class_examples(self) -> None:
        source = COLLECTIONS_DEMO.read_text(encoding="utf-8")

        self.assertIn("Private Sub WritePrimitiveCollectionExamples()", source)
        self.assertIn("Private Sub WriteUserClassLinqExamples()", source)
        self.assertIn("Private Sub RunCollectionBenchmark()", source)
        self.assertIn('Private Const USER_CLASS_SHEET As String = "User Class LINQ"', source)

    def test_user_class_model_exposes_demo_fields(self) -> None:
        source = CUSTOMER.read_text(encoding="utf-8")

        for property_name in ("CustomerName", "Age", "City"):
            self.assertIn(f"Public Property Get {property_name}", source)
            self.assertIn(f"Public Property Let {property_name}", source)

    def test_user_class_query_exposes_predicates_and_selectors(self) -> None:
        self.assertTrue(CUSTOMER_QUERY.is_file())
        source = CUSTOMER_QUERY.read_text(encoding="utf-8")

        for member in (
            "MeetsMinimumAge",
            "IsInRequiredCity",
            "SelectName",
            "SelectAge",
        ):
            self.assertIn(f"Public Function {member}", source)

    def test_collections_workbook_includes_user_class_linq_sheet(self) -> None:
        source = COLLECTIONS_BUILDER.read_text(encoding="utf-8")

        self.assertIn('workbook.worksheets.add("User Class LINQ")', source)
        self.assertIn("Deferred class Where", source)
        self.assertIn("Object ordering", source)
        self.assertIn("Quantifiers", source)
        self.assertIn("Aggregate projection", source)

    def test_collections_packager_embeds_query_helper(self) -> None:
        source = PACKAGER.read_text(encoding="utf-8")

        self.assertIn('"DemoCustomerQuery.cls"', source)
        self.assertIn('"DemoCustomerQuery"', source)

    def test_demo_modules_are_ascii_explicit_and_readable(self) -> None:
        for path in (COLLECTIONS_DEMO, DELEGATES_DEMO, CUSTOMER, CUSTOMER_QUERY):
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
