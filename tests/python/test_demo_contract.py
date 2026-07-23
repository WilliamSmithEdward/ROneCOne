from __future__ import annotations

import unittest
from pathlib import Path


ROOT = Path(__file__).resolve().parents[2]
DEMO_VBA = ROOT / "demo" / "vba"
COLLECTIONS_DEMO = DEMO_VBA / "CollectionsDemoUsage.bas"
DELEGATES_DEMO = DEMO_VBA / "DemoUsage.bas"
EVENTS_DEMO = DEMO_VBA / "EventsDemoUsage.bas"
EXCEPTIONS_DEMO = DEMO_VBA / "ExceptionsDemoUsage.bas"
TASKS_DEMO = DEMO_VBA / "TasksDemoUsage.bas"
DATA_DEMO = DEMO_VBA / "DataDemoUsage.bas"
HTTP_DEMO = DEMO_VBA / "HttpDemoUsage.bas"
CUSTOMER = DEMO_VBA / "DemoCustomer.cls"
COLLECTIONS_BUILDER = ROOT / "tools" / "build_collections_demo_workbook.cjs"
CAPABILITY_BUILDER = ROOT / "tools" / "build_capability_demo_workbooks.cjs"
PACKAGER = ROOT / "tools" / "package_demo_workbook.py"
RENDERER = ROOT / "tools" / "render_demo_workbook.ps1"
DEMO_RUNNER = ROOT / "tools" / "run_demo_workbook.ps1"
README = ROOT / "README.md"
USER_GUIDE = ROOT / "docs" / "user-guide" / "README.md"


class DemoContractTests(unittest.TestCase):
    def test_collection_demo_separates_primitive_and_user_class_examples(self) -> None:
        source = COLLECTIONS_DEMO.read_text(encoding="utf-8")

        self.assertIn("Private Sub WritePrimitiveCollectionExamples()", source)
        self.assertIn("Private Sub WriteUserClassLinqExamples()", source)
        self.assertIn("Private Sub RunCollectionBenchmark()", source)
        self.assertIn('Private Const USER_CLASS_SHEET As String = "User Class LINQ"', source)

    def test_delegate_demo_leads_with_a_practical_pricing_rule(self) -> None:
        source = DELEGATES_DEMO.read_text(encoding="utf-8")

        self.assertIn("Set price = ROneCOne.Var(vbDouble)", source)
        self.assertIn("Set applyDiscount = price.Multiply(0.9).AsFunc", source)
        self.assertIn("Set orderTotal = amount.Add(shipping).AsFunc", source)

    def test_delegate_demo_exercises_the_universal_surface(self) -> None:
        source = DELEGATES_DEMO.read_text(encoding="utf-8")
        builder = (ROOT / "tools" / "build_demo_workbook.cjs").read_text(
            encoding="utf-8"
        )

        required_syntax = (
            'ROneCOne.Func(worksheetFunctions, "Max")',
            'ROneCOne.Func("DemoUsage.CalculateOrderTotal")',
            "calculateTotal.DynamicInvoke",
            "ROneCOne.Combine(updateDashboard, writeAudit)",
            "ROneCOne.NativeAction",
            "ROneCOne.RefLong(orderNumber)",
            "calculateTotal.Signature",
        )
        for syntax in required_syntax:
            self.assertIn(syntax, source)
        self.assertNotIn("FromMethod", source)
        self.assertIn("DynamicInvoke", builder)
        self.assertIn("Notify two features", builder)
        self.assertIn("Update the original number", builder)
        self.assertNotIn("FromMethod", builder)

    def test_delegate_demo_uses_execute_and_inline_byref_sugar(self) -> None:
        source = DELEGATES_DEMO.read_text(encoding="utf-8")

        self.assertIn('notify.Execute "Order 1042 approved"', source)
        self.assertIn("increment.Execute ROneCOne.RefLong(orderNumber)", source)
        self.assertNotIn("Dim ignored As Variant", source)

    def test_user_class_model_exposes_demo_fields(self) -> None:
        source = CUSTOMER.read_text(encoding="utf-8")

        for property_name in ("CustomerName", "Age", "City"):
            self.assertIn(f"Public Property Get {property_name}", source)
            self.assertIn(f"Public Property Let {property_name}", source)

    def test_user_class_demo_uses_runtime_syntax_sugar(self) -> None:
        source = COLLECTIONS_DEMO.read_text(encoding="utf-8")

        required_syntax = (
            'customers.Where("Age").AtLeast(40)',
            '.Map("CustomerName", vbString)',
            ".Order",
            'customers.Condition("City").EqualTo("London")',
            '.OrderBy("City")',
            '.ThenByDescending("Age")',
            '.Where("CustomerName").StartsWith("G")',
            '.DistinctBy("City")',
            '"CollectionsDemoUsage.IsExperiencedCustomer").ToList',
            '.Where("Manager?.Age").AtLeast(40)',
            '.Where("City").IsIn(',
            "allowedCities.Contains(customers!City)",
            '.ContainsIgnoreCase("THER")',
            "customers.Count(",
            "customers.SingleItem(",
            "customers.WhereAny(",
            '"Reports", reportPredicate',
            "strings.Distinct(",
            ".Both(",
            ".Either(",
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
        self.assertIn(".Subscribe(updateDashboard)", source)
        self.assertIn('orderStatusChanged.Emit "Order 1042 shipped"', source)
        self.assertIn("orderStatusChanged.Unsubscribe(writeAudit)", source)

    def test_exception_demo_leads_with_structured_flow(self) -> None:
        source = EXCEPTIONS_DEMO.read_text(encoding="utf-8")

        self.assertIn("ROneCOne.Try(importWork)", source)
        self.assertIn(".Catch(INVALID_AMOUNT_ERROR, skipBadRow)", source)
        self.assertIn(".Finally(closeImportFile)", source)
        self.assertIn("importAttempt.Execute", source)
        self.assertIn("errorInfo.ErrorNumber = INVALID_AMOUNT_ERROR", source)
        self.assertNotIn("Err.Raise", source)
        self.assertNotIn("On Error Resume Next", source)

    def test_task_demo_leads_with_task_coordination(self) -> None:
        source = TASKS_DEMO.read_text(encoding="utf-8")

        self.assertIn("ROneCOne.Task.Run(forecastWork)", source)
        self.assertIn("ROneCOne.Task.Run(reorderWork)", source)
        self.assertIn("ROneCOne.Task.WhenAll(forecastTask, reorderTask)", source)
        self.assertIn("ROneCOne.Task.Run(countOpenOrders)", source)
        self.assertIn("Set results = allWork.Await", source)
        self.assertIn(".WaitAsync(100)", source)
        self.assertIn("ROneCOne.Task.YieldOnce", source)
        self.assertIn("allWork.ContinueWith(buildSummary)", source)
        self.assertIn("ROneCOne.CancellationTokenSource", source)
        self.assertIn("ROneCOne.TaskCompletionSourceOf(vbLong)", source)
        self.assertNotIn("RunOnExcel", source)
        self.assertNotIn("WorkerThreadId", source)
        self.assertNotIn("FromResult", source)

    def test_public_guides_lead_with_practical_workflows(self) -> None:
        guide_root = ROOT / "docs" / "user-guide"
        tasks = (guide_root / "tasks-and-async.md").read_text(encoding="utf-8")
        delegates = (guide_root / "delegates-and-expressions.md").read_text(
            encoding="utf-8"
        )
        events = (guide_root / "events-and-exceptions.md").read_text(
            encoding="utf-8"
        )

        self.assertIn("CountOpenOrders", tasks)
        self.assertIn("Task.Run(", tasks)
        self.assertLess(tasks.index("Task.Run("), tasks.index("Task.FromResult"))
        self.assertIn("applyDiscount", delegates)
        self.assertIn("Order 1042 approved", delegates)
        self.assertIn("Order 1042 shipped", events)
        self.assertIn("ImportSales", events)

    def test_http_demo_leads_with_awaitable_requests(self) -> None:
        source = HTTP_DEMO.read_text(encoding="utf-8")
        builder = CAPABILITY_BUILDER.read_text(encoding="utf-8")

        self.assertIn("ROneCOne.HttpClient()", source)
        self.assertIn(
            'client.BaseAddress = "https://pokeapi.co/api/v2/"', source
        )
        self.assertIn(
            'client.GetAsync("pokemon/pikachu").Await', source
        )
        self.assertIn("client.GetStringAsync(", source)
        self.assertIn("ROneCOne.Task.WhenAll(", source)
        self.assertIn("ROneCOne.HttpRequestError", source)
        self.assertIn("source.Token", source)
        # The failure example catches at the await site, mirroring C#'s
        # try / await / catch; a named-procedure raise cannot cross the
        # Application.Run boundary.
        self.assertIn("On Error Resume Next", source)
        self.assertNotIn("Application.Run", source)
        self.assertIn('"http"', builder)
        self.assertIn("ROneCOne_Http_Demo.xlsx", builder)
        self.assertIn("pokeapi.co", builder)

    def test_data_demo_leads_with_typed_data_and_provider_sugar(self) -> None:
        source = DATA_DEMO.read_text(encoding="utf-8")

        self.assertIn('table.Column("Id", vbLong).AutoNumber', source)
        self.assertIn(".AsPrimaryKey", source)
        self.assertIn('table.Row("Ada", 90, ROneCOne.DBNull).Add', source)
        self.assertIn("ROneCOne.DataView(table)", source)
        self.assertIn("ROneCOne.DataRelation", source)
        self.assertIn("ROneCOne.DbDataAdapter(command)", source)
        self.assertIn("ExecuteScalarAsync", source)
        self.assertIn("connection.AsyncMode", source)

    def test_capability_builder_and_packager_ship_separate_workbooks(self) -> None:
        builder = CAPABILITY_BUILDER.read_text(encoding="utf-8")
        packager = PACKAGER.read_text(encoding="utf-8")

        for name in (
            "ROneCOne_Events_Demo",
            "ROneCOne_Exceptions_Demo",
            "ROneCOne_Tasks_Demo",
            "ROneCOne_Data_Demo",
        ):
            self.assertIn(name, builder)
            self.assertIn(name, packager)

    def test_collections_workbook_includes_user_class_linq_sheet(self) -> None:
        source = COLLECTIONS_BUILDER.read_text(encoding="utf-8")

        self.assertIn('workbook.worksheets.add("User Class LINQ")', source)
        self.assertIn("Include customers added later", source)
        self.assertIn("Sort by city, then age", source)
        self.assertIn("Ask yes-or-no questions", source)
        self.assertIn("Calculate the average age", source)
        self.assertIn('.Where("Age").AtLeast(40)', source)
        self.assertIn('.Map("CustomerName", vbString)', source)
        self.assertIn('.OrderBy("City")', source)
        self.assertIn('.ThenByDescending("Age")', source)
        self.assertIn('.DistinctBy("City")', source)

    def test_renderer_has_a_non_cim_excel_ownership_fallback(self) -> None:
        source = RENDERER.read_text(encoding="utf-8")

        self.assertIn("Get-Process EXCEL", source)
        self.assertIn("CreationTimeUtc", source)
        self.assertIn("lock-matched Excel process", source)

    def test_demo_runner_fails_fast_after_dismissing_a_blocking_popup(self) -> None:
        source = DEMO_RUNNER.read_text(encoding="utf-8")

        self.assertIn("Get-BlockingDemoDialog", source)
        self.assertIn('dismissal_action -ne "none"', source)
        self.assertIn("Stop-OwnedDemoProcesses", source)
        self.assertIn("A blocking Excel or VBE dialog was observed", source)

    def test_demo_modules_are_ascii_explicit_and_readable(self) -> None:
        for path in (
            COLLECTIONS_DEMO,
            DELEGATES_DEMO,
            EVENTS_DEMO,
            EXCEPTIONS_DEMO,
            TASKS_DEMO,
            DATA_DEMO,
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
            *sorted((ROOT / "docs").rglob("*.md")),
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

    def test_readme_is_a_concise_code_led_product_surface(self) -> None:
        source = README.read_text(encoding="utf-8")

        self.assertEqual(2, source.count("```"), "README carries exactly one code sample")
        self.assertIn("```vba", source)
        self.assertNotIn("## Development", source)
        self.assertNotIn("## Quality gates", source)
        self.assertIn("Give VBA the love it deserves", source)
        self.assertIn("docs/user-guide/", source)
        self.assertIn("releases/latest", source)
        self.assertLessEqual(len(source.split()), 700)

    def test_docs_index_orients_every_documentation_surface(self) -> None:
        index = (ROOT / "docs" / "README.md").read_text(encoding="utf-8")

        self.assertIn("user-guide/README.md", index)
        for page_name in (
            "architecture.md",
            "collections.md",
            "delegates.md",
            "events.md",
            "exceptions.md",
            "tasks.md",
            "data.md",
            "development.md",
        ):
            self.assertIn(f"]({page_name})", index)

    def test_user_guide_is_indexed_and_code_led(self) -> None:
        guide_root = ROOT / "docs" / "user-guide"
        expected_pages = (
            "getting-started.md",
            "collections-and-linq.md",
            "delegates-and-expressions.md",
            "events-and-exceptions.md",
            "tasks-and-async.md",
            "data-and-providers.md",
            "reference.md",
        )

        index = USER_GUIDE.read_text(encoding="utf-8")
        self.assertIn("# ROneCOne user guide", index)
        for page_name in expected_pages:
            page = guide_root / page_name
            self.assertTrue(page.is_file(), page_name)
            self.assertIn(f"]({page_name})", index)
            self.assertIn("```vba", page.read_text(encoding="utf-8"), page_name)


if __name__ == "__main__":
    unittest.main()
