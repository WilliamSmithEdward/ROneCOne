from __future__ import annotations

import re
import unittest
from pathlib import Path


ROOT = Path(__file__).resolve().parents[2]
SOURCE = ROOT / "src" / "ROneCOne.cls"
BUILD_TOOL = ROOT / "tools" / "build_test_workbook.py"
GITIGNORE = ROOT / ".gitignore"


class SourceContractTests(unittest.TestCase):
    @classmethod
    def setUpClass(cls) -> None:
        cls.source = SOURCE.read_text(encoding="utf-8") if SOURCE.exists() else ""

    def test_runtime_source_exists(self) -> None:
        self.assertTrue(SOURCE.is_file(), "src/ROneCOne.cls must be the shipped runtime")

    def test_repository_ignores_local_and_generated_state(self) -> None:
        ignored = GITIGNORE.read_text(encoding="utf-8").splitlines()
        required_rules = (
            ".venv/",
            "node_modules/",
            "tests/output/",
            "demo/.working/",
            "*.candidate.xls*",
            "*.jsonl",
            "~$*",
        )
        for rule in required_rules:
            self.assertIn(rule, ignored)

    def test_runtime_embeds_the_mit_license(self) -> None:
        self.assertIn("' MIT License", self.source)
        self.assertIn("' Copyright (c) 2026 William Smith", self.source)
        self.assertIn("' THE SOFTWARE IS PROVIDED \"AS IS\"", self.source)

    def test_runtime_is_one_predeclared_class(self) -> None:
        class_files = sorted((ROOT / "src").glob("*.cls")) if (ROOT / "src").exists() else []
        self.assertEqual([SOURCE], class_files)
        self.assertIn('Attribute VB_Name = "ROneCOne"', self.source)
        self.assertIn("Attribute VB_PredeclaredId = True", self.source)
        self.assertIn("Option Explicit", self.source)

    def test_binary_builder_strips_only_the_class_export_preamble(self) -> None:
        builder = BUILD_TOOL.read_text(encoding="utf-8")
        self.assertIn("def prepare_class_source", builder)
        self.assertIn('"VERSION 1.0 CLASS"', builder)
        self.assertIn("prepare_class_source(ROOT / \"src\" / \"ROneCOne.cls\")", builder)

    def test_delegate_public_contract_is_present(self) -> None:
        required_members = (
            "Parameter",
            "ParameterLike",
            "Var",
            "VarLike",
            "Value",
            "Lambda",
            "AsFunc",
            "Func",
            "Action",
            "Native",
            "Run",
            "Execute",
            "DynamicInvoke",
            "Takes",
            "Returns",
            "Combine",
            "GetInvocationList",
            "PipeTo",
            "Add",
            "Subtract",
            "Multiply",
            "Divide",
            "Modulo",
            "Negate",
            "Concat",
            "EqualTo",
            "NotEqualTo",
            "LessThan",
            "LessThanOrEqual",
            "GreaterThan",
            "GreaterThanOrEqual",
            "AndAlso",
            "OrElse",
            "NotExpression",
            "Arity",
            "InvocationCount",
            "Target",
            "MethodName",
            "IsAction",
            "Signature",
            "RefOf",
            "RefLong",
            "RefByte",
            "RefInteger",
            "RefLongLong",
            "RefSingle",
            "RefDouble",
            "RefCurrency",
            "RefDate",
            "RefBoolean",
            "ByRefInvocationError",
            "InvalidOperationError",
            "TypeMismatchError",
            "EventOf",
            "Subscribe",
            "Unsubscribe",
            "Emit",
            "HandlerCount",
            "Try",
            "Catch",
            "Finally",
            "Exception",
            "ErrorNumber",
            "Message",
        )
        for member in required_members:
            pattern = rf"Public\s+(?:Function|Sub|Property\s+Get)\s+{member}\b"
            self.assertRegex(self.source, re.compile(pattern, re.IGNORECASE), member)
        self.assertNotRegex(
            self.source,
            re.compile(r"Public\s+Function\s+FromMethod\b", re.IGNORECASE),
        )

    def test_generic_list_and_linq_contract_is_present(self) -> None:
        required_members = (
            "ListOf",
            "ListLike",
            "ListFrom",
            "AddRange",
            "Clear",
            "Contains",
            "Count",
            "IndexOf",
            "Insert",
            "Item",
            "Remove",
            "RemoveAt",
            "Where",
            "SelectItems",
            "Take",
            "Skip",
            "Distinct",
            "DistinctBy",
            "Order",
            "OrderDescending",
            "OrderBy",
            "OrderByDescending",
            "ThenBy",
            "ThenByDescending",
            "Append",
            "Prepend",
            "Reverse",
            "AnyItem",
            "All",
            "First",
            "Last",
            "Sum",
            "Average",
            "Min",
            "Max",
            "ToList",
            "ToArray",
            "Range",
            "Repeat",
            "GenericTypeName",
            "Element",
            "Member",
            "Map",
            "Exists",
            "ForEach",
            "JoinText",
            "AtLeast",
            "AtMost",
            "Between",
            "OneOf",
            "StartsWith",
            "EndsWith",
            "ContainsText",
            "IsNothing",
            "IsNotNothing",
            "IsNullOrEmpty",
            "MatchesPattern",
            "IsTrue",
            "IsFalse",
            "Condition",
            "Predicate",
            "WhereMethod",
            "MinBy",
            "MaxBy",
            "MemberAccessError",
        )
        for member in required_members:
            pattern = (
                rf"Public\s+(?:Function|Sub|Property\s+(?:Get|Let|Set))\s+"
                rf"\[?{member}\]?(?![A-Za-z0-9_])"
            )
            self.assertRegex(self.source, re.compile(pattern, re.IGNORECASE), member)

        for removed_member in ("Sorted", "SortedDescending"):
            pattern = (
                rf"Public\s+(?:Function|Sub|Property\s+(?:Get|Let|Set))\s+"
                rf"\[?{removed_member}\]?(?![A-Za-z0-9_])"
            )
            self.assertNotRegex(
                self.source,
                re.compile(pattern, re.IGNORECASE),
                removed_member,
            )

    def test_complete_generic_collection_family_contract_is_present(self) -> None:
        factories = (
            "DictionaryOf",
            "HashSetOf",
            "QueueOf",
            "StackOf",
            "LinkedListOf",
            "SortedListOf",
            "SortedDictionaryOf",
            "SortedSetOf",
            "OrderedDictionaryOf",
            "PriorityQueueOf",
            "ObservableCollectionOf",
            "ReadOnlyCollectionOf",
            "KeyedCollectionOf",
            "ConcurrentDictionaryOf",
            "ConcurrentQueueOf",
            "ConcurrentStackOf",
            "ConcurrentBagOf",
            "BlockingCollectionOf",
            "ImmutableArrayOf",
            "ImmutableListOf",
            "ImmutableDictionaryOf",
            "ImmutableHashSetOf",
            "ImmutableQueueOf",
            "ImmutableStackOf",
            "ImmutableSortedDictionaryOf",
            "ImmutableSortedSetOf",
        )
        members = (
            "ContainsKey",
            "TryAdd",
            "TryGetValue",
            "Keys",
            "Values",
            "Enqueue",
            "Dequeue",
            "TryDequeue",
            "Push",
            "Pop",
            "TryPop",
            "Peek",
            "TryPeek",
            "AddFirst",
            "AddLast",
            "RemoveFirst",
            "RemoveLast",
            "FirstNode",
            "LastNode",
            "MinValue",
            "MaxValue",
            "GetViewBetween",
            "CollectionChanged",
            "IsReadOnly",
            "ToBuilder",
            "ToImmutable",
            "EnsureCapacity",
            "TrimExcess",
            "Capacity",
        )
        for member in factories + members:
            pattern = (
                rf"Public\s+(?:Function|Sub|Property\s+(?:Get|Let|Set))\s+"
                rf"\[?{member}\]?(?![A-Za-z0-9_])"
            )
            self.assertRegex(self.source, re.compile(pattern, re.IGNORECASE), member)

        self.assertIn("ROLE_DICTIONARY", self.source)
        self.assertIn("ROLE_HASH_SET", self.source)
        self.assertIn("ROLE_IMMUTABLE", self.source)
        self.assertIn("ROLE_CONCURRENT", self.source)

    def test_tasks_data_and_provider_contract_is_present(self) -> None:
        required_members = (
            "Task",
            "RunOnExcel",
            "FromResult",
            "Delay",
            "WhenAll",
            "WhenAny",
            "Await",
            "Wait",
            "WaitAsync",
            "YieldOnce",
            "ContinueWith",
            "Status",
            "IsCompleted",
            "IsCanceled",
            "IsFaulted",
            "Result",
            "CancellationTokenSource",
            "Token",
            "Cancel",
            "IsCancellationRequested",
            "Register",
            "Dispose",
            "InnerExceptions",
            "Flatten",
            "Handle",
            "ExceptionType",
            "DataTable",
            "DataColumn",
            "DataSet",
            "DataRelation",
            "DataView",
            "NewRow",
            "AddColumn",
            "AddRow",
            "AddTable",
            "AddRelation",
            "Columns",
            "Rows",
            "Tables",
            "Relations",
            "PrimaryKey",
            "Find",
            "AcceptChanges",
            "RejectChanges",
            "RowState",
            "GetChildRows",
            "GetParentRow",
            "DbConnection",
            "DbCommand",
            "DbParameter",
            "DbDataAdapter",
            "Connect",
            "Disconnect",
            "ExecuteReader",
            "ExecuteNonQuery",
            "ExecuteScalar",
            "Fill",
            "Update",
            "BeginTransaction",
            "OpenAsync",
            "ExecuteReaderAsync",
            "ExecuteNonQueryAsync",
            "ExecuteScalarAsync",
            "FillAsync",
            "UpdateAsync",
            "GetOrdinal",
            "GetValues",
            "WithParameter",
            "WithTimeout",
            "FromColumn",
            "DBNull",
            "AsPrimaryKey",
            "Row",
            "Using",
            "SupportsNativeAsync",
            "AsyncMode",
            "UseTransaction",
            "ContinueUpdateOnError",
            "LastUpdateErrors",
        )
        for member in required_members:
            pattern = (
                rf"Public\s+(?:Function|Sub|Property\s+(?:Get|Let|Set))\s+"
                rf"\[?{member}\]?(?![A-Za-z0-9_])"
            )
            self.assertRegex(self.source, re.compile(pattern, re.IGNORECASE), member)

    def test_tasks_are_cooperative_only_with_no_native_execution(self) -> None:
        self.assertIn("Public Function RunOnExcel(", self.source)
        for removed_syntax in (
            "CreateThreadpoolWork",
            "SubmitThreadpoolWork",
            "VirtualAlloc",
            "VirtualProtect",
            "FlushInstructionCache",
            "NativeTaskKernelHex",
            "Public Function TaskRun(",
            "Public Property Get ExecutionMode(",
            "Public Property Get WorkerThreadId(",
            "Public Property Get CurrentThreadId(",
            "TASK_NATIVE_WORK",
        ):
            self.assertNotIn(removed_syntax, self.source)

    def test_unproven_parallel_query_surface_is_not_public(self) -> None:
        for syntax in (
            "Public Function AsParallel(",
            "Public Function WithDegreeOfParallelism(",
            "Public Property Get IsParallel(",
        ):
            self.assertNotIn(syntax, self.source)

    def test_task_roles_and_combinators_are_present(self) -> None:
        for role_name in (
            "ROLE_TASK",
            "ROLE_DATA_TABLE",
            "ROLE_DATA_ROW",
            "ROLE_DATA_SET",
            "ROLE_DB_CONNECTION",
        ):
            self.assertIn(role_name, self.source)

        self.assertRegex(
            self.source,
            re.compile(r"Public\s+Function\s+WhenAll\s*\(ParamArray", re.IGNORECASE),
        )
        self.assertRegex(
            self.source,
            re.compile(r"Public\s+Function\s+WhenAny\s*\(ParamArray", re.IGNORECASE),
        )

    def test_complete_linq_materialization_contract_is_present(self) -> None:
        members = (
            "TakeWhile",
            "SkipWhile",
            "TakeLast",
            "SkipLast",
            "ConcatSequence",
            "Union",
            "Intersect",
            "Except",
            "DefaultIfEmpty",
            "Chunk",
            "SelectMany",
            "ElementAt",
            "ElementAtOrDefault",
            "Aggregate",
            "ToDictionary",
            "ToHashSet",
            "ToLookup",
            "GroupBy",
            "Join",
            "GroupJoin",
            "Zip",
        )
        for member in members:
            pattern = (
                rf"Public\s+(?:Function|Sub|Property\s+(?:Get|Let|Set))\s+"
                rf"\[?{member}\]?(?![A-Za-z0-9_])"
            )
            self.assertRegex(self.source, re.compile(pattern, re.IGNORECASE), member)

    def test_contextual_linq_contract_accepts_member_names(self) -> None:
        contextual_signatures = (
            r"Public Function Where\(Optional ByVal predicateOrMember As Variant\)",
            r"Public Function SelectItems\(\s*_\s*\n\s*ByVal selector As Variant",
            r"Public Function Map\(\s*_\s*\n\s*ByVal selector As Variant",
            r"Public Function Order\(Optional ByVal comparer As Variant\)",
            r"Public Function OrderDescending\(Optional ByVal comparer As Variant\)",
            r"Public Function OrderBy\(\s*_\s*\n\s*ByVal keySelector As Variant",
            r"Public Function OrderByDescending\(\s*_\s*\n\s*"
            r"ByVal keySelector As Variant",
            r"Public Function ThenBy\(\s*_\s*\n\s*ByVal keySelector As Variant",
            r"Public Function ThenByDescending\(\s*_\s*\n\s*"
            r"ByVal keySelector As Variant",
        )
        for signature in contextual_signatures:
            self.assertRegex(self.source, re.compile(signature, re.IGNORECASE))

    def test_predicate_system_contract_is_present(self) -> None:
        required_members = (
            "Both",
            "Either",
            "Negated",
            "WhereNot",
            "IsIn",
            "NotIn",
            "ContainsMember",
            "EqualToIgnoreCase",
            "NotEqualToIgnoreCase",
            "StartsWithIgnoreCase",
            "EndsWithIgnoreCase",
            "ContainsIgnoreCase",
            "MatchesPatternIgnoreCase",
            "FirstOrDefault",
            "LastOrDefault",
            "SingleItem",
            "SingleOrDefault",
            "None",
            "AnyMatch",
            "AllMatch",
            "NoneMatch",
            "WhereAny",
            "WhereAll",
            "WhereNone",
            "SequenceEqual",
            "EqualityComparer",
            "Comparer",
            "Always",
            "Never",
            "Match",
            "NotMatch",
        )
        for member in required_members:
            pattern = rf"Public\s+Function\s+\[?{member}\]?(?![A-Za-z0-9_])"
            self.assertRegex(self.source, re.compile(pattern, re.IGNORECASE), member)

    def test_predicate_terminals_accept_optional_predicates(self) -> None:
        self.assertRegex(
            self.source,
            re.compile(
                r"Public Property Get Count\(Optional ByVal predicate As Variant\)",
                re.IGNORECASE,
            ),
        )

    def test_null_safe_paths_and_cached_member_plans_are_present(self) -> None:
        self.assertIn('Replace(memberPath, "?.", ".?")', self.source)
        self.assertIn("ConfigureMemberExpression", self.source)
        self.assertIn("InternalMethodName", self.source)

    def test_sequence_default_member_can_select_contextual_members(self) -> None:
        self.assertIn("Set result = Condition(", self.source)
        self.assertIn("Attribute Run.VB_UserMemId = 0", self.source)

    def test_collections_expose_vba_foreach_enumeration(self) -> None:
        self.assertRegex(
            self.source,
            re.compile(r"Attribute\s+NewEnum\.VB_UserMemId\s*=\s*-4", re.IGNORECASE),
        )

    def test_runtime_does_not_depend_on_vbide_or_external_processes(self) -> None:
        lowered = self.source.lower()
        forbidden = (
            "vbproject",
            "vbcomponents",
            "vbe.",
            "getobject(",
            "shell(",
            "excel.application",
        )
        for term in forbidden:
            self.assertNotIn(term, lowered)

        created_prog_ids = re.findall(
            r'CreateObject\("([^"]+)"\)', self.source, flags=re.IGNORECASE
        )
        self.assertEqual(
            {"adodb.connection", "adodb.command"},
            {value.lower() for value in created_prog_ids},
        )

    def test_delegate_run_is_the_default_member(self) -> None:
        self.assertIn("Public Function Run(ParamArray arguments() As Variant)", self.source)
        self.assertIn("Attribute Run.VB_UserMemId = 0", self.source)
        self.assertNotRegex(self.source, re.compile(r"^Public Function Invoke\b", re.MULTILINE))

    def test_source_is_ascii_and_uses_short_lines(self) -> None:
        self.source.encode("ascii")
        long_lines = [
            (number, len(line))
            for number, line in enumerate(self.source.splitlines(), start=1)
            if len(line) > 100
        ]
        self.assertEqual([], long_lines)

    def test_public_members_have_intellisense_descriptions(self) -> None:
        public_members = re.findall(
            r"^Public\s+(?:Function|Property\s+Get)\s+([A-Za-z][A-Za-z0-9_]*)\b",
            self.source,
            flags=re.IGNORECASE | re.MULTILINE,
        )
        missing = [
            name
            for name in public_members
            if f"Attribute {name}.VB_Description = " not in self.source
        ]
        self.assertEqual([], missing)


if __name__ == "__main__":
    unittest.main()
