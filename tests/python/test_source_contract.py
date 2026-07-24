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
            "ToDisplayString",
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
        self.assertIn("Friend Function ScheduleTaskWork(", self.source)
        self.assertIn('RaiseContractError ERROR_ARITY_MISMATCH, "Task.Run"', self.source)
        for removed_syntax in (
            "RunOnExcel",
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
            {
                "adodb.connection",
                "adodb.command",
                "scripting.filesystemobject",
            },
            {value.lower() for value in created_prog_ids},
        )
        # The const-based CreateObjects are the in-process WinHTTP client and
        # the ADODB stream used for encoded file transput.
        self.assertIn(
            'Private Const HTTP_PROG_ID As String = "WinHttp.WinHttpRequest.5.1"',
            self.source,
        )
        self.assertIn(
            'Private Const STREAM_PROG_ID As String = "ADODB.Stream"',
            self.source,
        )
        self.assertIn(
            'Private Const WSHELL_PROG_ID As String = "WScript.Shell"',
            self.source,
        )
        const_created = re.findall(
            r"CreateObject\((\w+)\)", self.source, flags=re.IGNORECASE
        )
        self.assertEqual(
            {"HTTP_PROG_ID", "STREAM_PROG_ID", "WSHELL_PROG_ID"},
            set(const_created),
        )

    def test_file_system_surface_is_present(self) -> None:
        for member in (
            "Public Property Get File()",
            "Public Property Get Directory()",
            "Public Property Get Path()",
            "Public Function ReadAllText(",
            "Public Sub WriteAllText(",
            "Public Sub AppendAllText(",
            "Public Function ReadAllLines(",
            "Public Sub WriteAllLines(",
            "Public Function ReadAllBytes(",
            "Public Sub WriteAllBytes(",
            "Public Sub Move(",
            "Public Sub CreateDirectory(",
            "Public Function GetFiles(",
            "Public Function GetDirectories(",
            "Public Function GetFileName(",
            "Public Function GetDirectoryName(",
            "Public Function GetExtension(",
            "Public Function GetFileNameWithoutExtension(",
            "Public Function ChangeExtension(",
            "Public Function GetFullPath(",
            "Public Function GetTempPath(",
            "Public Property Get IOError()",
        ):
            self.assertIn(member, self.source)

    def test_csv_surface_is_present(self) -> None:
        for member in (
            "Public Property Get Csv()",
            "Public Function ToCsv(",
            "Public Property Get CsvError()",
            'Err.Raise ERROR_CSV_PARSE, "ROneCOne.CsvException"',
        ):
            self.assertIn(member, self.source)

    def test_process_surface_is_present(self) -> None:
        for member in (
            "Public Property Get Process()",
            "Public Function RunAsync(",
            "Public Property Get ExitCode()",
            "Public Property Get StandardOutput()",
            "Public Property Get StandardError()",
            "Public Property Get ProcessError()",
        ):
            self.assertIn(member, self.source)

    def test_json_surface_is_present_and_runtime_native(self) -> None:
        for member in (
            "Public Property Get Json()",
            "Public Function Serialize(",
            "Public Function Deserialize(",
            "Public Function DeserializeTable(",
            "Public Function DeserializeInto(",
            "Public Function DeserializeObjects(",
            "Public Function DataTableFromObjects(",
            "Public Function ToObjects(",
            "Public Function ToJson(",
            "Public Property Get JsonError()",
        ):
            self.assertIn(member, self.source)
        # The parser reads character codes from a byte snapshot, returns
        # escape-free strings with a single copy, accumulates short integers
        # inline, and writes numbers with an invariant decimal separator.
        for mechanism in (
            "Private Type JsonReader",
            "Private Type JsonTextBuilder",
            "reader.bytes = jsonText",
            "Private Function JsonReadNumber(",
            "Private Function JsonReadString(",
            "Private Function JsonNumberText(",
            "mJsonDecimalSeparator",
            '"ROneCOne.JsonException"',
        ):
            self.assertIn(mechanism, self.source)
        # The model is runtime-native: objects deserialize into ordered
        # dictionaries and arrays into Variant lists.
        self.assertIn(
            "ROneCOne.OrderedDictionaryOf(vbString, vbVariant)", self.source
        )
        self.assertIn("ROneCOne.ListOf(vbVariant)", self.source)

    def test_http_client_surface_is_present_and_awaitable(self) -> None:
        for member in (
            "Public Function HttpClient()",
            "Public Function GetAsync(",
            "Public Function GetStringAsync(",
            "Public Function GetByteArrayAsync(",
            "Public Function PostAsync(",
            "Public Function PutAsync(",
            "Public Function PatchAsync(",
            "Public Function DeleteAsync(",
            "Public Function SendAsync(",
            "Public Function DefaultRequestHeader(",
            "Public Property Get BaseAddress()",
            "Public Property Get Timeout()",
            "Public Property Get StatusCode()",
            "Public Property Get ReasonPhrase()",
            "Public Property Get IsSuccessStatusCode()",
            "Public Function EnsureSuccessStatusCode()",
            "Public Property Get Content()",
            "Public Function Header(",
            "Public Property Get AllHeaders()",
            "Public Property Get HttpRequestError()",
        ):
            self.assertIn(member, self.source)
        # Requests open asynchronously and complete on the cooperative
        # scheduler by polling WaitForResponse(0); cancellation aborts the
        # in-flight transport.
        self.assertIn("Private Const TASK_HTTP_SEND As Long", self.source)
        self.assertIn(".WaitForResponse(0)", self.source)
        self.assertIn("Case TASK_HTTP_SEND", self.source)
        self.assertIn("InternalHttpAbort", self.source)
        self.assertIn('"HttpRequestException"', self.source)

    def test_delegate_run_is_the_default_member(self) -> None:
        self.assertIn("Public Function Run(ParamArray arguments() As Variant)", self.source)
        self.assertIn("Attribute Run.VB_UserMemId = 0", self.source)
        self.assertNotRegex(self.source, re.compile(r"^Public Function Invoke\b", re.MULTILINE))

    def test_element_storage_is_array_backed(self) -> None:
        self.assertIn("Private mItems() As ROneCOne", self.source)
        self.assertIn("Private mItemsCount As Long", self.source)
        self.assertIn("Private mKeys() As ROneCOne", self.source)
        self.assertIn("Private mPriorities() As ROneCOne", self.source)
        self.assertIn("Private mOriginalItems() As ROneCOne", self.source)
        for helper in (
            "Private Sub ArrAppend(",
            "Private Sub ArrInsert(",
            "Private Sub ArrRemoveAt(",
            "Private Sub ArrReplaceAt(",
            "Private Sub ArrReset(",
            "Private Function ArrSnapshot(",
        ):
            self.assertIn(helper, self.source)
        self.assertNotIn("Private mItems As Collection", self.source)
        self.assertNotIn("Private mKeys As Collection", self.source)

    def test_keyed_mutation_maintains_the_hash_index_in_place(self) -> None:
        for member in (
            "Private mHashDirty As Boolean",
            "Private Sub MarkHashIndexDirty()",
            "Private Sub EnsureHashIndexCurrent()",
            "Private Sub RefreshHashSlotValue(",
            "Private Sub RemoveHashIndexEntry(",
            "Private Sub DeleteHashSlot(",
            "Private Sub ShiftHashIndexesAbove(",
        ):
            self.assertIn(member, self.source)

        def sub_body(name: str) -> str:
            match = re.search(
                rf"^(?:Friend|Private|Public) Sub {name}\(.*?^End Sub",
                self.source,
                flags=re.MULTILINE | re.DOTALL,
            )
            self.assertIsNotNone(match, name)
            assert match is not None
            return match.group(0)

        def function_body(name: str) -> str:
            match = re.search(
                rf"^(?:Friend|Private|Public) Function {name}\(.*?^End Function",
                self.source,
                flags=re.MULTILINE | re.DOTALL,
            )
            self.assertIsNotNone(match, name)
            assert match is not None
            return match.group(0)

        # A value replacement refreshes one slot and a removal repairs its
        # probe cluster in place; neither may rescan the whole index.
        replace_body = sub_body("ReplaceCollectionItem")
        self.assertIn("RefreshHashSlotValue", replace_body)
        self.assertNotIn("RebuildHashIndex", replace_body)
        remove_body = sub_body("RemoveCollectionAt")
        self.assertIn("RemoveHashIndexEntry", remove_body)
        self.assertNotIn("RebuildHashIndex", remove_body)
        # Every probe or insert flushes a deferred rebuild first, so bulk
        # removal paths may mark the index dirty exactly once.
        self.assertIn("EnsureHashIndexCurrent", sub_body("PrepareHashInsert"))
        self.assertIn("EnsureHashIndexCurrent", function_body("FindHashSlot"))
        self.assertIn(
            "EnsureHashIndexCurrent", function_body("FindHashSlotByKey")
        )

    def test_positional_access_and_snapshots_are_version_cached(self) -> None:
        def sub_body(name: str) -> str:
            match = re.search(
                rf"^(?:Friend|Private|Public) Sub {name}\(.*?^End Sub",
                self.source,
                flags=re.MULTILINE | re.DOTALL,
            )
            self.assertIsNotNone(match, name)
            assert match is not None
            return match.group(0)

        def function_body(name: str) -> str:
            match = re.search(
                rf"^(?:Friend|Private|Public) Function {name}\(.*?^End Function",
                self.source,
                flags=re.MULTILINE | re.DOTALL,
            )
            self.assertIsNotNone(match, name)
            assert match is not None
            return match.group(0)

        # Lists enumerate through the lazy version-checked mirror; no eager
        # mirror maintenance remains anywhere in the runtime.
        self.assertNotIn("AddUnwrappedValue mEnumerationValues", self.source)
        self.assertNotIn("mEnumerationValues.Remove", self.source)
        self.assertIn(
            "If mRole = ROLE_LIST Or mRole = ROLE_QUERY Or _", self.source
        )

        # Data snapshot properties cache against the structural version, and
        # field edits advance only the data version so those caches survive.
        for member in (
            "Private mDataVersion As Long",
            "Private mRowsSnapshot As ROneCOne",
            "Private mColumnsSnapshot As ROneCOne",
            "Private mTablesSnapshot As ROneCOne",
            "Private mRelationsSnapshot As ROneCOne",
            "Private mPrimaryKeySnapshot As ROneCOne",
        ):
            self.assertIn(member, self.source)
        self.assertIn("mRowsSnapshotVersion <> mVersion", self.source)
        touch_body = sub_body("TouchDataVersion")
        self.assertIn("mDataVersion = mDataVersion + 1", touch_body)
        self.assertNotIn("mVersion = mVersion + 1", touch_body)

        # Materialized generic collections index their arrays directly;
        # only deferred queries and live views materialize per read.
        wrapped_body = function_body("WrappedItem")
        self.assertIn("IsGenericCollectionRole(mRole)", wrapped_body)
        self.assertIn("CreateDictionaryEntry", wrapped_body)

        # Replacing a view's filter or sort advances the view version, and a
        # late column backfills every existing row with its default cell.
        self.assertIn("mVersion = mVersion + 1", function_body("WithFilter"))
        self.assertIn("mVersion = mVersion + 1", function_body("WithSort"))
        self.assertIn("InternalAppendRowCell", sub_body("AddColumn"))

    def test_constraint_indexes_are_incremental_and_lazy(self) -> None:
        def sub_body(name: str) -> str:
            match = re.search(
                rf"^(?:Friend|Private|Public) Sub {name}\(.*?^End Sub",
                self.source,
                flags=re.MULTILINE | re.DOTALL,
            )
            self.assertIsNotNone(match, name)
            assert match is not None
            return match.group(0)

        for member in (
            "Private mDataIndexDirty As Boolean",
            "Private mUniqueIndexes As Collection",
            "Private Sub EnsureDataIndexesCurrent()",
            "Private Sub IndexDataRow(",
            "Friend Sub InternalMarkDataIndexDirty()",
            "Friend Sub InternalDataRowKeyEdited(",
        ):
            self.assertIn(member, self.source)

        # A new row slots into the constraint indexes incrementally, a
        # single-field edit validates only its own column with index probes
        # instead of row scans, and key edits defer one rebuild.
        add_row = sub_body("AddRow")
        self.assertIn("IndexDataRow row", add_row)
        self.assertNotIn("RebuildPrimaryKeyIndex", add_row)
        validate = sub_body("ValidateDataRowConstraints")
        self.assertIn("EnsureDataIndexesCurrent", validate)
        self.assertNotIn("ArrSnapshot(mItems", validate)
        set_item = sub_body("SetDataRowItem")
        self.assertIn("InternalDataRowKeyEdited", set_item)
        self.assertNotIn("RebuildPrimaryKeyIndex", set_item)

    def test_worksheet_range_bridge_is_present(self) -> None:
        for member in (
            "DataTableFromRange",
            "ListFromRange",
            "LoadFromRange",
            "ToRange",
        ):
            pattern = rf"Public\s+Function\s+{member}\b"
            self.assertRegex(self.source, re.compile(pattern, re.IGNORECASE), member)
        # Bulk single-call I/O, never a per-cell loop.
        self.assertIn("raw = source.Value", self.source)
        self.assertIn(".Value = output", self.source)

    def test_lossless_numeric_widening_is_present_and_wired(self) -> None:
        self.assertIn("Private Function IsLosslessWidening(", self.source)
        self.assertIn("Private Function WidenScalarToward(", self.source)
        self.assertIn("Private Function CoerceScalarToType(", self.source)
        # Widening must be applied at every admission point, not a subset:
        # element, key, parameter binding, delegate arguments, delegate result,
        # DataColumn value, primary key, progress, and completion.
        self.assertGreaterEqual(self.source.count("WidenScalarToward("), 12)

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
