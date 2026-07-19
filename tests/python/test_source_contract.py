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
            ".ronecone.env",
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
        self.assertIn("!.ronecone.env.example", ignored)

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
            "OrderBy",
            "OrderByDescending",
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
            "Sorted",
            "SortedDescending",
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

    def test_contextual_linq_contract_accepts_member_names(self) -> None:
        contextual_signatures = (
            r"Public Function Where\(ByVal predicateOrMember As Variant\)",
            r"Public Function SelectItems\(\s*_\s*\n\s*ByVal selector As Variant",
            r"Public Function Map\(\s*_\s*\n\s*ByVal selector As Variant",
            r"Public Function OrderBy\(ByVal keySelector As Variant\)",
            r"Public Function OrderByDescending\(ByVal keySelector As Variant\)",
        )
        for signature in contextual_signatures:
            self.assertRegex(self.source, re.compile(signature, re.IGNORECASE))

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
            "createobject(",
            "getobject(",
            "shell(",
            "excel.application",
        )
        for term in forbidden:
            self.assertNotIn(term, lowered)

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
