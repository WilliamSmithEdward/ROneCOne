"""Security scan rules. tools/security_scan.py holds olevba's findings and MacroRaptor's
flags for every shipped file to tools/security_baseline.json, and the Security workflow
runs it over the real files, where every result matches. These tests keep each failure
the scan exists for failing, on small modules written for the purpose.
"""

from __future__ import annotations

import contextlib
import importlib.util
import io
import tempfile
import unittest
from pathlib import Path


ROOT = Path(__file__).resolve().parents[2]
SPEC = importlib.util.spec_from_file_location(
    "security_scan", ROOT / "tools" / "security_scan.py"
)
security_scan = importlib.util.module_from_spec(SPEC)
SPEC.loader.exec_module(security_scan)

# Writes and runs things only when called, as the runtime does: W and X, never A.
CALLED_ONLY = (
    'Attribute VB_Name = "Report"\r\n'
    "Public Sub WriteReport(ByVal Path As String)\r\n"
    "    Kill Path\r\n"
    '    Shell "cmd.exe /c dir"\r\n'
    "End Sub\r\n"
)
# Runs a command when the workbook opens: A with X, which MacroRaptor calls suspicious.
AUTO_RUN = (
    'Attribute VB_Name = "ThisWorkbook"\r\n'
    "Private Sub Workbook_Open()\r\n"
    '    Shell "cmd.exe /c dir"\r\n'
    "End Sub\r\n"
)
# Runs when the workbook opens but only shows a message: A alone, which MacroRaptor
# passes. The scan still refuses it, because ROneCOne runs nothing on its own.
QUIET_AUTO_RUN = (
    'Attribute VB_Name = "ThisWorkbook"\r\n'
    "Private Sub Workbook_Open()\r\n"
    '    MsgBox "Ready"\r\n'
    "End Sub\r\n"
)


class SecurityScanTests(unittest.TestCase):
    @classmethod
    def setUpClass(cls) -> None:
        cls.folder = tempfile.TemporaryDirectory()
        cls.paths = {}
        cls.results = {}
        for name, source in (
            ("called", CALLED_ONLY),
            ("auto_run", AUTO_RUN),
            ("quiet", QUIET_AUTO_RUN),
        ):
            path = Path(cls.folder.name) / f"{name}.bas"
            path.write_bytes(source.encode("ascii"))
            cls.paths[name] = path
            cls.results[name] = security_scan.scan(path)

    @classmethod
    def tearDownClass(cls) -> None:
        cls.folder.cleanup()

    def own_baseline(self, name: str) -> tuple[set[tuple[str, str]], str]:
        result = self.results[name]
        return set(result["findings"]), result["mraptor_flags"]

    def test_code_that_acts_only_when_called_matches_its_own_baseline(self) -> None:
        result = self.results["called"]

        self.assertEqual(security_scan.verdict(result), "Macro OK (-WX)")
        self.assertIn(("Suspicious", "Shell"), result["findings"])
        self.assertEqual(security_scan.auto_run(result), [])
        self.assertEqual(
            security_scan.differences("called", result, *self.own_baseline("called")), []
        )

    def test_a_finding_that_appears_or_disappears_fails(self) -> None:
        result = self.results["called"]
        findings, flags = self.own_baseline("called")
        reviewed = (findings - {("Suspicious", "Shell")}) | {("IOC", "other.dll")}

        self.assertEqual(
            security_scan.differences("called", result, reviewed, flags),
            [
                "called: unexpected Suspicious 'Shell'",
                "called: expected IOC 'other.dll' is gone",
            ],
        )

    def test_changed_mraptor_flags_fail(self) -> None:
        result = self.results["called"]
        findings, _ = self.own_baseline("called")

        self.assertEqual(
            security_scan.differences("called", result, findings, "-W-"),
            ["called: mraptor flags -WX where the baseline has -W-; X from `Shell`"],
        )

    def test_a_file_the_baseline_does_not_know_fails(self) -> None:
        self.assertEqual(
            security_scan.differences("new", self.results["called"], None, None),
            [
                "new: no olevba findings in the baseline",
                "new: no mraptor flags in the baseline",
            ],
        )

    def test_code_that_runs_on_its_own_fails_even_when_the_baseline_records_it(self) -> None:
        for name, verdict in (
            ("auto_run", "SUSPICIOUS (A-X)"),
            ("quiet", "Macro OK (A--)"),
        ):
            with self.subTest(name):
                result = self.results[name]
                refusal = [
                    f"{name}: code that runs on its own, `Workbook_Open`; "
                    f"mraptor calls it {verdict}"
                ]

                self.assertEqual(security_scan.verdict(result), verdict)
                self.assertIn(("AutoExec", "Workbook_Open"), result["findings"])
                self.assertEqual(security_scan.never_accepted(name, result), refusal)
                self.assertEqual(
                    security_scan.differences(name, result, *self.own_baseline(name)), refusal
                )

    def test_a_type_suffix_alone_explains_a_pcode_name_the_source_lacks(self) -> None:
        source = 'Dim payload As String\r\npayload = "abc"\r\nMsgBox payload\r\n'
        cases = (
            ("matching P-code", '\tLd payload \r\n\tLitStr 0x0003 "abc"\r\n', [], []),
            ("type suffix", "\tLd payload$ \r\n", ["payload$"], []),
            ("name only in P-code", "\tArgsCall Shell 0x0001\r\n", [], ["name Shell"]),
            ("string only in P-code", '\tLitStr 0x0004 "calc"\r\n', [], ['string "calc"']),
            ("suffix on a name the source lacks", "\tLd other$ \r\n", [], ["name other$"]),
        )
        for label, pcode, suffixed, unexplained in cases:
            with self.subTest(label):
                self.assertEqual(
                    security_scan.stomping_evidence(pcode, source), (suffixed, unexplained)
                )

    def test_the_report_shows_auto_run_code_and_every_difference(self) -> None:
        out = Path(self.folder.name) / "report.md"
        with contextlib.redirect_stdout(io.StringIO()):
            security_scan.report(out, "Test", [self.paths["called"], self.paths["auto_run"]])
        text = out.read_text(encoding="utf-8")

        self.assertIn("| Workbook_Open | not applicable, source file | SUSPICIOUS (A-X) | no |",
                      text)
        self.assertIn("| none | not applicable, source file | Macro OK (-WX) | no |", text)
        self.assertIn("- auto_run.bas: no olevba findings in the baseline", text)
        self.assertIn("- auto_run.bas: code that runs on its own, 'Workbook_Open'; mraptor "
                      "calls it SUSPICIOUS (A-X)", text)
        self.assertIn("Code that runs on its own was found in auto_run.bas.", text)
        self.assertIn("olevba flags no VBA stomping.", text)


if __name__ == "__main__":
    unittest.main()
