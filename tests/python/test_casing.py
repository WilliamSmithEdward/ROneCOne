"""Identifier casing guard: importing ROneCOne must not recase the names in a
host project's modules (issue #6).

VBA keeps one spelling per identifier across a whole project, and a
declaration anywhere sets it: a parameter named `value` in ROneCOne.cls turns
every `.Value` in the host's modules into `.value`, which lands as noise in
every exported diff. The same rule made ROneCOne and ModernJsonInVBA recase
each other when they shared a project. These tests hold the runtime's code
tokens to one spelling per name and to the spelling the default references
give any name they define. The second half reads the registered type
libraries, so it needs Windows and pywin32 and skips elsewhere: CI enforces
one spelling per name, and a local run enforces both.
tools/run_casing_roundtrip.ps1 checks the outcome directly in Excel.
"""

from __future__ import annotations

import functools
import re
import unittest
from collections import defaultdict
from pathlib import Path


ROOT = Path(__file__).resolve().parents[2]
RUNTIME = ROOT / "src" / "ROneCOne.cls"

# The references every new Excel VBA project carries, in reference order.
REFERENCES = (
    "{000204EF-0000-0000-C000-000000000046}",  # VBA
    "{00020813-0000-0000-C000-000000000046}",  # Excel
    "{00020430-0000-0000-C000-000000000046}",  # stdole
    "{2DF8D04C-5BFA-101B-BDE5-00AA0044DE52}",  # Office
)

# Keywords keep their own spelling and never enter the name table. Text is
# deliberately absent: it is a keyword only after Option Compare, and a
# declared `text` does recase every `.Text` in the project.
KEYWORDS = {word.lower() for word in """
    AddressOf And Any As Attribute Base Binary Boolean ByRef Byte ByVal Call
    Case Compare Const Currency Database Date Decimal Declare Dim Do Double
    Each Else ElseIf Empty End Enum Eqv Erase Error Event Exit Explicit False
    For Friend Function Get Global GoSub GoTo If Imp Implements In Integer Is
    Let Lib Like Long LongLong LongPtr Loop LSet Me Mod Module New Next Not
    Nothing Null Object On Option Optional Or ParamArray Preserve Print
    Private Property PtrSafe Public RaiseEvent ReDim Rem Resume Return RSet
    Select Set Single Static Step Stop String Sub Then To True Type
    TypeOf Until Variant Wend While With WithEvents Xor
""".split()}

# ROneCOne's own members named after their .NET types, kept over the
# all-capitals spelling Office gives the same words. A host that writes GUID
# or .XML sees those two recased; every other name keeps its library form.
EXEMPT = {"guid": "Guid", "xml": "Xml"}

# Identifiers start with a letter, so a line-continuation underscore, which
# the VBE drops when it rejoins a wrapped Attribute line, is not a token.
WORD = re.compile(r"[A-Za-z][A-Za-z0-9_]*")
NUMBER = re.compile(
    r"&[Hh][0-9A-Fa-f]+&?|&[Oo][0-7]+&?|"
    r"(?<![A-Za-z0-9_])\d+(?:\.\d+)?(?:[eE][+-]?\d+)?[#!@&%^]?"
)


def code_tokens(text: str) -> list[str]:
    """Identifier tokens after Option Explicit: comments, strings, and
    numeric literals are blanked first, so only code participates."""
    lines = text.replace("\r\n", "\n").split("\n")
    start = next(
        (i for i, line in enumerate(lines) if line.startswith("Option Explicit")),
        0,
    )
    found = []
    for line in lines[start:]:
        kept = []
        in_string = False
        for ch in line:
            if ch == '"':
                in_string = not in_string
                kept.append(" ")
            elif ch == "'" and not in_string:
                break
            else:
                kept.append(" " if in_string else ch)
        found.extend(WORD.findall(NUMBER.sub(" ", "".join(kept))))
    return found


def _registered_typelib(guid: str):
    import pythoncom
    import winreg

    def subkeys(path: str) -> list[str]:
        with winreg.OpenKey(winreg.HKEY_CLASSES_ROOT, path) as key:
            return [winreg.EnumKey(key, i) for i in range(winreg.QueryInfoKey(key)[0])]

    version = max(
        subkeys(rf"TypeLib\{guid}"),
        key=lambda v: tuple(int(part, 16) for part in v.split(".")),
    )
    lcid = next(
        int(name, 16)
        for name in subkeys(rf"TypeLib\{guid}\{version}")
        if re.fullmatch(r"[0-9A-Fa-f]+", name)
    )
    major, minor = (int(part, 16) for part in version.split("."))
    return pythoncom.LoadRegTypeLib(guid, major, minor, lcid)


@functools.lru_cache(maxsize=1)
def typelib_spellings() -> dict[str, str]:
    """The spelling the default references give each name: a type, member,
    or constant spelling beats a parameter spelling, and references rank in
    project order."""
    best: dict[str, tuple[tuple[int, int], str]] = {}

    def offer(name: str, rank: tuple[int, int]) -> None:
        key = name.lower()
        if key not in best or rank < best[key][0]:
            best[key] = (rank, name)

    for order, guid in enumerate(REFERENCES):
        library = _registered_typelib(guid)
        for index in range(library.GetTypeInfoCount()):
            info = library.GetTypeInfo(index)
            offer(library.GetDocumentation(index)[0], (0, order))
            attr = info.GetTypeAttr()
            for f in range(attr.cFuncs):
                names = info.GetNames(info.GetFuncDesc(f).memid)
                for position, name in enumerate(names):
                    offer(name, (1 if position else 0, order))
            for v in range(attr.cVars):
                for name in info.GetNames(info.GetVarDesc(v).memid):
                    offer(name, (0, order))
    return {key: spelling for key, (_, spelling) in best.items()}


def spellings_by_name(paths: list[Path]) -> dict[str, dict[str, list[str]]]:
    found: dict[str, dict[str, list[str]]] = defaultdict(lambda: defaultdict(list))
    for path in paths:
        for token in code_tokens(path.read_text(encoding="utf-8")):
            found[token.lower()][token].append(path.name)
    return found


class CasingTests(unittest.TestCase):
    def test_runtime_spells_every_name_one_way(self) -> None:
        split = [
            f"{lower}: spelled {sorted(forms)}"
            for lower, forms in sorted(spellings_by_name([RUNTIME]).items())
            if len(forms) > 1
        ]
        self.assertEqual([], split, "VBA would recase one spelling to the other")

    def test_runtime_uses_the_type_library_spelling(self) -> None:
        try:
            canon = typelib_spellings()
        except Exception as exc:  # noqa: BLE001 - no oracle without the libraries
            self.skipTest(f"default type libraries unavailable: {exc}")
        canon = {**canon, **EXEMPT}
        wrong = []
        for lower, forms in sorted(spellings_by_name([RUNTIME]).items()):
            wanted = canon.get(lower)
            if wanted is None or lower in KEYWORDS:
                continue
            wrong.extend(f"{name}: {wanted} is canonical" for name in forms if name != wanted)
        self.assertEqual([], wrong, "ROneCOne would recase these names in a host project")


if __name__ == "__main__":
    unittest.main()
