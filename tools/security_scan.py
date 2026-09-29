"""Scan the shipped VBA with olevba and mraptor and hold the results to a reviewed baseline.

The runtime reads and writes files, runs processes, calls native code, and talks HTTP by
design, so olevba always reports Suspicious keywords and IOCs for it, and MacroRaptor
(mraptor) always flags it W, for writing files or memory, and X, for running something
outside VBA. Those are expected. The check compares each shipped file's olevba findings
and mraptor flags with tools/security_baseline.json and fails on any difference, in either
direction. Two results fail whatever the baseline says: code that runs on its own, which
olevba reports as AutoExec and mraptor flags A, and P-code its source does not explain.
ROneCOne has neither, and mraptor's SUSPICIOUS verdict needs A.

olevba's own VBA stomping flag is kept out of the comparison. It reports stomping when a
name in the P-code is missing from the source, and pcodedmp prints some names with a
trailing type character (payload$) that the source never spells. This scan applies the
same test with that suffix set aside, and a P-code keyword the source still lacks fails.

Usage:
    python tools/security_scan.py                    compare the repository's shipped files
    python tools/security_scan.py --update-baseline  accept the current results
    python tools/security_scan.py --report OUT.md --title TITLE FILE...
                                                     write a report for release assets
"""

from __future__ import annotations

import argparse
import datetime
import hashlib
import json
import re
import sys
from collections.abc import Iterable
from pathlib import Path

import oletools.olevba
from oletools import mraptor
from oletools.olevba import VBA_Parser

ROOT = Path(__file__).resolve().parents[1]
BASELINE = ROOT / "tools" / "security_baseline.json"
STOMPING_KEYWORD = "VBA Stomping"
TYPE_SUFFIX = re.compile(r"^([A-Za-z_][A-Za-z0-9_]*)[$%&!#@^]$")
RAPTOR_FLAGS = {"A": mraptor.re_autoexec, "W": mraptor.re_write, "X": mraptor.re_execute}


def shipped_files() -> list[Path]:
    """The files a release publishes: the runtime and every demo workbook."""
    return [ROOT / "src" / "ROneCOne.cls", *sorted((ROOT / "demo").glob("*.xlsm"))]


def pcode_keywords(pcode: str) -> set[tuple[str, str]]:
    """The names and string literals olevba's stomping test reads from the P-code."""
    keywords: set[tuple[str, str]] = set()
    for line in pcode.splitlines():
        if not line.startswith("\t"):
            continue
        tokens = line.split(None, 1)
        mnemonic = tokens[0]
        args = tokens[1].strip() if len(tokens) == 2 else ""
        if mnemonic in ("ArgsCall", "ArgsLd", "St", "Ld", "MemSt", "Label"):
            if args.startswith("(Call) "):
                args = args[7:]
            name = args.split(None, 1)[0]
            if not name.startswith("id_"):
                keywords.add(("name", name))
        if mnemonic == "LitStr":
            text = args.split(None, 1)[1]
            if len(text) >= 2:
                text = '"' + text[1:-1].replace('"', '""') + '"'
            keywords.add(("string", text))
    return keywords


def stomping_evidence(pcode: str, source: str) -> tuple[list[str], list[str]]:
    """P-code keywords the source lacks: those a trailing type character explains, and
    the rest, which mean the P-code says something the source does not."""
    suffixed: list[str] = []
    unexplained: list[str] = []
    for kind, value in sorted(pcode_keywords(pcode)):
        if value in source:
            continue
        match = TYPE_SUFFIX.match(value) if kind == "name" else None
        if match and re.search(rf"\b{re.escape(match.group(1))}\b", source):
            suffixed.append(value)
        else:
            unexplained.append(f"{kind} {value}")
    return suffixed, unexplained


def distinct(words: Iterable[str]) -> list[str]:
    """One spelling per word, ignoring case, in alphabetical order."""
    spellings: dict[str, str] = {}
    for word in sorted(words):
        spellings.setdefault(word.lower(), word)
    return sorted(spellings.values(), key=str.lower)


def cell(text: str) -> str:
    """Text safe inside a Markdown table cell."""
    return " ".join(text.split()).replace("|", "\\|").replace("`", "'")


def word_list(words: list[str]) -> str:
    """The words behind one MacroRaptor flag, with Declare statements counted, not listed."""
    declares = [word for word in words if word.lower().startswith("declare")]
    shown = [f"`{cell(word)}`" for word in words if word not in declares]
    if declares:
        plural = "s" if len(declares) != 1 else ""
        shown.append(f"{len(declares)} `Declare ... Lib` statement{plural}")
    if len(shown) > 2:
        return ", ".join(shown[:-1]) + ", and " + shown[-1]
    return " and ".join(shown)


def scan(path: Path) -> dict:
    """olevba's findings for one file, the stomping evidence behind its flag, and
    MacroRaptor's verdict with every word behind each of its flags."""
    parser = VBA_Parser(str(path))
    try:
        results = parser.analyze_macros() or []
        olevba_stomping = bool(parser.detect_vba_stomping())
        findings: dict[tuple[str, str], str] = {}
        for kind, keyword, description in results:
            if kind == "Suspicious" and keyword == STOMPING_KEYWORD:
                continue
            findings.setdefault((kind, keyword), description)
        source = parser.get_vba_code_all_modules()
        suffixed: list[str] = []
        unexplained: list[str] = []
        if parser.pcodedmp_output:
            suffixed, unexplained = stomping_evidence(parser.pcodedmp_output, source)
        # The same code, and the same call, as the mraptor command line.
        raptor = mraptor.MacroRaptor(source)
        raptor.scan()
        return {
            "findings": findings,
            "olevba_stomping": olevba_stomping,
            "has_pcode": bool(parser.pcodedmp_output),
            "suffixed": suffixed,
            "unexplained": unexplained,
            "mraptor_flags": raptor.get_flags(),
            "mraptor_suspicious": bool(raptor.suspicious),
            "mraptor_words": {
                flag: distinct(match.group() for match in pattern.finditer(raptor.vba_code))
                for flag, pattern in RAPTOR_FLAGS.items()
            },
        }
    finally:
        parser.close()


def verdict(result: dict) -> str:
    """MacroRaptor's verdict and flags, in the words its command line uses."""
    outcome = "SUSPICIOUS" if result["mraptor_suspicious"] else "Macro OK"
    return f"{outcome} ({result['mraptor_flags']})"


def auto_run(result: dict) -> list[str]:
    """Code that runs on its own, by olevba's AutoExec findings and MacroRaptor's A flag."""
    return distinct([keyword for kind, keyword in result["findings"] if kind == "AutoExec"]
                    + result["mraptor_words"]["A"])


def never_accepted(name: str, result: dict) -> list[str]:
    """Results no baseline may record: code that runs on its own, and P-code that its
    source does not explain."""
    problems: list[str] = []
    running = auto_run(result)
    if running:
        problems.append(f"{name}: code that runs on its own, {word_list(running)}; "
                        f"mraptor calls it {verdict(result)}")
    problems += [f"{name}: P-code keyword missing from the source: {item}"
                 for item in result["unexplained"]]
    return problems


def differences(name: str, result: dict, findings: set[tuple[str, str]] | None,
                flags: str | None) -> list[str]:
    """What one file's scan shows that its reviewed findings and flags do not expect."""
    problems: list[str] = []
    found = set(result["findings"])
    if findings is None:
        problems.append(f"{name}: no olevba findings in the baseline")
    else:
        problems += [f"{name}: unexpected {kind} {keyword!r}"
                     for kind, keyword in sorted(found - findings)]
        problems += [f"{name}: expected {kind} {keyword!r} is gone"
                     for kind, keyword in sorted(findings - found)]
    if flags is None:
        problems.append(f"{name}: no mraptor flags in the baseline")
    elif result["mraptor_flags"] != flags:
        added = [f"{flag} from {word_list(result['mraptor_words'][flag])}"
                 for flag in RAPTOR_FLAGS if flag in result["mraptor_flags"] and flag not in flags]
        problems.append(f"{name}: mraptor flags {result['mraptor_flags']} where the baseline has "
                        f"{flags}" + "".join(f"; {item}" for item in added))
    return problems + never_accepted(name, result)


def load_baseline() -> tuple[dict[str, set[tuple[str, str]]], dict[str, str]]:
    """The reviewed olevba findings and MacroRaptor flags, by repository path."""
    data = json.loads(BASELINE.read_text(encoding="utf-8"))
    findings = {name: {tuple(pair) for pair in pairs} for name, pairs in data["olevba"].items()}
    return findings, dict(data["mraptor"])


def write_baseline(scans: dict[str, dict]) -> None:
    """One entry per line, so a baseline change reads as a plain diff."""
    names = sorted(scans)
    lines = ["{", f'  "oletools": "{oletools.olevba.__version__}",', '  "mraptor": {']
    for index, name in enumerate(names):
        comma = "," if index < len(names) - 1 else ""
        lines.append(f"    {json.dumps(name)}: {json.dumps(scans[name]['mraptor_flags'])}{comma}")
    lines += ["  },", '  "olevba": {']
    for index, name in enumerate(names):
        pairs = sorted(scans[name]["findings"])
        lines.append(f"    {json.dumps(name)}: [")
        for pair_index, pair in enumerate(pairs):
            comma = "," if pair_index < len(pairs) - 1 else ""
            lines.append(f"      {json.dumps(list(pair))}{comma}")
        lines.append("    ]" + ("," if index < len(names) - 1 else ""))
    lines += ["  }", "}", ""]
    BASELINE.write_text("\n".join(lines), encoding="utf-8", newline="\n")


def relative(path: Path) -> str:
    return path.resolve().relative_to(ROOT).as_posix()


def check() -> int:
    """Compare the repository's shipped files with the baseline; 1 on any difference."""
    findings, flags = load_baseline()
    problems: list[str] = []
    seen: set[str] = set()
    for path in shipped_files():
        name = relative(path)
        seen.add(name)
        result = scan(path)
        problems += differences(name, result, findings.get(name), flags.get(name))
        print(f"{name}: {len(result['findings'])} olevba findings, stomping evidence clean: "
              f"{not result['unexplained']}, mraptor {verdict(result)}")
    for name in sorted((set(findings) | set(flags)) - seen):
        problems.append(f"{name}: in the baseline but no longer shipped")
    if problems:
        print("\nThe scan found something the reviewed baseline does not expect:")
        for problem in problems:
            print(f"  {problem}")
        print("\nIf the change is intended, review it and run "
              "python tools/security_scan.py --update-baseline")
        return 1
    print("olevba and mraptor results match the reviewed baseline.")
    return 0


def sha256(path: Path) -> str:
    return hashlib.sha256(path.read_bytes()).hexdigest()


def findings_table(result: dict, keys: list[tuple[str, str]]) -> str:
    """A Markdown table of the given olevba findings for one file."""
    rows = ["| Type | Keyword | Description |", "|---|---|---|"]
    rows += [f"| {kind} | `{cell(keyword)}` | {cell(result['findings'][(kind, keyword)])} |"
             for kind, keyword in keys]
    return "\n".join(rows)


def report(out: Path, title: str, files: list[Path]) -> int:
    """Write a Markdown report of olevba's and MacroRaptor's results for release assets."""
    baseline_findings, baseline_flags = load_baseline()
    findings_by_file = {Path(name).name: pairs for name, pairs in baseline_findings.items()}
    flags_by_file = {Path(name).name: flags for name, flags in baseline_flags.items()}
    scans = {path: scan(path) for path in sorted(files, key=lambda p: p.name)}
    runtime = next((result for path, result in scans.items() if path.suffix == ".cls"), None)
    runtime_set = set(runtime["findings"]) if runtime else set()
    lines = [
        f"# {title}",
        "",
        f"Generated {datetime.datetime.now(datetime.timezone.utc):%Y-%m-%d %H:%M} UTC by olevba "
        f"and MacroRaptor (oletools {oletools.olevba.__version__}) from the release assets. "
        "Every result below is compared with the reviewed baseline in "
        "`tools/security_baseline.json` at the release tag.",
        "",
        "| File | SHA-256 | Auto-run code | VBA stomping | MacroRaptor | Matches the baseline |",
        "|---|---|---|---|---|---|",
    ]
    all_differences: list[str] = []
    for path, result in scans.items():
        problems = differences(path.name, result, findings_by_file.get(path.name),
                               flags_by_file.get(path.name))
        all_differences += problems
        if not result["has_pcode"]:
            stomping = "not applicable, source file"
        elif result["unexplained"]:
            stomping = "P-code differs from the source"
        elif result["olevba_stomping"]:
            stomping = ("olevba flags it; no P-code lacks source"
                        + (" (see below)" if result["suffixed"] else ""))
        else:
            stomping = "none"
        lines.append(f"| {path.name} | `{sha256(path)}` | "
                     f"{cell(', '.join(auto_run(result))) or 'none'} | {stomping} | "
                     f"{verdict(result)} | {'no' if problems else 'yes'} |")
    lines += ["", "## Differences from the reviewed baseline", ""]
    lines += [f"- {cell(item)}" for item in all_differences] or ["None."]
    running = [path.name for path, result in scans.items() if auto_run(result)]
    raptor_note = ("MacroRaptor sets A for code that runs on its own, W for code that writes "
                   "files or memory, and X for code that runs something outside VBA, such as a "
                   "COM object, a command, or a Windows API function. It calls a file "
                   "suspicious only when A appears with W or X.")
    if runtime:
        raptor_note += " In ROneCOne.cls, " + "; ".join(
            f"{flag} comes from {word_list(words)}"
            for flag, words in runtime["mraptor_words"].items() if words) + "."
    stomping_notes: list[str] = []
    if any(result["olevba_stomping"] and not result["unexplained"] for result in scans.values()):
        stomping_notes.append(
            "olevba flags VBA stomping when a name in a workbook's P-code is missing from its "
            "source. In every workbook it flags without a difference above, the only such names "
            "are ones pcodedmp prints with a trailing type character that the source never "
            "spells, listed below per workbook. Each name without its suffix is in the source, "
            "so no P-code lacks source. Excel compiled that P-code from the published source "
            "while building the workbooks.")
    stomped = [path.name for path, result in scans.items() if result["unexplained"]]
    if stomped:
        stomping_notes.append(
            "In " + ", ".join(stomped) + ", the P-code holds names or strings that the source "
            "lacks, listed under the differences above, which is the mark of VBA stomping.")
    lines += [
        "",
        "## What the findings mean",
        "",
        "ROneCOne reads and writes files, runs commands through `cmd.exe`, calls Windows APIs "
        "in `kernel32`, `oleaut32`, `ole32`, and `bcrypt`, and sends HTTP requests, each only "
        "when the calling code asks it to. olevba reports the keywords behind those "
        "capabilities as Suspicious and the library names as IOCs. Its Hex and Base64 "
        "entries are text in the code that happens to decode; the Keyword column shows the "
        "decoded bytes and the Description column the text in the code. "
        + ("Neither olevba nor MacroRaptor finds code that runs on its own in any file."
           if not running else
           "Code that runs on its own was found in " + ", ".join(running) + "."),
        "",
        raptor_note,
        "",
        " ".join(stomping_notes) or "olevba flags no VBA stomping.",
        "",
        "## Findings",
        "",
    ]
    if runtime:
        lines += ["### ROneCOne.cls", "", findings_table(runtime, sorted(runtime_set)), ""]
    for path, result in scans.items():
        if result is runtime:
            continue
        found = set(result["findings"])
        extra = sorted(found - runtime_set)
        missing = sorted(runtime_set - found)
        blocks: list[str] = []
        if runtime:
            blocks.append("The findings of ROneCOne.cls, which this workbook contains"
                          + (", plus:" if extra else ", and nothing else."))
        if extra:
            blocks.append(findings_table(result, extra))
        if missing:
            blocks.append("Absent here though ROneCOne.cls has them: "
                          + ", ".join(f"`{cell(keyword)}`" for _, keyword in missing) + ".")
        if runtime:
            beyond: list[str] = []
            for flag, words in result["mraptor_words"].items():
                known = {word.lower() for word in runtime["mraptor_words"][flag]}
                new = [word for word in words if word.lower() not in known]
                if new:
                    beyond.append(f"{flag} from {word_list(new)}")
            if beyond:
                blocks.append("MacroRaptor also matches words that ROneCOne.cls lacks: "
                              + "; ".join(beyond) + ".")
        if result["olevba_stomping"] and result["suffixed"]:
            reach = "partly" if result["unexplained"] else "only"
            blocks.append(f"olevba's stomping flag rests {reach} on these suffixed names: "
                          + ", ".join(f"`{name}`" for name in result["suffixed"]) + ".")
        lines += [f"### {path.name}", "", "\n\n".join(blocks) or "No findings.", ""]
    out.write_text("\n".join(lines), encoding="utf-8")
    print(f"wrote {out}")
    return 0


def main(argv: list[str]) -> int:
    parser = argparse.ArgumentParser(description=__doc__.split("\n\n")[0])
    parser.add_argument("--update-baseline", action="store_true")
    parser.add_argument("--report", type=Path)
    parser.add_argument("--title", default="Security report")
    parser.add_argument("files", nargs="*", type=Path)
    args = parser.parse_args(argv)
    if args.report:
        return report(args.report, args.title, args.files)
    if args.update_baseline:
        scans = {relative(path): scan(path) for path in shipped_files()}
        refused = [problem for name, result in scans.items()
                   for problem in never_accepted(name, result)]
        if refused:
            print("refusing to record results the scan never accepts:")
            for problem in refused:
                print(f"  {problem}")
            return 1
        write_baseline(scans)
        print(f"wrote {BASELINE.relative_to(ROOT).as_posix()}")
        return 0
    return check()


if __name__ == "__main__":
    sys.exit(main(sys.argv[1:]))
