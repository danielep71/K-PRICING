#!/usr/bin/env python3
"""Tests for what keeps the VBE export of the VBA source equal to the committed source (#52)."""

from __future__ import annotations

import re
import unittest
from collections import defaultdict
from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]
CHECK = (ROOT / "tools" / "Check-KPR.ps1").read_text(encoding="ascii").replace("\r\n", "\n")

# Identifiers spelled with different case in two places that the Excel round trip does not change:
# Cleanup/CleanUp are names of procedures in separate modules, Filename is the named argument of
# Workbook.SaveAs. A new spelling difference is a new entry here only after an Excel round trip passed.
TOLERATED = {"cleanup", "filename"}


def sources() -> list[Path]:
    return sorted(path for folder in ("src", "examples", "tests") for path in (ROOT / folder).rglob("*.bas"))


class RoundTripTests(unittest.TestCase):
    def test_macro_options_attributes_are_dropped_from_the_export(self) -> None:
        """Excel adds procedure attributes when MacroOptions runs; the round trip ignores exactly those."""
        found = re.search(r"\$macroOptionsAttributes = '([^']+)'", CHECK)
        self.assertIsNotNone(found, "Check-KPR.ps1 names the pattern")
        pattern = found.group(1) if found else ""
        exported = (
            'Attribute VB_Name = "KPR_DATES_DAYS"\n'
            "Public Function KPR_Dates_DayOfWeek() As Variant\n"
            '    As Variant\n'
            'Attribute KPR_Dates_DayOfWeek.VB_Description = "Returns the weekday number"\n'
            'Attribute KPR_Dates_DayOfWeek.VB_ProcData.VB_Invoke_Func = " \\n20"\n'
            "End Function\n"
        )
        source = exported.replace(
            'Attribute KPR_Dates_DayOfWeek.VB_Description = "Returns the weekday number"\n', "").replace(
            'Attribute KPR_Dates_DayOfWeek.VB_ProcData.VB_Invoke_Func = " \\n20"\n', "")
        self.assertEqual(re.sub(pattern, "", exported), source)
        # a source that carries a procedure attribute still differs from the stripped export
        self.assertNotEqual(re.sub(pattern, "", source), exported)
        self.assertIn('Attribute VB_Name = "KPR_DATES_DAYS"', re.sub(pattern, "", exported))
        self.assertIn("[regex]::Replace($actual, $macroOptionsAttributes", CHECK)

    def test_the_committed_source_has_no_procedure_attributes(self) -> None:
        pattern = re.compile(r"^Attribute [A-Za-z_]\w*\.VB_", re.M)
        for path in sources():
            self.assertIsNone(pattern.search(path.read_text(encoding="ascii")), path.name)

    def test_no_identifier_is_spelled_two_ways(self) -> None:
        """The VBE conforms the case of identifiers that differ only by case, so the export would differ."""
        spellings: dict[str, dict[str, set[str]]] = defaultdict(lambda: defaultdict(set))
        for path in sources():
            for line in path.read_text(encoding="ascii").splitlines():
                code = re.sub(r"'.*$", "", re.sub(r'"(?:[^"]|"")*"', '""', line))
                if code.lstrip().lower().startswith("attribute"):
                    continue
                for token in re.findall(r"(?<![\w.])[A-Za-z_]\w*", code):
                    spellings[token.lower()][token].add(path.name)
        clash = {name: {spelling: sorted(files) for spelling, files in forms.items()}
                 for name, forms in spellings.items() if len(forms) > 1 and name not in TOLERATED}
        self.assertEqual(clash, {})


if __name__ == "__main__":
    unittest.main()
