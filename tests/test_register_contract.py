"""Tests that src/workbook/M_Register.bas registers what the facade documents.

M_Register hands Excel a description for each of the 22 KPR_Dates_* functions and one
description per argument. Nothing compares them with the facade, because the VBA runs only
in Excel, so these tests read the VBA source and fail when the texts, the names, the order or
the number of arguments differ from docs/PUBLIC_API.txt and KPR_Dates.bas.
"""

from __future__ import annotations

import re
import unittest
from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]
REGISTER = ROOT / "src/workbook/M_Register.bas"
FACADE = ROOT / "src/modules/KPR_Dates.bas"
API = ROOT / "docs/PUBLIC_API.txt"

HOST_LIMIT = 255


def read(path: Path) -> str:
    return path.read_text(encoding="utf-8").replace("\r\n", "\n")


def need(found: re.Match[str] | None, what: str) -> re.Match[str]:
    """The match, or a failure that names what the source lacks."""
    if found is None:
        raise AssertionError(f"the source has no {what}")
    return found


def unquote(literal: str) -> str:
    return literal[1:-1].replace('""', '"')


def constants() -> dict[str, str]:
    """The Private Const strings of M_Register, by name."""
    return {name: unquote(value) for name, value in
            re.findall(r'^\s*Private Const (\w+)\s+As String = ("(?:[^"]|"")*")$',
                       read(REGISTER), re.M)}


def registered() -> list[tuple[str, str, list[str]]]:
    """Name, description and argument descriptions of each Case of Describe, in order."""
    text = read(REGISTER)
    body = text[text.index("Select Case Index"):text.index("End Select")]
    texts = constants()
    found = []
    for case in re.split(r"^\s*Case \d+\n", body, flags=re.M)[1:]:
        name = unquote(need(re.search(r'FunctionName = ("(?:[^"]|"")*")', case), "name").group(1))
        summary = unquote(need(re.search(r'Summary = ("(?:[^"]|"")*")', case), "summary").group(1))
        arrays = re.search(r"ArgTexts = VBA\.Array\((.*?)\)\n", case, re.S)
        args: list[str] = []
        if arrays is not None:
            for item in re.findall(r'("(?:[^"]|"")*"|\w+)', arrays.group(1).replace("_\n", "")):
                args.append(unquote(item) if item.startswith('"') else texts[item])
        else:
            assert "ArgTexts = Empty" in case, name
        found.append((name, summary, args))
    return found


def supported() -> list[tuple[str, list[str]]]:
    """Name and parameter names of each [supported] function, in inventory order."""
    rows = []
    section = read(API).split("\n[supported]\n")[1].split("\n[project-public]\n")[0]
    for name, parameters in re.findall(r"^# SIG\t\w+\tFunction\t(\w+)\tPublic Function \w+\((.*?)\) As",
                                       section, re.M):
        rows.append((name, re.findall(r"(\w+) As Variant", parameters)))
    return rows


def facade() -> dict[str, tuple[str, list[tuple[str, str]]]]:
    """Purpose and (argument, text) pairs of each function's documentation header."""
    found = {}
    for part in re.split(r"(?=^Public Function )", read(FACADE), flags=re.M)[1:]:
        name = need(re.match(r"Public Function (\w+)", part), "function").group(1)
        purpose = need(re.search(r"' PURPOSE\n(.*?)\n'\n' WHY", part, re.S), "PURPOSE").group(1)
        purpose = " ".join(line[1:].strip() for line in purpose.split("\n"))
        inputs = need(re.search(r"' INPUTS\n(.*?)\n' RETURNS", part, re.S), "INPUTS").group(1)
        found[name] = (purpose, re.findall(r"^'   (\w+)\n'     (.*)$", inputs, re.M))
    return found


class RegisterContractTest(unittest.TestCase):
    def test_every_supported_function_is_registered_once_in_order(self) -> None:
        self.assertEqual([name for name, _, _ in registered()],
                         [name for name, _ in supported()])
        self.assertEqual(len(registered()), 22)

    def test_the_count_constant_is_the_number_of_functions(self) -> None:
        found = need(re.search(r"Public Const FUNCTION_COUNT As Long = (\d+)", read(REGISTER)),
                     "FUNCTION_COUNT")
        self.assertEqual(int(found.group(1)), len(supported()))

    def test_each_function_has_one_description_per_argument(self) -> None:
        signatures = dict(supported())
        for name, _, args in registered():
            self.assertEqual(len(args), len(signatures[name]), name)

    def test_the_descriptions_are_those_of_the_facade(self) -> None:
        documented = facade()
        signatures = dict(supported())
        for name, summary, args in registered():
            purpose, inputs = documented[name]
            self.assertEqual(summary, purpose, name)
            if signatures[name]:
                self.assertEqual([argument for argument, _ in inputs], signatures[name], name)
                self.assertEqual(args, [text for _, text in inputs], name)

    def test_no_text_passes_the_limit_of_the_host(self) -> None:
        for name, summary, args in registered():
            for text in [summary, *args]:
                self.assertGreater(len(text), 0, name)
                self.assertLessEqual(len(text), HOST_LIMIT, name)

    def test_the_category_name_fits_the_host(self) -> None:
        category = constants()["CATEGORY_NAME"]
        self.assertTrue(category)
        self.assertLessEqual(len(category), HOST_LIMIT)

    def test_the_module_stays_out_of_the_function_list(self) -> None:
        self.assertIn("Option Private Module", read(REGISTER))


if __name__ == "__main__":
    unittest.main()
