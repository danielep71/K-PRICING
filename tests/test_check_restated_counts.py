"""Tests for tools/check_restated_counts.py: counts restated outside their source."""

from __future__ import annotations

import contextlib
import io
import json
import shutil
import subprocess
import sys
import tempfile
import unittest
from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(ROOT / "tools"))

import check_data_contract  # noqa: E402
import check_restated_counts as C  # noqa: E402

CONTRACT = "docs/methodology/data-contract.json"
# The restatements the guides held before this check, as they were written.
OLD_INSTALLATION = ("line. One case reads the illustrative test cases from that folder and\n"
                    "records their totals as results. The two entry points below take the same\n"
                    "argument. A passing run ends with:\n")
OLD_HARNESS = ("## Modes and certification\n\nThe example suite has three modes:\n\n"
               "- `harness` and `suite`: the harness module;\n"
               "- `modes`: the three modes in run order, the log each one saves and whether\n"
               "  it must pass (only `NORMAL`);\n")


class Fixture:
    """A repository with the evidence policy, the data contract and some guides."""

    def __init__(self, root: Path, docs: dict[str, str],
                 assertions: int | None = None) -> None:
        self.root = root
        for path in (C.POLICY, CONTRACT):
            (root / path).parent.mkdir(parents=True, exist_ok=True)
            shutil.copyfile(ROOT / path, root / path)
        if assertions is not None:
            policy = json.loads((root / C.POLICY).read_text(encoding="utf-8"))
            policy["suite"]["expected_assertions"] = assertions
            (root / C.POLICY).write_text(json.dumps(policy), encoding="utf-8")
        for path, text in docs.items():
            (root / path).parent.mkdir(parents=True, exist_ok=True)
            (root / path).write_text(text, encoding="utf-8")
        subprocess.run(["git", "init", "-q", str(root)], check=True)
        subprocess.run(["git", "-C", str(root), "add", "-A"], check=True)


class RestatedCounts(unittest.TestCase):
    def findings(self, docs: dict[str, str], assertions: int | None = None,
                 allowed: tuple[C.Allowed, ...] = ()) -> list[str]:
        with tempfile.TemporaryDirectory() as folder:
            Fixture(Path(folder), docs, assertions)
            return C.findings(Path(folder), allowed=allowed)

    def assert_found(self, text: str, phrase: str, source: str) -> None:
        found = self.findings({"GUIDE.md": text})
        self.assertEqual(len(found), 1, (text, found))
        self.assertIn(f"GUIDE.md:1: {phrase!r} restates the {source}", found[0])

    def test_the_repository_passes(self) -> None:
        self.assertEqual(C.findings(ROOT), [])

    def test_every_source_reads_its_current_value(self) -> None:
        policy = json.loads((ROOT / C.POLICY).read_text(encoding="utf-8"))
        values = {source.name: source.read(ROOT) for source in C.SOURCES}
        self.assertEqual(values["cases of the example suite"],
                         policy["suite"]["expected_cases"])
        self.assertEqual(values["assertions of the example suite"],
                         policy["suite"]["expected_assertions"])
        self.assertEqual(values["modes of the example suite"], len(policy["modes"]))
        self.assertEqual(values["failure entry points of the example suite"],
                         sum(1 for mode in policy["modes"] if not mode["must_pass"]))
        self.assertEqual(values["finding codes of the data-contract validator"],
                         len(check_data_contract.CODES))
        self.assertGreater(values["Repository integrity checks"],
                           values["static checks of exported VBA source"])

    def test_the_old_restatements_fail_with_file_line_phrase_and_source(self) -> None:
        found = self.findings({"INSTALLATION.md": OLD_INSTALLATION,
                               "docs/methodology/TEST_HARNESS.md": OLD_HARNESS})
        self.assertEqual(len(found), 3, found)
        self.assertTrue(found[0].startswith(
            "INSTALLATION.md:2: 'two entry points' restates the failure entry points of the "
            "example suite; its source is the modes in .github/excel-evidence-policy.json "
            "that must not pass, which holds 2;"), found[0])
        self.assertTrue(found[1].startswith(
            "docs/methodology/TEST_HARNESS.md:3: 'three modes' restates the modes of the "
            "example suite; its source is modes in .github/excel-evidence-policy.json"),
            found[1])
        self.assertTrue(found[2].startswith("docs/methodology/TEST_HARNESS.md:6: 'three modes'"),
                        found[2])
        for finding in found:
            self.assertTrue(finding.endswith("link the source or add a renderer to "
                                             "tools/render_docs.py"), finding)

    def test_counts_in_digits_or_words_fail_whatever_their_value(self) -> None:
        for text, phrase in (("It makes 34 assertions.\n", "34 assertions"),
                             ("It makes thirty-four assertions.\n", "thirty-four assertions"),
                             ("It makes Thirty four assertions.\n", "thirty four assertions"),
                             ("It makes `34` assertions.\n", "34 assertions"),
                             ("It makes **34** assertions.\n", "34 assertions"),
                             ("It makes 33 assertions.\n", "33 assertions"),
                             ("It makes 1,034 expected assertions.\n",
                              "1,034 expected assertions")):
            with self.subTest(text=text):
                self.assert_found(text, phrase, "assertions of the example suite")
        self.assert_found("Run all 14 cases.\n", "14 cases", "cases of the example suite")
        self.assert_found("The 39 **Repository integrity** checks.\n",
                          "39 repository integrity checks", "Repository integrity checks")
        self.assert_found("Nine static checks run.\n", "nine static checks",
                          "static checks of exported VBA source")
        self.assert_found("It reads two datasets.\n", "two datasets",
                          "datasets of the data contract")
        self.assert_found("It has twelve finding codes.\n", "twelve finding codes",
                          "finding codes of the data-contract validator")

    def test_a_stale_copy_says_so_and_a_current_one_does_not(self) -> None:
        stale = self.findings({"GUIDE.md": "It makes 33 assertions.\n"})
        current = self.findings({"GUIDE.md": "It makes 34 assertions.\n"})
        self.assertIn("(a stale copy)", stale[0])
        self.assertNotIn("stale", current[0])

    def test_a_count_and_its_phrase_may_span_a_line_break(self) -> None:
        found = self.findings({"GUIDE.md": "Title\n\nThe example suite has three\nmodes.\n"})
        self.assertEqual(len(found), 1, found)
        self.assertIn("GUIDE.md:3: 'three modes'", found[0])

    def test_other_numbers_pass(self) -> None:
        text = ("It took 34 ms. See issue #34 and v3.4, commit `a34bc5d` and 34bc5d1, the "
                "file case34 assertions, 3.4 assertions, step 6 checks them, run 34-cases, "
                "path tools/34 cases, one case, the two example cases and 34 test cases "
                "are ignored. 2 assertion-level errors.\n")
        self.assertEqual(self.findings({"GUIDE.md": text}), [])

    def test_generated_blocks_and_the_changelog_are_out_of_scope(self) -> None:
        block = ("Before.\n\n<!-- BEGIN generated: example-suite-results -->\n"
                 "cases=14; 34 assertions; Two more entry points show the failure paths\n"
                 "<!-- END generated: example-suite-results -->\n\nAfter: three modes.\n")
        found = self.findings({"INSTALLATION.md": block,
                               "CHANGELOG.md": "Ran 14 cases and 34 assertions.\n"})
        self.assertEqual(len(found), 1, found)
        self.assertIn("INSTALLATION.md:7: 'three modes'", found[0])

    def test_untracked_markdown_is_out_of_scope(self) -> None:
        with tempfile.TemporaryDirectory() as folder:
            Fixture(Path(folder), {"GUIDE.md": "clean\n"})
            (Path(folder) / "NOTES.md").write_text("three modes\n", encoding="utf-8")
            self.assertEqual(C.findings(Path(folder), allowed=()), [])

    def test_the_policy_decides_what_is_reported(self) -> None:
        text = {"GUIDE.md": "It makes 34 assertions.\n"}
        current = self.findings(text)
        changed = self.findings(text, assertions=35)
        self.assertIn("which holds 34;", current[0])
        self.assertIn("which holds 35;", changed[0])
        self.assertIn("(a stale copy)", changed[0])

    def test_an_unreadable_source_fails(self) -> None:
        with tempfile.TemporaryDirectory() as folder:
            Fixture(Path(folder), {"GUIDE.md": "clean\n"})
            (Path(folder) / C.POLICY).write_text("{", encoding="utf-8")
            found = C.findings(Path(folder), allowed=())
        self.assertTrue(found)
        self.assertTrue(all("cannot be read" in finding for finding in found), found)

    def test_allowed_restatements(self) -> None:
        docs = {"GUIDE.md": "The example suite has three modes.\n"}
        kept = C.Allowed("GUIDE.md", "three  modes", "The table below lists them.")
        self.assertEqual(self.findings(docs, allowed=(kept,)), [])
        elsewhere = C.Allowed("OTHER.md", "three modes", "Another guide.")
        found = self.findings(docs, allowed=(elsewhere,))
        self.assertEqual(len(found), 2, found)
        self.assertIn("GUIDE.md:1: 'three modes'", found[0])
        self.assertIn("allowed restatement 'three modes' in OTHER.md matches nothing", found[1])
        unreasoned = C.Allowed("GUIDE.md", "three modes", " ")
        self.assertEqual(self.findings(docs, allowed=(unreasoned,)),
                         ["allowed restatement 'three modes' in GUIDE.md has no reason"])

    def test_number_values(self) -> None:
        for text, value in (("34", 34), ("1,034", 1034), ("thirty-four", 34),
                            ("Twenty one", 21), ("nineteen", 19), ("ninety", 90)):
            self.assertEqual(C.number_value(text), value, text)

    def test_main_reports_and_exits(self) -> None:
        with tempfile.TemporaryDirectory() as folder:
            Fixture(Path(folder), {"GUIDE.md": "three modes\n"})
            output = io.StringIO()
            with contextlib.redirect_stdout(output):
                code = C.main(["--root", folder])
        self.assertEqual(code, 1)
        self.assertIn("FAIL: GUIDE.md:1: 'three modes'", output.getvalue())
        self.assertIn("restated-counts: FAIL", output.getvalue())


if __name__ == "__main__":
    unittest.main()
