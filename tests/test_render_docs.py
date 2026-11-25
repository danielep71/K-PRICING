"""Tests for tools/render_docs.py: generated documentation blocks."""

from __future__ import annotations

import contextlib
import io
import json
import re
import shutil
import sys
import tempfile
import unittest
from collections.abc import Callable
from pathlib import Path
from typing import Any
from unittest import mock

ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(ROOT / "tools"))

import check_excel_run  # noqa: E402
import check_issues  # noqa: E402
import check_repository  # noqa: E402
import render_docs  # noqa: E402
from render_docs import Block  # noqa: E402

try:
    import yaml
except ImportError:  # pragma: no cover - exercised only without the parser
    yaml = None


def suite_counts() -> tuple[int, int, int]:
    """The example suite's expected cases and assertions, read from its source as
    tests/test_vba_harness.py reads them, and the assertions of the last case its
    runner calls."""
    policy = json.loads((ROOT / ".github/excel-evidence-policy.json").read_text("utf-8"))
    source = (ROOT / policy["suite"]["source"]).read_text("ascii").replace("\r\n", "\n")

    def constant(name: str) -> int:
        match = re.search(rf"^\s*Private Const {name}\s+As Long = (\d+)\b", source, re.M)
        assert match, name
        return int(match.group(1))

    def body(name: str) -> str:
        match = re.search(rf"^Private Sub {name}\(.*?^End Sub$", source, re.M | re.S)
        assert match, name
        return match.group(0)

    last = re.findall(r"^\s+(Case\w+)$", body("RunSuite"), re.M)[-1]
    return (constant("EXPECTED_CASES"), constant("EXPECTED_ASSERTIONS"),
            len(re.findall(r"TEST_Harness\.Assert\w+ ", body(last))))


def begin(name: str) -> str:
    return f"<!-- BEGIN generated: {name} -->"


def end(name: str) -> str:
    return f"<!-- END generated: {name} -->"


class RenderTest(unittest.TestCase):
    def setUp(self) -> None:
        directory = tempfile.TemporaryDirectory(prefix="render-docs-")
        self.addCleanup(directory.cleanup)
        self.root = Path(directory.name)
        blocks = {"facts": Block("GUIDE.md", lambda root: "| fact |\n| --- |\n| one |")}
        patcher = mock.patch.dict(render_docs.BLOCKS, blocks, clear=True)
        patcher.start()
        self.addCleanup(patcher.stop)

    def write(self, name: str, text: str) -> None:
        (self.root / name).write_text(text, encoding="utf-8")

    def problems(self, write: bool = False) -> list[str]:
        with contextlib.redirect_stdout(io.StringIO()):
            return render_docs.run(self.root, write)

    def test_a_fresh_block_passes(self) -> None:
        self.write("GUIDE.md", f"Intro\n{begin('facts')}\n| fact |\n| --- |\n| one |\n"
                               f"{end('facts')}\nOutro\n")
        self.assertEqual(self.problems(), [])

    def test_a_stale_block_fails_and_write_renders_it(self) -> None:
        self.write("GUIDE.md", f"Intro\n{begin('facts')}\nold\n{end('facts')}\nOutro\n")
        self.assertEqual(self.problems(),
                         ["GUIDE.md: block facts differs from a fresh render; run "
                          "python tools/render_docs.py --write"])
        self.assertEqual(self.problems(write=True), [])
        self.assertEqual((self.root / "GUIDE.md").read_text(encoding="utf-8"),
                         f"Intro\n{begin('facts')}\n| fact |\n| --- |\n| one |\n{end('facts')}"
                         "\nOutro\n")
        self.assertEqual(self.problems(), [])

    def test_a_block_that_cannot_be_rendered_is_a_problem(self) -> None:
        def broken_policy(root: Path) -> str:
            raise check_excel_run.PolicyError("policy: host office must be one of 64-bit, 32-bit")
        self.write("GUIDE.md", f"Intro\n{begin('facts')}\nold\n{end('facts')}\nOutro\n")
        with mock.patch.dict(render_docs.BLOCKS, {"facts": Block("GUIDE.md", broken_policy)}):
            for write in (False, True):
                with self.subTest(write=write):
                    self.assertEqual(self.problems(write), [
                        "GUIDE.md: block facts cannot be rendered: policy: host office must be "
                        "one of 64-bit, 32-bit"])
        self.assertIn("\nold\n", (self.root / "GUIDE.md").read_text(encoding="utf-8"))

    def test_a_file_a_block_cannot_read_is_a_problem(self) -> None:
        # The evidence policy names a suite source that does not exist, as after the
        # module is renamed or deleted; or a renderer cannot read a file at all.
        policy = json.loads((ROOT / check_excel_run.POLICY).read_text(encoding="utf-8"))
        policy["suite"]["source"] = "tests/modules/TEST_Missing.bas"
        (self.root / check_excel_run.POLICY).parent.mkdir(parents=True)
        (self.root / check_excel_run.POLICY).write_text(json.dumps(policy), encoding="utf-8")

        def unreadable(root: Path) -> str:
            raise PermissionError(13, "Permission denied", "data.json")
        blocks = {"results": Block("GUIDE.md", render_docs.example_suite_results),
                  "facts": Block("GUIDE.md", unreadable)}
        self.write("GUIDE.md", f"{begin('results')}\nold\n{end('results')}\n"
                               f"{begin('facts')}\nold\n{end('facts')}\n")
        with mock.patch.dict(render_docs.BLOCKS, blocks, clear=True):
            for write in (False, True):
                with self.subTest(write=write):
                    self.assertEqual(sorted(self.problems(write)), [
                        "GUIDE.md: block facts cannot be rendered: [Errno 13] Permission "
                        "denied: 'data.json'",
                        "GUIDE.md: block results cannot be rendered: "
                        "tests/modules/TEST_Missing.bas: missing"])
        self.assertEqual((self.root / "GUIDE.md").read_text(encoding="utf-8").count("\nold\n"), 2)

    def test_marker_problems(self) -> None:
        cases = {
            f"{begin('facts')}\n": "block facts is not closed",
            f"{end('facts')}\n": "END facts without its BEGIN",
            f"{begin('facts')}\n{begin('other')}\n{end('other')}\n{end('facts')}\n":
                "BEGIN other inside block facts",
            f"{begin('facts')}\nx\n{end('facts')}\n{begin('facts')}\nx\n{end('facts')}\n":
                "block facts appears twice",
            "no markers\n": "block facts is missing",
        }
        for text, expected in cases.items():
            with self.subTest(expected=expected):
                self.write("GUIDE.md", text)
                found = self.problems()
                self.assertTrue(any(expected in problem for problem in found), found)

    def test_an_unknown_block_anywhere_fails(self) -> None:
        self.write("GUIDE.md", f"{begin('facts')}\n| fact |\n| --- |\n| one |\n{end('facts')}\n")
        (self.root / "docs").mkdir()
        self.write("docs/OTHER.md", f"{begin('nothing')}\nx\n{end('nothing')}\n")
        self.assertEqual(self.problems(), ["docs/OTHER.md: block nothing has no renderer"])

    def test_a_missing_file_fails(self) -> None:
        self.assertEqual(self.problems(), ["GUIDE.md: missing"])

    def test_main_exit_codes(self) -> None:
        self.write("GUIDE.md", f"{begin('facts')}\nold\n{end('facts')}\n")
        with contextlib.redirect_stdout(io.StringIO()) as output:
            self.assertEqual(render_docs.main(["--check", "--root", str(self.root)]), 1)
        self.assertIn("generated-docs: FAIL", output.getvalue())
        with contextlib.redirect_stdout(io.StringIO()):
            self.assertEqual(render_docs.main(["--write", "--root", str(self.root)]), 0)
            self.assertEqual(render_docs.main(["--check", "--root", str(self.root)]), 0)


class DataContractTest(unittest.TestCase):
    """The tables of the data contract follow docs/methodology/data-contract.json."""

    GUIDE = "docs/methodology/DATA_CONTRACT.md"
    CONTRACT = "docs/methodology/data-contract.json"

    def setUp(self) -> None:
        directory = tempfile.TemporaryDirectory(prefix="render-contract-")
        self.addCleanup(directory.cleanup)
        self.root = Path(directory.name)
        (self.root / "docs/methodology").mkdir(parents=True)
        for path in (self.GUIDE, self.CONTRACT):
            shutil.copyfile(ROOT / path, self.root / path)
        blocks = {name: block for name, block in render_docs.BLOCKS.items()
                  if block.path == self.GUIDE}
        patcher = mock.patch.dict(render_docs.BLOCKS, blocks, clear=True)
        patcher.start()
        self.addCleanup(patcher.stop)

    def edit_contract(self, edit: Callable[[dict[str, Any]], None]) -> None:
        path = self.root / self.CONTRACT
        contract = json.loads(path.read_text(encoding="utf-8"))
        edit(contract)
        path.write_text(json.dumps(contract), encoding="utf-8")

    def problems(self, write: bool = False) -> list[str]:
        with contextlib.redirect_stdout(io.StringIO()):
            return render_docs.run(self.root, write)

    def test_the_guide_has_the_three_blocks(self) -> None:
        self.assertEqual(sorted(render_docs.BLOCKS), ["data-contract-codes",
                                                      "data-contract-datasets",
                                                      "data-contract-fields"])
        self.assertEqual(self.problems(), [])

    def test_a_changed_contract_makes_the_tables_stale(self) -> None:
        def optional_first_field(contract: dict[str, Any]) -> None:
            contract["datasets"][0]["fields"][0]["required"] = False
        self.edit_contract(optional_first_field)
        self.assertEqual(self.problems(), [
            f"{self.GUIDE}: block data-contract-fields differs from a fresh render; run "
            "python tools/render_docs.py --write"])
        self.edit_contract(lambda contract: contract["datasets"][0].update(provider="Other"))
        self.assertIn(f"{self.GUIDE}: block data-contract-datasets differs from a fresh "
                      "render; run python tools/render_docs.py --write", self.problems())
        self.assertEqual(self.problems(write=True), [])
        self.assertIn("| Other |", (self.root / self.GUIDE).read_text(encoding="utf-8"))

    def test_an_edited_table_fails(self) -> None:
        path = self.root / self.GUIDE
        text = path.read_text(encoding="utf-8")
        path.write_text(text.replace("| Yes |", "| No |", 1), encoding="utf-8")
        self.assertEqual(self.problems(), [
            f"{self.GUIDE}: block data-contract-fields differs from a fresh render; run "
            "python tools/render_docs.py --write"])

    def test_an_invalid_contract_cannot_be_rendered(self) -> None:
        self.edit_contract(lambda contract: contract.update(kind="other"))
        reason = (f"{self.CONTRACT}: envelope does not match: schema_version 1, kind "
                  "'data-contract'; python tools/check_data_contract.py lists every problem")
        # The codes come from the validator, not the contract, and still render.
        self.assertEqual(sorted(self.problems()), [
            f"{self.GUIDE}: block data-contract-datasets cannot be rendered: {reason}",
            f"{self.GUIDE}: block data-contract-fields cannot be rendered: {reason}"])

    def test_how_each_kind_of_dataset_and_field_is_shown(self) -> None:
        base = {"illustrative": False, "description": "D", "provider": "P",
                "frequency": "F", "granularity": "G", "fields": [
                    {"name": "kind", "type": "text", "unit": None, "required": True,
                     "allowed": {"values": ["A", "B"]}, "notes": "Two kinds"},
                    {"name": "weight", "type": "decimal", "unit": "Decimal fraction",
                     "required": False, "allowed": {"min": 0, "max": 1}, "notes": ""},
                    {"name": "cap", "type": "integer", "unit": None, "required": False,
                     "allowed": {"max": 9}, "notes": ""}]}
        self.edit_contract(lambda contract: contract.update(datasets=[
            {**base, "id": "DS-02", "source": "a.tsv", "encoding": None, "delimiter": "\t",
             "decimal_separator": ","},
            {**base, "id": "DS-03", "source": "b.xlsx", "encoding": None, "delimiter": None,
             "decimal_separator": None},
            {**base, "id": "DS-04", "source": "Inputs", "encoding": None, "delimiter": None,
             "decimal_separator": None}]))
        datasets = render_docs.data_contract_datasets(self.root)
        self.assertIn("| `DS-02` | D | P | F | G | `a.tsv` |  |", datasets)
        fields = render_docs.data_contract_fields(self.root)
        self.assertIn("### `DS-02` D\n\nRead from `a.tsv`, delimited text in `UTF-8`, with a "
                      "tab between cells and `,` as the decimal separator.", fields)
        self.assertIn("Read from `b.xlsx`, which is not delimited text", fields)
        self.assertIn("Read from the sheet `Inputs`.", fields)
        self.assertIn("| `kind` | Text | — | Yes | `A`, `B` | Two kinds |", fields)
        self.assertIn("| `weight` | Decimal | Decimal fraction | No | 0 to 1 |  |", fields)
        self.assertIn("| `cap` | Integer | — | No | ≤ 9 |  |", fields)
        self.assertNotIn("Illustrative", fields)


class TriageRulesTest(unittest.TestCase):
    """The triage rules table of the issue guide follows check_issues.RULES."""

    GUIDE = "docs/repo_governance/ISSUE_LABELS.md"

    def setUp(self) -> None:
        directory = tempfile.TemporaryDirectory(prefix="render-triage-")
        self.addCleanup(directory.cleanup)
        self.root = Path(directory.name)
        (self.root / self.GUIDE).parent.mkdir(parents=True)
        shutil.copyfile(ROOT / self.GUIDE, self.root / self.GUIDE)
        patcher = mock.patch.dict(render_docs.BLOCKS, {"triage-rules": render_docs.BLOCKS[
            "triage-rules"]}, clear=True)
        patcher.start()
        self.addCleanup(patcher.stop)

    def problems(self, write: bool = False) -> list[str]:
        with contextlib.redirect_stdout(io.StringIO()):
            return render_docs.run(self.root, write)

    def test_the_rendered_table_lists_every_registered_rule(self) -> None:
        text = render_docs.triage_rules(ROOT)
        rows = re.findall(r"^\| `([a-z0-9-]+)` \| (.+) \|$", text, re.M)
        self.assertEqual(rows, list(check_issues.RULES.items()))
        self.assertIn(f"{begin('triage-rules')}\n{text}\n{end('triage-rules')}",
                      (ROOT / self.GUIDE).read_text(encoding="utf-8"))

    def test_a_new_rule_makes_the_table_stale(self) -> None:
        self.assertEqual(self.problems(), [])
        rules = {**check_issues.RULES, "estimate": "An estimate of the work."}
        with mock.patch.object(check_issues, "RULES", rules):
            self.assertEqual(self.problems(), [
                f"{self.GUIDE}: block triage-rules differs from a fresh render; run "
                "python tools/render_docs.py --write"])
            self.assertEqual(self.problems(write=True), [])
        self.assertIn("| `estimate` | An estimate of the work. |",
                      (self.root / self.GUIDE).read_text(encoding="utf-8"))


@unittest.skipIf(yaml is None, "PyYAML is not installed (.github/requirements.txt)")
class WorkflowTriggersTest(unittest.TestCase):
    """The trigger table of the README follows each workflow's on: block."""

    def setUp(self) -> None:
        directory = tempfile.TemporaryDirectory(prefix="render-triggers-")
        self.addCleanup(directory.cleanup)
        self.root = Path(directory.name)
        shutil.copyfile(ROOT / "README.md", self.root / "README.md")
        shutil.copytree(ROOT / render_docs.WORKFLOWS, self.root / render_docs.WORKFLOWS)
        patcher = mock.patch.dict(render_docs.BLOCKS, {"workflow-triggers": render_docs.BLOCKS[
            "workflow-triggers"]}, clear=True)
        patcher.start()
        self.addCleanup(patcher.stop)

    def problems(self, write: bool = False) -> list[str]:
        with contextlib.redirect_stdout(io.StringIO()):
            return render_docs.run(self.root, write)

    def edit(self, name: str, old: str, new: str) -> None:
        path = self.root / render_docs.WORKFLOWS / name
        text = path.read_text(encoding="utf-8")
        self.assertEqual(text.count(old), 1, old)
        path.write_text(text.replace(old, new), encoding="utf-8")

    def test_a_changed_trigger_makes_the_block_stale(self) -> None:
        self.assertEqual(self.problems(), [])
        stale = ["README.md: block workflow-triggers differs from a fresh render; run "
                 "python tools/render_docs.py --write"]
        changes = [
            ("link-check.yml", '- cron: "43 6 * * 1"', '- cron: "43 6 * * 2"'),
            ("repository-integrity.yml", "  pull_request:\n",
             "  pull_request:\n    types: [opened]\n"),
            ("minimum-python.yml", '      - "tools/**"\n', ""),
            ("release.yml", "  workflow_dispatch:\n", "  workflow_call:\n"),
        ]
        for name, old, new in changes:
            with self.subTest(workflow=name):
                path = self.root / render_docs.WORKFLOWS / name
                original = path.read_text(encoding="utf-8")
                self.edit(name, old, new)
                self.assertEqual(self.problems(), stale)
                path.write_text(original, encoding="utf-8")
                self.assertEqual(self.problems(), [])

    def test_a_new_workflow_makes_the_block_stale_and_write_lists_it(self) -> None:
        (self.root / render_docs.WORKFLOWS / "nightly.yml").write_text(
            "name: Nightly\non:\n  schedule:\n    - cron: '5 1 * * *'\njobs: {}\n",
            encoding="utf-8")
        self.assertEqual(len(self.problems()), 1)
        self.assertEqual(self.problems(write=True), [])
        self.assertIn("| [**Nightly**](.github/workflows/nightly.yml) | Daily at 01:05 UTC |",
                      (self.root / "README.md").read_text(encoding="utf-8"))

    def test_a_workflow_with_a_duplicate_key_cannot_be_rendered(self) -> None:
        self.edit("link-check.yml", "  workflow_dispatch:\n",
                  "  workflow_dispatch:\n  workflow_dispatch:\n")
        found = self.problems()
        self.assertEqual(len(found), 1)
        self.assertIn("README.md: block workflow-triggers cannot be rendered: "
                      ".github/workflows/link-check.yml:", found[0])
        self.assertIn("found duplicate key 'workflow_dispatch'", found[0])

    def test_how_each_event_and_filter_reads(self) -> None:
        def text(event: str, config: object) -> str:
            return render_docs.event_text(event, config, "w.yml")
        self.assertEqual(text("push", None), "every push")
        self.assertEqual(text("push", {"branches": ["main"], "paths": ["a/**"]}),
                         "pushes to `main` that change `a/**`")
        self.assertEqual(text("push", {"tags": ["v*"]}), "pushes of the tags `v*`")
        self.assertEqual(text("pull_request", {"branches-ignore": ["x"]}),
                         "pull requests except into `x`")
        self.assertEqual(text("pull_request", {"paths-ignore": ["docs/**"]}),
                         "pull requests unless they change only `docs/**`")
        self.assertEqual(text("pull_request_target", {"types": ["opened", "edited"]}),
                         "pull requests (opened, edited), run from the base branch")
        self.assertEqual(text("workflow_dispatch", {"inputs": {"ref": {}}}), "manual runs")
        self.assertEqual(text("workflow_run", {"workflows": ["CI"]}),
                         "`workflow_run` events with `workflows` `CI`")
        self.assertEqual(text("schedule", [{"cron": "0 3 * * *"}, {"cron": "30 4 * * 0"},
                                           {"cron": "0 */6 * * *"}]),
                         "daily at 03:00 UTC, weekly on Sunday at 04:30 UTC, on the schedule "
                         "`0 */6 * * *`")
        with self.assertRaisesRegex(ValueError, "w.yml: on.push.branches is not a list"):
            text("push", {"branches": []})

    def test_without_pyyaml_the_block_is_skipped_and_left_as_it_is(self) -> None:
        before = (self.root / "README.md").read_text(encoding="utf-8")
        self.edit("link-check.yml", '- cron: "43 6 * * 1"', '- cron: "43 6 * * 2"')
        with mock.patch.dict(sys.modules, {"strict_yaml": None}), \
                contextlib.redirect_stdout(io.StringIO()) as output:
            self.assertEqual(render_docs.run(self.root, write=True), [])
        self.assertIn("skipped: README.md: block workflow-triggers: PyYAML is not installed",
                      output.getvalue())
        self.assertEqual((self.root / "README.md").read_text(encoding="utf-8"), before)


class RepositoryTest(unittest.TestCase):
    def test_the_repository_blocks_are_fresh(self) -> None:
        with contextlib.redirect_stdout(io.StringIO()):
            self.assertEqual(render_docs.run(ROOT, write=False), [])

    def test_the_suite_results_follow_the_suite(self) -> None:
        cases, assertions, last = suite_counts()
        text = render_docs.example_suite_results(ROOT)
        self.assertIn("RESULT=PASS; suite=examples; operation=PASS; completeness=COMPLETE; "
                      f"cases={cases}; assertions={assertions}; failures=0; cleanup=PASS; "
                      "recovery=NONE", text)
        # The skipped-case run leaves out the last case and its assertions.
        self.assertIn("`TEST_Examples.RunExampleTestsWithSkippedCase` | `RESULT=FAIL`, "
                      "`operation=PASS`, `completeness=INCOMPLETE`, "
                      f"`cases={cases - 1}`, `assertions={assertions - last}`, `failures=1`",
                      text)
        self.assertIn("Two more entry points", text)

    def test_every_check_is_listed_once_and_described(self) -> None:
        text = render_docs.integrity_checks(ROOT)
        context = check_repository.Context(ROOT, repository=None, default_branch=None,
                                           is_template=None)
        for check in check_repository.checks(context):
            with self.subTest(check=check.name):
                self.assertEqual(text.count(f"| `{check.name}` |"), 1)
        self.assertNotIn(sys.executable, text)
        self.assertNotIn("Built in: |", text)

    def test_the_changelog_check_is_shown_as_ci_runs_it(self) -> None:
        text = render_docs.integrity_checks(ROOT)
        self.assertIn("| `changelog` | `python tools/check_changelog.py --repository "
                      "OWNER/REPOSITORY` |", text)

    def test_the_probe_a_check_must_fail_first_is_shown(self) -> None:
        text = render_docs.integrity_checks(ROOT)
        self.assertIn("| `actionlint` | `actionlint -no-color`, once the probe "
                      "`actionlint -no-color .github/actionlint/invalid-workflow.yml` has "
                      "failed with `unexpected key \"bogus-key\"` |", text)

    def test_a_check_without_a_description_is_refused(self) -> None:
        def undocumented(context: check_repository.Context) -> list[tuple[str, str]]:
            return []

        listed = [check_repository.Check("silent", function=undocumented)]
        with mock.patch.object(check_repository, "checks", return_value=listed), \
                self.assertRaisesRegex(ValueError, "silent: undocumented has no docstring"):
            render_docs.integrity_checks(ROOT)

    def test_commands_are_shown_as_typed(self) -> None:
        self.assertEqual(render_docs.shown_command([sys.executable, "-m", "x"]), "python -m x")
        self.assertEqual(render_docs.shown_command(["/usr/bin/node", "a.mjs"]), "node a.mjs")
        # The same table on Windows, whose executables carry a suffix.
        self.assertEqual(render_docs.shown_command([r"C:\Program Files\nodejs\node.exe", "a.mjs"]),
                         "node a.mjs")
        self.assertEqual(render_docs.shown_command([r"C:\tools\actionlint.EXE", "-no-color"]),
                         "actionlint -no-color")
        self.assertEqual(render_docs.shown_command(["git", "diff"]), "git diff")


if __name__ == "__main__":
    unittest.main()
