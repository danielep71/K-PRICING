"""Tests for tools/release.py, in temporary Git repositories with an origin."""

from __future__ import annotations

import argparse
import contextlib
import gzip
import hashlib
import io
import json
import shutil
import subprocess
import sys
import tarfile
import tempfile
import unittest
import zipfile
from collections.abc import Sequence
from pathlib import Path
from typing import Any, ClassVar
from unittest import mock

ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(ROOT / "tools"))
sys.path.insert(0, str(ROOT / "tests"))

import check_excel_run  # noqa: E402
import release  # noqa: E402
from gitfixture import git, init  # noqa: E402

try:
    import yaml
except ImportError:  # pragma: no cover - installed from .github/requirements.txt
    yaml = None

CHANGELOG = """# Changelog

## [Unreleased]

## [1.2.0] - 2026-11-02

### Added

- A feature.

## [1.1.0] - 2026-10-01

### Fixed

- A bug.

[Unreleased]: https://github.com/owner/project/compare/v1.2.0...HEAD
[1.2.0]: https://github.com/owner/project/releases/tag/v1.2.0
[1.1.0]: https://github.com/owner/project/releases/tag/v1.1.0
"""


@unittest.skipIf(shutil.which("git") is None, "git is not installed")
class ReleaseTest(unittest.TestCase):
    def setUp(self) -> None:
        directory = tempfile.TemporaryDirectory(prefix="release-")
        self.addCleanup(directory.cleanup)
        base = Path(directory.name)
        self.origin = base / "origin.git"
        self.root = base / "work"
        init(self.origin, "--bare")
        init(self.root)
        git(self.root, "remote", "add", "origin", str(self.origin))
        self.commit({"CHANGELOG.md": CHANGELOG, "VERSION": "1.2.0\n", "README.md": "x\n"})
        git(self.root, "push", "-q", "origin", "main")
        self.sha = git(self.root, "rev-parse", "HEAD").strip()

    def commit(self, files: dict[str, str]) -> None:
        for name, text in files.items():
            (self.root / name).write_text(text, encoding="utf-8")
        git(self.root, "add", "-A")
        git(self.root, "commit", "-qm", "Fixture")

    def tag(self, name: str = "v1.2.0", message: str | None = None) -> None:
        text = message if message is not None else (
            f"{name} - Release\n\nCandidate {git(self.root, 'rev-parse', 'HEAD').strip()}.\n"
            "Deviations: none.\n")
        git(self.root, "tag", "-a", name, "-m", text)

    def problems(self, tag: str = "v1.2.0") -> list[str]:
        return release.verify(self.root, tag, "main")

    def test_a_consistent_tag_passes(self) -> None:
        self.tag()
        self.assertEqual(self.problems(), [])

    def test_the_tag_must_be_annotated_and_well_named(self) -> None:
        git(self.root, "tag", "v1.2.0")
        self.assertTrue(any("not an annotated tag" in p for p in self.problems()))
        self.assertEqual(self.problems("1.2.0"), ["tag '1.2.0' is not vX.Y.Z"])

    def test_the_version_must_match_the_changelog_and_version(self) -> None:
        self.tag("v1.3.0", f"v1.3.0\n\nCandidate {self.sha}.\nDeviations: none.\n")
        found = self.problems("v1.3.0")
        self.assertTrue(any("releases 1.2.0 as its newest version, not 1.3.0" in p
                            for p in found), found)
        self.assertTrue(any("VERSION at v1.3.0 is 1.2.0, not 1.3.0" in p for p in found), found)

    def test_version_must_exist(self) -> None:
        git(self.root, "rm", "-q", "VERSION")
        git(self.root, "commit", "-qm", "Drop VERSION")
        git(self.root, "push", "-q", "origin", "main")
        self.tag()
        self.assertEqual(self.problems(), ["VERSION at v1.2.0 is missing, not 1.2.0"])

    def test_the_commit_must_be_on_the_default_branch(self) -> None:
        git(self.root, "switch", "-q", "-c", "side")
        self.commit({"README.md": "side\n"})
        self.tag()
        self.assertTrue(any("is not on origin/main" in p for p in self.problems()))

    def test_the_message_must_name_the_candidate_and_its_deviations(self) -> None:
        self.tag(message="v1.2.0\n\nCandidate " + "0" * 40 + ".\n")
        found = self.problems()
        self.assertTrue(any(f"'Candidate {self.sha}.'" in p for p in found), found)
        self.assertTrue(any("no 'Deviations:' line" in p for p in found), found)

    def test_deviations_may_be_a_list_on_the_next_lines(self) -> None:
        head = git(self.root, "rev-parse", "HEAD").strip()
        self.tag(message=f"v1.2.0 - Release\n\nCandidate {head}.\nDeviations:\n"
                         "- Invariant 8: the published workbook was not checked.\n"
                         "* Invariant 4: the run names another build.\n")
        self.assertEqual(self.problems(), [])

    def test_deviations_without_text_or_a_list_fail(self) -> None:
        head = git(self.root, "rev-parse", "HEAD").strip()
        self.tag(message=f"v1.2.0 - Release\n\nCandidate {head}.\nDeviations:\n\nSigned.\n")
        self.assertEqual(self.problems(),
                         ["the message of v1.2.0 has no 'Deviations:' line followed by text, "
                          "such as 'none', or by a list"])

    def with_vba(self, path: str = "src/modules/MExample.bas") -> str:
        """Commit a VBA component on main, push it and return the new commit."""
        (self.root / path).parent.mkdir(parents=True, exist_ok=True)
        self.commit({path: 'Attribute VB_Name = "MExample"\nOption Explicit\n'})
        git(self.root, "push", "-q", "origin", "main")
        return git(self.root, "rev-parse", "HEAD").strip()

    def message(self, head: str, *lines: str, deviations: str = "Deviations: none.") -> str:
        return "\n".join(["v1.2.0 - Release", "", f"Candidate {head}.", *lines, deviations, ""])

    def retag(self, message: str) -> None:
        if git(self.root, "tag", "-l", "v1.2.0"):
            git(self.root, "tag", "-d", "v1.2.0")
        self.tag(message=message)

    def test_a_vba_component_needs_excel_evidence_or_an_invariant_4_deviation(self) -> None:
        head = self.with_vba()
        self.tag(message=self.message(head))
        self.assertEqual(self.problems(), [
            f"v1.2.0 ({head}) holds VBA components, but its message has neither an "
            "'Excel evidence: <digest> (run <sha>).' line nor a '- Invariant 4: <text>' "
            "item under 'Deviations:'"])

    def test_a_vba_component_is_a_bas_cls_or_frm_file_at_any_depth(self) -> None:
        for path, component in (("src/modules/MExample.bas", True),
                                ("deep/er/Sheet.CLS", True), ("forms/frmX.frm", True),
                                ("forms/frmX.frx", False), ("src/MExample.bas.orig", False)):
            with self.subTest(path=path):
                head = self.with_vba(path)
                self.assertEqual(release.holds_vba(self.root, head), component)
                git(self.root, "rm", "-q", path)
                git(self.root, "commit", "-qm", "Remove")

    def test_another_deviation_does_not_waive_the_excel_run(self) -> None:
        head = self.with_vba()
        self.tag(message=self.message(head, deviations="Deviations:\n"
                                      "- Invariant 8: the workbook was not checked.\n"
                                      "Invariant 4: outside the list."))
        self.assertEqual(len(self.problems()), 1, self.problems())

    def test_an_excel_evidence_line_is_accepted(self) -> None:
        head = self.with_vba()
        for run in (head, "1" * 40):
            with self.subTest(run=run):
                self.retag(self.message(head, f"Excel evidence: {'a' * 64} (run {run})."))
                self.assertEqual(self.problems(), [])

    def test_an_invariant_4_deviation_is_accepted(self) -> None:
        head = self.with_vba()
        for deviations in ("Deviations:\n- Invariant 4: no Excel host was available.",
                           "Deviations:\n- Invariant 8: not checked.\n"
                           "* Invariant 4: the run names\n  another build."):
            with self.subTest(deviations=deviations):
                self.retag(self.message(head, deviations=deviations))
                self.assertEqual(self.problems(), [])

    def test_a_malformed_excel_evidence_line_fails(self) -> None:
        head = self.with_vba()
        digest = "a" * 64
        for line in (f"Excel evidence: {digest[:-1]} (run {head}).",
                     f"Excel evidence: {digest.upper()} (run {head}).",
                     f"Excel evidence: {digest} (run {head[:12]}).",
                     f"Excel evidence: {digest} (run {head})",
                     f"Excel evidence: {digest} (run {head}). Checked.",
                     f"Excel evidence: sha256:{digest} (run {head}).",
                     "Excel evidence: none."):
            with self.subTest(line=line):
                # An invariant 4 deviation does not excuse a malformed line.
                self.retag(self.message(head, line,
                                        deviations="Deviations:\n- Invariant 4: no host."))
                found = self.problems()
                self.assertEqual(len(found), 1, found)
                self.assertIn(f"has a malformed line {line!r}", found[0])
        git(self.root, "tag", "-d", "v1.2.0")
        line = f"Excel evidence: {digest} (run {head})."
        self.tag(message=self.message(head, line, line))
        self.assertEqual(self.problems(),
                         ["the message of v1.2.0 has 2 'Excel evidence:' lines, not one"])

    def test_without_a_vba_component_no_excel_evidence_is_needed(self) -> None:
        self.commit({"MExample.bas.orig": "x\n", "frmX.frx": "x\n"})
        git(self.root, "push", "-q", "origin", "main")
        head = git(self.root, "rev-parse", "HEAD").strip()
        self.tag(message=self.message(head))
        self.assertEqual(self.problems(), [])
        # A malformed declaration fails all the same.
        git(self.root, "tag", "-d", "v1.2.0")
        self.tag(message=self.message(head, "Excel evidence: pending."))
        self.assertEqual(len(self.problems()), 1)

    def test_verify_with_evidence_checks_the_declared_digest_and_run(self) -> None:
        head = self.with_vba()
        folder, workbook = self.evidence_files()
        checker = FakeChecker()
        gate = release.excel_gate(self.root, head, folder, workbook, checker=checker)
        self.assertEqual(gate.problems, [])
        self.tag(message=self.message(head, gate.line))
        self.assertEqual(release.verify(self.root, "v1.2.0", "main", (folder, workbook),
                                        checker), [])
        self.assertEqual(checker.calls[-1], (head, None))
        # The evidence changed after the tag line was written.
        (folder / "host.txt").write_text("operator=Someone else\n", encoding="utf-8")
        found = release.verify(self.root, "v1.2.0", "main", (folder, workbook), checker)
        self.assertEqual(len(found), 1, found)
        self.assertIn(f"declares Excel evidence {gate.digest}, but", found[0])
        # A tag with only a deviation declares no evidence to check.
        git(self.root, "tag", "-d", "v1.2.0")
        self.tag(message=self.message(head, deviations="Deviations:\n- Invariant 4: x."))
        found = release.verify(self.root, "v1.2.0", "main", (folder, workbook), checker)
        self.assertEqual(found, [f"the message of v1.2.0 declares no Excel evidence to check "
                                 f"against {folder}"])

    def test_verify_with_evidence_checks_an_equivalent_run(self) -> None:
        run = self.with_vba()
        self.commit({"README.md": "later\n"})
        git(self.root, "push", "-q", "origin", "main")
        head = git(self.root, "rev-parse", "HEAD").strip()
        folder, workbook = self.evidence_files()
        checker = FakeChecker(equivalence="identical")
        gate = release.excel_gate(self.root, head, folder, workbook, run, checker)
        self.assertEqual((gate.problems, gate.run, gate.candidate), ([], run, head))
        self.tag(message=self.message(head, gate.line))
        self.assertEqual(release.verify(self.root, "v1.2.0", "main", (folder, workbook),
                                        checker), [])
        self.assertEqual(checker.calls[-1], (run, head))

    def evidence_files(self) -> tuple[Path, Path]:
        folder = self.root.parent / "excel-run"
        (folder / "exported").mkdir(parents=True)
        (folder / "host.txt").write_text("operator=Fixture\n", encoding="utf-8")
        (folder / "exported" / "MExample.bas").write_bytes(b"Option Explicit\r\n")
        workbook = self.root.parent / "Workbook-development.xlsm"
        workbook.write_bytes(b"workbook")
        return folder, workbook

    def test_notes_hold_the_message_and_the_changelog_section(self) -> None:
        self.tag()
        text = release.notes(self.root, "v1.2.0")
        self.assertTrue(text.startswith("v1.2.0 - Release\n\nCandidate "))
        self.assertIn("## Changes\n\n### Added\n\n- A feature.\n", text)
        self.assertNotIn("A bug.", text)
        self.assertNotIn("releases/tag", text)
        with self.assertRaisesRegex(release.ReleaseError, "no section for 1.3.0"):
            git(self.root, "tag", "-a", "v1.3.0", "-m", "x")
            release.notes(self.root, "v1.3.0")

    def test_bundle_is_reproducible(self) -> None:
        self.tag()
        first = release.bundle(self.root, "v1.2.0", "project", self.root.parent / "one")
        second = release.bundle(self.root, "v1.2.0", "project", self.root.parent / "two")
        self.assertEqual([path.name for path in first],
                         ["project-1.2.0-source.tar.gz", "SHA256SUMS"])
        self.assertEqual(first[0].read_bytes(), second[0].read_bytes())
        digest = hashlib.sha256(first[0].read_bytes()).hexdigest()
        self.assertEqual(first[1].read_text(encoding="utf-8"),
                         f"{digest}  project-1.2.0-source.tar.gz\n")
        with gzip.open(first[0]) as packed, tarfile.open(fileobj=io.BytesIO(packed.read())) as tar:
            self.assertIn("project-1.2.0/CHANGELOG.md", tar.getnames())
        other = release.bundle(self.root, "HEAD", "project", self.root.parent / "three")
        self.assertEqual(other[0].name, f"project-{self.sha[:12]}-source.tar.gz")

    def test_verify_published_reads_the_bundle_and_checksums_that_bundle_writes(
            self) -> None:
        self.tag()
        archive, checksums = release.bundle(self.root, "v1.2.0", "project",
                                            self.root.parent / "dist")
        gh = FakeGh()
        gh.files = {"1": archive.read_bytes(), "2": checksums.read_bytes()}
        self.assertEqual([path.name for path in (archive, checksums)],
                         [BUNDLE_NAME, release.SUMS])
        facts = release.verify_published(REPOSITORY, "v1.2.0", runner=gh)
        self.assertEqual([status for status, _ in facts], ["PASS"] * 7, facts)

    def test_main_reports_and_exits(self) -> None:
        self.tag()
        with contextlib.redirect_stdout(io.StringIO()) as output:
            self.assertEqual(release.main(["--root", str(self.root), "verify", "v1.2.0",
                                           "--default-branch", "main"]), 0)
            self.assertEqual(release.main(["--root", str(self.root), "verify", "v9.9.9",
                                           "--default-branch", "main"]), 1)
        self.assertIn("release-verify: PASS (v1.2.0)", output.getvalue())
        self.assertIn("FAIL: tag v9.9.9 is missing", output.getvalue())
        self.assertIn("release-verify: FAIL (v9.9.9)", output.getvalue())

    def test_notes_and_bundle_report_a_failure_without_a_traceback(self) -> None:
        # Neither subcommand has --sha: the summary line names its own argument.
        with contextlib.redirect_stdout(io.StringIO()) as output:
            self.assertEqual(release.main(["--root", str(self.root), "notes", "v9.9.9"]), 1)
            self.assertEqual(release.main(["--root", str(self.root), "bundle", "no-such-ref",
                                           "--name", "Project", "--out", str(self.root / "out")]),
                             1)
        self.assertIn("release-notes: FAIL (v9.9.9)", output.getvalue())
        self.assertIn("release-bundle: FAIL (no-such-ref)", output.getvalue())


class FakeChecker:
    """check_excel_run.check, returning set evidence and recording the commit checked
    and the commit it was asked to be equivalent to."""

    def __init__(self, findings: Sequence[str] = (), binding: str = "VERIFIED",
                 equivalence: str = "", carried: str = "") -> None:
        self.evidence = check_excel_run.Evidence(findings=list(findings), binding=binding,
                                                 equivalence=equivalence)
        # Set as an attribute: the check names carried evidence this way once it
        # can carry a run across comment-only changes.
        if carried:
            setattr(self.evidence, "carried", carried)
        self.calls: list[tuple[str, str | None]] = []

    def __call__(self, root: Path, commit: str, folder: Path, workbook: Path | None = None,
                 require_binding: bool = False, equivalent_to: str | None = None,
                 **others: object) -> check_excel_run.Evidence:
        assert require_binding and workbook is not None and not others, others
        self.calls.append((commit, equivalent_to))
        return self.evidence


@unittest.skipIf(shutil.which("git") is None, "git is not installed")
class ExcelGateTest(unittest.TestCase):
    """The evidence subcommand, with check_excel_run.check replaced at its seam."""

    def setUp(self) -> None:
        directory = tempfile.TemporaryDirectory(prefix="release-gate-")
        self.addCleanup(directory.cleanup)
        self.base = Path(directory.name)
        self.root = self.base / "work"
        init(self.root)
        self.commit("first")
        self.run_sha = git(self.root, "rev-parse", "HEAD").strip()
        self.commit("second")
        self.head = git(self.root, "rev-parse", "HEAD").strip()
        self.folder = self.base / "excel-run"
        (self.folder / "exported").mkdir(parents=True)
        (self.folder / "host.txt").write_bytes(b"outcome=PASS\n")
        (self.folder / "exported" / "MExample.bas").write_bytes(b"Option Explicit\r\n")
        self.workbook = self.base / "Workbook-development.xlsm"
        self.workbook.write_bytes(b"workbook")

    def commit(self, text: str) -> None:
        (self.root / "src").mkdir(exist_ok=True)
        (self.root / "src" / "MExample.bas").write_text(text, encoding="utf-8")
        git(self.root, "add", "-A")
        git(self.root, "commit", "-qm", text)

    def gate(self, checker: FakeChecker, run: str | None = None) -> release.ExcelGate:
        return release.excel_gate(self.root, "HEAD", self.folder, self.workbook, run, checker)

    def test_good_evidence_gives_the_digest_and_the_tag_line(self) -> None:
        checker = FakeChecker()
        gate = self.gate(checker)
        self.assertEqual(gate.problems, [])
        self.assertEqual(checker.calls, [(self.head, None)])
        listing = "".join(f"{hashlib.sha256(data).hexdigest()}  {name}\n" for name, data in (
            ("evidence/exported/MExample.bas", b"Option Explicit\r\n"),
            ("evidence/host.txt", b"outcome=PASS\n"), ("workbook", b"workbook")))
        self.assertEqual(gate.digest, hashlib.sha256(listing.encode()).hexdigest())
        self.assertEqual(gate.line, f"Excel evidence: {gate.digest} (run {self.head}).")
        self.assertTrue(release.EXCEL_LINE.match(gate.line))
        # The same files elsewhere, with the workbook under another name, give the
        # same digest; another byte gives another.
        copy = self.base / "copy"
        shutil.copytree(self.folder, copy)
        renamed = self.base / "other.xlsm"
        shutil.copyfile(self.workbook, renamed)
        self.assertEqual(release.evidence_digest(copy, renamed), gate.digest)
        (copy / "host.txt").write_bytes(b"outcome=PASS \n")
        self.assertNotEqual(release.evidence_digest(copy, renamed), gate.digest)

    def test_an_earlier_run_counts_only_when_shown_equivalent(self) -> None:
        checker = FakeChecker(equivalence=f"{self.head} (2 run inputs identical)")
        gate = self.gate(checker, self.run_sha[:12])
        self.assertEqual((gate.problems, gate.run), ([], self.run_sha))
        self.assertEqual(checker.calls, [(self.run_sha, self.head)])
        self.assertEqual(gate.line, f"Excel evidence: {gate.digest} (run {self.run_sha}).")
        gate = self.gate(FakeChecker(findings=["run input src/MExample.bas: differs"]),
                         self.run_sha)
        self.assertEqual(gate.problems, [
            "excel-run: run input src/MExample.bas: differs",
            f"the run of {self.run_sha[:12]} is not shown equivalent to {self.head[:12]}"])
        self.assertEqual(gate.digest, "")

    def test_bad_evidence_is_refused(self) -> None:
        cases: list[tuple[FakeChecker, str]] = [
            (FakeChecker(findings=["host.txt: outcome is FAIL"]),
             "excel-run: host.txt: outcome is FAIL"),
            (FakeChecker(binding="UNVERIFIED"), "the binding is UNVERIFIED, not VERIFIED"),
            (FakeChecker(binding="FAIL"), "the binding is FAIL, not VERIFIED"),
            (FakeChecker(carried="Carried from 0123456789ab: src/MExample.bas differ only "
                                 "in comments"),
             "the evidence is carried from another commit; carried evidence never "
             "certifies a release")]
        for checker, problem in cases:
            with self.subTest(problem=problem):
                gate = self.gate(checker)
                self.assertEqual((gate.problems, gate.digest), ([problem], ""))

    def test_missing_evidence_is_refused_without_a_check(self) -> None:
        checker = FakeChecker()
        gate = release.excel_gate(self.root, "HEAD", self.base / "none",
                                  self.base / "none.xlsm", checker=checker)
        self.assertEqual(gate.problems, [f"the evidence folder {self.base / 'none'} is missing",
                                         f"the workbook {self.base / 'none.xlsm'} is missing"])
        self.assertEqual(checker.calls, [])
        with self.assertRaisesRegex(release.ReleaseError, "commit nothing not found"):
            self.gate(checker, "nothing")

    def test_main_prints_the_report_and_the_tag_line(self) -> None:
        arguments = ["--root", str(self.root), "evidence", "HEAD", "--excel-evidence",
                     str(self.folder), "--workbook", str(self.workbook)]
        with mock.patch.object(check_excel_run, "check", FakeChecker()), \
                contextlib.redirect_stdout(io.StringIO()) as output:
            self.assertEqual(release.main(arguments), 0)
        text = output.getvalue()
        digest = release.evidence_digest(self.folder, self.workbook)
        self.assertIn(f"Candidate commit: {self.head}\nEnvironment:", text)
        self.assertIn("Binding: VERIFIED\n", text)
        self.assertIn(f"Excel evidence digest: {digest} (a run of the candidate)\n", text)
        self.assertIn(f"Tag message line: Excel evidence: {digest} (run {self.head}).\n", text)
        self.assertTrue(text.endswith(f"release-evidence: PASS ({self.head[:12]})\n"))
        with mock.patch.object(check_excel_run, "check", FakeChecker(binding="FAIL")), \
                contextlib.redirect_stdout(io.StringIO()) as output:
            self.assertEqual(release.main([*arguments, "--run-commit", "nothing"]), 2)
            self.assertEqual(release.main(arguments), 1)
        self.assertIn("FAIL: commit nothing not found\nrelease-evidence: FAIL (HEAD)",
                      output.getvalue())
        self.assertIn("FAIL: the binding is FAIL, not VERIFIED\n", output.getvalue())
        self.assertNotIn("Tag message line", output.getvalue())

    def test_verify_takes_the_evidence_and_the_workbook_together(self) -> None:
        with contextlib.redirect_stderr(io.StringIO()), self.assertRaises(SystemExit):
            release.main(["--root", str(self.root), "verify", "v1.2.0", "--default-branch",
                          "main", "--excel-evidence", str(self.folder)])


@unittest.skipIf(shutil.which("git") is None, "git is not installed")
class ExcelGateRunTest(unittest.TestCase):
    """The evidence subcommand on the evidence fixtures of test_check_excel_run.py,
    with the real check: a run of the candidate, an equivalent run and a run whose
    inputs changed."""

    fixtures: ClassVar[Any]
    root: ClassVar[Path]
    built: ClassVar[str]

    @classmethod
    def setUpClass(cls) -> None:
        import test_check_excel_run as fixtures

        # The fixture module builds its repository in setUpModule, which runs only
        # with its own tests, and removes it once they have run.
        if not fixtures.COMMIT:
            fixtures.setUpModule()
            cls.addClassCleanup(fixtures.tearDownModule)
        if not fixtures.COMMIT:
            raise unittest.SkipTest("needs the modules of tests/modules/")
        cls.fixtures = fixtures
        cls.root, cls.built = fixtures.make_repository()
        cls.addClassCleanup(shutil.rmtree, cls.root, True)
        # The fixture helpers read the example suite from their own repository,
        # which their module removes once its tests have run.
        for name, value in (("ROOT", cls.root), ("COMMIT", cls.built)):
            patcher = mock.patch.object(fixtures, name, value)
            patcher.start()
            cls.addClassCleanup(patcher.stop)

    def gate(self, candidate: str, run: str | None = None) -> release.ExcelGate:
        fixtures = self.fixtures
        with fixtures.evidence(root=self.root, commit=self.built) as folder:
            workbook = folder.parent / f"{folder.name}-Workbook.xlsm"
            workbook.write_bytes(fixtures.workbook_bytes(self.built))
            try:
                fixtures.export_manifest.write(workbook, folder)
                return release.excel_gate(self.root, candidate, folder, workbook, run)
            finally:
                workbook.unlink(missing_ok=True)

    def test_a_run_of_the_candidate_passes(self) -> None:
        gate = self.gate(self.built)
        self.assertEqual(gate.problems, [])
        self.assertIn("Binding: VERIFIED", gate.report)
        self.assertTrue(release.DIGEST.match(gate.digest))

    def test_an_equivalent_run_passes_and_a_changed_input_fails(self) -> None:
        later = self.fixtures.commit_change(self.root, self.built,
                                            {"README.md": b"Documentation only.\n"})
        gate = self.gate(later, self.built)
        self.assertEqual(gate.problems, [])
        self.assertIn(f"Equivalent: {later} (", gate.report)
        self.assertEqual(gate.line, f"Excel evidence: {gate.digest} (run {self.built}).")
        harness = (self.root / self.fixtures.POLICY.harness).read_bytes()
        changed = self.fixtures.commit_change(self.root, self.built,
                                              {self.fixtures.POLICY.harness: harness + b"' x\n"})
        gate = self.gate(changed, self.built)
        self.assertIn(f"the run of {self.built[:12]} is not shown equivalent to "
                      f"{changed[:12]}", gate.problems)
        self.assertTrue(any(f"run input {self.fixtures.POLICY.harness}: differs" in problem
                            for problem in gate.problems), gate.problems)
        # Evidence of the run is no evidence of the changed commit by itself.
        gate = self.gate(changed)
        self.assertNotEqual(gate.problems, [])
        self.assertEqual(gate.digest, "")


class IntegrityTest(unittest.TestCase):
    def runs(self, payload: object, returncode: int = 0) -> list[str]:
        completed = subprocess.CompletedProcess(
            [], returncode, stdout=json.dumps({"check_runs": payload}).encode(),
            stderr=b"denied")
        with mock.patch("release.subprocess.run", return_value=completed) as run:
            problems = release.integrity("owner/project", "a" * 40)
        self.assertEqual(run.call_args.args[0][:2], ["gh", "api"])
        return problems

    def test_a_successful_run_passes(self) -> None:
        self.assertEqual(self.runs([{"name": "Repository integrity", "conclusion": "success"},
                                    {"name": "CodeQL", "conclusion": "failure"}]), [])

    def test_missing_failed_or_unreadable_runs_fail(self) -> None:
        self.assertEqual(self.runs([{"name": "CodeQL", "conclusion": "success"}]),
                         ["Repository integrity has not run on aaaaaaaaaaaa"])
        self.assertEqual(self.runs([{"name": "Repository integrity", "conclusion": "failure"}]),
                         ["Repository integrity on aaaaaaaaaaaa did not succeed: failure"])
        self.assertEqual(self.runs([], returncode=1),
                         ["cannot read the check runs of aaaaaaaaaaaa: denied"])


REPOSITORY = "owner/project"
COMMIT = "c" * 40
BUNDLE_NAME = "project-1.2.0-source.tar.gz"
BUNDLE = b"the source bundle"
WRITES = {"-X", "--method", "-f", "-F", "--field", "--raw-field", "--input"}


def sums(content: bytes, name: str = BUNDLE_NAME) -> bytes:
    return f"{hashlib.sha256(content).hexdigest()}  {name}\n".encode()


def office_package(parts: dict[str, str] | None = None) -> bytes:
    """A minimal Office Open XML package with the given extra parts."""
    buffer = io.BytesIO()
    with zipfile.ZipFile(buffer, "w") as archive:
        archive.writestr("[Content_Types].xml", '<Types xmlns="http://schemas.openxmlformats.org/'
                                                'package/2006/content-types"/>')
        archive.writestr("xl/workbook.xml", "<workbook/>")
        for name, text in (parts or {}).items():
            archive.writestr(name, text)
    return buffer.getvalue()


# A workbook last saved to a OneDrive for Business folder, with a threaded-comment person.
PRIVATE_WORKBOOK = office_package({
    "xl/workbook.xml": ('<workbook xmlns:x15ac="http://schemas.microsoft.com/office/spreadsheetml/'
                        '2010/11/ac"><x15ac:absPath url="https://contoso-my.sharepoint.com/'
                        'personal/avery_contoso_com/Documents/"/></workbook>'),
    "xl/persons/person.xml": '<personList><person displayName="Avery Contoso"/></personList>'})


def asset(name: str, number: int) -> dict[str, object]:
    return {"name": name,
            "apiUrl": f"https://api.github.com/repos/{REPOSITORY}/releases/assets/{number}",
            "url": f"https://github.com/{REPOSITORY}/releases/download/v1.2.0/{name}"}


class FakeGh:
    """gh for one published release, as JSON fixtures that a test changes."""

    def __init__(self) -> None:
        self.tag_ref: dict[str, object] = {"object": {"type": "tag", "sha": "d" * 40}}
        self.tag: dict[str, object] = {
            "message": f"v1.2.0 - Release\n\nCandidate {COMMIT}.\nDeviations: none.\n",
            "object": {"type": "commit", "sha": COMMIT}}
        self.check_runs: list[dict[str, object]] = [
            {"name": "Repository integrity", "conclusion": "success"}]
        self.release: dict[str, object] | None = {
            "tagName": "v1.2.0", "isDraft": False, "publishedAt": "2026-11-04T10:00:00Z",
            "url": f"https://github.com/{REPOSITORY}/releases/tag/v1.2.0",
            "body": f"v1.2.0 - Release\n\nCandidate {COMMIT}.\nDeviations: none.\n",
            "assets": [asset(BUNDLE_NAME, 1), asset("SHA256SUMS", 2)]}
        self.files: dict[str, bytes] = {"1": BUNDLE, "2": sums(BUNDLE)}
        self.attested = True
        self.milestone: dict[str, object] = {"title": "v1.2.0 - Release", "state": "closed",
                                             "open_issues": 0}
        self.calls: list[list[str]] = []
        self.attested_files: list[tuple[Path, bytes]] = []

    @staticmethod
    def answer(stdout: bytes = b"", error: str = "") -> subprocess.CompletedProcess[bytes]:
        return subprocess.CompletedProcess([], 1 if error else 0, stdout=stdout,
                                           stderr=error.encode())

    def json(self, payload: object) -> subprocess.CompletedProcess[bytes]:
        return self.answer(json.dumps(payload).encode())

    def __call__(self, arguments: Sequence[str]) -> subprocess.CompletedProcess[bytes]:
        args = list(arguments)
        self.calls.append(args)
        base = f"repos/{REPOSITORY}/"
        if args[0] == "api":
            path = args[-1]
            # GitHub resolves an owner and a name in any letter case.
            if path.lower().startswith(base.lower()):
                path = base + path[len(base):]
            if path.lower() == base.rstrip("/").lower():
                return self.json({"full_name": REPOSITORY, "name": "project"})
            if path == base + "git/ref/tags/v1.2.0":
                return self.json(self.tag_ref)
            if path == base + "git/tags/" + "d" * 40:
                return self.json(self.tag)
            if path.startswith(base + "commits/"):
                return self.json({"check_runs": self.check_runs})
            if path.startswith(base + "releases/assets/"):
                number = path.rsplit("/", 1)[-1]
                if number in self.files:
                    return self.answer(self.files[number])
            if path == base + "milestones/7":
                return self.json(self.milestone)
            return self.answer(error="gh: Not Found (HTTP 404)")
        if args[:2] == ["release", "view"]:
            if self.release is None:
                return self.answer(error="release not found")
            return self.json(self.release)
        if args[:2] == ["attestation", "verify"]:
            self.attested_files.append((Path(args[2]), Path(args[2]).read_bytes()))
            if not self.attested:
                return self.answer(error="Error: no matching attestations found")
            return self.json([{"verificationResult": {}}])
        return self.answer(error=f"unexpected gh {' '.join(args)}")


class VerifyPublishedTest(unittest.TestCase):
    def setUp(self) -> None:
        self.gh = FakeGh()

    def facts(self, milestone: int | None = 7,
              extra: dict[str, str] | None = None) -> list[tuple[str, str]]:
        found = release.verify_published(REPOSITORY, "v1.2.0", milestone, extra, self.gh)
        for args in self.gh.calls:
            with self.subTest(call=args):
                self.assertIn(args[0] if args[0] == "api" else " ".join(args[:2]),
                              {"api", "release view", "attestation verify"})
                self.assertFalse(WRITES & set(args), args)
        return found

    def failures(self, milestone: int | None = 7,
                 extra: dict[str, str] | None = None) -> list[str]:
        return [text for status, text in self.facts(milestone, extra) if status == "FAIL"]

    def test_a_published_release_passes_with_one_line_per_fact(self) -> None:
        facts = self.facts()
        self.assertEqual([status for status, _ in facts], ["PASS"] * 8, facts)
        text = "\n".join(line for _, line in facts)
        for expected in (f"tag v1.2.0 is annotated and points at {COMMIT}",
                         f"Repository integrity succeeded on {COMMIT}",
                         "the release of v1.2.0 is published, not a draft",
                         f"the release notes name the candidate {COMMIT}",
                         f"exactly the expected assets: {BUNDLE_NAME}, SHA256SUMS",
                         f"{BUNDLE_NAME} has SHA-256 {hashlib.sha256(BUNDLE).hexdigest()}, "
                         "as SHA256SUMS records",
                         f"gh attestation verify passed on {BUNDLE_NAME}",
                         "milestone 7 (v1.2.0 - Release) is closed and has no open issue"):
            with self.subTest(expected=expected):
                self.assertIn(expected, text)

    def test_the_attestation_is_verified_on_a_removed_copy_of_the_published_bundle(
            self) -> None:
        self.facts()
        [(copy, content)] = self.gh.attested_files
        self.assertEqual(content, BUNDLE)
        self.assertFalse(copy.exists())
        call = next(args for args in self.gh.calls if args[0] == "attestation")
        for option in (["--repo", REPOSITORY],
                       ["--signer-workflow", f"{REPOSITORY}/.github/workflows/release.yml"],
                       ["--source-ref", "refs/tags/v1.2.0"], ["--format", "json"]):
            with self.subTest(option=option):
                index = call.index(option[0])
                self.assertEqual(call[index:index + 2], option)

    def test_a_draft_release_fails(self) -> None:
        assert self.gh.release is not None
        self.gh.release["isDraft"] = True
        self.assertEqual(self.failures(), ["the release of v1.2.0 is a draft, not published"])

    def test_a_missing_release_fails(self) -> None:
        self.gh.release = None
        self.assertEqual(self.failures(),
                         ["cannot read the release of v1.2.0: release not found"])

    def test_a_tag_that_points_elsewhere_than_its_candidate_fails(self) -> None:
        self.gh.tag["object"] = {"type": "commit", "sha": "e" * 40}
        self.assertIn(f"tag v1.2.0 points at {'e' * 40}, but its Candidate line names "
                      f"{COMMIT}", self.failures())

    def test_a_lightweight_or_missing_tag_fails(self) -> None:
        self.gh.tag_ref = {"object": {"type": "commit", "sha": COMMIT}}
        self.assertIn("tag v1.2.0 is commit, not an annotated tag", self.failures())
        self.gh.tag_ref = {}
        self.assertIn("tag v1.2.0 is missing, not an annotated tag", self.failures())
        self.assertEqual(release.verify_published(REPOSITORY, "1.2.0", runner=self.gh),
                         [("FAIL", "tag '1.2.0' is not vX.Y.Z")])

    def test_a_failed_integrity_run_fails(self) -> None:
        self.gh.check_runs = [{"name": "Repository integrity", "conclusion": "failure"}]
        self.assertEqual(self.failures(),
                         [f"Repository integrity on {COMMIT[:12]} did not succeed: failure"])

    def test_notes_that_do_not_name_the_candidate_fail(self) -> None:
        assert self.gh.release is not None
        self.gh.release["body"] = "Notes."
        self.assertEqual(self.failures(),
                         [f"the release notes do not name the candidate {COMMIT}"])

    def test_a_missing_or_unexpected_asset_fails(self) -> None:
        assert self.gh.release is not None
        self.gh.release["assets"] = [asset(BUNDLE_NAME, 1), asset("notes.txt", 3)]
        self.assertEqual(self.failures(), ["asset SHA256SUMS is missing from the release",
                                           "asset notes.txt is not an expected asset"])

    def test_an_altered_asset_fails(self) -> None:
        self.gh.files["1"] = b"an altered bundle"
        digest = hashlib.sha256(b"an altered bundle").hexdigest()
        self.assertIn(f"{BUNDLE_NAME} has SHA-256 {digest}, but SHA256SUMS records "
                      f"{hashlib.sha256(BUNDLE).hexdigest()}", self.failures())

    def test_checksums_must_record_exactly_the_bundle(self) -> None:
        self.gh.files["2"] = sums(BUNDLE) + sums(b"x", "other.zip")
        self.assertEqual(self.failures(), [f"SHA256SUMS records other.zip, {BUNDLE_NAME}, not "
                                           f"exactly {BUNDLE_NAME}"])

    def test_an_unreadable_asset_fails(self) -> None:
        del self.gh.files["2"]
        self.assertEqual(self.failures(),
                         ["cannot read asset SHA256SUMS: gh: Not Found (HTTP 404)"])

    def test_other_assets_are_expected_with_their_recorded_digests(self) -> None:
        assert self.gh.release is not None
        workbook = office_package()
        self.gh.release["assets"] = [asset(BUNDLE_NAME, 1), asset("SHA256SUMS", 2),
                                     asset("Workbook.xlsm", 4)]
        self.gh.files["4"] = workbook
        digest = hashlib.sha256(workbook).hexdigest()
        self.assertEqual(self.failures(extra={"Workbook.xlsm": digest}), [])
        self.assertEqual(self.failures(extra={"Workbook.xlsm": "0" * 64}),
                         [f"Workbook.xlsm has SHA-256 {digest}, not the recorded {'0' * 64}"])
        self.assertEqual(self.failures(), ["asset Workbook.xlsm is not an expected asset"])
        self.assertIn("asset Manual.pdf is missing from the release",
                      self.failures(extra={"Workbook.xlsm": digest, "Manual.pdf": "1" * 64}))

    def attach(self, name: str, content: bytes) -> str:
        """Attach another asset to the release; its SHA-256."""
        assert self.gh.release is not None
        self.gh.release["assets"] = [asset(BUNDLE_NAME, 1), asset("SHA256SUMS", 2),
                                     asset(name, 4)]
        self.gh.files["4"] = content
        return hashlib.sha256(content).hexdigest()

    def test_an_attached_workbook_is_checked_for_private_data(self) -> None:
        digest = self.attach("Workbook.xlsm", office_package())
        self.assertIn(("PASS", "Workbook.xlsm has no finding of tools/check_workbook_privacy.py"),
                      self.facts(extra={"Workbook.xlsm": digest}))

    def test_an_attached_workbook_with_private_data_fails_whatever_its_digest(self) -> None:
        digest = self.attach("Workbook.xlsm", PRIVATE_WORKBOOK)
        private = ["Workbook.xlsm carries private data: 2 finding(s) of "
                   "tools/check_workbook_privacy.py",
                   "Workbook.xlsm: xl/workbook.xml: path or URL (x15ac:absPath holds a OneDrive "
                   "for Business URL)",
                   "Workbook.xlsm: xl/persons/person.xml: person or organization "
                   "(person@displayName is not blank)"]
        self.assertEqual(self.failures(extra={"Workbook.xlsm": digest}), private)
        self.assertEqual(self.failures(extra={"Workbook.xlsm": "0" * 64}),
                         [f"Workbook.xlsm has SHA-256 {digest}, not the recorded {'0' * 64}",
                          *private])
        for _, text in self.facts(extra={"Workbook.xlsm": digest}):
            with self.subTest(text=text):
                self.assertNotIn("avery", text.lower())
                self.assertNotIn("contoso", text.lower())

    def test_an_asset_named_like_a_package_must_be_one(self) -> None:
        digest = self.attach("Workbook.xlsx", b"an encrypted workbook")
        [failure] = self.failures(extra={"Workbook.xlsx": digest})
        self.assertTrue(failure.startswith("Workbook.xlsx cannot be checked for private data: "
                                           "it is not a ZIP"), failure)
        digest = self.attach("Manual.pdf", b"%PDF-1.7 C:\\Users\\avery\\Manual.pdf")
        facts = self.facts(extra={"Manual.pdf": digest})
        self.assertEqual([text for status, text in facts if status == "FAIL"], [])
        self.assertFalse([text for _, text in facts if "check_workbook_privacy" in text], facts)

    def test_main_fails_on_an_attached_workbook_with_private_data(self) -> None:
        digest = self.attach("Workbook.xlsm", PRIVATE_WORKBOOK)
        arguments = ["verify-published", "v1.2.0", "--repository", REPOSITORY,
                     "--asset", f"Workbook.xlsm={digest}"]
        with mock.patch("release.run_gh", self.gh), \
                contextlib.redirect_stdout(io.StringIO()) as output:
            self.assertEqual(release.main(arguments), 1)
        lines = output.getvalue().splitlines()
        self.assertIn("FAIL: Workbook.xlsm carries private data: 2 finding(s) of "
                      "tools/check_workbook_privacy.py", lines)
        self.assertEqual(lines[-1], "release-verify-published: FAIL (v1.2.0)")

    def test_a_failed_attestation_fails(self) -> None:
        self.gh.attested = False
        self.assertEqual(self.failures(), [f"gh attestation verify failed on {BUNDLE_NAME}: "
                                           "Error: no matching attestations found"])

    def test_an_unattested_bundle_stated_in_the_notes_is_a_note(self) -> None:
        assert self.gh.release is not None
        self.gh.release["body"] = f"{self.gh.release['body']}\n{release.UNATTESTED}: this ...\n"
        self.gh.attested = False
        facts = self.facts()
        self.assertIn(("NOTE", f"{BUNDLE_NAME} is not attested, as the release notes state: "
                               "record its absence as a deviation"), facts)
        self.assertNotIn("FAIL", [status for status, _ in facts])
        self.assertEqual(self.gh.attested_files, [])

    def test_an_open_milestone_fails(self) -> None:
        self.gh.milestone = {"title": "v1.2.0 - Release", "state": "open", "open_issues": 0}
        self.assertEqual(self.failures(),
                         ["milestone 7 (v1.2.0 - Release) is open, with 0 open issue(s)"])
        self.gh.milestone["state"] = "closed"
        self.gh.milestone["open_issues"] = 2
        self.assertEqual(self.failures(),
                         ["milestone 7 (v1.2.0 - Release) is closed, with 2 open issue(s)"])
        self.assertEqual(self.failures(milestone=8),
                         ["cannot read milestone 8: gh: Not Found (HTTP 404)"])

    def test_the_repository_is_checked_under_its_canonical_name(self) -> None:
        found = release.verify_published("Owner/Project", "v1.2.0", 7, None, self.gh)
        self.assertEqual([status for status, _ in found], ["PASS"] * 8, found)
        self.assertIn(f"exactly the expected assets: {BUNDLE_NAME}, SHA256SUMS",
                      "\n".join(text for _, text in found))
        call = next(args for args in self.gh.calls if args[0] == "attestation")
        self.assertEqual(call[call.index("--repo") + 1], REPOSITORY)
        self.assertEqual(call[call.index("--signer-workflow") + 1],
                         f"{REPOSITORY}/.github/workflows/release.yml")
        self.gh.calls.clear()
        self.assertEqual(release.verify_published("owner/missing", "v1.2.0", runner=self.gh),
                         [("FAIL", "cannot read repository owner/missing: "
                                   "gh: Not Found (HTTP 404)")])

    def test_without_a_milestone_none_is_read(self) -> None:
        self.assertEqual(self.failures(milestone=None), [])
        self.assertFalse(any("milestones" in arg for args in self.gh.calls for arg in args))

    def test_main_prints_each_fact_and_exits(self) -> None:
        arguments = ["verify-published", "v1.2.0", "--repository", REPOSITORY,
                     "--milestone", "7"]
        with mock.patch("release.run_gh", self.gh), \
                contextlib.redirect_stdout(io.StringIO()) as output:
            self.assertEqual(release.main(arguments), 0)
            self.gh.attested = False
            self.assertEqual(release.main(arguments), 1)
        lines = output.getvalue().splitlines()
        self.assertEqual(lines[8], "release-verify-published: PASS (v1.2.0)")
        self.assertTrue(all(line.startswith("PASS: ") for line in lines[:8]), lines)
        self.assertIn(f"FAIL: gh attestation verify failed on {BUNDLE_NAME}: "
                      "Error: no matching attestations found", lines)
        self.assertEqual(lines[-1], "release-verify-published: FAIL (v1.2.0)")

    def test_an_asset_record_is_file_equals_sha256(self) -> None:
        self.assertEqual(release.asset_record("a=b.xlsm=" + "A" * 64), ("a=b.xlsm", "a" * 64))
        for value in ("Workbook.xlsm", "=" + "a" * 64, "Workbook.xlsm=abc"):
            with self.subTest(value=value), self.assertRaises(argparse.ArgumentTypeError):
                release.asset_record(value)


class WorkflowTest(unittest.TestCase):
    def test_the_workflow_states_an_unattested_bundle_as_verify_published_reads_it(
            self) -> None:
        text = (ROOT / ".github/workflows/release.yml").read_text(encoding="utf-8")
        self.assertIn(release.UNATTESTED, text)
        self.assertIn(f"dist/{release.SUMS}\n", text)

    def test_the_workflow_calls_the_tool_it_is_tested_with(self) -> None:
        text = (ROOT / ".github/workflows/release.yml").read_text(encoding="utf-8")
        for command in ("python tools/release.py verify", "python tools/release.py integrity",
                        "python tools/release.py bundle", "python tools/release.py notes",
                        "gh release create \"$TAG\" --draft --verify-tag"):
            with self.subTest(command=command):
                self.assertIn(command, text)

    @unittest.skipIf(yaml is None, "needs PyYAML from .github/requirements.txt")
    def test_a_repository_without_attestations_still_gets_its_draft(self) -> None:
        assert yaml is not None
        jobs = yaml.safe_load((ROOT / ".github/workflows/release.yml")
                              .read_text(encoding="utf-8"))["jobs"]
        attest, draft = jobs["attest"]["if"], jobs["draft"]["if"]
        self.assertIn("github.event.repository.private == false", attest)
        self.assertIn("vars.RELEASE_ATTESTATIONS == 'enabled'", attest)
        self.assertEqual(jobs["draft"]["needs"], ["build", "attest"])
        for condition in ("!cancelled()", "needs.build.result == 'success'",
                          "needs.attest.result == 'success'",
                          "needs.attest.result == 'skipped'"):
            with self.subTest(condition=condition):
                self.assertIn(condition, draft)
        script = jobs["draft"]["steps"][-1]["run"]
        self.assertIn('if [ "$ATTESTED" != "success" ]', script)
        self.assertLess(script.index("no build provenance attestation"),
                        script.index("gh release create"))


if __name__ == "__main__":
    unittest.main()
