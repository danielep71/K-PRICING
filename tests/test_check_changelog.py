"""Valid and invalid CHANGELOG.md and VERSION fixtures for tools/check_changelog.py.

Each fixture is a temporary directory holding the two files, plus RELEASING.md
or README.md where a test gives them; some are Git repositories with tags.
The versions and dates are synthetic; no release of this repository is
implied.
"""

from __future__ import annotations

import contextlib
import io
import shutil
import sys
import tempfile
import unittest
from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(ROOT / "tools"))
sys.path.insert(0, str(ROOT / "tests"))

import check_changelog  # noqa: E402
from gitfixture import git, init  # noqa: E402

REPOSITORY = "example/repository"
TAGS = f"https://github.com/{REPOSITORY}/releases/tag"
UNRELEASED_ONLY = (
    "# Changelog\n\nNo version has been released.\n\n## Policy\n\nText.\n\n"
    "## [Unreleased]\n\n### Added\n\n- A change.\n\n"
    f"[Unreleased]: https://github.com/{REPOSITORY}/commits/main\n"
)
RELEASED = (
    "# Changelog\n\n## [Unreleased]\n\n## [0.2.0] - 2026-11-02\n\n### Fixed\n\n- A fix.\n\n"
    "## [0.1.0] - 2026-10-30\n\n### Added\n\n- A start.\n\n"
    f"[Unreleased]: https://github.com/{REPOSITORY}/compare/v0.2.0...HEAD\n"
    f"[0.2.0]: {TAGS}/v0.2.0\n[0.1.0]: {TAGS}/v0.1.0\n"
)


STATUS_BEFORE = ("# Releasing\n\n## 🧭 Current status\n\n| Item | State |\n| --- | --- |\n"
                 "| Released versions | None |\n"
                 "| `VERSION` | Does not exist yet; it is created when the first release is "
                 "prepared |\n\n## 🏷️ Versioning\n\n| Item | Convention |\n| --- | --- |\n"
                 "| `VERSION` | One line |\n")
STATUS_AFTER = (STATUS_BEFORE.replace("| Released versions | None |",
                                      "| Released versions | 0.2.0, 0.1.0 |")
                .replace("Does not exist yet; it is created when the first release is prepared",
                         "0.2.0"))


def findings(changelog: str | None, version: str | None = None,
             repository: str | None = None, **files: str) -> list[str]:
    """Check a directory holding the given CHANGELOG.md, VERSION and other files.

    files maps RELEASING and README to the text of RELEASING.md and README.md.
    """
    with tempfile.TemporaryDirectory(prefix="changelog-") as temporary:
        root = Path(temporary)
        if changelog is not None:
            (root / "CHANGELOG.md").write_text(changelog, encoding="utf-8")
        if version is not None:
            (root / "VERSION").write_text(version, encoding="utf-8")
        for name, text in files.items():
            (root / f"{name}.md").write_text(text, encoding="utf-8")
        return check_changelog.check(root, repository)[0]


class ChangelogTest(unittest.TestCase):
    def assertFinding(self, changelog: str | None, text: str, version: str | None = None,
                      repository: str | None = None) -> None:
        found = findings(changelog, version, repository)
        self.assertTrue(any(text in item for item in found), found)

    def test_valid_files_pass(self) -> None:
        self.assertEqual(findings(UNRELEASED_ONLY), [])
        self.assertEqual(findings(RELEASED, "0.2.0\n"), [])
        self.assertEqual(findings(RELEASED, "0.2.0\n", "Example/Repository"), [])

    def test_repository_files_pass(self) -> None:
        found, _ = check_changelog.check(ROOT)
        self.assertEqual(found, [])

    def test_missing_changelog(self) -> None:
        self.assertFinding(None, "CHANGELOG.md is missing")

    def test_unreleased_comes_first_without_a_date(self) -> None:
        self.assertFinding(RELEASED.replace("## [Unreleased]\n\n", ""),
                           "first version heading must be '## [Unreleased]'", "0.2.0\n")
        self.assertFinding(UNRELEASED_ONLY.replace("## [Unreleased]", "## [Unreleased] - 2026-10-30"),
                           "takes no date or suffix")
        self.assertFinding(UNRELEASED_ONLY.replace("[Unreleased]: ", "[Other]: "),
                           "missing link reference for [Unreleased]")

    def test_release_headings(self) -> None:
        for heading, text in (("## [0.2.0]", "must be '## [X.Y.Z] - YYYY-MM-DD'"),
                              ("## [v0.2.0] - 2026-11-02", "must be '## [X.Y.Z] - YYYY-MM-DD'"),
                              ("## [0.2] - 2026-11-02", "must be '## [X.Y.Z] - YYYY-MM-DD'"),
                              ("## [0.2.0] - 2026-02-30", "is not a calendar date")):
            with self.subTest(heading=heading):
                self.assertFinding(RELEASED.replace("## [0.2.0] - 2026-11-02", heading), text,
                                   "0.2.0\n")

    def test_unbracketed_version_headings(self) -> None:
        for heading in ("## 0.2.0 - 2026-11-02", "## v0.2.0 - 2026-11-02", "## Unreleased",
                        "##  [0.2.0] - 2026-11-02"):
            with self.subTest(heading=heading):
                self.assertFinding(RELEASED.replace("## [0.2.0] - 2026-11-02", heading),
                                   "version heading must be", "0.1.0\n")

    def test_order_and_duplicates(self) -> None:
        swapped = RELEASED.replace("0.2.0", "X").replace("0.1.0", "0.2.0").replace("X", "0.1.0")
        self.assertFinding(swapped, "newest first", "0.1.0\n")
        self.assertFinding(RELEASED.replace("2026-10-30", "2026-11-05"), "newest first", "0.2.0\n")
        self.assertFinding(RELEASED + "\n## [Unreleased]\n", "duplicate heading [Unreleased]",
                           "0.2.0\n")

    def test_release_link_references(self) -> None:
        self.assertFinding(RELEASED.replace(f"[0.1.0]: {TAGS}/v0.1.0\n", ""),
                           "missing link reference for [0.1.0]", "0.2.0\n")
        self.assertFinding(RELEASED.replace(f"{TAGS}/v0.1.0", f"{TAGS}/v0.0.9"),
                           "must end with /releases/tag/v0.1.0", "0.2.0\n")

    def test_link_references_point_to_the_repository(self) -> None:
        # Without --repository only the format is checked.
        self.assertEqual(findings(RELEASED, "0.2.0\n"), [])
        for label in ("Unreleased", "0.2.0", "0.1.0"):
            with self.subTest(label=label):
                self.assertFinding(RELEASED, f"the link reference for [{label}] does not point "
                                   "to other/project", "0.2.0\n", "other/project")
        copied = UNRELEASED_ONLY.replace(REPOSITORY, "example/repository-template")
        self.assertFinding(copied, "does not point to example/repository", None, REPOSITORY)

    def test_every_github_link_reference_points_to_the_repository(self) -> None:
        # Not only the version headings: a copied issue link fails too.
        foreign = UNRELEASED_ONLY + "[#12]: https://github.com/other/template/issues/12\n"
        self.assertFinding(foreign, "the link reference for [#12] does not point to "
                           f"{REPOSITORY}", None, REPOSITORY)
        self.assertEqual(findings(foreign), [])
        own = UNRELEASED_ONLY + f"[#12]: https://github.com/{REPOSITORY}/issues/12\n"
        outside = UNRELEASED_ONLY + "[semver]: https://semver.org/\n"
        for text in (own, outside):
            with self.subTest(text=text.splitlines()[-1]):
                self.assertEqual(findings(text, None, REPOSITORY), [])


class ReleaseStateTest(unittest.TestCase):
    def assertFinding(self, found: list[str], text: str) -> None:
        self.assertTrue(any(text in item for item in found), found)

    def test_consistent_state_passes(self) -> None:
        readme = "# Project\n\n- [x] Done\n"
        self.assertEqual(findings(UNRELEASED_ONLY, RELEASING=STATUS_BEFORE, README=readme), [])
        self.assertEqual(findings(RELEASED, "0.2.0\n", RELEASING=STATUS_AFTER, README=readme),
                         [])

    def test_changelog_release_statement(self) -> None:
        self.assertFinding(findings(RELEASED.replace("# Changelog\n", "# Changelog\n\n"
                                                     "No version has been released.\n"),
                                    "0.2.0\n"),
                           "the introduction says 'No version has been released' but the "
                           "changelog releases [0.2.0]")
        self.assertFinding(findings(UNRELEASED_ONLY.replace("No version has been released.", "")),
                           "its introduction does not say 'No version has been released'")

    def test_release_statement_is_read_from_the_introduction_only(self) -> None:
        # An entry may quote the statement after the first release; the
        # statement outside the introduction does not count before it either.
        quoted = "- The check reads the 'No version has been released' sentence.\n"
        self.assertEqual(findings(RELEASED.replace("- A fix.\n", "- A fix.\n" + quoted),
                                  "0.2.0\n"), [])
        moved = UNRELEASED_ONLY.replace("No version has been released.\n\n", "").replace(
            "- A change.\n", "- A change.\n" + quoted)
        self.assertFinding(findings(moved), "its introduction does not say")

    def test_placeholder_lines(self) -> None:
        for line in ("- TBD", "_Nothing yet._", "- TODO", "..."):
            with self.subTest(line=line):
                self.assertFinding(findings(UNRELEASED_ONLY.replace("- A change.", line)),
                                   f"placeholder line in a version section: {line}")
        # A placeholder word inside a sentence is not a placeholder line.
        self.assertEqual(findings(UNRELEASED_ONLY.replace("- A change.", "- Remove a TODO.")), [])

    def test_releasing_current_status(self) -> None:
        self.assertFinding(findings(RELEASED, "0.2.0\n", RELEASING=STATUS_BEFORE),
                           "'Released versions' does not name the newest release [0.2.0]: None")
        self.assertFinding(findings(RELEASED, "0.2.0\n", RELEASING=STATUS_BEFORE),
                           "'`VERSION`' does not name the newest release [0.2.0]")
        for other in ("10.2.0", "0.2.0.1", "v0.2.00", "0.2.0-rc.1", "0.2.0b"):
            with self.subTest(other=other):
                wrong = STATUS_AFTER.replace("0.2.0", other)
                self.assertFinding(findings(RELEASED, "0.2.0\n", RELEASING=wrong),
                                   "'Released versions' does not name the newest release")
        for named in ("0.2.0.", "v0.2.0", "`0.2.0`", "0.2.0 (2026-11-02)"):
            with self.subTest(named=named):
                status = STATUS_AFTER.replace("0.2.0, 0.1.0", named).replace(
                    "| `VERSION` | 0.2.0 |", f"| `VERSION` | {named} |")
                self.assertEqual(findings(RELEASED, "0.2.0\n", RELEASING=status), [])
        self.assertFinding(findings(UNRELEASED_ONLY, RELEASING=STATUS_AFTER),
                           "'Released versions' must be None before the first release")
        self.assertFinding(findings(UNRELEASED_ONLY, RELEASING=STATUS_AFTER),
                           "'`VERSION`' must say it does not exist before the first release")
        self.assertFinding(findings(UNRELEASED_ONLY, RELEASING=STATUS_BEFORE.replace(
            "| Released versions | None |\n", "")), "Current status has no 'Released versions' row")
        self.assertFinding(findings(UNRELEASED_ONLY, RELEASING="# Releasing\n"),
                           "RELEASING.md: has no Current status section")

    def test_readme_release_state(self) -> None:
        self.assertFinding(findings(UNRELEASED_ONLY, README="Badge: Release_Candidate\n"),
                           "README.md: mentions a Release Candidate")
        self.assertFinding(findings(UNRELEASED_ONLY, README="## Release\n\n- [ ] Tag it\n"),
                           "README.md: holds an unchecked task-list item")


class VersionTest(unittest.TestCase):
    def test_version_only_from_the_first_release(self) -> None:
        self.assertIn("VERSION exists but CHANGELOG.md has no dated release; it is created "
                      "when a release is prepared (RELEASING.md)",
                      findings(UNRELEASED_ONLY, "0.1.0\n"))
        self.assertIn("VERSION is missing; CHANGELOG.md releases [0.2.0]", findings(RELEASED))

    def test_version_names_the_newest_release(self) -> None:
        self.assertIn("VERSION 0.1.0 differs from the newest CHANGELOG.md release [0.2.0]",
                      findings(RELEASED, "0.1.0\n"))
        for raw in ("0.2.0", "v0.2.0\n", "0.2.0\n\n", "0.2\n"):
            with self.subTest(raw=raw):
                self.assertIn("VERSION must hold one X.Y.Z line ending in a newline",
                              findings(RELEASED, raw))

    def test_main_reports_the_verdict(self) -> None:
        for version, code, verdict in (("0.2.0\n", 0, "changelog: PASS (2 dated release(s)"),
                                       ("0.1.0\n", 1, "changelog: FAIL (2 dated release(s)")):
            with self.subTest(version=version), \
                    tempfile.TemporaryDirectory(prefix="changelog-") as temporary, \
                    contextlib.redirect_stdout(io.StringIO()) as output:
                (Path(temporary) / "CHANGELOG.md").write_text(RELEASED, encoding="utf-8")
                (Path(temporary) / "VERSION").write_text(version, encoding="utf-8")
                self.assertEqual(check_changelog.main(["--root", temporary,
                                                       "--repository", REPOSITORY]), code)
                self.assertIn(verdict, output.getvalue())


@unittest.skipUnless(shutil.which("git"), "needs Git")
class FrozenSectionTest(unittest.TestCase):
    """Released sections against CHANGELOG.md at their tags, in temporary repositories."""

    def frozen(self, edit: str | None = None, tags: tuple[str, ...] = ("v0.1.0", "v0.2.0"),
               tagged: str = RELEASED) -> tuple[list[str], list[str]]:
        """Commit ``tagged`` with the tags, then check ``edit`` (or ``tagged``) against them."""
        with tempfile.TemporaryDirectory(prefix="changelog-") as temporary:
            root = Path(temporary)
            (root / "CHANGELOG.md").write_text(tagged, encoding="utf-8", newline="")
            (root / "VERSION").write_text("0.2.0\n", encoding="utf-8")
            init(root)
            git(root, "add", "-A")
            git(root, "commit", "-qm", "Release")
            for tag in tags:
                git(root, "tag", "-a", tag, "-m", tag)
            if edit is not None:
                (root / "CHANGELOG.md").write_text(edit, encoding="utf-8", newline="")
            found, _, notes = check_changelog.check_with_notes(root, REPOSITORY)
            return found, notes

    def test_an_unchanged_release_passes(self) -> None:
        self.assertEqual(self.frozen(), ([], []))

    def test_a_new_release_and_line_endings_do_not_unfreeze_older_sections(self) -> None:
        newer = RELEASED.replace(
            "## [Unreleased]\n\n", "## [Unreleased]\n\n## [0.3.0] - 2026-11-20\n\n"
            "### Added\n\n- More.\n\n").replace(
            "compare/v0.2.0...HEAD\n", f"compare/v0.3.0...HEAD\n[0.3.0]: {TAGS}/v0.3.0\n")
        found, notes = self.frozen(newer.replace("\n", "\r\n"))
        self.assertEqual(notes, ["CHANGELOG.md: [0.3.0] has no tag v0.3.0 yet, so it is not "
                                 "frozen"])
        self.assertEqual([item for item in found if "tag" in item], [])

    def test_an_edited_release_fails_with_its_line(self) -> None:
        cases = {"- A start.": ("- A different start.", "CHANGELOG.md:15: [0.1.0]"),
                 "## [0.2.0] - 2026-11-02": ("## [0.2.0] - 2026-11-03", "CHANGELOG.md:5: [0.2.0]"),
                 "### Fixed\n\n- A fix.\n": ("### Fixed\n\n- A fix.\n- Another.\n",
                                              "CHANGELOG.md:10: [0.2.0]")}
        for old, (new, where) in cases.items():
            with self.subTest(old=old):
                found, _ = self.frozen(RELEASED.replace(old, new))
                self.assertEqual(found, [f"{where} differs from its content at tag v"
                                         f"{where.split('[')[1][:5]}; a released section is "
                                         "frozen (add 'Corrected after release on YYYY-MM-DD: "
                                         "<reason>' only "
                                         "for a reviewed correction)"])

    def test_a_deleted_or_renamed_release_fails(self) -> None:
        deleted = RELEASED.replace("## [0.1.0] - 2026-10-30\n\n### Added\n\n- A start.\n\n",
                                   "").replace(f"[0.1.0]: {TAGS}/v0.1.0\n", "")
        gone = ("CHANGELOG.md: [0.1.0] was released with tag v0.1.0 but its section is gone; "
                "a released section is frozen")
        self.assertEqual(self.frozen(deleted), ([gone], []))
        renamed = RELEASED.replace("[0.1.0]", "[0.1.1]").replace("/v0.1.0", "/v0.1.1")
        found, notes = self.frozen(renamed)
        self.assertEqual(found, [gone])
        self.assertEqual(notes, ["CHANGELOG.md: [0.1.1] has no tag v0.1.1 yet, so it is not "
                                 "frozen"])
        # A tag that no released section ever named is not a release of this changelog.
        self.assertEqual(self.frozen(tags=("v0.1.0", "v0.2.0", "v9.9.9", "vnext")), ([], []))

    def test_an_older_release_checked_out_ignores_newer_tags(self) -> None:
        first = (
            "# Changelog\n\n## [Unreleased]\n\n## [0.1.0] - 2026-10-30\n\n### Added\n\n"
            f"- A start.\n\n[Unreleased]: https://github.com/{REPOSITORY}/compare/v0.1.0...HEAD\n"
            f"[0.1.0]: {TAGS}/v0.1.0\n")
        with tempfile.TemporaryDirectory(prefix="changelog-") as temporary:
            root = Path(temporary)
            init(root)
            for text, version in ((first, "0.1.0"), (RELEASED, "0.2.0")):
                (root / "CHANGELOG.md").write_text(text, encoding="utf-8")
                (root / "VERSION").write_text(f"{version}\n", encoding="utf-8")
                git(root, "add", "-A")
                git(root, "commit", "-qm", f"Release {version}")
                git(root, "tag", "-a", f"v{version}", "-m", f"v{version}")
            git(root, "checkout", "-q", "v0.1.0")
            self.assertEqual(check_changelog.check_with_notes(root, REPOSITORY), ([], 1, []))
            git(root, "checkout", "-q", "v0.2.0")
            self.assertEqual(check_changelog.check_with_notes(root, REPOSITORY), ([], 2, []))
            (root / "CHANGELOG.md").write_text(
                RELEASED.replace("## [0.1.0] - 2026-10-30\n\n### Added\n\n- A start.\n\n", "")
                .replace(f"[0.1.0]: {TAGS}/v0.1.0\n", ""), encoding="utf-8")
            found, _, _ = check_changelog.check_with_notes(root, REPOSITORY)
            self.assertIn("CHANGELOG.md: [0.1.0] was released with tag v0.1.0 but its section "
                          "is gone; a released section is frozen", found)

    def test_a_dated_correction_is_a_note(self) -> None:
        note = "Corrected after release on 2026-11-05: the entry named the wrong file."
        corrected = RELEASED.replace("- A start.\n", f"- A corrected start.\n\n{note}\n")
        self.assertEqual(self.frozen(corrected), (
            [], [f"CHANGELOG.md: [0.1.0] differs from its tag v0.1.0: {note}"]))
        for invalid in ("Corrected after release: no date.",
                        "Corrected after release on 2026-02-30: not a calendar date."):
            with self.subTest(note=invalid):
                found, _ = self.frozen(RELEASED.replace("- A start.\n",
                                                        f"- A corrected start.\n\n{invalid}\n"))
                self.assertEqual(len(found), 1)
                self.assertIn("[0.1.0] differs from its content at tag v0.1.0", found[0])

    def test_an_untagged_release_is_a_note(self) -> None:
        self.assertEqual(self.frozen(tags=("v0.2.0",)), (
            [], ["CHANGELOG.md: [0.1.0] has no tag v0.1.0 yet, so it is not frozen"]))

    def test_a_tag_whose_changelog_lacks_the_section_fails(self) -> None:
        found, _ = self.frozen(RELEASED, tagged=UNRELEASED_ONLY)
        self.assertIn("CHANGELOG.md: [0.2.0] is not in CHANGELOG.md at its tag v0.2.0", found)

    def test_without_git_nothing_is_compared(self) -> None:
        with tempfile.TemporaryDirectory(prefix="changelog-") as temporary:
            root = Path(temporary)
            (root / "CHANGELOG.md").write_text(RELEASED, encoding="utf-8")
            (root / "VERSION").write_text("0.2.0\n", encoding="utf-8")
            self.assertEqual(check_changelog.frozen_problems(root, RELEASED, ["0.2.0"]), (
                [], ["CHANGELOG.md: not a Git checkout, so released sections were not compared "
                     "with their tags"]))


if __name__ == "__main__":
    unittest.main()
