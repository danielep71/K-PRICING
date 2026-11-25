"""Tests for tools/initialize_repository.py.

Each adaptation is tested on a small text, and each option also on a small
tree, with a negative control. One test copies the tracked tree to a
temporary Git repository, initializes it with every option, --remove-traffic
and --assignee included, and runs the Repository integrity checks there with
--template false, as a generated repository's CI does. The initializer has
deleted itself and this test by then, as it does once it has run. That run
leaves out three checks: the test suite, which would run this test again
inside itself, and Ruff and mypy. They read only the Python code of tools/
and tests/, which the initializer only deletes, never edits, so the
template's own run of them already covers the copy. In place of the suite it
runs the test modules that compare the adapted files with each other; the
integrity checks of the copy cover the rest. Once a repository is
initialized, this test is gone, and the copy test is skipped anywhere else.
"""

from __future__ import annotations

import contextlib
import io
import json
import os
import shutil
import subprocess
import sys
import tempfile
import unittest
from pathlib import Path
from unittest import mock

ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(ROOT / "tools"))
sys.path.insert(0, str(ROOT / "tests"))

import check_changelog  # noqa: E402
import check_issues  # noqa: E402
import check_repository  # noqa: E402
import gitio  # noqa: E402
import initialize_repository as initialize  # noqa: E402
import vba_lexer  # noqa: E402
from check_repository import Check, Context  # noqa: E402
from gitfixture import git, init  # noqa: E402
from initialize_repository import InitializeError, Options  # noqa: E402

REPOSITORY = "owner/new-workbook"
OPTIONS = Options(REPOSITORY, "LCR", "release/0.1.0")
# The checks a generated repository fails until it is initialized.
CLEANUP_CHECKS = ("template-only-content", "issue-contact-links", "readme-traffic-badges",
                  "dependabot-target-branches", "changelog")
# Read only code the initializer never changes; see the module docstring.
NOT_RUN = {"tests", "ruff", "mypy"}
ADAPTED_FILE_TESTS = ("test_repository_consistency", "test_workflow_controls",
                      "test_check_issues", "test_setup_repository", "test_check_doc_commands")
INITIALIZER = initialize.INITIALIZER
FORMS = [f"{initialize.ISSUE_FORMS}/{name}" for name in sorted(check_issues.FORM_PRIORITIES)]
# Only the template still holds the template-only section; once it is
# initialized, a copy of this tree has nothing left to adapt.
IS_TEMPLATE = (ROOT / "AGENTS.md").is_file() and \
    initialize.SECTION_START in (ROOT / "AGENTS.md").read_text(encoding="utf-8")

AGENTS = """# Instructions

- A rule.

<!-- template-only:start -->
## Template maintenance

- A template rule.
<!-- template-only:end -->
"""

CONFIG = """blank_issues_enabled: false
contact_links:
  - name: "Security"
    url: "https://github.com/template-owner/template/security/policy"
  - name: "Labels"
    url: "https://github.com/template-owner/template/blob/HEAD/docs/LABELS.md"
"""

BADGES = "https://raw.githubusercontent.com/template-owner/template/traffic-history/data/"
README = f"""# Template

[![Clones]({BADGES}clones.json)](docs/TRAFFIC.md)
[![Visitors]({BADGES}visitors.json)](docs/TRAFFIC.md)

I use this template.
"""

DEPENDABOT = """version: 2
updates:
  - package-ecosystem: pip
    directory: /tools
    target-branch: release/1.0.0
    schedule:
      interval: weekly
  - package-ecosystem: github-actions
    directory: /
    target-branch: "release/1.0.0"
"""

# The template after its first release: a dated section, VERSION, and a
# Current status that names the release.
CHANGELOG = """# Changelog

Notable changes. It is neither a commit log nor release evidence.

## Maintenance policy

- Add a material change under `[Unreleased]`.

## [Unreleased]

### Added

- Something new (#9).

## [1.0.0] - 2026-12-01

### Added

- The first release (#1).

[Unreleased]: https://github.com/template-owner/template/compare/v1.0.0...HEAD
[1.0.0]: https://github.com/template-owner/template/releases/tag/v1.0.0
"""

RELEASING = """# Releasing

## Current status

| Item | State |
| --- | --- |
| Released versions | 1.0.0 |
| `VERSION` | 1.0.0 |
"""

HOUSE_STYLE = """# House style

## Naming

Replace `<REGIME>` with the project's short name, such as `SACCR` or `LCR`,
once, in this table. The checks accept any prefix while it is `<REGIME>_`.

| Element | Convention | Example |
| --- | --- | --- |
| Public facade module | `<REGIME>_` plus subject | `<REGIME>_Functions` |
| Enum and member | `<REGIME>` plus noun | `<REGIME>RunStatus`, `rsFailed` |

A later paragraph that names `<REGIME>` keeps it.
"""

# The template's evidence: a tested environment and a host behaviour row that
# cite its own commit and issue.
INSTALLATION = """# Installation

Record each host on which a commit has been compiled and run:

| Tested environment | Commit | Evidence |
| --- | --- | --- |
| Excel 16.0, 64-bit | `0123456789abcdef0123456789abcdef01234567` | #204: `excel-run: PASS` |
| Excel 16.0, 32-bit | `fedcba9876543210fedcba9876543210fedcba98` | #300: `excel-run: PASS` |

## Next section
"""

VALIDATION_PLAN = """# Validation plan

| ID | Behaviour | Status | Probe | Host | Date | Result | Issue |
| --- | --- | --- | --- | --- | --- | --- | --- |
| `HB-01` | A name is re-cased | `assumed` | By hand | \u2014 | \u2014 | Not measured | #245 |
| `HB-02` | Macros start disabled | `measured` | `Run` it | Excel 16.0 | 2026-11-03 | It failed | #204 |

| Other | Issue |
| --- | --- |
| Not host behaviour | #9 |
"""


# A section of whole lines with a replacement, one within a paragraph with a
# replacement, and one within a table cell without one.
README_SECTIONS = """# Project

[Intro](#intro) \u00b7<!-- template-only:start -->
[Start](#start)<!-- template-only:replacement
[Guides](#guides)--><!-- template-only:end -->
\u00b7 [End](#end)

<!-- template-only:start -->
## Start

Template steps.
<!-- template-only:replacement
## Guides

| I want to | Read |
| --- | --- |
| Build | [Guide](GUIDE.md) |
-->
<!-- template-only:end -->

## Later

| Document | Covers |
| --- | --- |
| [Guide](GUIDE.md) | All<!-- template-only:start --> |
| [Wiki](https://example.com/wiki) | Tour<!-- template-only:end --> |
"""

# The traffic collection and what names it, in small.
TRAFFIC_TREE = {
    **{path: "traffic\n" for path in initialize.TRAFFIC_FILES},
    "README.md": "# Project\n\n| Area | What |\n| --- | --- |\n| Traffic | Clones |\n"
                 "| Links | Checked |\n",
    initialize.SETTINGS_CONFIG: json.dumps({"rulesets": [
        {"name": "Protect main", "conditions": {"ref_name": {"include": ["~DEFAULT_BRANCH"]}}},
        {"name": "Protect traffic history",
         "conditions": {"ref_name": {"include": [initialize.TRAFFIC_REF]}}}]}, indent=2) + "\n",
    initialize.SETUP_GUIDE: "| **Protection \u00b7 2 rulesets** | Default branch. |\n",
    initialize.VERIFICATION_DEPTH: json.dumps({"checks": [
        {"check": "changelog"}, {"check": initialize.TRAFFIC_CHECK}]}, indent=2) + "\n",
    initialize.NOTICE: f"Files:\n\n{check_repository.TRAFFIC_WORKFLOW}\nSECURITY.md\n",
    initialize.PROVENANCE: json.dumps({"sources": [{"files": [
        "SECURITY.md", check_repository.TRAFFIC_WORKFLOW]}]}, indent=2) + "\n",
    "docs/GUIDE.md": "Collect the traffic weekly.\n",
}

# The issue forms, Dependabot and the label guide, which name the assignee.
ASSIGNEE_TREE = {
    **{form: 'name: "Form"\nlabels: ["P2"]\nassignees: ["maintainer"]\nbody: []\n'
       for form in FORMS},
    initialize.DEPENDABOT: "version: 2\nupdates:\n  - package-ecosystem: pip\n"
                           "    assignees:\n      - maintainer\n      - other\n"
                           "    commit-message:\n      prefix: ci\n",
    initialize.ISSUE_LABELS: "My account, `maintainer`, is the assignee.\n",
}


def write(root: Path, files: dict[str, str]) -> None:
    for name, text in files.items():
        path = root / name
        path.parent.mkdir(parents=True, exist_ok=True)
        path.write_text(text, encoding="utf-8")


class AdaptationTests(unittest.TestCase):
    def assertIdempotent(self, adapt: initialize.Adapter, text: str,
                         options: Options = OPTIONS) -> str:
        adapted, notes = adapt(text, options)
        self.assertTrue(notes)
        self.assertEqual(adapt(adapted, options), (adapted, []))
        return adapted

    def test_agents_loses_the_template_only_section_and_its_markers(self) -> None:
        adapt = initialize.template_only("AGENTS.md")
        adapted = self.assertIdempotent(adapt, AGENTS)
        self.assertEqual(adapted, "# Instructions\n\n- A rule.\n")
        middle = AGENTS.replace("<!-- template-only:end -->\n",
                                "<!-- template-only:end -->\n\n## Later\n\n- Kept.\n")
        self.assertEqual(adapt(middle, OPTIONS)[0],
                         "# Instructions\n\n- A rule.\n\n## Later\n\n- Kept.\n")
        with self.assertRaisesRegex(InitializeError, "AGENTS.md: .* has no "
                                    "<!-- template-only:end -->"):
            adapt("<!-- template-only:start -->\n", OPTIONS)
        # Without a section, nothing changes.
        self.assertEqual(adapt("# Instructions\n", OPTIONS), ("# Instructions\n", []))

    def test_a_section_with_a_replacement_is_replaced_by_it(self) -> None:
        adapt = initialize.template_only("README.md")
        adapted, notes = adapt(README_SECTIONS, OPTIONS)
        self.assertEqual(adapted, (
            "# Project\n\n[Intro](#intro) \u00b7\n[Guides](#guides)\n\u00b7 [End](#end)\n\n"
            "## Guides\n\n| I want to | Read |\n| --- | --- |\n| Build | [Guide](GUIDE.md) |\n\n"
            "## Later\n\n| Document | Covers |\n| --- | --- |\n| [Guide](GUIDE.md) | All |\n"))
        self.assertEqual(notes, [
            "replace 2 template-only section(s) with the replacement each holds, with their "
            "markers",
            "delete 1 template-only section(s), with their markers"])
        self.assertEqual(adapt(adapted, OPTIONS), (adapted, []))
        # The replacement must close just before the end marker.
        unclosed = README_SECTIONS.replace("[Guides](#guides)-->", "[Guides](#guides)")
        with self.assertRaisesRegex(InitializeError, "README.md: <!-- template-only:replacement "
                                    "is not closed by -->"):
            adapt(unclosed, OPTIONS)

    @unittest.skipUnless(IS_TEMPLATE, "this repository is already initialized")
    def test_the_real_readme_gets_a_table_of_guides_in_place_of_its_template_sections(
            self) -> None:
        text = (ROOT / "README.md").read_text(encoding="utf-8")
        self.assertIn(initialize.SECTION_START, text)
        adapted = self.assertIdempotent(initialize.template_only("README.md"), text)
        self.assertNotIn("template-only", adapted)
        self.assertNotIn("/wiki", adapted)
        self.assertNotIn("## \U0001f680 Start a new repository", adapted)
        self.assertIn("\n[Find your way](#-find-your-way)\n", adapted)
        table = adapted.split("## \U0001f4cd Find your way\n\n", 1)[1].split("\n\n", 1)[0]
        rows = table.splitlines()[2:]
        self.assertTrue(6 <= len(rows) <= 8, rows)
        # Each row links a guide that a new repository keeps.
        for row in rows:
            with self.subTest(row=row):
                target = row.rsplit("](", 1)[1].split(")", 1)[0]
                self.assertTrue((ROOT / target).is_file(), target)

    def test_contact_links_point_at_the_new_repository(self) -> None:
        adapted = self.assertIdempotent(initialize.adapt_contact_links, CONFIG)
        self.assertIn('url: "https://github.com/owner/new-workbook/security/policy"', adapted)
        self.assertIn("https://github.com/owner/new-workbook/blob/HEAD/docs/LABELS.md", adapted)
        self.assertNotIn("template-owner", adapted)
        with self.assertRaisesRegex(InitializeError, "not a GitHub repository link"):
            initialize.adapt_contact_links('    url: "https://example.com/policy"\n', OPTIONS)

    def test_traffic_badges_are_repointed_or_removed(self) -> None:
        adapted = self.assertIdempotent(initialize.adapt_badges, README)
        self.assertEqual(adapted, README.replace("template-owner/template", REPOSITORY))
        removing = Options(REPOSITORY, "LCR", remove_traffic_badges=True)
        removed = self.assertIdempotent(initialize.adapt_badges, README, removing)
        self.assertEqual(removed, "# Template\n\nI use this template.\n")
        # Removing the collection removes its badges too.
        everything = Options(REPOSITORY, "LCR", remove_traffic=True)
        self.assertEqual(initialize.adapt_badges(README, everything)[0], removed)

    def test_dependabot_target_branch_is_set_added_or_removed(self) -> None:
        adapted = self.assertIdempotent(initialize.adapt_dependabot, DEPENDABOT)
        self.assertEqual(adapted.count("    target-branch: release/0.1.0\n"), 2)
        self.assertNotIn("release/1.0.0", adapted)
        default = Options(REPOSITORY, "LCR")
        removed = self.assertIdempotent(initialize.adapt_dependabot, DEPENDABOT, default)
        self.assertNotIn("target-branch", removed)
        added = self.assertIdempotent(initialize.adapt_dependabot, removed)
        self.assertEqual(added, adapted.replace('"release/0.1.0"', "release/0.1.0"))
        self.assertIn("    directory: /tools\n    target-branch: release/0.1.0\n", added)

    def test_dependabot_target_branch_is_added_after_a_list_of_directories(self) -> None:
        removed = initialize.adapt_dependabot(DEPENDABOT, Options(REPOSITORY, "LCR"))[0]
        multiple = removed.replace("    directory: /tools\n",
                                   "    directories:\n      - /.github\n      - /tools\n")
        self.assertNotEqual(multiple, removed)
        added = self.assertIdempotent(initialize.adapt_dependabot, multiple)
        self.assertIn("    directories:\n      - /.github\n      - /tools\n"
                      "    target-branch: release/0.1.0\n    schedule:", added)
        self.assertEqual(added.count("target-branch:"), 2)

    def test_house_style_takes_the_regime_in_its_table_and_the_request_above_it(self) -> None:
        adapted = self.assertIdempotent(initialize.adapt_house_style, HOUSE_STYLE)
        self.assertIn("| `LCR_` plus subject | `LCR_Functions` |", adapted)
        self.assertIn("`LCRRunStatus`", adapted)
        # The request to replace the placeholder becomes a statement of the prefix.
        self.assertIn("## Naming\n\nThe project's short name is `LCR`. The [static checks]"
                      "(#static-checks)\nread the facade prefix, `LCR_`, from this table; the "
                      "reserved prefixes\nare `CORE_`, `TEST_`, `EX_` and `M_`.\n\n| Element",
                      adapted)
        self.assertNotIn("Replace `<REGIME>`", adapted)
        self.assertNotIn("while it is", adapted)
        self.assertIn("A later paragraph that names `<REGIME>` keeps it.", adapted)
        with self.assertRaisesRegex(InitializeError, "already LCR_, not FRTB_"):
            initialize.adapt_house_style(adapted, Options(REPOSITORY, "FRTB"))
        # A table set by hand still has its request replaced, once.
        by_hand = HOUSE_STYLE.replace("| `<REGIME>_` plus subject | `<REGIME>_Functions` |",
                                      "| `LCR_` plus subject | `LCR_Functions` |") \
            .replace("| `<REGIME>` plus noun | `<REGIME>RunStatus`", "| `LCR` plus noun | "
                     "`LCRRunStatus`")
        self.assertEqual(initialize.adapt_house_style(by_hand, OPTIONS)[0], adapted)

    def test_the_real_house_style_names_the_prefix_and_no_placeholder(self) -> None:
        text = (ROOT / vba_lexer.HOUSE_STYLE).read_text(encoding="utf-8")
        if "<REGIME>" not in text:
            self.skipTest("the house style already names the project's prefix")
        adapted, notes = initialize.adapt_house_style(text, OPTIONS)
        self.assertEqual(len(notes), 2, notes)
        self.assertNotIn("<REGIME>", adapted)
        self.assertIn("The project's short name is `LCR`.", adapted)

    def test_tested_environments_lose_their_rows_but_keep_the_header(self) -> None:
        adapted = self.assertIdempotent(initialize.adapt_tested_environments, INSTALLATION)
        self.assertEqual(adapted, INSTALLATION.split("| Excel 16.0, 64-bit")[0]
                         + "\n## Next section\n")
        self.assertIn("| Tested environment | Commit | Evidence |\n| --- | --- | --- |\n", adapted)
        self.assertNotIn("#204", adapted)

    def test_host_behaviour_rows_cite_no_issue_and_other_tables_stay(self) -> None:
        adapted = self.assertIdempotent(initialize.adapt_host_behaviour, VALIDATION_PLAN)
        self.assertEqual(adapted, VALIDATION_PLAN.replace("| #245 |", "| \u2014 |")
                         .replace("| #204 |", "| \u2014 |"))
        self.assertIn("| Not host behaviour | #9 |", adapted)

    def test_the_template_evidence_goes_with_the_changelog_entries(self) -> None:
        with tempfile.TemporaryDirectory(prefix="initialize-evidence-") as directory:
            root = Path(directory)
            files = {"INSTALLATION.md": INSTALLATION,
                     "docs/methodology/VALIDATION_PLAN.md": VALIDATION_PLAN}
            # A repository whose changelog is its own keeps its evidence.
            write(root, files)
            self.assertEqual(initialize.plan(root, OPTIONS), [])
            write(root, {"CHANGELOG.md": CHANGELOG})
            changes = initialize.plan(root, OPTIONS)
            self.assertEqual([change.path for change in changes],
                             ["CHANGELOG.md", *files])
            initialize.apply(root, changes)
            for name in files:
                self.assertNotIn("#2", (root / name).read_text(encoding="utf-8"))
            self.assertEqual(initialize.plan(root, OPTIONS), [])

    def test_options_are_validated(self) -> None:
        for options in (Options("no-slash", "LCR"), Options(REPOSITORY, "CORE"),
                        Options(REPOSITORY, "lcr"), Options(REPOSITORY, "LCR", "bad branch")):
            with self.subTest(options=options), self.assertRaises(InitializeError):
                initialize.validate(options)
        initialize.validate(OPTIONS)

    def test_the_assignee_must_be_a_github_login(self) -> None:
        for login in ("", "-lead", "trail-", "two--hyphens", "under_score", "a space", "a" * 40,
                      "\u00e9t\u00e9"):
            with self.subTest(login=login), \
                    self.assertRaisesRegex(InitializeError, "--assignee must be a GitHub login"):
                initialize.validate(Options(REPOSITORY, "LCR", assignee=login))
        for login in ("a", "someone", "Some-One-2", "a" * 39):
            with self.subTest(login=login):
                initialize.validate(Options(REPOSITORY, "LCR", assignee=login))
        refused = io.StringIO()
        with contextlib.redirect_stderr(refused):
            code = initialize.main(["--repository", REPOSITORY, "--regime", "LCR",
                                    "--assignee", "bad--login", "--root", str(ROOT)])
        self.assertEqual(code, 2)
        self.assertIn("--assignee must be a GitHub login", refused.getvalue())

    def test_the_assignee_is_set_in_the_forms_dependabot_and_the_label_guide(self) -> None:
        with tempfile.TemporaryDirectory(prefix="initialize-assignee-") as directory:
            root = Path(directory)
            write(root, ASSIGNEE_TREE)
            # Without --assignee, nothing that names the assignee changes.
            self.assertEqual(initialize.plan(root, Options(REPOSITORY, "LCR")), [])
            options = Options(REPOSITORY, "LCR", assignee="someone")
            changes = initialize.plan(root, options)
            self.assertEqual([change.path for change in changes],
                             [*FORMS, initialize.DEPENDABOT, initialize.ISSUE_LABELS])
            initialize.apply(root, changes)
            self.assertEqual(check_issues.configured_assignee(root), "someone")
            for form in FORMS:
                text = (root / form).read_text(encoding="utf-8")
                self.assertEqual(check_issues.flow_list(text, "assignees"), ["someone"], form)
            self.assertEqual((root / initialize.DEPENDABOT).read_text(encoding="utf-8"),
                             "version: 2\nupdates:\n  - package-ecosystem: pip\n"
                             "    assignees:\n      - someone\n"
                             "    commit-message:\n      prefix: ci\n")
            self.assertEqual((root / initialize.ISSUE_LABELS).read_text(encoding="utf-8"),
                             "My account, `someone`, is the assignee.\n")
            self.assertEqual(initialize.plan(root, options), [])

    def test_remove_traffic_removes_the_collection_and_what_names_it(self) -> None:
        with tempfile.TemporaryDirectory(prefix="initialize-traffic-") as directory:
            root = Path(directory)
            write(root, TRAFFIC_TREE)
            # Without --remove-traffic, the collection and its ruleset stay.
            self.assertEqual(initialize.plan(root, Options(REPOSITORY, "LCR")), [])
            options = Options(REPOSITORY, "LCR", remove_traffic=True)
            changes = initialize.plan(root, options)
            notes = [f"{change.path}: {note}" for change in changes for note in change.notes]
            self.assertIn(f"{initialize.SETTINGS_CONFIG}: a ruleset \"Protect traffic history\" "
                          "already created on GitHub is left to you: tools/setup_repository.py "
                          "reports it as unmanaged and never deletes it; delete it in Settings > "
                          "Rules > Rulesets if it exists", notes)
            # A mention it cannot place is a warning, to edit by hand.
            self.assertEqual(initialize.traffic_mentions(root, changes), ["docs/GUIDE.md:1"])
            initialize.apply(root, changes)
            for path in initialize.TRAFFIC_FILES:
                self.assertFalse((root / path).exists(), path)
            self.assertEqual((root / "README.md").read_text(encoding="utf-8"),
                             "# Project\n\n| Area | What |\n| --- | --- |\n| Links | Checked |\n")
            config = json.loads((root / initialize.SETTINGS_CONFIG).read_text(encoding="utf-8"))
            self.assertEqual([ruleset["name"] for ruleset in config["rulesets"]], ["Protect main"])
            self.assertEqual((root / initialize.SETUP_GUIDE).read_text(encoding="utf-8"),
                             "| **Protection \u00b7 1 rulesets** | Default branch. |\n")
            depth = json.loads((root / initialize.VERIFICATION_DEPTH).read_text(encoding="utf-8"))
            self.assertEqual(depth, {"checks": [{"check": "changelog"}]})
            self.assertEqual((root / initialize.NOTICE).read_text(encoding="utf-8"),
                             "Files:\n\nSECURITY.md\n")
            register = json.loads((root / initialize.PROVENANCE).read_text(encoding="utf-8"))
            self.assertEqual(register["sources"][0]["files"], ["SECURITY.md"])
            self.assertEqual(initialize.plan(root, options), [])

    def test_the_initializer_removes_itself_last_and_only_with_apply(self) -> None:
        with tempfile.TemporaryDirectory(prefix="initialize-itself-") as directory:
            root = Path(directory)
            write(root, {"AGENTS.md": AGENTS, **dict.fromkeys(INITIALIZER, "# initializer\n")})
            init(root)
            git(root, "add", "-A")
            git(root, "commit", "-qm", "Created")
            changes = initialize.plan(root, OPTIONS)
            self.assertEqual([(change.path, change.last) for change in changes],
                             [("AGENTS.md", False), *((path, True) for path in INITIALIZER)])

            def run(write: bool) -> tuple[int | None, str, str]:
                printed, refused = io.StringIO(), io.StringIO()
                code = None
                with contextlib.redirect_stdout(printed), contextlib.redirect_stderr(refused):
                    code = initialize.main(["--repository", REPOSITORY, "--regime", "LCR",
                                            "--root", str(root), *(["--apply"] if write else [])])
                return code, printed.getvalue(), refused.getvalue()

            def kept() -> list[str]:
                return [path for path in INITIALIZER if (root / path).is_file()]

            # The dry run lists the deletion and deletes nothing.
            code, printed, _ = run(False)
            self.assertEqual(code, 0)
            self.assertIn("- tools/initialize_repository.py: delete it, last", printed)
            self.assertEqual(kept(), list(INITIALIZER))
            # A write that fails, or generated blocks that cannot be rendered, keep it.
            with mock.patch.object(Path, "write_bytes", side_effect=OSError("disk full")):
                code, _, refused = run(True)
            self.assertEqual((code, kept()), (2, list(INITIALIZER)))
            self.assertIn("AGENTS.md cannot be written: disk full", refused)
            code, _, refused = run(True)
            self.assertEqual((code, kept()), (2, list(INITIALIZER)))
            self.assertIn("the generated blocks cannot be rendered", refused)
            self.assertIn("CONTRIBUTING.md: missing", refused)
            git(root, "checkout", "-q", "--", ".")
            # Once every other change is written and the blocks rendered, it goes.
            with mock.patch.object(initialize, "render", return_value=[]) as rendered:
                code, printed, _ = run(True)
            self.assertEqual((code, kept()), (0, []))
            rendered.assert_called_once()
            self.assertEqual((root / "AGENTS.md").read_text(encoding="utf-8"),
                             "# Instructions\n\n- A rule.\n")

    def test_a_released_template_changelog_becomes_a_fresh_one(self) -> None:
        with tempfile.TemporaryDirectory(prefix="initialize-changelog-") as directory:
            root = Path(directory)
            write(root, {"CHANGELOG.md": CHANGELOG, "VERSION": "1.0.0\n",
                         "RELEASING.md": RELEASING})
            before, _ = check_changelog.check(root, REPOSITORY)
            self.assertTrue(any("does not point to" in finding for finding in before), before)
            changes = initialize.plan(root, OPTIONS)
            self.assertEqual([change.path for change in changes],
                             ["CHANGELOG.md", "VERSION", "RELEASING.md"])
            initialize.apply(root, changes)
            self.assertFalse((root / "VERSION").exists())
            changelog = (root / "CHANGELOG.md").read_text(encoding="utf-8")
            self.assertEqual(changelog, (
                "# Changelog\n\nNotable changes. It is neither a commit log nor release "
                "evidence. **No version has been released.**\n\n## Maintenance policy\n\n"
                "- Add a material change under `[Unreleased]`.\n\n## [Unreleased]\n\n"
                "[Unreleased]: https://github.com/owner/new-workbook/commits/main\n"))
            self.assertEqual(check_changelog.check(root, REPOSITORY), ([], 0))
            self.assertEqual(initialize.plan(root, OPTIONS), [])


def copy_tracked_tree(destination: Path) -> None:
    for name in gitio.tracked(ROOT):
        if (ROOT / name).is_file():
            (destination / name).parent.mkdir(parents=True, exist_ok=True)
            shutil.copy2(ROOT / name, destination / name)


def adapted_checks(context: Context) -> list[Check]:
    listed = [check for check in check_repository.checks(context) if check.name not in NOT_RUN]
    tests = [sys.executable, "-m", "unittest", *[f"tests.{name}" for name in ADAPTED_FILE_TESTS]]
    return [Check("tests-of-adapted-files", tests), *listed]


@unittest.skipIf(shutil.which("git") is None, "git is not installed")
class ReleaseBranchWarningTests(unittest.TestCase):
    """--release-branch is checked against origin as dependabot-target-branches checks it."""

    def setUp(self) -> None:
        directory = tempfile.TemporaryDirectory(prefix="initialize-warning-")
        self.addCleanup(directory.cleanup)
        self.base = Path(directory.name)
        self.origin = self.base / "origin.git"
        work = self.base / "work"
        work.mkdir()
        git(self.base, "init", "-q", "--bare", str(self.origin))
        init(work)
        git(work, "commit", "-q", "--allow-empty", "-m", "Created")
        git(work, "push", "-q", str(self.origin), "HEAD:refs/heads/main",
            "HEAD:refs/heads/release/0.1.0")

    def clone(self, *options: str) -> Path:
        target = self.base / "clone"
        git(self.base, "clone", "-q", *options, str(self.origin), str(target))
        return target

    def test_a_branch_only_on_origin_gives_no_warning(self) -> None:
        clone = self.clone("--single-branch", "--branch", "main")
        self.assertEqual(initialize.warnings(clone, OPTIONS), [])

    def test_a_branch_origin_lacks_is_a_warning(self) -> None:
        clone = self.clone()
        found = initialize.warnings(clone, Options(REPOSITORY, "LCR", "release/0.2.0"))
        self.assertEqual(len(found), 1)
        self.assertIn("origin has no branch release/0.2.0", found[0])
        self.assertIn("dependabot-target-branches check fails", found[0])

    def test_no_answer_from_origin_is_a_warning_that_says_so(self) -> None:
        clone = self.clone()
        git(clone, "remote", "remove", "origin")
        found = initialize.warnings(clone, OPTIONS)
        self.assertEqual(len(found), 1)
        self.assertIn("origin cannot be asked whether branch release/0.1.0 exists", found[0])

    def test_no_release_branch_needs_no_check(self) -> None:
        self.assertEqual(initialize.warnings(self.base, Options(REPOSITORY, "LCR")), [])


@unittest.skipIf(shutil.which("git") is None, "git is not installed")
@unittest.skipUnless(IS_TEMPLATE, "this repository is already initialized")
class GeneratedRepositoryTest(unittest.TestCase):
    def setUp(self) -> None:
        directory = tempfile.TemporaryDirectory(prefix="initialize-repository-")
        self.addCleanup(directory.cleanup)
        self.root = Path(directory.name)
        copy_tracked_tree(self.root)
        init(self.root)
        self.commit("Created from the template")
        # An origin with the release branch that Dependabot targets; the push
        # also makes it a fetched remote branch.
        origin = tempfile.TemporaryDirectory(prefix="initialize-repository-origin-")
        self.addCleanup(origin.cleanup)
        git(self.root, "init", "-q", "--bare", origin.name)
        git(self.root, "remote", "add", "origin", origin.name)
        git(self.root, "push", "-q", "origin", "HEAD:refs/heads/main",
            "HEAD:refs/heads/release/0.1.0")

    def commit(self, message: str) -> None:
        git(self.root, "add", "-A")
        git(self.root, "commit", "-qm", message)

    def initialize(self, *extra: str) -> subprocess.CompletedProcess[str]:
        return subprocess.run([sys.executable, "tools/initialize_repository.py", "--repository",
                               REPOSITORY, "--regime", "LCR", "--release-branch",
                               "release/0.1.0", *extra], cwd=self.root, capture_output=True,
                              text=True, check=False)

    def integrity(self, listed: list[str] | None = None) -> tuple[int, str, dict[str, str]]:
        """Run the checks with --template false; return the code, output and statuses."""
        def chosen(context: Context) -> list[Check]:
            return [check for check in adapted_checks(context)
                    if listed is None or check.name in listed]
        printed = io.StringIO()
        with contextlib.redirect_stdout(printed):
            code = check_repository.main(["--root", str(self.root), "--template", "false",
                                          "--repository", REPOSITORY, "--default-branch",
                                          "main"], chosen)
        report = json.loads((self.root / check_repository.REPORT_DIR / "report.json")
                            .read_text(encoding="utf-8"))
        return code, printed.getvalue(), {check["name"]: check["status"]
                                          for check in report["checks"]}

    def test_the_dry_run_writes_no_file_not_even_an_ignored_one(self) -> None:
        def every_path() -> list[str]:
            return sorted(path.relative_to(self.root).as_posix() for path in self.root.rglob("*")
                          if ".git" not in path.relative_to(self.root).parts)
        before = every_path()
        # Without this variable Python writes the bytecode of each module it imports.
        environment = {key: value for key, value in os.environ.items()
                       if key != "PYTHONDONTWRITEBYTECODE"}
        dry_run = subprocess.run([sys.executable, "tools/initialize_repository.py", "--repository",
                                  REPOSITORY, "--regime", "LCR"], cwd=self.root, env=environment,
                                 capture_output=True, text=True, check=False)
        self.assertEqual(dry_run.returncode, 0, dry_run.stderr)
        self.assertIn("dry run, nothing written", dry_run.stdout)
        self.assertEqual(every_path(), before)

    def test_an_initialized_copy_passes_repository_integrity(self) -> None:
        code, _, statuses = self.integrity(list(CLEANUP_CHECKS))
        self.assertEqual((code, statuses), (1, dict.fromkeys(CLEANUP_CHECKS, "fail")))
        options = ("--remove-traffic", "--assignee", "someone")
        traffic_ruleset = "Protect traffic history"

        # Negative controls: without the options, the collection and the assignee stay.
        plain = {change.path for change in initialize.plan(self.root, OPTIONS)}
        self.assertFalse(plain & {*initialize.TRAFFIC_FILES, initialize.SETTINGS_CONFIG,
                                  initialize.VERIFICATION_DEPTH, *FORMS}, plain)
        self.assertTrue(set(INITIALIZER) <= plain)

        dry_run = self.initialize(*options)
        self.assertEqual(dry_run.returncode, 0, dry_run.stderr)
        self.assertIn("- AGENTS.md: delete 1 template-only section(s)", dry_run.stdout)
        self.assertIn(f"- {check_repository.TRAFFIC_WORKFLOW}: delete it", dry_run.stdout)
        self.assertIn(f'a ruleset "{traffic_ruleset}" already created on GitHub is left to you',
                      dry_run.stdout)
        self.assertIn("- tools/initialize_repository.py: delete it, last", dry_run.stdout)
        self.assertIn("dry run, nothing written", dry_run.stdout)
        self.assertNotIn("WARNING", dry_run.stdout)
        self.assertEqual(git(self.root, "status", "--porcelain", "--untracked-files=all"), "")

        (self.root / "notes.txt").write_text("unrelated\n", encoding="utf-8")
        refused = self.initialize("--apply", *options)
        self.assertEqual(refused.returncode, 2)
        self.assertIn("the working tree differs from the commit", refused.stderr)
        self.assertEqual(git(self.root, "status", "--porcelain"), "?? notes.txt\n")
        (self.root / "notes.txt").unlink()

        applied = self.initialize("--apply", *options)
        self.assertEqual(applied.returncode, 0, applied.stderr)
        self.assertIn("Nothing on GitHub was changed", applied.stdout)
        self.assertNotIn("WARNING", applied.stdout)
        # Last, it removed itself and its test.
        for path in (*INITIALIZER, *initialize.TRAFFIC_FILES):
            self.assertFalse((self.root / path).exists(), path)
        # No evidence row cites a commit or an issue of the template.
        installation = (self.root / initialize.INSTALLATION).read_text(encoding="utf-8")
        rows = installation.split(initialize.TESTED_HEADER, 1)[1].split("\n")[2]
        self.assertFalse(rows.startswith("|"), rows)
        plan = (self.root / initialize.VALIDATION_PLAN).read_text(encoding="utf-8")
        host = [line for line in plan.split("\n") if line.startswith("| `HB-")]
        self.assertTrue(host)
        self.assertTrue(all(line.endswith(f"| {initialize.NO_EVIDENCE} |") for line in host), host)
        # The README's template-only sections gave way to the table of guides.
        readme = (self.root / "README.md").read_text(encoding="utf-8")
        self.assertNotIn("template-only", readme)
        self.assertIn("Find your way", readme)
        # Nothing but the setup guide's record of this initializer names the collection.
        self.assertEqual(initialize.traffic_mentions(self.root, []), [])
        self.assertNotIn(traffic_ruleset, (self.root / initialize.SETTINGS_CONFIG)
                         .read_text(encoding="utf-8"))
        self.assertEqual(check_issues.configured_assignee(self.root), "someone")
        for form in FORMS:
            text = (self.root / form).read_text(encoding="utf-8")
            self.assertEqual(check_issues.flow_list(text, "assignees"), ["someone"], form)
        self.assertNotIn(check_issues.configured_assignee(ROOT),
                         (self.root / initialize.DEPENDABOT).read_text(encoding="utf-8"))
        self.commit("Initialize the repository")
        # A second run, from the template, finds nothing left to change.
        again = subprocess.run([sys.executable, str(ROOT / INITIALIZER[0]), "--repository",
                                REPOSITORY, "--regime", "LCR", "--release-branch",
                                "release/0.1.0", "--root", str(self.root), "--apply", *options],
                               capture_output=True, text=True, check=False)
        self.assertEqual(again.returncode, 0, again.stderr)
        self.assertIn("nothing to change", again.stdout)
        self.assertEqual(git(self.root, "status", "--porcelain", "--untracked-files=all"), "")

        code, printed, statuses = self.integrity()
        self.assertEqual(code, 0, printed)
        self.assertNotIn("fail", statuses.values(), printed)
        self.assertNotIn("readme-traffic-badges", statuses)
        self.assertEqual(statuses["issue-configuration"], "pass", printed)

if __name__ == "__main__":
    unittest.main()
