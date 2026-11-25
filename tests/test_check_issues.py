"""Tests for the issue triage rules (tools/check_issues.py)."""

from __future__ import annotations

import contextlib
import io
import json
import shutil
import sys
import tempfile
import unittest
import unittest.mock
from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(ROOT / "tools"))

import check_issues  # noqa: E402
from check_issues import Issue  # noqa: E402

REPOSITORY = "example-org/sample-project"
BLOCKED_LABEL = check_issues.BLOCKED
READY = "## Dependencies and blocking status\n\n**Ready to start**: none.\n"
PLANNING = "## Planning\n- Target end date: {} (Europe/Rome)\n\n"
START = "- Start date: {} (Europe/Rome)\n\n"
# Both planning lines, for a body built by hand.
PLANNED = PLANNING + START.format("2026-11-01")


def issue(number: int, *labels: str, state: str = "open", body: str = READY,
          assignees: tuple[str, ...] = ("triage-person",), milestone: str | None = "v0.1.0",
          target: str | None = "2026-11-04", milestone_due: str | None = "2026-11-30",
          start: str | None = "2026-11-01") -> Issue:
    """An issue whose body starts with planning lines for ``target`` and
    ``start``, each left out when it is None."""
    planned = ((PLANNING.format(target) if target is not None else "")
               + (START.format(start) if start is not None else "") + body)
    return Issue(number, state, labels or ("P3",), assignees, milestone, planned, milestone_due)


def blocked_by(*numbers: int, heading: str = "## Dependencies and blocking status") -> str:
    lines = "\n".join(f"- #{n} — prerequisite" for n in numbers)
    return f"## Objective\n\nx\n\n{heading}\n\n**blocked**:\n\n{lines}\n\n## Tracking\n\n#99\n"


def rules(issues: list[Issue], states: dict[int, str] | None = None,
          pulls: frozenset[int] = frozenset()) -> list[tuple[int, str]]:
    known = states if states is not None else {i.number: i.state for i in issues}
    return [(f.number, f.rule)
            for f in check_issues.audit(issues, known, REPOSITORY, pulls=pulls, assignee="triage-person")]


class MetadataTests(unittest.TestCase):
    def test_compliant_issue_has_no_findings(self) -> None:
        self.assertEqual(rules([issue(1)]), [])

    def test_assignee_priority_and_milestone_are_required(self) -> None:
        self.assertEqual(rules([issue(1, assignees=())]), [(1, "assignee")])
        self.assertEqual(rules([issue(1, "documentation")]), [(1, "priority")])
        self.assertEqual(rules([issue(1, "P2", "P3")]), [(1, "priority")])
        self.assertEqual(rules([issue(1, milestone=None)]), [(1, "milestone")])

    def test_p1_is_reserved_for_bugs(self) -> None:
        self.assertEqual(rules([issue(1, "P1", "documentation")]), [(1, "p1-reserved")])
        self.assertEqual(rules([issue(1, "P1", "bug")]), [])

    def test_closed_issues_are_not_audited(self) -> None:
        self.assertEqual(rules([issue(1, "documentation", state="closed", assignees=())]), [])

    def test_unknown_labels_are_reported_against_the_catalogue(self) -> None:
        findings = check_issues.audit([issue(1, "P3", "ready")], {1: "open"}, REPOSITORY,
                                      catalogue={"P3"}, assignee="triage-person")
        self.assertEqual([(f.number, f.rule) for f in findings], [(1, "catalogue")])
        # The catalogue is what the manifest resolves, not every label the file defines:
        # a label of an unselected overlay is in the file but outside the catalogue.
        self.assertEqual(findings[0].message,
                         "has label 'ready', which is not in the catalogue that "
                         ".github/labels.json resolves (core and selected overlays)")


class RuleRegistryTests(unittest.TestCase):
    """Each finding names a rule of check_issues.RULES, which the guide renders."""

    def test_a_finding_with_an_unregistered_code_is_refused(self) -> None:
        with self.assertRaisesRegex(ValueError, "'ready' is not registered in RULES"):
            check_issues.Finding(1, "ready", "has a status label")

    def test_every_finding_code_is_registered_and_every_rule_is_reported(self) -> None:
        blocked = (BLOCKED_LABEL, "P3")
        issues = [
            issue(1, assignees=()),
            issue(2, "P2", "P3"),
            issue(3, "P1"),
            issue(4, milestone=None),
            issue(5, target="soon"),
            issue(6, milestone_due=None),
            issue(7, start="soon"),
            issue(8, body="## Objective\n\nx\n"),
            issue(9, *blocked),
            issue(10, "P3", "unlisted"),
            issue(11, *blocked, body=blocked_by(12)),
            issue(12, *blocked, body=blocked_by(11)),
        ]
        findings = check_issues.audit(issues, {i.number: i.state for i in issues}, REPOSITORY,
                                      catalogue={"P1", "P2", "P3", "bug", BLOCKED_LABEL},
                                      assignee="triage-person")
        reported = {finding.rule for finding in findings}
        self.assertLessEqual(reported, set(check_issues.RULES))
        self.assertEqual(set(check_issues.RULES) - reported, set())

    def test_every_rule_is_described_on_one_line(self) -> None:
        for code, rule in check_issues.RULES.items():
            with self.subTest(code=code):
                self.assertRegex(code, r"^[a-z0-9]+(?:-[a-z0-9]+)*$")
                self.assertTrue(rule.endswith("."), rule)
                self.assertNotIn("\n", rule)
                self.assertNotIn("|", rule)


class TargetDateTests(unittest.TestCase):
    def test_target_end_date_must_be_a_real_iso_date(self) -> None:
        for date in ("2026-02-30", "04/11/2026", "2026-11-4", "soon"):
            with self.subTest(date=date):
                self.assertEqual(rules([issue(1, target=date)]), [(1, "target-date")])
        self.assertEqual(rules([issue(1, target=None)]), [(1, "target-date")])

    def test_crlf_bodies_are_read(self) -> None:
        planned = PLANNED.format("2026-11-04").replace("\n", "\r\n") + READY.replace("\n", "\r\n")
        self.assertEqual(check_issues.target_end_date(planned), "2026-11-04")
        form = "### 🗓️ Target end date (Europe/Rome)\r\n\r\n2026-11-04\r\n\r\n" + READY.replace("\n", "\r\n")
        self.assertEqual(check_issues.target_end_date(form), "2026-11-04")
        item = {"number": 1, "state": "open", "labels": [{"name": "P3"}],
                "assignees": [{"login": "triage-person"}], "body": planned,
                "milestone": {"title": "v0.1.0", "due_on": "2026-11-30T08:00:00Z", "state": "open"}}
        self.assertEqual(rules([check_issues.issue_from_api(item)]), [])

    def test_issue_form_field_is_read(self) -> None:
        body = "### 🗓️ Target end date (Europe/Rome)\n\n2026-11-04\n\n" + READY
        self.assertEqual(check_issues.target_end_date(body), "2026-11-04")
        self.assertEqual(rules([issue(1, body=body, target=None)]), [])
        self.assertEqual(rules([issue(1, body="### 🗓️ Target end date (Europe/Rome)\n\n_No response_\n\n"
                                       + READY, target=None)]), [(1, "target-date")])

    def test_planning_line_takes_precedence_even_when_invalid(self) -> None:
        form = "### 🗓️ Target end date (Europe/Rome)\n\n2026-11-04\n\n" + READY
        self.assertEqual(rules([issue(1, body=form, target="2026-11-10")]), [])
        self.assertEqual(check_issues.target_end_date(PLANNING.format("2026-11-10") + form),
                         "2026-11-10")
        for value in ("to be decided", "", "2026-11-10 or later"):
            with self.subTest(value=value):
                body = f"- Target end date: {value}\n\n" + form
                self.assertEqual(rules([issue(1, body=body, target=None)]), [(1, "target-date")])

    def test_start_date_must_be_a_real_iso_date(self) -> None:
        for date in ("2026-02-30", "16/11/2026", "soon", ""):
            with self.subTest(date=date):
                self.assertEqual(rules([issue(1, start=date)]), [(1, "start-date")])
        self.assertEqual(rules([issue(1, start=None)]), [(1, "start-date")])

    def test_start_date_must_not_be_after_the_target_end_date(self) -> None:
        self.assertEqual(rules([issue(1, start="2026-11-05", target="2026-11-04")]),
                         [(1, "start-date")])
        self.assertEqual(rules([issue(1, start="2026-11-04", target="2026-11-04")]), [])

    def test_start_date_form_field_is_read_and_the_line_takes_precedence(self) -> None:
        form = ("### 🗓️ Start date (Europe/Rome)\n\n2026-11-02\n\n"
                "### 🗓️ Target end date (Europe/Rome)\n\n2026-11-04\n\n" + READY)
        self.assertEqual(check_issues.start_date(form), "2026-11-02")
        self.assertEqual(rules([issue(1, body=form, target=None, start=None)]), [])
        self.assertEqual(check_issues.start_date(START.format("2026-11-03") + form), "2026-11-03")
        empty = form.replace("2026-11-02", "_No response_")
        self.assertEqual(rules([issue(1, body=empty, target=None, start=None)]),
                         [(1, "start-date")])

    def test_contradictory_planning_lines_are_reported(self) -> None:
        later_start = START.format("2026-11-02")
        later_target = "- Target end date: 2026-11-05 (Europe/Rome)\n\n"
        self.assertEqual(rules([issue(1, body=later_start + READY)]), [(1, "start-date")])
        self.assertEqual(rules([issue(1, body=later_target + READY)]), [(1, "target-date")])
        crlf = (later_start + later_target + READY).replace("\n", "\r\n")
        self.assertEqual(rules([issue(1, body=crlf)]), [(1, "start-date"), (1, "target-date")])
        finding = check_issues.planning_conflict_findings(issue(1, body=later_start + READY))
        self.assertIn("2026-11-01, 2026-11-02", finding[0].message)

    def test_repeated_planning_lines_and_a_different_form_field_are_not_contradictions(self) -> None:
        repeated = START.format("2026-11-01") + PLANNING.format("2026-11-04") + READY
        self.assertEqual(rules([issue(1, body=repeated)]), [])
        self.assertEqual(rules([issue(1, body="- Start date: 2026-11-01\n\n" + READY)]), [])
        form = ("### 🗓️ Start date (Europe/Rome)\n\n2026-11-03\n\n"
                "### 🗓️ Target end date (Europe/Rome)\n\n2026-11-20\n\n" + READY)
        self.assertEqual(rules([issue(1, body=form)]), [])

    def test_target_date_must_be_before_the_milestone_due_date(self) -> None:
        self.assertEqual(rules([issue(1, milestone_due="2026-11-05")]), [])
        self.assertEqual(rules([issue(1, milestone_due="2026-11-04")]), [(1, "milestone-due")])
        self.assertEqual(rules([issue(1, milestone_due=None)]), [(1, "milestone-due")])

    def test_invalid_target_date_does_not_hide_a_milestone_without_due_date(self) -> None:
        for target in (None, "soon"):
            with self.subTest(target=target):
                self.assertEqual(rules([issue(1, target=target, milestone_due=None)]),
                                 [(1, "target-date"), (1, "milestone-due")])
        # A milestone with a due date adds nothing while the target end date is invalid.
        self.assertEqual(rules([issue(1, target="soon")]), [(1, "target-date")])

    def test_missing_milestone_is_reported_once(self) -> None:
        self.assertEqual(rules([issue(1, milestone=None, milestone_due=None)]), [(1, "milestone")])

    def test_open_issue_in_a_closed_milestone_is_reported(self) -> None:
        closed = Issue(1, "open", ("P3",), ("triage-person",), "v0.1.0",
                       PLANNED.format("2026-11-04") + READY, "2026-11-30", "closed")
        self.assertEqual(rules([closed]), [(1, "milestone")])

    def test_due_date_comes_from_the_api_milestone(self) -> None:
        item = {"number": 1, "state": "open", "labels": [], "assignees": [], "body": "",
                "milestone": {"title": "v0.1.0", "due_on": "2026-11-30T08:00:00Z"}}
        self.assertEqual(check_issues.issue_from_api(item).milestone_due, "2026-11-30")
        self.assertEqual(check_issues.issue_from_api(item).milestone_state, "open")
        item["milestone"] = {"title": "v0.1.0", "due_on": None, "state": "closed"}
        self.assertEqual(check_issues.issue_from_api(item).milestone_state, "closed")
        item["milestone"] = {"title": "v0.1.0", "due_on": None}
        self.assertIsNone(check_issues.issue_from_api(item).milestone_due)


class DependencyTests(unittest.TestCase):
    def test_blocked_by_an_open_prerequisite(self) -> None:
        self.assertEqual(rules([issue(2, "P3", BLOCKED_LABEL, body=blocked_by(1)), issue(1)]), [])

    def test_blocked_with_only_closed_prerequisites_must_be_unblocked(self) -> None:
        closed = issue(1, state="closed")
        self.assertEqual(rules([issue(2, "P3", BLOCKED_LABEL, body=blocked_by(1)), closed]),
                         [(2, "blocked")])

    def test_open_prerequisite_requires_blocked(self) -> None:
        self.assertEqual(rules([issue(2, "P3", body=blocked_by(1)), issue(1)]), [(2, "blocked")])

    def test_external_condition_counts_as_a_blocker(self) -> None:
        body = "### 🔗 Dependencies and blocking status\n\nExternal condition: Excel host access\n"
        self.assertEqual(rules([issue(2, "P3", BLOCKED_LABEL, body=body)]), [])
        self.assertEqual(rules([issue(2, "P3", body=body)]), [(2, "blocked")])

    def test_references_outside_the_section_are_ignored(self) -> None:
        # blocked_by() also mentions #99 under Tracking; it must not count.
        self.assertEqual(rules([issue(2, "P3", BLOCKED_LABEL, body=blocked_by(1)), issue(1)],
                               {1: "open", 2: "open"}), [])

    def test_issue_urls_and_form_headings_are_recognised(self) -> None:
        body = ("### 🔗 Dependencies and blocking status\n\n"
                f"- https://github.com/{REPOSITORY}/issues/1\n")
        self.assertEqual(rules([issue(2, "P3", BLOCKED_LABEL, body=body), issue(1)]), [])

    def test_a_dependency_cycle_is_reported_on_each_member(self) -> None:
        cycle = [issue(1, "P3", BLOCKED_LABEL, body=blocked_by(2)),
                 issue(2, "P3", BLOCKED_LABEL, body=blocked_by(3)),
                 issue(3, "P3", BLOCKED_LABEL, body=blocked_by(1))]
        self.assertEqual(rules(cycle), [(1, "dependency-cycle"), (2, "dependency-cycle"),
                                        (3, "dependency-cycle")])
        findings = check_issues.audit(cycle, {1: "open", 2: "open", 3: "open"}, REPOSITORY,
                                      assignee="triage-person")
        self.assertIn("is in a dependency cycle: #1, #2, #3", findings[0].message)
        pair = [issue(4, "P3", BLOCKED_LABEL, body=blocked_by(5)),
                issue(5, "P3", BLOCKED_LABEL, body=blocked_by(4))]
        self.assertEqual(rules(pair), [(4, "dependency-cycle"), (5, "dependency-cycle")])

    def test_chains_and_issues_pointing_into_a_cycle_are_not_cycles(self) -> None:
        chain = [issue(1, "P3", BLOCKED_LABEL, body=blocked_by(2)),
                 issue(2, "P3", BLOCKED_LABEL, body=blocked_by(3)), issue(3)]
        self.assertEqual(rules(chain), [])
        into = [issue(1, "P3", BLOCKED_LABEL, body=blocked_by(2)),
                issue(2, "P3", BLOCKED_LABEL, body=blocked_by(3)),
                issue(3, "P3", BLOCKED_LABEL, body=blocked_by(2)),
                issue(6, "P3", BLOCKED_LABEL, body=blocked_by(7)),
                issue(7, "P3", BLOCKED_LABEL, body=blocked_by(6))]
        self.assertEqual(rules(into), [(2, "dependency-cycle"), (3, "dependency-cycle"),
                                       (6, "dependency-cycle"), (7, "dependency-cycle")])

    def test_closed_issues_pull_requests_and_text_outside_the_section_break_a_cycle(self) -> None:
        through_closed = [issue(1, body=blocked_by(2)),
                          issue(2, state="closed", body=blocked_by(1))]
        self.assertEqual(rules(through_closed), [])
        through_pull = [issue(1, "P3", BLOCKED_LABEL, body=blocked_by(2)),
                        issue(2, "P3", BLOCKED_LABEL, body=blocked_by(9))]
        self.assertEqual(rules(through_pull, {1: "open", 2: "open"}, frozenset({9})),
                         [(2, "blocked")])
        # blocked_by() names #99 under Tracking, outside the section.
        outside = [issue(1, "P3", BLOCKED_LABEL, body=blocked_by(99)), issue(99)]
        self.assertEqual(rules(outside), [])
        # A self-reference is not a cycle.
        self.assertEqual(rules([issue(1, body=blocked_by(1))]), [])

    def test_a_long_chain_does_not_exceed_the_recursion_limit(self) -> None:
        count = 2500
        chain = [issue(n, "P3", BLOCKED_LABEL, body=blocked_by(n + 1)) for n in range(1, count)]
        chain.append(issue(count, "P3", BLOCKED_LABEL, body=blocked_by(1)))
        found = rules(chain)
        self.assertEqual(len(found), count)
        self.assertEqual(found[0], (1, "dependency-cycle"))

    def test_missing_section_and_unknown_reference(self) -> None:
        self.assertEqual(rules([issue(1, body="## Objective\n\nx\n")]), [(1, "dependencies")])
        self.assertEqual(rules([issue(2, "P3", BLOCKED_LABEL, body=blocked_by(7))]),
                         [(2, "dependencies"), (2, "blocked")])

    def test_empty_dependency_section_is_reported(self) -> None:
        for body in ("## Dependencies and blocking status\n",
                     "## Dependencies and blocking status\n\n  \n\n## Tracking\n\n#99\n",
                     "### 🔗 Dependencies and blocking status\n\n_No response_\n"):
            with self.subTest(body=body):
                findings = check_issues.audit([issue(1, body=body)], {1: "open"}, REPOSITORY,
                                              assignee="triage-person")
                self.assertEqual([(f.number, f.rule) for f in findings], [(1, "dependencies")])
                self.assertEqual(findings[0].message,
                                 "has an empty dependency section: write None, link each "
                                 "prerequisite or add an External condition: line")
        for none in ("None", "None."):
            with self.subTest(none=none):
                body = f"## Dependencies and blocking status\n\n{none}\n"
                self.assertEqual(rules([issue(1, body=body)]), [])

    def test_self_reference_is_ignored(self) -> None:
        self.assertEqual(rules([issue(2, "P3", body=blocked_by(2))]), [])

    def test_named_pull_request_is_a_reference_not_a_prerequisite(self) -> None:
        body = "## Dependencies and blocking status\n\nNone. #1 was delivered in #22.\n"
        closed = issue(1, state="closed")
        self.assertEqual(rules([issue(2, "P3", body=body), closed], pulls=frozenset({22})), [])
        # A pull request never justifies blocked, and unknown numbers are still reported.
        self.assertEqual(rules([issue(2, "P3", BLOCKED_LABEL, body=blocked_by(22, 7))],
                               pulls=frozenset({22})),
                         [(2, "dependencies"), (2, "blocked")])


class GuideExampleTests(unittest.TestCase):
    """The dependency-section example of the label guide reads as the guide says."""

    def example(self) -> str:
        guide = (ROOT / "docs/repo_governance/ISSUE_LABELS.md").read_text(encoding="utf-8")
        section = guide.split("### Dependency section", 1)[1]
        return section.split("```markdown\n", 1)[1].split("```", 1)[0]

    def findings(self, body: str, *labels: str) -> list[str]:
        # #12 is an open issue and #34 a pull request, as the example says.
        found = rules([issue(1, "P3", *labels, body=body)], {12: "open"}, pulls=frozenset({34}))
        return [rule for _, rule in found if rule in {"blocked", "dependencies"}]

    def test_the_example_is_a_correctly_blocked_issue(self) -> None:
        self.assertEqual(self.findings(self.example(), BLOCKED_LABEL), [])

    def test_each_blocker_of_the_example_requires_blocked(self) -> None:
        self.assertEqual(self.findings(self.example()), ["blocked", "blocked"])

    def test_the_pull_request_alone_does_not_block(self) -> None:
        only_pull = "\n".join(line for line in self.example().splitlines()
                              if "#12" not in line and "External condition" not in line)
        self.assertIn("#34", only_pull)
        self.assertEqual(self.findings(only_pull, BLOCKED_LABEL), ["blocked"])


class ConfigTests(unittest.TestCase):
    def test_repository_configuration_passes(self) -> None:
        self.assertEqual(check_issues.config_findings(ROOT), [])

    def test_forms_must_ask_for_the_start_date(self) -> None:
        with tempfile.TemporaryDirectory() as directory:
            root = Path(directory)
            shutil.copytree(ROOT / ".github", root / ".github")
            form = root / ".github/ISSUE_TEMPLATE/bug.yml"
            form.write_text(form.read_text(encoding="utf-8").replace(
                "id: start_date", "id: begin_date"), encoding="utf-8")
            self.assertEqual(check_issues.config_findings(root),
                             ["bug.yml has no start date field"])

    def test_forms_must_ask_for_the_target_end_date(self) -> None:
        with tempfile.TemporaryDirectory() as directory:
            root = Path(directory)
            shutil.copytree(ROOT / ".github", root / ".github")
            form = root / ".github/ISSUE_TEMPLATE/enhancement.yml"
            form.write_text(form.read_text(encoding="utf-8").replace(
                "id: target_end_date", "id: planned_date"), encoding="utf-8")
            self.assertEqual(check_issues.config_findings(root),
                             ["enhancement.yml has no target end date field"])

    @staticmethod
    def catalogue_with_overlay(root: Path, selected: list[str]) -> set[str]:
        """The resolved catalogue of a copy of the manifest with one overlay, ``regime``."""
        shutil.copytree(ROOT / ".github/scripts", root / ".github/scripts")
        manifest = json.loads((ROOT / ".github/labels.json").read_text(encoding="utf-8"))
        manifest["overlays"] = {"regime": [
            {"name": "zz-regime", "color": "ABCDEF", "description": "overlay label"}]}
        manifest["selected_overlays"] = selected
        (root / ".github/labels.json").write_text(json.dumps(manifest), encoding="utf-8")
        return check_issues.catalogue_labels(root)

    def test_catalogue_is_the_core_plus_the_selected_overlays(self) -> None:
        core = {label["name"] for label in
                json.loads((ROOT / ".github/labels.json").read_text(encoding="utf-8"))["core"]}
        with tempfile.TemporaryDirectory() as directory:
            self.assertEqual(self.catalogue_with_overlay(Path(directory), ["regime"]),
                             core | {"zz-regime"})
        with tempfile.TemporaryDirectory() as directory:
            self.assertEqual(self.catalogue_with_overlay(Path(directory), []), core)

    def test_an_unknown_or_duplicated_overlay_selection_is_rejected(self) -> None:
        for selected, message in ((["missing"], 'selects unknown overlay "missing"'),
                                  (["regime", "regime"], 'selects overlay "regime" again')):
            with self.subTest(selected=selected), tempfile.TemporaryDirectory() as directory:
                with self.assertRaisesRegex(ValueError, message):
                    self.catalogue_with_overlay(Path(directory), selected)

    def test_the_previous_manifest_schema_is_rejected(self) -> None:
        with tempfile.TemporaryDirectory() as directory:
            root = Path(directory)
            shutil.copytree(ROOT / ".github/scripts", root / ".github/scripts")
            manifest = json.loads((ROOT / ".github/labels.json").read_text(encoding="utf-8"))
            previous = {"schema_version": 2, "prune": manifest["prune"], "labels": manifest["core"]}
            (root / ".github/labels.json").write_text(json.dumps(previous), encoding="utf-8")
            with self.assertRaisesRegex(ValueError, "schema_version 2 with a labels list is no "
                                                    "longer supported"):
                check_issues.catalogue_labels(root)

    def test_assignee_comes_from_configuration_not_repository_identity(self) -> None:
        with tempfile.TemporaryDirectory() as directory:
            root = Path(directory)
            forms = root / ".github/ISSUE_TEMPLATE"
            forms.mkdir(parents=True)
            (forms / "bug.yml").write_text('assignees: [triage-person]\n')
            self.assertEqual(check_issues.configured_assignee(root), "triage-person")
            (forms / "bug.yml").write_text('assignees: []\n')
            with self.assertRaises(ValueError):
                check_issues.configured_assignee(root)

    def test_flow_list(self) -> None:
        text = 'name: x\nlabels: ["bug", "P2"]\nassignees: [triage-person]\n'
        self.assertEqual(check_issues.flow_list(text, "labels"), ["bug", "P2"])
        self.assertEqual(check_issues.flow_list(text, "assignees"), ["triage-person"])
        self.assertEqual(check_issues.flow_list(text, "title"), [])


class MainTests(unittest.TestCase):
    def test_input_mode_reports_and_exits_nonzero_on_findings(self) -> None:
        items = [
            {"number": 1, "state": "open", "labels": [{"name": "P3"}],
             "assignees": [{"login": "triage-person"}], "milestone": None,
             "body": PLANNED.format("2026-11-04") + READY + "Delivered in #2, a pull request.\n"},
            {"number": 2, "state": "open", "labels": [], "assignees": [],
             "milestone": None, "body": READY, "pull_request": {}},
        ]
        with tempfile.TemporaryDirectory() as directory:
            source = Path(directory) / "issues.json"
            source.write_text(json.dumps(items), encoding="utf-8")
            output = Path(directory) / "report.json"
            with contextlib.redirect_stdout(io.StringIO()) as stdout, \
                    unittest.mock.patch.object(check_issues, "configured_assignee",
                                               return_value="triage-person"):
                code = check_issues.main(["--root", str(ROOT), "--input", str(source),
                                          "--output", str(output), "--repo-name", REPOSITORY])
            self.assertEqual(code, 1)
            report = json.loads(output.read_text(encoding="utf-8"))
            self.assertEqual(report["status"], "fail")
            self.assertEqual([f["rule"] for f in report["findings"]], ["milestone"])
            self.assertIn("open issues audited: 1", stdout.getvalue())

    def test_input_mode_reports_a_dependency_cycle(self) -> None:
        def item(number: int, prerequisite: int) -> dict[str, object]:
            return {"number": number, "state": "open",
                    "labels": [{"name": "P3"}, {"name": BLOCKED_LABEL}],
                    "assignees": [{"login": "triage-person"}],
                    "milestone": {"title": "v0.1.0", "state": "open",
                                  "due_on": "2026-11-30T00:00:00Z"},
                    "body": PLANNED.format("2026-11-04") + blocked_by(prerequisite)}

        with tempfile.TemporaryDirectory() as directory:
            source = Path(directory) / "issues.json"
            source.write_text(json.dumps([item(1, 2), item(2, 1)]), encoding="utf-8")
            output = Path(directory) / "report.json"
            with contextlib.redirect_stdout(io.StringIO()), \
                    unittest.mock.patch.object(check_issues, "configured_assignee",
                                               return_value="triage-person"):
                code = check_issues.main(["--root", str(ROOT), "--input", str(source),
                                          "--output", str(output), "--repo-name", REPOSITORY])
            self.assertEqual(code, 1)
            report = json.loads(output.read_text(encoding="utf-8"))
            self.assertEqual([(f["number"], f["rule"]) for f in report["findings"]],
                             [(1, "dependency-cycle"), (2, "dependency-cycle")])

    def test_live_mode_reads_every_page_and_refuses_partial_lists(self) -> None:
        def item(number: int) -> dict[str, object]:
            return {"number": number, "state": "open", "labels": [], "assignees": [],
                    "milestone": None, "body": ""}

        pages = {1: [item(n) for n in range(1, 101)], 2: [item(101)]}
        with unittest.mock.patch.object(
                check_issues, "fetch",
                side_effect=lambda url, token: pages[int(url.rsplit("page=", 1)[1])]) as fetch:
            issues, states, pulls = check_issues.load_live("o/r", "token")
        self.assertEqual(len(issues), 101)
        self.assertEqual(fetch.call_count, 2)
        with unittest.mock.patch.object(check_issues, "fetch",
                                        return_value=[item(1)] * 100) as fetch, \
                self.assertRaisesRegex(RuntimeError, "10000 or more issues"):
            check_issues.load_live("o/r", "token")
        self.assertEqual(fetch.call_count, 100)

    def test_offline_mode_requires_explicit_repository(self) -> None:
        with contextlib.redirect_stderr(io.StringIO()), self.assertRaises(SystemExit) as error:
            check_issues.parse_args(["--input", "issues.json"])
        self.assertEqual(error.exception.code, 2)

    def test_live_mode_requires_a_token(self) -> None:
        with contextlib.redirect_stderr(io.StringIO()), \
                unittest.mock.patch.dict("os.environ", {"GITHUB_TOKEN": ""}):
            self.assertEqual(check_issues.main(["--repository", REPOSITORY]), 2)


if __name__ == "__main__":
    unittest.main()
