#!/usr/bin/env python3
"""Render the documentation facts that the tools already know.

A generated block sits between two markers in a Markdown file:

    <!-- BEGIN generated: NAME -->
    ...
    <!-- END generated: NAME -->

Each NAME has a renderer here that computes the block from the repository:

- example-suite-results, in INSTALLATION.md: what the example suite's runs
  end with, read from the suite that .github/excel-evidence-policy.json names
  by the same code that checks Excel-run evidence (tools/check_excel_run.py);
- integrity-checks, in CONTRIBUTING.md: the checks that
  tools/check_repository.py runs, in order, with what each runs, as CI runs
  it for the repository OWNER/REPOSITORY, the probe it must fail first, if
  any, and whether it is advisory;
- data-contract-datasets and data-contract-fields, in
  docs/methodology/DATA_CONTRACT.md: the Datasets table and one field table
  per dataset, read from docs/methodology/data-contract.json by the code that
  validates input files against it (tools/check_data_contract.py);
- data-contract-codes, in the same guide: the finding codes of that validator;
- triage-rules, in docs/repo_governance/ISSUE_LABELS.md: the rules of the
  issue triage check, by the code each finding names, from the registry of
  tools/check_issues.py;
- workflow-triggers, in README.md: when each workflow of .github/workflows/
  runs, read from its on: block through tools/strict_yaml.py.

The workflow-triggers block needs PyYAML (.github/requirements.txt). Without
it, the block is reported as skipped and left as it is, as the tests that read
YAML are skipped; CI installs it, so a stale block fails there.

--check fails when a block differs from a fresh render, when its markers are
missing, unbalanced or repeated, or when a marker names no renderer. --write
replaces each block with its fresh render. Repository integrity runs --check.
"""

from __future__ import annotations

import argparse
import re
import sys
from collections.abc import Callable, Sequence
from dataclasses import dataclass
from pathlib import Path, PureWindowsPath

sys.path.insert(0, str(Path(__file__).resolve().parent))

import check_data_contract  # noqa: E402
import check_excel_run  # noqa: E402
import check_issues  # noqa: E402
import check_repository  # noqa: E402

ROOT = Path(__file__).resolve().parents[1]
MARKER = re.compile(r"^<!-- (BEGIN|END) generated: ([a-z0-9-]+) -->$", re.M)
# An entry point that runs the suite in one mode, with or without arguments.
ENTRY_POINT = re.compile(r'^Public Sub (\w+)\((?:(?!^End Sub).)*?^\s+RunSuite "(\w+)"',
                         re.M | re.S)
WINDOWS_SUFFIXES = frozenset({"exe", "cmd", "bat"})
COUNTS = {1: "One", 2: "Two", 3: "Three", 4: "Four"}
RESULT_ORDER = ("RESULT", "suite", "operation", "completeness", "cases", "assertions",
                "failures", "cleanup", "recovery")
WORKFLOWS = Path(".github/workflows")
# How each event reads with no filter, and with one; and what follows its filters.
EVENTS = {
    "push": ("every push", "pushes", ""),
    "pull_request": ("every pull request", "pull requests", ""),
    "pull_request_target": ("every pull request", "pull requests",
                            ", run from the base branch"),
    "workflow_dispatch": ("manual runs", "manual runs", ""),
    "issues": ("issue events", "issue events", ""),
    "milestone": ("milestone events", "milestone events", ""),
    "branch_protection_rule": ("branch protection rule changes",
                               "branch protection rule changes", ""),
}
# How each filter of an event reads after it.
FILTERS = {
    "branches": "{to}",
    "branches-ignore": "except {to}",
    "tags": "of the tags",
    "tags-ignore": "except the tags",
    "paths": "that change",
    "paths-ignore": "unless they change only",
}
# A manual run's inputs say what it does, not when it runs.
IGNORED_FILTERS = frozenset({"inputs"})
WEEKDAYS = ("Sunday", "Monday", "Tuesday", "Wednesday", "Thursday", "Friday", "Saturday",
            "Sunday")


class Unavailable(Exception):
    """A block that cannot be rendered here, for want of a tool CI installs."""


@dataclass(frozen=True)
class Block:
    path: str
    render: Callable[[Path], str]


def example_suite_results(root: Path) -> str:
    """The passing RESULT line and the expected end of each failure mode."""
    policy = check_excel_run.working_policy(root)
    path = root / policy.suite_source
    if not path.is_file():
        raise ValueError(f"{policy.suite_source}: missing")
    source = path.read_bytes()
    suite = check_excel_run.read_suite(source, policy.office)
    text = source.decode("ascii", errors="replace").replace("\r\n", "\n")
    entry_points = {mode: name for name, mode in ENTRY_POINT.findall(text)}
    plans = check_excel_run.plans(suite, policy.suite)
    passing = next(plan for plan in plans if plan.mode in policy.must_pass)
    module = Path(policy.suite_source).stem
    line = "; ".join(f"{key}={passing.result[key]}" for key in RESULT_ORDER)
    rows = []
    for plan in plans:
        if plan.mode in policy.must_pass:
            continue
        shown = ["RESULT", "operation", "completeness"]
        shown += [key for key in ("cases", "assertions")
                  if plan.result[key] != passing.result[key]]
        shown.append("failures")
        fields = ", ".join(f"`{key}={plan.result[key]}`" for key in shown)
        if plan.mode not in entry_points:
            raise ValueError(f"{policy.suite_source} has no entry point for {plan.mode}")
        rows.append(f"| `{module}.{entry_points[plan.mode]}` | {fields}, then *Suite "
                    f"{policy.suite} did not pass* |")
    return "\n".join([
        "```text", line, "```", "",
        f"Anything else, or the error *Suite {policy.suite} did not pass*, is a "
        "failure.",
        f"{COUNTS.get(len(rows), str(len(rows)))} more entry points show the failure paths "
        "and are expected to fail:", "",
        "| Entry point | Expected end |", "| --- | --- |", *rows])


def shown_command(command: Sequence[str]) -> str:
    """A check's command as a reader types it, whatever the local paths."""
    words = list(command)
    if words and words[0] == sys.executable:
        words[0] = "python"
    elif words:
        # PureWindowsPath splits on both separators; the suffix that Windows
        # adds to an executable is not part of the command a reader types.
        name = PureWindowsPath(words[0]).name
        stem, _, suffix = name.rpartition(".")
        words[0] = stem if stem and suffix.lower() in WINDOWS_SUFFIXES else name
    return " ".join(words)


def integrity_checks(root: Path) -> str:
    """The checks of tools/check_repository.py, in the order they run."""
    # With a repository, as CI runs them: changelog then shows --repository.
    context = check_repository.Context(root, repository="OWNER/REPOSITORY",
                                       default_branch=None, is_template=None)
    rows = []
    for check in check_repository.checks(context):
        if check.command is not None:
            runs = f"`{shown_command(check.command)}`"
        else:
            assert check.function is not None
            lines = (check.function.__doc__ or "").strip().splitlines()
            if not lines:
                raise ValueError(f"check {check.name}: {check.function.__name__} has no "
                                 "docstring to describe it")
            summary = lines[0]
            runs = f"Built in: {summary}"
        if check.probe is not None:
            runs += (f", once the probe `{shown_command(check.probe.command)}` has failed "
                     f"with `{check.probe.finding}`")
        if check.advisory:
            runs += " (advisory: findings warn, never fail the run)"
        rows.append(f"| `{check.name}` | {runs} |")
    return "\n".join(["| Check | Runs |", "| --- | --- |", *rows])


def contract_datasets(root: Path) -> list[check_data_contract.Dataset]:
    datasets, problems = check_data_contract.load(root)
    if problems:
        raise ValueError(f"{problems[0]}; python tools/check_data_contract.py lists every "
                         "problem")
    return datasets


def illustrative(dataset: check_data_contract.Dataset) -> str:
    return "*Illustrative:* " if dataset.illustrative else ""


def data_contract_datasets(root: Path) -> str:
    """The Datasets table of the data contract."""
    rows = [f"| {illustrative(d)}`{d.id}` | {d.description} | {d.provider} | {d.frequency} | "
            f"{d.granularity} | `{d.source}` | {f'`{d.encoding}`' if d.encoding else ''} |"
            for d in contract_datasets(root)]
    return "\n".join(["| ID | Description | Provider | Frequency | Granularity | File or sheet "
                      "| Encoding |", "| --- | --- | --- | --- | --- | --- | --- |", *rows])


def read_from(dataset: check_data_contract.Dataset) -> str:
    """Where a dataset is read from, and how, in one sentence."""
    if dataset.delimiter is not None:
        between = "a tab" if dataset.delimiter == "\t" else f"`{dataset.delimiter}`"
        encoding = dataset.encoding or check_data_contract.DEFAULT_ENCODING
        return (f"Read from `{dataset.source}`, delimited text in `{encoding}`, with {between} "
                f"between cells and `{dataset.decimal_separator}` as the decimal separator.")
    if dataset.is_file:
        return (f"Read from `{dataset.source}`, which is not delimited text: the reference "
                "validator does not read it.")
    return f"Read from the sheet `{dataset.source}`."


def data_contract_fields(root: Path) -> str:
    """One field table per dataset, under a heading naming the dataset."""
    parts = []
    for dataset in contract_datasets(root):
        rows = [f"| {illustrative(dataset)}`{field.name}` | "
                f"{check_data_contract.TYPE_NAMES[field.type]} | "
                f"{check_data_contract.unit_text(field)} | {'Yes' if field.required else 'No'} | "
                f"{check_data_contract.allowed_text(field)} | {field.notes} |"
                for field in dataset.fields]
        parts += [f"### {illustrative(dataset)}`{dataset.id}` {dataset.description}", "",
                  read_from(dataset), "",
                  "| Field | Type | Unit or currency | Required | Allowed values or range | Notes |",
                  "| --- | --- | --- | --- | --- | --- |", *rows, ""]
    return "\n".join(parts[:-1])


def data_contract_codes(root: Path) -> str:
    """The finding codes of the reference validator."""
    rows = [f"| `{code}` | {meaning} |" for code, meaning in check_data_contract.CODES.items()]
    return "\n".join(["| Code | Finding |", "| --- | --- |", *rows])


def triage_rules(root: Path) -> str:
    """The rules of the issue triage check, by the code each finding names."""
    rows = [f"| `{code}` | {rule} |" for code, rule in check_issues.RULES.items()]
    return "\n".join(["| Rule | What every open issue needs |", "| --- | --- |", *rows])


def code_list(values: object, what: str) -> str:
    if not isinstance(values, list) or not values:
        raise ValueError(f"{what} is not a list of values")
    return ", ".join(f"`{value}`" for value in values)


def schedule_text(entries: object, path: str) -> str:
    """Each cron schedule as a reader says it: daily, weekly, or the expression."""
    if not isinstance(entries, list) or not entries:
        raise ValueError(f"{path}: schedule is not a list of cron entries")
    parts = []
    for entry in entries:
        cron = entry.get("cron") if isinstance(entry, dict) else None
        if not isinstance(cron, str):
            raise ValueError(f"{path}: a schedule entry has no cron expression")
        fields = cron.split()
        if (len(fields) == 5 and fields[0].isdigit() and fields[1].isdigit()
                and fields[2:4] == ["*", "*"] and (fields[4] == "*" or fields[4] in "01234567")):
            time = f"{int(fields[1]):02d}:{int(fields[0]):02d} UTC"
            parts.append(f"daily at {time}" if fields[4] == "*"
                         else f"weekly on {WEEKDAYS[int(fields[4])]} at {time}")
        else:
            parts.append(f"on the schedule `{cron}`")
    return ", ".join(parts)


def event_text(event: str, config: object, path: str) -> str:
    """One event of a workflow's on: block, with its filters, in words."""
    if event == "schedule":
        return schedule_text(config, path)
    bare, filtered, suffix = EVENTS.get(event, (f"`{event}` events", f"`{event}` events", ""))
    if config is None:
        return bare + suffix
    if not isinstance(config, dict):
        raise ValueError(f"{path}: on.{event} is neither empty nor a mapping")
    to = "into" if event.startswith("pull_request") else "to"
    words = []
    for key, value in config.items():
        if key in IGNORED_FILTERS:
            continue
        if key == "types":
            if not isinstance(value, list) or not value:
                raise ValueError(f"{path}: on.{event}.types is not a list of values")
            words.append(f"({', '.join(str(item) for item in value)})")
            continue
        lead = FILTERS.get(key, f"with `{key}`").format(to=to)
        words.append(f"{lead} {code_list(value, f'{path}: on.{event}.{key}')}")
    return " ".join([filtered if words else bare, *words]) + suffix


def workflow_triggers(root: Path) -> str:
    """When each workflow runs, read from its on: block."""
    try:
        import strict_yaml
        import yaml
    except ImportError as error:  # PyYAML is missing
        raise Unavailable(f"PyYAML is not installed (.github/requirements.txt): {error}") \
            from error
    folder = root / WORKFLOWS
    paths = sorted([*folder.glob("*.yml"), *folder.glob("*.yaml")])
    if not paths:
        raise ValueError(f"{WORKFLOWS.as_posix()}: no workflow")
    rows = []
    for file in paths:
        path = file.relative_to(root).as_posix()
        try:
            document = strict_yaml.load(file.read_text(encoding="utf-8"))
        except yaml.YAMLError as error:
            raise ValueError(f"{path}: {error}") from error
        if not isinstance(document, dict):
            raise ValueError(f"{path}: not a mapping")
        # PyYAML reads the key on as the boolean true.
        triggers = document.get("on", document.get(True))
        if isinstance(triggers, str):
            triggers = {triggers: None}
        elif isinstance(triggers, list):
            triggers = dict.fromkeys(triggers)
        if not isinstance(triggers, dict) or not triggers:
            raise ValueError(f"{path}: no on: block")
        events = "; ".join(event_text(str(event), config, path)
                           for event, config in triggers.items())
        name = document.get("name") or file.name
        rows.append(f"| [**{name}**]({path}) | {events[0].upper()}{events[1:]} |")
    return "\n".join(["| Workflow | When it runs |", "| --- | --- |", *rows])


CONTRACT_GUIDE = "docs/methodology/DATA_CONTRACT.md"
BLOCKS: dict[str, Block] = {
    "example-suite-results": Block("INSTALLATION.md", example_suite_results),
    "integrity-checks": Block("CONTRIBUTING.md", integrity_checks),
    "data-contract-datasets": Block(CONTRACT_GUIDE, data_contract_datasets),
    "data-contract-fields": Block(CONTRACT_GUIDE, data_contract_fields),
    "data-contract-codes": Block(CONTRACT_GUIDE, data_contract_codes),
    "triage-rules": Block("docs/repo_governance/ISSUE_LABELS.md", triage_rules),
    "workflow-triggers": Block("README.md", workflow_triggers),
}


def blocks_in(text: str, path: str) -> tuple[dict[str, tuple[int, int]], list[str]]:
    """Each block's content span by name, and the problems with its markers."""
    spans: dict[str, tuple[int, int]] = {}
    problems: list[str] = []
    open_name: str | None = None
    start = 0
    for match in MARKER.finditer(text):
        kind, name = match.groups()
        line = text.count("\n", 0, match.start()) + 1
        if kind == "BEGIN":
            if open_name is not None:
                problems.append(f"{path}:{line}: BEGIN {name} inside block {open_name}")
            open_name, start = name, match.end()
        elif open_name != name:
            problems.append(f"{path}:{line}: END {name} without its BEGIN")
            open_name = None
        else:
            if name in spans:
                problems.append(f"{path}:{line}: block {name} appears twice")
            spans[name] = (start, match.start())
            open_name = None
    if open_name is not None:
        problems.append(f"{path}: block {open_name} is not closed")
    return spans, problems


def run(root: Path, write: bool) -> list[str]:
    """The problems of every block; write, if asked, the stale ones afresh."""
    problems: list[str] = []
    skipped = {".git", "test-results", "node_modules"}
    paths = sorted({block.path for block in BLOCKS.values()} | {
        path.relative_to(root).as_posix() for path in root.rglob("*.md")
        if not skipped & set(path.relative_to(root).parts)})
    for path in paths:
        file = root / path
        if not file.is_file():
            problems.append(f"{path}: missing")
            continue
        text = file.read_text(encoding="utf-8")
        spans, marker_problems = blocks_in(text, path)
        problems += marker_problems
        problems += [f"{path}: block {name} has no renderer" for name in spans
                     if name not in BLOCKS]
        expected = [name for name, block in BLOCKS.items() if block.path == path]
        problems += [f"{path}: block {name} is missing" for name in expected
                     if name not in spans]
        if marker_problems:
            continue
        updated = text
        for name in sorted((n for n in expected if n in spans), key=lambda n: -spans[n][0]):
            begin, end = spans[name]
            try:
                fresh = "\n" + BLOCKS[name].render(root) + "\n"
            except Unavailable as reason:
                print(f"skipped: {path}: block {name}: {reason}")
                continue
            except (OSError, ValueError, check_excel_run.PolicyError) as error:
                problems.append(f"{path}: block {name} cannot be rendered: {error}")
                continue
            if updated[begin:end] != fresh:
                if write:
                    updated = updated[:begin] + fresh + updated[end:]
                else:
                    problems.append(f"{path}: block {name} differs from a fresh render; run "
                                    "python tools/render_docs.py --write")
        if write and updated != text:
            file.write_text(updated, encoding="utf-8")
            print(f"rendered: {path}")
    return problems


def main(argv: Sequence[str] | None = None) -> int:
    parser = argparse.ArgumentParser(description=(__doc__ or "").splitlines()[0])
    mode = parser.add_mutually_exclusive_group(required=True)
    mode.add_argument("--check", action="store_true", help="fail on a stale block")
    mode.add_argument("--write", action="store_true", help="rewrite stale blocks")
    parser.add_argument("--root", type=Path, default=ROOT)
    args = parser.parse_args(argv)
    problems = run(args.root.resolve(), args.write)
    for problem in problems:
        print(f"FAIL: {problem}")
    status = "FAIL" if problems else "PASS"
    print(f"generated-docs: {status} ({len(BLOCKS)} block(s))")
    return 1 if problems else 0


if __name__ == "__main__":
    sys.exit(main())
