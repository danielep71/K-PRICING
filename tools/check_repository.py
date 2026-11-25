#!/usr/bin/env python3
"""Run every Repository integrity check, the same way locally and in CI.

Each check is a command or a small function. The run writes, under
test-results/integrity/:

- one <check>.log per check, with its command and output, and those of its
  probe if it has one;
- report.json: the commit, whether the tree is dirty, the pull request head,
  the Python version and platform, the SHA-256 of every tracked file, and each
  check's command, status and exit code;
- summary.md, which --summary also appends to a file such as
  $GITHUB_STEP_SUMMARY.

A check with a probe first runs its tool on an input the tool must reject,
and fails without running its command unless the tool exits non-zero and
reports the expected finding: a tool that accepts everything cannot pass.

A check whose tool or GitHub context is missing is skipped and reported as
such, as is a check that cannot reach the origin remote; --strict, which CI
passes, makes it fail instead. An advisory check reports its findings as a
warning and does not fail the run. --ci, which CI also passes, adds a last
check that fails when the working tree differs from the commit, so that the
result describes the commit and nothing a step changed. The run covers the
repository only: it compiles no VBA, runs no Excel and certifies no figure.

The repository is --repository, as CI passes it, or else the GitHub repository
of the origin remote; without either, the changelog check runs without
--repository, and the checks that need the repository are skipped.

The command checks run in parallel, up to four at a time; the checks written
as functions run one at a time, and that last check runs alone, once every
other check has finished. The console, the logs and the report list the
checks in the same order whichever finishes first.
"""

from __future__ import annotations

import argparse
import concurrent.futures
import datetime
import fnmatch
import hashlib
import importlib.util
import json
import os
import platform
import re
import shutil
import subprocess
import sys
from collections.abc import Callable, Iterator, Sequence
from dataclasses import dataclass, field
from pathlib import Path
from typing import Any

import gitio
from gitio import GitError

ROOT = Path(__file__).resolve().parents[1]
REPORT_DIR = Path("test-results/integrity")
SCOPE = "Repository checks only: no VBA compilation, no Excel run, no numerical certification"
TIMEOUT_SECONDS = 600
# The most command checks that run at once, and never more than the CPUs.
MAX_WORKERS = 4
# How long to wait for the origin remote to list a branch.
REMOTE_TIMEOUT_SECONDS = 60
# A workflow that actionlint must reject; it lives outside .github/workflows/,
# so GitHub never runs it and the normal lint never reads it.
ACTIONLINT_PROBE = ".github/actionlint/invalid-workflow.yml"
ACTIONLINT_PROBE_FINDING = 'unexpected key "bogus-key"'
REQUIRED_DOCUMENTS = ("README.md", "LICENSE", "docs/repo_governance/SETTINGS.md")
BADGE = re.compile(r"raw\.githubusercontent\.com/([^/\s]+/[^/\s]+)/traffic-history/")
# The optional traffic collection, whose badges readme-traffic-badges checks.
TRAFFIC_WORKFLOW = ".github/workflows/traffic-history.yml"
# The files whose template-only sections only the template keeps.
TEMPLATE_ONLY_FILES = ("AGENTS.md", "README.md")
CONTACT_URL = re.compile(r'^\s*url:\s*"?([^"\s]+)"?\s*$', re.M)
TARGET_BRANCH = re.compile(r'^\s*target-branch:\s*"?([^"\s]+)"?\s*$', re.M)
GITHUB_REMOTE = re.compile(r"github\.com[:/]([\w.-]+/[\w.-]+?)(?:\.git)?/?$")
# File names that mark a secret or a private key, matched case-insensitively
# against the last path component. A reviewed list: .gitignore ignores the same
# names, but git add -f bypasses it, and secret scanning matches token
# patterns, not file types.
SECRET_FILES = (".env", ".env.*", "*.pem", "*.key", "*.p12", "*.pfx", "*.pvk", "*.jks",
                "*.keystore", "id_rsa*", "id_ed25519*", "*.kdbx")
# A sanitized example and public SSH keys stay trackable.
SECRET_FILE_EXCEPTIONS = (".env.example", "id_rsa*.pub", "id_ed25519*.pub")

Finding = tuple[str, str]


class MissingContext(Exception):
    """Raised by a check function that cannot run here; the message says why.

    The run reports it as it reports a Check's missing reason: skipped, or
    failed under --strict.
    """


@dataclass
class Context:
    """What the checks may need to know about the repository and the run."""

    root: Path
    repository: str | None = None
    default_branch: str | None = None
    is_template: bool | None = None
    # Whether to fail a working tree that differs from the commit (--ci).
    ci: bool = False


@dataclass
class Probe:
    """A command that must fail and report a finding, proving the check can fail."""

    command: list[str]
    # Text the probe's output must contain.
    finding: str


@dataclass
class Check:
    """A command to run, or a function returning findings as (file, message)."""

    name: str
    command: list[str] | None = None
    function: Callable[[Context], list[Finding]] | None = None
    # Why the check cannot run here, if it cannot.
    missing: str | None = None
    # Findings are reported as a warning and do not fail the run.
    advisory: bool = False
    # Run first; the check fails without running its command unless the probe fails.
    probe: Probe | None = None
    # Start once every check listed before it has finished, and finish before any
    # check listed after it starts.
    alone: bool = False


@dataclass
class Result:
    name: str
    status: str
    command: list[str] | None
    exit_code: int | None
    detail: str = ""
    log: str = field(default="", repr=False)


# ---------------------------------------------------------------------------
# Checks that read files or Git, without a separate tool.

def template_only(context: Context) -> list[Finding]:
    """Only the template keeps the template-only sections of AGENTS.md and README.md."""
    if context.is_template is not False:
        return []
    findings = []
    for name in TEMPLATE_ONLY_FILES:
        path = context.root / name
        if path.is_file() and "template-only:start" in path.read_text(encoding="utf-8"):
            findings.append((name, f"Replace or delete the template-only sections of {name}, "
                                   "including their markers: only the template keeps them."))
    return findings


def contact_links(context: Context) -> list[Finding]:
    """The issue forms' contact links point to this repository."""
    path = ".github/ISSUE_TEMPLATE/config.yml"
    if not (context.root / path).is_file():
        return []
    expected = f"https://github.com/{str(context.repository).lower()}/"
    return [(path, f"Contact link {url} does not point to {context.repository}; point it "
                   "there.")
            for url in CONTACT_URL.findall((context.root / path).read_text(encoding="utf-8"))
            if not url.lower().startswith(expected)]


def has_traffic(root: Path) -> bool:
    """Whether the repository keeps the traffic collection or a badge that reads one."""
    readme = root / "README.md"
    return (root / TRAFFIC_WORKFLOW).is_file() or \
        (readme.is_file() and BADGE.search(readme.read_text(encoding="utf-8")) is not None)


def traffic_badges(context: Context) -> list[Finding]:
    """The README's traffic badges read this repository's traffic."""
    path = context.root / "README.md"
    if not path.is_file():
        return []
    read = sorted(set(BADGE.findall(path.read_text(encoding="utf-8"))))
    return [("README.md", f"A traffic badge reads {name}; point it at {context.repository} or "
                          "remove it.")
            for name in read if name.lower() != str(context.repository).lower()]


def secret_pattern(path: str) -> str | None:
    """The SECRET_FILES pattern a tracked path matches, unless an exception covers it."""
    name = path.rsplit("/", 1)[-1].lower()
    if any(fnmatch.fnmatchcase(name, pattern) for pattern in SECRET_FILE_EXCEPTIONS):
        return None
    return next((pattern for pattern in SECRET_FILES if fnmatch.fnmatchcase(name, pattern)), None)


def secret_files(context: Context) -> list[Finding]:
    """No tracked file is named as a secret or private key file."""
    findings = []
    for path in gitio.tracked(context.root):
        pattern = secret_pattern(path)
        if pattern:
            findings.append((path, f"is tracked and named as a secret or key file ({pattern}); "
                                   "remove it from the repository, rotate the secret it holds, "
                                   "and follow Handling a finding in "
                                   "docs/repo_governance/DISTRIBUTION.md."))
    return findings


def dependabot_branches(context: Context) -> list[Finding]:
    """Every Dependabot target-branch exists."""
    path = ".github/dependabot.yml"
    if not (context.root / path).is_file():
        return []
    findings = []
    for branch in sorted(set(TARGET_BRANCH.findall((context.root / path).read_text("utf-8")))):
        if not origin_has_branch(context.root, branch):
            findings.append((path, f"target-branch {branch} does not exist; set it to the active "
                                   "release branch or remove it."))
    return findings


def origin_has_branch(root: Path, branch: str) -> bool:
    """Whether origin has the branch, by its remote-tracking branch or else by asking it.

    A single-branch or shallow clone lacks the remote-tracking branch of a
    branch that exists. Raises MissingContext when origin cannot answer: no
    such remote, no network or no access.
    """
    try:
        gitio.git(root, "rev-parse", "--verify", "--quiet", f"refs/remotes/origin/{branch}")
        return True
    except GitError:
        pass
    ref = f"refs/heads/{branch}"
    try:
        # --exit-code exits with 2 when origin answers that it has no such branch.
        asked = subprocess.run(["git", "ls-remote", "--exit-code", "--heads", "origin", ref],
                               cwd=root, stdin=subprocess.DEVNULL, capture_output=True,
                               text=True, timeout=REMOTE_TIMEOUT_SECONDS, check=False,
                               env={**os.environ, "GIT_TERMINAL_PROMPT": "0"})
    except subprocess.TimeoutExpired:
        raise MissingContext(f"origin did not answer within {REMOTE_TIMEOUT_SECONDS} s whether "
                             f"branch {branch} exists") from None
    if asked.returncode == 2:
        return False
    if asked.returncode:
        error = next((line for line in asked.stderr.splitlines() if line.strip()),
                     f"exit code {asked.returncode}")
        raise MissingContext(f"origin cannot be asked whether branch {branch} exists "
                             f"(git ls-remote: {error.strip()})")
    # A pattern also matches the end of a longer name: keep only the exact branch.
    return any(line.split("\t")[-1] == ref for line in asked.stdout.splitlines())


# The workflows that must run on the default branch, as their branch filters read.
DEFAULT_BRANCH_WORKFLOWS = (".github/workflows/codeql.yml",
                            ".github/workflows/minimum-python.yml",
                            ".github/workflows/repository-integrity.yml")
# The events whose branch filters select the branch: the pushed one, or the
# base of a pull request.
BRANCH_EVENTS = ("push", "pull_request", "pull_request_target")
BRANCH_CLASS = re.compile(r"(?:[A-Za-z0-9](?:-[A-Za-z0-9])?)+")


def branch_pattern(pattern: str) -> re.Pattern[str]:
    """One branch pattern of a workflow filter, as GitHub's filter pattern rules
    read it: ``*`` matches within a path segment, ``**`` across them (``**/``
    also matches no segment, so ``**/main`` matches ``main``), ``?`` and
    ``+`` repeat the character before them, ``[...]`` lists characters and
    ranges, and ``\\`` escapes the next character. ValueError when malformed."""
    atoms: list[str] = []
    repeatable = False
    index = 0
    while index < len(pattern):
        char = pattern[index]
        if char == "\\":
            if index + 1 == len(pattern):
                raise ValueError(f"{pattern!r} ends with an escape")
            atoms.append(re.escape(pattern[index + 1]))
            repeatable, index = True, index + 2
        elif pattern.startswith("**/", index):
            atoms.append("(?:.*/)?")  # zero or more whole path segments
            repeatable, index = False, index + 3
        elif char == "*":
            double = pattern.startswith("**", index)
            atoms.append(".*" if double else "[^/]*")
            repeatable, index = False, index + (2 if double else 1)
        elif char in "?+":
            if not repeatable:
                raise ValueError(f"{pattern!r} has {char!r} with no character to repeat")
            atoms[-1] = f"(?:{atoms[-1]}){char}"
            repeatable, index = False, index + 1
        elif char == "[":
            close = pattern.find("]", index + 1)
            listed = pattern[index + 1:close] if close > 0 else ""
            if not BRANCH_CLASS.fullmatch(listed) or any(
                    first > last for first, last in re.findall(r"(.)-(.)", listed)):
                raise ValueError(f"{pattern!r} has a malformed or unclosed [...]")
            atoms.append(f"[{listed}]")
            repeatable, index = True, close + 1
        else:
            atoms.append(re.escape(char))
            repeatable, index = True, index + 1
    return re.compile("".join(atoms))


def branch_patterns(value: Any, key: str) -> list[str]:
    """The patterns of one filter key: a string or a list of non-empty strings."""
    patterns = [value] if isinstance(value, str) else value
    if not isinstance(patterns, list) or not patterns or not all(
            isinstance(item, str) and item for item in patterns):
        raise ValueError(f"{key} must be a non-empty list of non-empty patterns")
    return patterns


def runs_on_branch(event: str, config: Any, branch: str | None) -> bool:
    """Whether ``event`` with this configuration runs for ``branch`` (pushed, or
    the base of a pull request). ValueError for a malformed branch filter; with
    ``branch`` None, only the filter is checked."""
    if config is None:
        return True
    if not isinstance(config, dict):
        raise ValueError("must be a mapping of filters")
    if "branches" in config and "branches-ignore" in config:
        raise ValueError("sets both branches and branches-ignore, which GitHub rejects")
    if "branches" in config:
        patterns = branch_patterns(config["branches"], "branches")
        if all(pattern.startswith("!") for pattern in patterns):
            raise ValueError("branches lists only negations; it needs a pattern without !")
        regexes = [(pattern.startswith("!"), branch_pattern(pattern.removeprefix("!")))
                   for pattern in patterns]
        included = False
        for negated, regex in regexes:  # the last matching pattern decides
            if branch is not None and regex.fullmatch(branch):
                included = not negated
        return included
    if "branches-ignore" in config:
        patterns = branch_patterns(config["branches-ignore"], "branches-ignore")
        if any(pattern.startswith("!") for pattern in patterns):
            raise ValueError("branches-ignore cannot negate a pattern; use branches with !")
        ignored = [branch_pattern(pattern) for pattern in patterns]
        return not any(branch is not None and regex.fullmatch(branch) for regex in ignored)
    # A push filtered by tags alone runs for tags only, never for a branch.
    return not (event == "push" and ("tags" in config or "tags-ignore" in config))


def codeql_default_branch(context: Context) -> list[Finding]:
    """CodeQL, Minimum Python and Repository integrity run on the default branch, and every workflow's branch filters are well formed.

    Each branch filter of push, pull_request and pull_request_target is read as
    GitHub reads it: patterns, ! negations in order, branches-ignore, and a push
    filtered by tags alone. Branch filters cannot name the default branch
    generically, so a rename must update them.
    """
    workflows = context.root / ".github" / "workflows"
    paths = sorted(path.relative_to(context.root).as_posix() for path in
                   [*workflows.glob("*.yml"), *workflows.glob("*.yaml")]) \
        if workflows.is_dir() else []
    if not paths:
        return []
    # Imported here, not at the top: it needs PyYAML, and without PyYAML the
    # check is skipped (codeql_missing) while the others still run.
    import strict_yaml
    findings = []
    for path in paths:
        document = strict_yaml.load((context.root / path).read_text(encoding="utf-8")) or {}
        triggers = document.get("on", document.get(True)) if isinstance(document, dict) else None
        if not isinstance(triggers, dict):
            continue  # a list or a single event name sets no filter
        required = path in DEFAULT_BRANCH_WORKFLOWS
        for event in BRANCH_EVENTS:
            if event not in triggers:
                continue
            config = triggers[event]
            try:
                runs = runs_on_branch(event, config, context.default_branch)
            except ValueError as error:
                findings.append((path, f"{event} branch filter {error}."))
                continue
            if required and not runs:
                shown = {key: config[key] for key in ("branches", "branches-ignore", "tags",
                                                      "tags-ignore") if key in config}
                findings.append((path, f"{event} filter {shown} excludes the default branch "
                                       f"{context.default_branch}; make it match."))
    return findings


def required_documents(context: Context) -> list[Finding]:
    """The documents the repository requires exist and are not empty."""
    return [(name, "is missing or empty") for name in REQUIRED_DOCUMENTS
            if not (context.root / name).is_file() or not (context.root / name).stat().st_size]


def clean_working_tree(context: Context) -> list[Finding]:
    """No file differs from the commit, after every other check has run."""
    changes = gitio.text(context.root, "status", "--porcelain", "--untracked-files=all")
    return [(line[3:], f"differs from the commit (git status {line[:2].strip()}); the checks "
                       "must describe the commit as it is")
            for line in changes.splitlines() if line]


def whitespace_commands(root: Path) -> list[tuple[str, list[str]]]:
    """Check every committed file, then local changes, without needing history."""
    empty_tree = gitio.text(root, "hash-object", "-t", "tree", "--stdin", input="").strip()
    return [("whitespace-committed", ["git", "diff", "--check", empty_tree, "HEAD"]),
            ("whitespace-local", ["git", "diff", "--check", "HEAD"])]


# ---------------------------------------------------------------------------
# The list of checks.

def tool_command(name: str, *args: str, module: bool = False) -> tuple[list[str], str | None]:
    """A command for an installed tool, and why it is missing if it is."""
    if module:
        found = importlib.util.find_spec(name) is not None
        return [sys.executable, "-m", name, *args], None if found else f"{name} is not installed"
    executable = shutil.which(name)
    return [executable or name, *args], None if executable else f"{name} is not on PATH"


def checks(context: Context) -> list[Check]:
    python = sys.executable
    changelog = [python, "tools/check_changelog.py"]
    if context.repository:
        changelog += ["--repository", context.repository]
    ruff, ruff_missing = tool_command("ruff", "check", "--no-cache", "tools", "tests",
                                      module=True)
    mypy, mypy_missing = tool_command("mypy", "--no-incremental", module=True)
    actionlint, actionlint_missing = tool_command("actionlint", "-no-color")
    actionlint_probe = Probe([*actionlint, ACTIONLINT_PROBE], ACTIONLINT_PROBE_FINDING)
    node = shutil.which("node")
    node_missing = None if node else "node is not on PATH"
    no_repository = None if context.repository else "the repository is unknown (--repository)"
    # The CodeQL check reads its workflow through tools/strict_yaml.py, which needs PyYAML.
    codeql_missing = "; ".join(reason for found, reason in (
        (context.default_branch, "the default branch is unknown (--default-branch)"),
        (importlib.util.find_spec("yaml"), "PyYAML is not installed (.github/requirements.txt)"),
    ) if not found) or None
    listed: list[Check] = [
        Check("tests", [python, "-m", "unittest", "discover", "-s", "tests", "-v"]),
        Check("ruff", ruff, missing=ruff_missing),
        Check("mypy", mypy, missing=mypy_missing),
        Check("actionlint", actionlint, missing=actionlint_missing, probe=actionlint_probe),
        Check("labels-sync-self-test", [node or "node", ".github/scripts/labels-sync.mjs",
                                        "--self-test"], missing=node_missing),
        Check("labels-catalogue", [node or "node", ".github/scripts/labels-sync.mjs", "--mode",
                                   "validate"], missing=node_missing),
        Check("labels-drift-self-test", [node or "node", ".github/scripts/labels-drift.mjs",
                                         "--self-test"], missing=node_missing),
        # check_issues.py validates the label catalogue with the label script, under Node.js.
        Check("issue-configuration", [python, "tools/check_issues.py", "--check-config"],
              missing=node_missing),
        Check("template-only-content", function=template_only,
              missing=None if context.is_template is not None
              else "whether this is the template is unknown (--template)"),
        Check("issue-contact-links", function=contact_links, missing=no_repository),
        # Listed while the traffic collection, or a badge of one, is there.
        *([Check("readme-traffic-badges", function=traffic_badges, missing=no_repository)]
          if has_traffic(context.root) else []),
        Check("binaries", [python, "tools/check_binaries.py"]),
        Check("secret-files", function=secret_files),
        Check("vba-format", [python, "tools/format_vba.py", "--check"]),
        *[Check(f"vba-{name.replace('_', '-')}", [python, f"tools/check_vba_{name}.py"])
          for name in ("source", "layout", "conditionals", "jumps", "public_api",
                       "dependencies", "identifiers", "errors", "declares")],
        Check("vba-calls", [python, "tools/check_vba_calls.py"], advisory=True),
        Check("data-contract", [python, "tools/check_data_contract.py"]),
        Check("test-cases", [python, "tools/check_test_cases.py"]),
        Check("derivations", [python, "tools/check_derivations.py"]),
        Check("independent-implementations", [python, "tools/check_independent.py"]),
        Check("accuracy", [python, "tools/check_accuracy.py"]),
        Check("workbook-protection", [python, "tools/check_workbook_protection.py"]),
        Check("changelog", changelog),
        Check("generated-docs", [python, "tools/render_docs.py", "--check"]),
        Check("restated-counts", [python, "tools/check_restated_counts.py"]),
        Check("template-wording", [python, "tools/check_template_wording.py"]),
        Check("markdown-links", [python, "tools/check_links.py"]),
        Check("documented-commands", [python, "tools/check_doc_commands.py"]),
        # Every check of this list that can fail the run, this one included, needs a
        # negative proof in tools/verification_depth.json.
        Check("verification-depth", [python, "tools/check_verification_depth.py"]),
        Check("dependabot-target-branches", function=dependabot_branches),
        Check("codeql-default-branch", function=codeql_default_branch, missing=codeql_missing),
        Check("required-documents", function=required_documents),
    ]
    listed += [Check(name, command) for name, command in whitespace_commands(context.root)]
    if context.ci:
        # Last and alone, so that it also sees what the other checks changed.
        listed.append(Check("clean-working-tree", function=clean_working_tree, alone=True))
    return listed


# ---------------------------------------------------------------------------
# Running, reporting.

def annotate(finding: Finding) -> str:
    path, message = finding
    if os.environ.get("GITHUB_ACTIONS") == "true":
        return f"::error file={path}::{message}"
    return f"FAIL: {path}: {message}"


def run_check(check: Check, context: Context, strict: bool) -> Result:
    try:
        result = run_one(check, context, strict)
    except Exception as error:  # noqa: BLE001 - reported as a failed check, not raised
        # A command that cannot be started, like a function that raises, fails
        # its check and lets the other checks run.
        result = Result(check.name, "fail", check.command, None, "raised an error",
                        f"{type(error).__name__}: {error}")
    if check.advisory and result.status == "fail":
        result.status = "warn"
        result.detail = f"advisory, not a failure: {result.detail}"
    return result


def run_one(check: Check, context: Context, strict: bool) -> Result:
    if check.missing:
        return missing_result(check, check.missing, strict)
    if check.function is not None:
        try:
            findings = check.function(context)
        except MissingContext as missing:
            return missing_result(check, str(missing), strict)
        except Exception as error:  # noqa: BLE001 - reported as a failed check, not raised
            # A malformed file must fail its check, not stop the run before the
            # other checks, the report and the summary are written.
            detail = f"{type(error).__name__}: {error}"
            return Result(check.name, "fail", None, 1, "raised an error", detail)
        log = "\n".join(annotate(finding) for finding in findings)
        return Result(check.name, "fail" if findings else "pass", None, 1 if findings else 0,
                      f"{len(findings)} finding(s)" if findings else "", log)
    assert check.command is not None
    prefix = ""
    if check.probe is not None:
        prefix, failure = run_probe(check.probe, context.root)
        if failure:
            return Result(check.name, "fail", check.command, None, failure, prefix)
    output, code = run_command(check.command, context.root)
    return Result(check.name, "pass" if code == 0 else "fail", check.command, code,
                  "" if code == 0 else f"exit code {code}", prefix + output)


def run_probe(probe: Probe, root: Path) -> tuple[str, str | None]:
    """The probe's log, and why it does not prove the check can fail, or None if it does."""
    output, code = run_command(probe.command, root)
    log = f"probe: $ {' '.join(probe.command)}\n{output.rstrip()}\n"
    if code != 0 and probe.finding in output:
        return log + f"probe: rejected as expected (exit code {code})\n\n", None
    found = "exited 0" if code == 0 else f"exited {code} without reporting {probe.finding!r}"
    log += (f"FAIL: the probe {found}; the tool must exit non-zero and report "
            f"{probe.finding!r} before the check can pass\n")
    return log, f"probe not rejected: {found}"


def run_command(command: list[str], root: Path) -> tuple[str, int]:
    """A command's combined output and exit code; 1 when it cannot run or times out."""
    try:
        completed = subprocess.run(command, cwd=root, capture_output=True, text=True,
                                   timeout=TIMEOUT_SECONDS, check=False)
    except (OSError, subprocess.TimeoutExpired) as error:
        return str(error), 1
    return completed.stdout + completed.stderr, completed.returncode


def missing_result(check: Check, reason: str, strict: bool) -> Result:
    """A check that cannot run here: skipped, or failed under --strict."""
    return Result(check.name, "fail" if strict else "skip", check.command, None, reason, reason)


def tracked_hashes(root: Path) -> dict[str, str]:
    names = gitio.tracked(root)
    return {name: hashlib.sha256((root / name).read_bytes()).hexdigest()
            for name in names if (root / name).is_file()}


def build_report(context: Context, results: Sequence[Result]) -> dict[str, Any]:
    root = context.root
    return {
        "schema_version": 1,
        "commit": gitio.text(root, "rev-parse", "HEAD").strip(),
        "dirty": bool(gitio.text(root, "status", "--porcelain")),
        "pull_request_head": os.environ.get("PR_HEAD_SHA") or None,
        "repository": context.repository,
        "generated_at": datetime.datetime.now(datetime.UTC).isoformat(timespec="seconds"),
        "python": sys.version.split()[0],
        "platform": platform.platform(),
        "scope": SCOPE,
        "status": "fail" if any(r.status == "fail" for r in results) else "pass",
        "checks": [{"name": r.name, "status": r.status, "command": r.command,
                    "exit_code": r.exit_code, "detail": r.detail, "log": f"{r.name}.log"}
                   for r in results],
        "files_sha256": tracked_hashes(root),
    }


def summary(report: dict[str, Any]) -> str:
    marks = {"pass": "✅ pass", "fail": "❌ fail", "warn": "⚠️ warn", "skip": "⏭️ skip"}
    lines = [f"## Repository integrity: {report['status'].upper()}", "",
             f"- Commit: `{report['commit']}`" + (" (dirty working tree)" if report["dirty"]
                                                   else ""),
             f"- Pull request head: `{report['pull_request_head']}`"
             if report["pull_request_head"] else "- Pull request head: none",
             f"- Scope: {report['scope']}.", "",
             "| Check | Result | Detail |", "| --- | --- | --- |"]
    lines += [f"| `{check['name']}` | {marks[check['status']]} | "
              f"{check['detail'].replace('|', '/')} |" for check in report["checks"]]
    lines += ["", "Logs and `report.json`, with the SHA-256 of every tracked file, are in "
              f"`{REPORT_DIR.as_posix()}/`, kept as the run's artifact in CI."]
    return "\n".join(lines) + "\n"


def results_in_order(context: Context, listed: Sequence[Check],
                     strict: bool) -> Iterator[Result]:
    """Each check's result, in the listed order, as soon as it and every check
    before it have finished.

    The command checks run in a pool of threads, up to MAX_WORKERS of them at
    once and no more than the CPUs; the function checks, which are quick and
    may change the process's state, run one at a time in this thread, as does
    a check whose tool or context is missing. A check marked alone splits the
    list: it runs once the checks before it have finished, and the checks
    after it start once it has finished.
    """
    workers = max(1, min(MAX_WORKERS, os.cpu_count() or 1))
    with concurrent.futures.ThreadPoolExecutor(max_workers=workers) as pool:
        start = 0
        while start < len(listed):
            if listed[start].alone:
                yield run_check(listed[start], context, strict)
                start += 1
                continue
            end = next((index for index in range(start, len(listed)) if listed[index].alone),
                       len(listed))
            futures = {index: pool.submit(run_check, listed[index], context, strict)
                       for index in range(start, end)
                       if listed[index].function is None and not listed[index].missing}
            for index in range(start, end):
                yield futures[index].result() if index in futures \
                    else run_check(listed[index], context, strict)
            start = end


def run(context: Context, listed: Sequence[Check], out: Path, strict: bool,
        summary_path: Path | None = None) -> int:
    out.mkdir(parents=True, exist_ok=True)
    results = []
    for check, result in zip(listed, results_in_order(context, listed, strict), strict=True):
        command = " ".join(result.command) if result.command else check.name
        (out / f"{check.name}.log").write_text(f"$ {command}\n{result.log}", encoding="utf-8")
        results.append(result)
        print(f"{check.name}: {result.status.upper()}"
              + (f" ({result.detail})" if result.detail else ""), flush=True)
        if result.status in ("fail", "warn") and result.log:
            print(result.log.rstrip(), flush=True)
    report = build_report(context, results)
    (out / "report.json").write_text(json.dumps(report, indent=2) + "\n", encoding="utf-8")
    text = summary(report)
    (out / "summary.md").write_text(text, encoding="utf-8")
    if summary_path is not None:
        with summary_path.open("a", encoding="utf-8") as handle:
            handle.write(text)
    skipped = sum(result.status == "skip" for result in results)
    print(f"repository: {report['status'].upper()} ({len(results)} check(s), {skipped} "
          f"skipped; {SCOPE})")
    return 1 if report["status"] == "fail" else 0


def detect_repository(root: Path) -> str | None:
    try:
        url = gitio.text(root, "remote", "get-url", "origin").strip()
    except GitError:
        return None
    match = GITHUB_REMOTE.search(url)
    return match.group(1) if match else None


def detect_default_branch(root: Path) -> str | None:
    try:
        ref = gitio.text(root, "symbolic-ref", "--quiet", "refs/remotes/origin/HEAD").strip()
    except GitError:
        return None
    return ref.removeprefix("refs/remotes/origin/") or None


def parse_flag(text: str | None) -> bool | None:
    if text is None or text == "":
        return None
    if text.lower() in ("true", "false"):
        return text.lower() == "true"
    raise argparse.ArgumentTypeError(f"expected true or false, not {text!r}")


def main(argv: Sequence[str] | None = None,
         listed: Callable[[Context], list[Check]] = checks) -> int:
    parser = argparse.ArgumentParser(description=(__doc__ or "").splitlines()[0])
    parser.add_argument("--root", type=Path, default=ROOT)
    parser.add_argument("--repository", metavar="OWNER/REPOSITORY",
                        help="default: the GitHub repository of the origin remote")
    parser.add_argument("--default-branch", help="default: origin's HEAD branch")
    parser.add_argument("--template", type=parse_flag, metavar="true|false",
                        help="whether this repository is the template (default: unknown)")
    parser.add_argument("--strict", action="store_true",
                        help="fail a check whose tool or context is missing, as CI does")
    parser.add_argument("--ci", action="store_true",
                        help="fail when the working tree differs from the commit, as CI does")
    parser.add_argument("--summary", type=Path, help="also append the summary to this file")
    args = parser.parse_args(argv)
    root = args.root.resolve()
    context = Context(root, args.repository or detect_repository(root),
                      args.default_branch or detect_default_branch(root), args.template,
                      args.ci)
    return run(context, listed(context), root / REPORT_DIR, args.strict, args.summary)


if __name__ == "__main__":
    sys.exit(main())
