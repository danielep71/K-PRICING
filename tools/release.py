#!/usr/bin/env python3
"""Verify, describe and package a release tag, and check the published release.

Subcommands, all read-only apart from the files that bundle writes:

- verify TAG --default-branch BRANCH [--excel-evidence DIR --workbook FILE]
  Fails unless:
  - TAG is vX.Y.Z and annotated;
  - X.Y.Z is the newest dated release of CHANGELOG.md and the content of
    VERSION at the tag;
  - the tag's commit is on origin/BRANCH;
  - the tag message names that commit on a "Candidate <sha>." line and has
    a "Deviations:" line followed by text on the same line, such as "none",
    or by a list on the next lines, as RELEASING.md requires;
  - when the tag's commit holds a VBA component (a .bas, .cls or .frm file
    anywhere in its tree, as tools/check_excel_run.py compares them), the
    message declares its Excel run on one line
    "Excel evidence: <digest> (run <run sha>)." or lists an
    "- Invariant 4: <text>" item under "Deviations:". "Deviations: none."
    does not waive the run. A line that starts "Excel evidence:" in any
    other form fails, with or without a VBA component.
  Without --excel-evidence, as in the Release workflow, the declaration is
  checked for its form only: the evidence never leaves the machine it was
  recorded on. With --excel-evidence and --workbook, on that machine before
  the tag is pushed, the message must declare the evidence, and the gate of
  the evidence subcommand must pass for the tag's commit and the declared run
  commit, with the declared digest.
- evidence COMMIT --excel-evidence DIR --workbook FILE [--run-commit SHA]
  The Excel evidence gate, run before tagging. It runs the checks of
  tools/check_excel_run.py with --require-binding on the evidence folder DIR
  and the workbook FILE, in process: against COMMIT, or against SHA, the
  commit the workbook was built from, with --equivalent-to COMMIT, so that
  every run input must have the same Git blob at both commits. It fails on
  any finding, on a binding other than VERIFIED, on a run of another commit
  without the Equivalent result, and on carried evidence, which it never
  asks for. On success it prints the evidence digest and the line to put in
  the tag message. The digest is the SHA-256 of a listing, one line
  "<sha256 of the file>  evidence/<path>" for each file under DIR, sorted
  by its path relative to DIR with "/" separators, then one line
  "<sha256 of FILE>  workbook", each line ending in a line feed. The same
  files give the same digest wherever they are; the workbook's name does not
  count.
- integrity --repository OWNER/NAME --sha SHA
  Fails unless the check run Repository integrity succeeded on SHA. It reads
  the check runs through the GitHub CLI.
- notes TAG
  Prints the release notes: the tag message, then the changelog section of
  the version.
- bundle REF --name NAME --out DIR
  Writes DIR/NAME-VERSION-source.tar.gz, a gzip of `git archive` of REF with
  no timestamp in the gzip header, and DIR/SHA256SUMS. The same REF gives the
  same bytes. VERSION is the tag's version, or the short commit for another
  ref, as in a rehearsal.
- verify-published TAG --repository OWNER/NAME [--milestone N]
  [--asset FILE=SHA256 ...]
  After publication, prints one line per fact and fails unless, with
  OWNER/NAME read back from GitHub in its canonical letter case:
  - TAG is annotated on GitHub and points at the commit its Candidate line
    names;
  - Repository integrity succeeded on that commit;
  - the GitHub Release of TAG is published, not a draft, and its notes name
    that commit;
  - its assets are exactly NAME-VERSION-source.tar.gz, SHA256SUMS and each
    FILE given; SHA256SUMS records exactly the bundle and its SHA-256, and
    each FILE has the SHA-256 given;
  - each FILE that is an Office Open XML package, or is named like one
    (.xlsx, .xlsm, .docx and the like), has no finding of
    tools/check_workbook_privacy.py, whatever its SHA-256: no OneDrive or
    SharePoint path or URL, absolute path under a user profile, person or
    organization, Purview sensitivity label or SharePoint metadata. A file
    named like a package that is not one, such as an encrypted workbook,
    cannot be checked and fails;
  - gh attestation verify passes on the bundle, as built by the Release
    workflow at TAG, unless the notes state that it is not attested;
  - with --milestone, milestone N is closed and has no open issue.
  It only reads, through gh: the assets are read into memory, and the bundle
  is copied to a temporary file, removed afterwards, for gh attestation.

Standard library only, apart from git, and gh for integrity and
verify-published. The Excel evidence gate imports tools/check_excel_run.py
only when evidence is given.
"""

from __future__ import annotations

import argparse
import gzip
import hashlib
import json
import re
import subprocess
import sys
import tempfile
import urllib.parse
from collections.abc import Callable, Sequence
from dataclasses import dataclass, field
from pathlib import Path
from typing import TYPE_CHECKING

sys.path.insert(0, str(Path(__file__).resolve().parent))

import check_workbook_privacy  # noqa: E402
import gitio  # noqa: E402
from check_changelog import CHANGELOG, VERSION, released_versions  # noqa: E402
from gitio import GitError  # noqa: E402
from vba_lexer import VBA_SUFFIXES  # noqa: E402

if TYPE_CHECKING:
    import check_excel_run

ROOT = Path(__file__).resolve().parents[1]
TAG = re.compile(r"^v(\d+\.\d+\.\d+)$")
CANDIDATE = re.compile(r"^Candidate ([0-9a-f]{40})\.\s*$", re.M)
# "Deviations:" followed by text on the same line, or by a "- " or "* " list
# that starts on the next line.
DEVIATIONS = re.compile(r"^Deviations:(?:[ \t]+\S|[ \t]*\n[ \t]*[-*][ \t]+\S)", re.M)
SECTION = re.compile(r"^## \[(?P<label>[^\]]+)\][^\n]*\n(?P<body>.*?)(?=^## \[|^\[[^\]]+\]: |\Z)",
                     re.M | re.S)
REQUIRED_CHECK = "Repository integrity"
WORKFLOW = ".github/workflows/release.yml"
SUMS = "SHA256SUMS"
# The sentence the Release workflow adds to the notes of an unattested draft.
UNATTESTED = "The source bundle has no build provenance attestation"
DIGEST = re.compile(r"^[0-9a-f]{64}$")
# The tag message's declaration of the Excel evidence, which the evidence
# subcommand prints; any other line that starts with its label is malformed.
EXCEL_EVIDENCE = "Excel evidence:"
EXCEL_LINE = re.compile(r"^Excel evidence: ([0-9a-f]{64}) \(run ([0-9a-f]{40})\)\.[ \t]*$")
INVARIANT_4 = re.compile(r"^Invariant 4:[ \t]*\S")
LIST_ITEM = re.compile(r"^[ \t]*[-*][ \t]+(.*)$")
# The names of Office Open XML packages: an asset named so must be one, so that
# its content can be checked for private data.
OFFICE_SUFFIXES = (".xlsx", ".xlsm", ".xltx", ".xltm", ".xlam", ".docx", ".docm", ".dotx",
                   ".dotm", ".pptx", ".pptm", ".potx", ".potm", ".ppsx", ".ppsm")

# Runs gh with the arguments given and returns the completed process, with
# stdout and stderr as bytes. Tests replace it with a fake.
Runner = Callable[[Sequence[str]], "subprocess.CompletedProcess[bytes]"]
Fact = tuple[str, str]  # PASS, FAIL or NOTE, and what was found
# check_excel_run.check, or a fake in the tests.
Checker = Callable[..., "check_excel_run.Evidence"]


def run_gh(args: Sequence[str]) -> subprocess.CompletedProcess[bytes]:
    return subprocess.run(["gh", *args], capture_output=True, check=False)


class ReleaseError(Exception):
    """The release cannot proceed; the message says why."""


def git(root: Path, *args: str) -> str:
    """Git's output as text; a ReleaseError naming the command when Git fails."""
    try:
        return gitio.git(root, *args).decode("utf-8")
    except GitError as error:
        raise ReleaseError(f"git {' '.join(args)}: {error.stderr.strip()}") from error


def version_of(tag: str) -> str:
    match = TAG.match(tag)
    if not match:
        raise ReleaseError(f"tag {tag!r} is not vX.Y.Z")
    return match.group(1)


def verify(root: Path, tag: str, default_branch: str,
           evidence: tuple[Path, Path] | None = None,
           checker: Checker | None = None) -> list[str]:
    """Every reason the tag cannot be released; empty when it can. With
    ``evidence``, the evidence folder and the workbook, the declared Excel
    evidence is checked too (``checker`` replaces check_excel_run.check)."""
    try:
        version = version_of(tag)
    except ReleaseError as error:
        return [str(error)]
    problems = []
    try:
        kind = git(root, "cat-file", "-t", f"refs/tags/{tag}").strip()
    except ReleaseError:
        return [f"tag {tag} is missing"]
    if kind != "tag":
        problems.append(f"tag {tag} is {kind}, not an annotated tag")
    commit = git(root, "rev-parse", f"{tag}^{{commit}}").strip()
    try:
        changelog = git(root, "show", f"{commit}:{CHANGELOG}")
        released = released_versions(changelog)
    except ReleaseError:
        released = []
    if not released or released[0] != version:
        newest = released[0] if released else "none"
        problems.append(f"{CHANGELOG} at {tag} releases {newest} as its newest version, "
                        f"not {version}")
    try:
        recorded = git(root, "show", f"{commit}:{VERSION}").strip()
    except ReleaseError:
        recorded = ""
    if recorded != version:
        problems.append(f"{VERSION} at {tag} is {recorded or 'missing'}, not {version}")
    try:
        gitio.git(root, "merge-base", "--is-ancestor", commit, f"origin/{default_branch}")
    except GitError:
        problems.append(f"{tag} ({commit[:12]}) is not on origin/{default_branch}")
    if kind == "tag":
        message = git(root, "tag", "-l", "--format=%(contents)", tag)
        candidates = CANDIDATE.findall(message)
        if candidates != [commit]:
            problems.append(f"the message of {tag} must name its commit on one line "
                            f"'Candidate {commit}.'")
        if not DEVIATIONS.search(message):
            problems.append(f"the message of {tag} has no 'Deviations:' line followed by "
                            "text, such as 'none', or by a list")
        problems += excel_declaration(root, tag, commit, message)
        if evidence is not None:
            problems += declared_evidence(root, tag, commit, message, evidence, checker)
    elif evidence is not None:
        problems.append(f"{tag} has no message to declare the Excel evidence in")
    return problems


def holds_vba(root: Path, commit: str) -> bool:
    """Whether the commit's tree holds a VBA component, at any depth."""
    listed = git(root, "ls-tree", "-r", "-z", "--name-only", commit).split("\0")
    return any(path.lower().endswith(VBA_SUFFIXES) for path in listed)


def deviation_items(message: str) -> list[str]:
    """The items listed under the first "Deviations:" line, each with the lines
    indented under it."""
    lines = message.splitlines()
    start = next((index for index, line in enumerate(lines)
                  if line.startswith("Deviations:")), None)
    if start is None:
        return []
    items: list[str] = []
    for line in lines[start + 1:]:
        item = LIST_ITEM.match(line)
        if item:
            items.append(item.group(1).strip())
        elif items and line[:1] in (" ", "\t") and line.strip():
            items[-1] += " " + line.strip()
        else:
            break
    return items


def excel_lines(message: str) -> list[str]:
    return [line for line in message.splitlines() if line.startswith(EXCEL_EVIDENCE)]


def excel_declaration(root: Path, tag: str, commit: str, message: str) -> list[str]:
    """Why the message does not declare the Excel run that the commit needs: an
    evidence line, or an invariant 4 deviation. Only the form is checked."""
    lines = excel_lines(message)
    problems = [f"the message of {tag} has a malformed line {line!r}: write "
                f"'{EXCEL_EVIDENCE} <64-hex digest> (run <40-hex commit>).', as "
                "tools/release.py evidence prints it" for line in lines
                if not EXCEL_LINE.match(line)]
    if len(lines) > 1:
        problems.append(f"the message of {tag} has {len(lines)} '{EXCEL_EVIDENCE}' lines, "
                        "not one")
    if problems or lines:
        return problems
    if not holds_vba(root, commit):
        return []
    if any(INVARIANT_4.match(item) for item in deviation_items(message)):
        return []
    return [f"{tag} ({commit}) holds VBA components, but its message has neither an "
            f"'{EXCEL_EVIDENCE} <digest> (run <sha>).' line nor a '- Invariant 4: <text>' "
            "item under 'Deviations:'"]


def declared_evidence(root: Path, tag: str, commit: str, message: str,
                      evidence: tuple[Path, Path], checker: Checker | None) -> list[str]:
    """The local check: the evidence given passes the gate for the commit and the
    run the message declares, and has the digest it declares."""
    declared = [match for match in map(EXCEL_LINE.match, excel_lines(message)) if match]
    if len(declared) != 1:
        return [f"the message of {tag} declares no Excel evidence to check against "
                f"{evidence[0]}"]
    digest, run = declared[0].groups()
    try:
        result = excel_gate(root, commit, evidence[0], evidence[1], run, checker)
    except ReleaseError as error:
        return [f"the Excel evidence of {tag}: {error}"]
    problems = list(result.problems)
    if not problems and result.digest != digest:
        problems.append(f"the message of {tag} declares Excel evidence {digest}, but "
                        f"{evidence[0]} and {evidence[1]} have the digest {result.digest}")
    return problems


@dataclass
class ExcelGate:
    """The result of the Excel evidence gate for one candidate commit."""

    candidate: str
    run: str
    report: str = ""
    digest: str = ""
    problems: list[str] = field(default_factory=list)

    @property
    def line(self) -> str:
        """The tag message's declaration of this evidence."""
        return f"{EXCEL_EVIDENCE} {self.digest} (run {self.run})."


def evidence_digest(folder: Path, workbook: Path) -> str:
    """The SHA-256 of the listing of the evidence folder's files and the workbook,
    as the module docstring defines it."""
    files = sorted((path.relative_to(folder).as_posix(), path)
                   for path in folder.rglob("*") if path.is_file())
    lines = [f"{hashlib.sha256(path.read_bytes()).hexdigest()}  evidence/{name}\n"
             for name, path in files]
    lines.append(f"{hashlib.sha256(workbook.read_bytes()).hexdigest()}  workbook\n")
    return hashlib.sha256("".join(lines).encode("utf-8")).hexdigest()


def commit_of(root: Path, name: str) -> str:
    try:
        return git(root, "rev-parse", "--verify", f"{name}^{{commit}}").strip()
    except ReleaseError as error:
        raise ReleaseError(f"commit {name} not found") from error


def excel_gate(root: Path, candidate: str, folder: Path, workbook: Path,
               run: str | None = None, checker: Checker | None = None) -> ExcelGate:
    """Run the checks of tools/check_excel_run.py with a required binding on the
    evidence of ``run`` (by default the candidate), with --equivalent-to the
    candidate when the run is of another commit. Carried evidence is never
    asked for, and refused if a checker reports it."""
    candidate = commit_of(root, candidate)
    run = commit_of(root, run) if run is not None else candidate
    gate = ExcelGate(candidate, run)
    for path, kind, exists in ((folder, "evidence folder", folder.is_dir()),
                               (workbook, "workbook", workbook.is_file())):
        if not exists:
            gate.problems.append(f"the {kind} {path} is missing")
    if gate.problems:
        return gate
    # Imported here, so that verify without evidence, as in CI, never loads it.
    import check_excel_run

    equivalent = candidate if run != candidate else None
    found = (checker or check_excel_run.check)(root, run, folder, workbook,
                                               require_binding=True,
                                               equivalent_to=equivalent)
    gate.report = check_excel_run.report(run, found, equivalent)
    gate.problems += [f"excel-run: {finding}" for finding in found.findings]
    if found.binding != "VERIFIED":
        gate.problems.append(f"the binding is {found.binding}, not VERIFIED")
    if equivalent is not None and not found.equivalence:
        gate.problems.append(f"the run of {run[:12]} is not shown equivalent to "
                             f"{candidate[:12]}")
    if getattr(found, "carried", ""):
        gate.problems.append("the evidence is carried from another commit; carried "
                             "evidence never certifies a release")
    if not gate.problems:
        gate.digest = evidence_digest(folder, workbook)
    return gate


def gh_output(runner: Runner, *args: str) -> bytes:
    """The standard output of gh args; a ReleaseError with its last error line."""
    completed = runner(args)
    if completed.returncode:
        error = completed.stderr.decode("utf-8", "replace").strip()
        raise ReleaseError(error.splitlines()[-1] if error else f"exit {completed.returncode}")
    return completed.stdout


def gh_json(runner: Runner, *args: str) -> object:
    try:
        return json.loads(gh_output(runner, *args).decode("utf-8") or "null")
    except ValueError as error:
        raise ReleaseError(f"gh {args[0]} printed no JSON: {error}") from error


def integrity(repository: str, sha: str, runner: Runner = run_gh) -> list[str]:
    """Why Repository integrity did not succeed on sha, if it did not."""
    try:
        payload = gh_json(runner, "api",
                          f"repos/{repository}/commits/{sha}/check-runs?per_page=100")
    except ReleaseError as error:
        return [f"cannot read the check runs of {sha[:12]}: {error}"]
    found = payload.get("check_runs") if isinstance(payload, dict) else None
    runs = [run for run in found or [] if isinstance(run, dict)
            and run.get("name") == REQUIRED_CHECK]
    if not runs:
        return [f"{REQUIRED_CHECK} has not run on {sha[:12]}"]
    if not any(run.get("conclusion") == "success" for run in runs):
        conclusions = sorted({str(run.get("conclusion")) for run in runs})
        return [f"{REQUIRED_CHECK} on {sha[:12]} did not succeed: {', '.join(conclusions)}"]
    return []


def notes(root: Path, tag: str) -> str:
    version = version_of(tag)
    commit = git(root, "rev-parse", f"{tag}^{{commit}}").strip()
    changelog = git(root, "show", f"{commit}:{CHANGELOG}")
    section = next((match.group("body").strip() for match in SECTION.finditer(changelog)
                    if match.group("label") == version), None)
    if section is None:
        raise ReleaseError(f"{CHANGELOG} at {tag} has no section for {version}")
    message = git(root, "tag", "-l", "--format=%(contents)", tag).strip()
    return f"{message}\n\n## Changes\n\n{section}\n"


def bundle(root: Path, ref: str, name: str, out: Path) -> list[Path]:
    """The source archive of ref and its checksums, byte for byte reproducible."""
    commit = git(root, "rev-parse", f"{ref}^{{commit}}").strip()
    match = TAG.match(ref)
    version = match.group(1) if match else commit[:12]
    stem = f"{name}-{version}"
    try:
        archive = gitio.git(root, "archive", "--format=tar", f"--prefix={stem}/", commit)
    except GitError as error:
        raise ReleaseError(f"git archive {ref}: {error.stderr}") from error
    out.mkdir(parents=True, exist_ok=True)
    target = out / f"{stem}-source.tar.gz"
    with open(target, "wb") as raw, gzip.GzipFile(filename="", mode="wb", fileobj=raw,
                                                  mtime=0, compresslevel=9) as packed:
        packed.write(archive)
    sums = out / "SHA256SUMS"
    digest = hashlib.sha256(target.read_bytes()).hexdigest()
    sums.write_text(f"{digest}  {target.name}\n", encoding="utf-8")
    return [target, sums]


def mapping(value: object) -> dict[str, object]:
    return value if isinstance(value, dict) else {}


def published_tag(repository: str, tag: str, runner: Runner) -> tuple[Fact, str | None]:
    """The fact about the tag on GitHub, and the commit it points at, if read."""
    try:
        target = mapping(mapping(gh_json(runner, "api",
                                         f"repos/{repository}/git/ref/tags/{tag}"))
                         .get("object"))
        if target.get("type") != "tag":
            commit = (str(target.get("sha")) if target.get("type") == "commit"
                      and target.get("sha") else None)
            return ("FAIL", f"tag {tag} is {target.get('type') or 'missing'}, "
                            "not an annotated tag"), commit
        annotated = mapping(gh_json(runner, "api",
                                    f"repos/{repository}/git/tags/{target.get('sha')}"))
    except ReleaseError as error:
        return ("FAIL", f"cannot read tag {tag}: {error}"), None
    pointed = mapping(annotated.get("object"))
    if pointed.get("type") != "commit":
        return ("FAIL", f"tag {tag} does not point at a commit"), None
    commit = str(pointed.get("sha"))
    candidates = CANDIDATE.findall(str(annotated.get("message") or ""))
    if candidates == [commit]:
        return ("PASS", f"tag {tag} is annotated and points at {commit}, the commit its "
                        "Candidate line names"), commit
    named = ", ".join(candidates) or "no commit"
    return ("FAIL", f"tag {tag} points at {commit}, but its Candidate line names "
                    f"{named}"), commit


def asset_content(repository: str, asset: dict[str, object], runner: Runner) -> bytes:
    """The bytes of a release asset, read through its API URL into memory."""
    path = urllib.parse.urlparse(str(asset.get("apiUrl") or "")).path.lstrip("/")
    if not path.startswith(f"repos/{repository}/releases/assets/"):
        raise ReleaseError(f"its API URL {asset.get('apiUrl')!r} is not an asset of "
                           f"{repository}")
    return gh_output(runner, "api", "-H", "Accept: application/octet-stream", path)


def checksums(text: str) -> dict[str, str]:
    """SHA256SUMS as file name to SHA-256, in the format bundle writes."""
    records = {}
    for line in text.splitlines():
        if line.strip():
            digest, _, name = line.partition("  ")
            records[name.removeprefix("*")] = digest
    return records


def attestation(repository: str, tag: str, name: str, content: bytes,
                runner: Runner) -> Fact:
    """gh attestation verify on a temporary copy of the bundle."""
    with tempfile.TemporaryDirectory(prefix="release-verify-") as directory:
        copy = Path(directory) / name
        copy.write_bytes(content)
        try:
            verified = gh_json(runner, "attestation", "verify", str(copy),
                               "--repo", repository,
                               "--signer-workflow", f"{repository}/{WORKFLOW}",
                               "--source-ref", f"refs/tags/{tag}", "--format", "json")
        except ReleaseError as error:
            return "FAIL", f"gh attestation verify failed on {name}: {error}"
    if not isinstance(verified, list) or not verified:
        return "FAIL", f"gh attestation verify found no attestation of {name}"
    return "PASS", (f"gh attestation verify passed on {name}: built by {WORKFLOW} "
                    f"at refs/tags/{tag}")


def privacy_facts(name: str, content: bytes) -> list[Fact]:
    """The facts about the private data of an attached Office package; none for
    another file."""
    try:
        findings = check_workbook_privacy.scan(content)
    except check_workbook_privacy.NotAPackage as error:
        if name.lower().endswith(OFFICE_SUFFIXES):
            return [("FAIL", f"{name} cannot be checked for private data: it is {error}")]
        return []
    if not findings:
        return [("PASS", f"{name} has no finding of tools/check_workbook_privacy.py")]
    return [("FAIL", f"{name} carries private data: {len(findings)} finding(s) of "
                     "tools/check_workbook_privacy.py"),
            *(("FAIL", f"{name}: {finding}") for finding in findings)]


def published_assets(repository: str, tag: str, release: dict[str, object],
                     extra: dict[str, str], runner: Runner) -> list[Fact]:
    """The facts about the assets of the release, their digests and attestation."""
    bundle_name = f"{repository.rsplit('/', 1)[-1]}-{version_of(tag)}-source.tar.gz"
    expected = [bundle_name, SUMS, *extra]
    listed = release.get("assets")
    assets = {str(asset.get("name")): asset
              for asset in map(mapping, listed if isinstance(listed, list) else [])}
    facts: list[Fact] = [("FAIL", f"asset {name} is missing from the release")
                         for name in expected if name not in assets]
    facts += [("FAIL", f"asset {name} is not an expected asset")
              for name in sorted(set(assets) - set(expected))]
    if not facts:
        facts.append(("PASS", "the release has exactly the expected assets: "
                              f"{', '.join(expected)}"))
    contents: dict[str, bytes] = {}
    for name in expected:
        if name in assets:
            try:
                contents[name] = asset_content(repository, assets[name], runner)
            except ReleaseError as error:
                facts.append(("FAIL", f"cannot read asset {name}: {error}"))
    digests = {name: hashlib.sha256(content).hexdigest() for name, content in contents.items()}
    if SUMS in contents:
        records = checksums(contents[SUMS].decode("utf-8", "replace"))
        if set(records) != {bundle_name}:
            facts.append(("FAIL", f"{SUMS} records {', '.join(sorted(records)) or 'nothing'}, "
                                  f"not exactly {bundle_name}"))
        elif bundle_name in digests:
            actual, recorded = digests[bundle_name], records[bundle_name]
            facts.append(("PASS", f"{bundle_name} has SHA-256 {actual}, as {SUMS} records")
                         if actual == recorded else
                         ("FAIL", f"{bundle_name} has SHA-256 {actual}, but {SUMS} "
                                  f"records {recorded}"))
    for name, recorded in extra.items():
        if name in digests:
            facts.append(("PASS", f"{name} has SHA-256 {digests[name]}, as recorded")
                         if digests[name] == recorded else
                         ("FAIL", f"{name} has SHA-256 {digests[name]}, not the recorded "
                                  f"{recorded}"))
    for name in extra:
        if name in contents:
            facts += privacy_facts(name, contents[name])
    if bundle_name in contents:
        if UNATTESTED in str(release.get("body") or ""):
            facts.append(("NOTE", f"{bundle_name} is not attested, as the release notes "
                                  "state: record its absence as a deviation"))
        else:
            facts.append(attestation(repository, tag, bundle_name, contents[bundle_name],
                                     runner))
    return facts


def published_milestone(repository: str, number: int, runner: Runner) -> Fact:
    try:
        milestone = mapping(gh_json(runner, "api", f"repos/{repository}/milestones/{number}"))
    except ReleaseError as error:
        return "FAIL", f"cannot read milestone {number}: {error}"
    title, state = milestone.get("title"), milestone.get("state")
    open_issues = milestone.get("open_issues")
    if state == "closed" and open_issues == 0:
        return "PASS", f"milestone {number} ({title}) is closed and has no open issue"
    return "FAIL", f"milestone {number} ({title}) is {state}, with {open_issues} open issue(s)"


def verify_published(repository: str, tag: str, milestone: int | None = None,
                     extra: dict[str, str] | None = None,
                     runner: Runner = run_gh) -> list[Fact]:
    """One fact per check of the published release of tag; it only reads."""
    try:
        version_of(tag)
    except ReleaseError as error:
        return [("FAIL", str(error))]
    # GitHub resolves OWNER/NAME in any letter case, but the asset URLs, the
    # bundle name and the attestation's signer carry the canonical one.
    try:
        canonical = mapping(gh_json(runner, "api", f"repos/{repository}")).get("full_name")
    except ReleaseError as error:
        return [("FAIL", f"cannot read repository {repository}: {error}")]
    if not isinstance(canonical, str) or canonical.lower() != repository.lower():
        return [("FAIL", f"repository {repository} resolves to {canonical}")]
    repository = canonical
    fact, commit = published_tag(repository, tag, runner)
    facts = [fact]
    if commit is not None:
        problems = integrity(repository, commit, runner)
        facts += [("FAIL", problem) for problem in problems] or [
            ("PASS", f"{REQUIRED_CHECK} succeeded on {commit}")]
    try:
        release = mapping(gh_json(runner, "release", "view", tag, "--repo", repository,
                                  "--json", "tagName,isDraft,publishedAt,body,assets,url"))
    except ReleaseError as error:
        facts.append(("FAIL", f"cannot read the release of {tag}: {error}"))
    else:
        facts.append(("FAIL", f"the release of {tag} is a draft, not published")
                     if release.get("isDraft") is not False else
                     ("PASS", f"the release of {tag} is published, not a draft: "
                              f"{release.get('publishedAt')}, {release.get('url')}"))
        if commit is not None:
            named = commit in str(release.get("body") or "")
            facts.append(("PASS", f"the release notes name the candidate {commit}") if named
                         else ("FAIL", f"the release notes do not name the candidate {commit}"))
        facts += published_assets(repository, tag, release, extra or {}, runner)
    if milestone is not None:
        facts.append(published_milestone(repository, milestone, runner))
    return facts


def asset_record(value: str) -> tuple[str, str]:
    """FILE=SHA256 from the command line."""
    name, _, digest = value.rpartition("=")
    if not name or not DIGEST.match(digest.lower()):
        raise argparse.ArgumentTypeError(f"{value!r} is not FILE=SHA256")
    return name, digest.lower()


def main(argv: Sequence[str] | None = None) -> int:
    parser = argparse.ArgumentParser(description=(__doc__ or "").splitlines()[0])
    parser.add_argument("--root", type=Path, default=ROOT)
    commands = parser.add_subparsers(dest="command", required=True)
    check = commands.add_parser("verify", help="verify a release tag")
    check.add_argument("tag")
    check.add_argument("--default-branch", required=True)
    check.add_argument("--excel-evidence", type=Path, metavar="DIR",
                       help="the Excel run's evidence folder, checked against the tag's "
                            "declaration on the machine that holds it; needs --workbook")
    check.add_argument("--workbook", type=Path, metavar="FILE",
                       help="the workbook of that run; needs --excel-evidence")
    gate = commands.add_parser("evidence", help="check Excel evidence for a commit and print "
                                                "the tag message line")
    gate.add_argument("commit", help="the candidate commit")
    gate.add_argument("--excel-evidence", type=Path, metavar="DIR", required=True,
                      help="the evidence folder of the Excel run")
    gate.add_argument("--workbook", type=Path, metavar="FILE", required=True,
                      help="the workbook the run used")
    gate.add_argument("--run-commit", metavar="SHA",
                      help="the commit the workbook was built from, when it is not the "
                           "candidate; every run input must be identical at both")
    runs = commands.add_parser("integrity", help="require a successful Repository integrity")
    runs.add_argument("--repository", required=True)
    runs.add_argument("--sha", required=True)
    describe = commands.add_parser("notes", help="print the release notes of a tag")
    describe.add_argument("tag")
    package = commands.add_parser("bundle", help="write the source archive and checksums")
    package.add_argument("ref")
    package.add_argument("--name", required=True)
    package.add_argument("--out", type=Path, required=True)
    published = commands.add_parser("verify-published",
                                    help="check a published release; read-only")
    published.add_argument("tag")
    published.add_argument("--repository", required=True)
    published.add_argument("--milestone", type=int)
    published.add_argument("--asset", type=asset_record, action="append", default=[],
                           metavar="FILE=SHA256",
                           help="another expected asset, with its recorded SHA-256")
    args = parser.parse_args(argv)
    root = args.root.resolve()
    if args.command == "verify" and (args.excel_evidence is None) != (args.workbook is None):
        parser.error("verify: --excel-evidence and --workbook go together")
    if args.command == "evidence":
        return evidence_main(root, args.commit, args.excel_evidence, args.workbook,
                             args.run_commit)
    if args.command == "verify-published":
        facts = verify_published(args.repository, args.tag, args.milestone, dict(args.asset),
                                 run_gh)
        for status, text in facts:
            print(f"{status}: {text}")
        failed = any(status == "FAIL" for status, _ in facts)
        print(f"release-verify-published: {'FAIL' if failed else 'PASS'} ({args.tag})")
        return 1 if failed else 0
    try:
        if args.command == "notes":
            sys.stdout.write(notes(root, args.tag))
            return 0
        if args.command == "bundle":
            for path in bundle(root, args.ref, args.name, args.out):
                print(f"wrote {path}")
            return 0
        given = ((args.excel_evidence, args.workbook)
                 if args.command == "verify" and args.excel_evidence is not None else None)
        problems = (verify(root, args.tag, args.default_branch, given)
                    if args.command == "verify" else integrity(args.repository, args.sha))
    except ReleaseError as error:
        problems = [str(error)]
    for problem in problems:
        print(f"FAIL: {problem}")
    # Each subcommand is named by its own argument: only integrity has --sha.
    label = (args.sha[:12] if args.command == "integrity"
             else args.ref if args.command == "bundle" else args.tag)
    print(f"release-{args.command}: {'FAIL' if problems else 'PASS'} ({label})")
    return 1 if problems else 0


def evidence_main(root: Path, commit: str, folder: Path, workbook: Path, run: str | None,
                  checker: Checker | None = None) -> int:
    """The evidence subcommand: the report, any failure, then the digest and the
    tag message line."""
    try:
        gate = excel_gate(root, commit, folder, workbook, run, checker)
    except ReleaseError as error:
        print(f"FAIL: {error}")
        print(f"release-evidence: FAIL ({commit})")
        return 2
    if gate.report:
        print(gate.report)
    for problem in gate.problems:
        print(f"FAIL: {problem}")
    if not gate.problems:
        kind = ("a run of the candidate" if gate.run == gate.candidate else
                f"a run of {gate.run}, equivalent to the candidate")
        print(f"Excel evidence digest: {gate.digest} ({kind})")
        print(f"Tag message line: {gate.line}")
    print(f"release-evidence: {'FAIL' if gate.problems else 'PASS'} ({gate.candidate[:12]})")
    return 1 if gate.problems else 0


if __name__ == "__main__":
    sys.exit(main())
