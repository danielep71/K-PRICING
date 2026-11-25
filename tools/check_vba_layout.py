#!/usr/bin/env python3
"""Check the documentation layout of exported VBA source.

docs/VBA_HOUSE_STYLE.md (*Layout*) defines a comment banner as an apostrophe
followed by 78 dashes, or 78 equals signs for a test module's component
header, and gives every component and every procedure a header with the same
nine fields, in this order:

    PURPOSE, WHY THIS EXISTS, INPUTS, RETURNS, BEHAVIOR, ERROR POLICY,
    DEPENDENCIES, NOTES, UPDATED

This check reads every tracked .bas, .cls and .frm file under the source
directories of tools/format_vba.py and reports:

- a banner line that is not exactly 79 characters long;
- a component without its header: the first comment block after the lines
  the VBE writes and the Option statements, which names the component as
  "' MODULE: <name>" or, in a class module, "' CLASS: <name>";
- a component or procedure header whose field labels are not the nine
  fields, each once and in order, or that holds another label;
- a field with no text under it (write "None" instead), and an UPDATED field
  whose text is not one YYYY-MM-DD date.

A field label is a comment line of one space and upper-case words, such as
"' ERROR POLICY"; its text follows on lines indented three spaces. A
procedure's header is the comment block that follows its declaration, after
the Attribute lines the VBE writes below it and at most one blank line. A
form's VBE header is exempt. The check reads comments only: it never
compiles or runs VBA.

It reads the baseline that tools/format_vba.py reads (tools/vba_baseline.py):
a project that imports existing modules lists, in
tools/vba_layout_baseline.json, the modules not yet fully in layout. A listed module is reported as baselined and does not fail;
a listed module that is in layout in both checks fails until its entry is
removed; a listed path that is not a tracked module fails. --baseline FILE
reads another file and --no-baseline ignores it. Exit status: 0 when
everything passes, 1 for a layout finding, 2 for an error: a missing or
malformed baseline, or an unreadable Git index.
"""

from __future__ import annotations

import argparse
import sys
from collections.abc import Sequence
from pathlib import Path

import format_vba
import vba_baseline
from vba_header_layout import check_text
from vba_lexer import Finding, report

ROOT = Path(__file__).resolve().parents[1]


def check(root: Path, baseline: Sequence[str],
          baseline_name: str) -> tuple[list[Finding], list[str], int, int]:
    """The findings, the baselined modules, the modules read and their procedures."""
    findings: list[Finding] = []
    baselined: list[str] = []
    listed = set(baseline)
    names = format_vba.tracked_modules(root)
    procedures = 0
    for name in names:
        if not (root / name).is_file():
            continue
        text = format_vba.read(root / name)
        found, count = check_text(name, text)
        procedures += count
        verdict = vba_baseline.judge(name, listed, not found,
                                     not found and format_vba.format_text(text) == text)
        if verdict == vba_baseline.FAIL:
            findings += found
        elif verdict == vba_baseline.STALE:
            findings.append(Finding(name, None, vba_baseline.stale_message(baseline_name)))
        elif verdict == vba_baseline.BASELINED:
            baselined.append(name)
    for name in vba_baseline.unknown_entries(baseline, names):
        findings.append(Finding(name, None, vba_baseline.unknown_message(baseline_name)))
    return findings, baselined, len(names), procedures


def main(argv: Sequence[str] | None = None) -> int:
    parser = argparse.ArgumentParser(description=(__doc__ or "").splitlines()[0])
    parser.add_argument("--root", type=Path, default=ROOT)
    source = parser.add_mutually_exclusive_group()
    source.add_argument("--baseline", type=Path,
                        help="baseline file, relative to the root (default: "
                             "tools/vba_layout_baseline.json, when it exists)")
    source.add_argument("--no-baseline", action="store_true",
                        help="ignore any baseline: every module must be in layout")
    args = parser.parse_args(argv)
    root = args.root.resolve()
    try:
        baseline, baseline_name = vba_baseline.read_baseline(root, args.baseline,
                                                             args.no_baseline)
        findings, baselined, components, procedures = check(root, baseline, baseline_name)
    except (vba_baseline.BaselineError, RuntimeError) as error:
        print(f"vba-layout: ERROR: {error}", file=sys.stderr)
        return 2
    for name in baselined:
        print(f"baselined: {name}: {vba_baseline.baselined_message(baseline_name)}")
    return report("vba-layout", findings,
                  f"{components} VBA component(s), {len(baselined)} baselined, "
                  f"{procedures} procedure(s)")


if __name__ == "__main__":
    sys.exit(main())
