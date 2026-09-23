#!/usr/bin/env python3
"""Create a disposable generated consumer for hosted reusable-workflow tests.

The destination must not exist. No remote repository or branch is changed.
"""

from __future__ import annotations

import argparse
import re
from pathlib import Path

import initialize_repository as initializer


def create_fixture(source: Path, destination: Path, profile: str, workflow_sha: str) -> None:
    if not re.fullmatch(r"[0-9a-f]{40}", workflow_sha):
        raise ValueError("workflow-sha must be a full lowercase commit SHA")
    if destination.exists():
        raise ValueError("destination must not exist")
    initializer._copy_fixture(source, destination)
    scalars, repeatable = initializer._fixture_arguments(profile)
    changes, _ = initializer._build_changes(destination, profile, scalars, repeatable)
    initializer._apply_changes(destination, changes)
    caller = f"""name: Reusable workflow consumer fixture
on:
  push:
  workflow_dispatch:
permissions:
  contents: read
jobs:
  generic:
    # Workflow interface version, independent of the product release.
    uses: danielep71/EXCEL-VBA-PROJECT-TEMPLATE/.github/workflows/static-checks.yml@{workflow_sha} # v1.0.0
    with:
      expected-profile: {profile}
  specialist:
    needs: generic
    runs-on: ubuntu-24.04
    steps:
      - uses: actions/checkout@3d3c42e5aac5ba805825da76410c181273ba90b1 # v7.0.1
        with:
          persist-credentials: false
      - name: Verify exact candidate output and preserve a stricter local requirement
        env:
          CHECKED_SHA: ${{{{ needs.generic.outputs.candidate-sha }}}}
          CANDIDATE_SHA: ${{{{ github.sha }}}}
        run: |
          test "$CHECKED_SHA" = "$CANDIDATE_SHA"
          test -f tests/modules/ProjectTests.bas
"""
    (destination / ".github/workflows/static-checks.yml").write_text(caller, encoding="utf-8")
    initializer._git(destination, "add", "--all")
    initializer._git(destination, "commit", "-m", f"Create {profile} reusable workflow consumer")


def main() -> int:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--root", type=Path, default=Path("."))
    parser.add_argument("--destination", type=Path, required=True)
    parser.add_argument("--profile", choices=initializer.SUPPORTED_PROFILES, required=True)
    parser.add_argument("--workflow-sha", required=True)
    options = parser.parse_args()
    create_fixture(options.root.resolve(), options.destination.resolve(),
                   options.profile, options.workflow_sha)
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
