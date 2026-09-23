#!/usr/bin/env bash
# Run the locally reproducible part of hosted CI, in the same order, with the
# same pinned tools.
#
# Purpose: let a maintainer find a failure before pushing instead of after a CI
# round trip. Every check below is exactly the command the hosted workflows run.
#
# This is NOT a claim of CI equivalence. It deliberately cannot cover:
#   - actionlint workflow validation, unless actionlint is already on PATH
#     (hosted CI installs a SHA-256-verified 1.7.12 binary);
#   - anything requiring Excel, a Windows host, or live GitHub state;
#   - hosted evidence upload, job summaries and the terminal CI verdict.
# Those remain the hosted gates' responsibility and are reported as SKIPPED
# rather than counted as passes.
#
# Usage:
#   python3 -m pip install -r tools/requirements-dev.txt
#   tools/dev_check.sh              # full local set
#   FAST=1 tools/dev_check.sh       # skip the slow initializer and coverage gates

set -uo pipefail

cd "$(dirname "$0")/.." || exit 2
FAST="${FAST:-0}"
PASS=0
FAIL=0
SKIP=0
FAILED_NAMES=()

bold() { printf '\033[1m%s\033[0m\n' "$1"; }

record() { # record <outcome> <name>
  case "$1" in
    pass) PASS=$((PASS + 1)); printf '  [PASS] %s\n' "$2" ;;
    fail) FAIL=$((FAIL + 1)); FAILED_NAMES+=("$2"); printf '  [FAIL] %s\n' "$2" ;;
    skip) SKIP=$((SKIP + 1)); printf '  [SKIP] %s\n' "$2" ;;
  esac
}

run() { # run <name> <command...>
  local name="$1"; shift
  local output
  if output=$("$@" 2>&1); then
    record pass "$name"
  else
    record fail "$name"
    printf '%s\n' "$output" | tail -20 | sed 's/^/         /'
  fi
}

pin() { grep -E "^$1==" tools/requirements-dev.txt | head -1 | sed 's/.*==//'; }

# --- CI pin-mirror guard ------------------------------------------------------
# requirements-dev.txt only means anything while it still mirrors the workflows.
# If a workflow pin moves and this file does not, a clean local run becomes a
# false negative, so the mirror is verified rather than trusted.
bold "CI pin mirror"
STATIC_CHECKS=.github/workflows/static-checks.yml
CHECKER_DEV=.github/workflows/checker-development.yml

workflow_env() { # workflow_env <name> <workflow>
  grep -E "^[[:space:]]+$1:[[:space:]]" "$2" | head -1 |
    sed -E 's/^[^:]*:[[:space:]]*"?([^"]*)"?[[:space:]]*$/\1/'
}
lock_pin() { # lock_pin <requirement name> <lock file>
  grep -E "^$1==" "$2" | head -1 | sed -E 's/.*==([^[:space:]\\]+).*/\1/'
}
comment_value() { # comment_value <key>
  grep -E "^#[[:space:]]+$1=" tools/requirements-dev.txt | head -1 | sed "s/.*$1=//"
}

mirror() { # mirror <label> <declared here> <declared in CI> <workflow>
  if [[ -z "$2" || -z "$3" ]]; then
    record fail "$1 pin unreadable (requirements-dev.txt='$2', $4='$3')"
  elif [[ "$2" == "$3" ]]; then
    record pass "$1 $2 mirrors $4"
  else
    record fail "$1 pinned $2 here but $4 installs $3"
  fi
}

QUALITY_LOCK=tools/requirements-quality-ci.txt
COVERAGE_LOCK=tools/requirements-coverage-ci.txt
PORTFOLIO_LOCK=tools/requirements-portfolio-ci.txt

mirror ruff     "$(pin ruff)" "$(workflow_env RUFF_VERSION "$STATIC_CHECKS")" "$STATIC_CHECKS"
mirror mypy     "$(pin mypy)" "$(workflow_env MYPY_VERSION "$STATIC_CHECKS")" "$STATIC_CHECKS"
mirror "ruff CI lock" "$(lock_pin ruff "$QUALITY_LOCK")" \
  "$(workflow_env RUFF_VERSION "$STATIC_CHECKS")" "$STATIC_CHECKS"
mirror "mypy CI lock" "$(lock_pin mypy "$QUALITY_LOCK")" \
  "$(workflow_env MYPY_VERSION "$STATIC_CHECKS")" "$STATIC_CHECKS"
mirror coverage "$(pin 'coverage\[toml\]')" \
  "$(lock_pin 'coverage\[toml\]' "$COVERAGE_LOCK")" "$COVERAGE_LOCK"
mirror PyYAML "$(pin PyYAML)" "$(lock_pin PyYAML "$COVERAGE_LOCK")" "$COVERAGE_LOCK"
mirror "PyYAML portfolio lock" "$(pin PyYAML)" \
  "$(lock_pin PyYAML "$PORTFOLIO_LOCK")" "$PORTFOLIO_LOCK"
# The actionlint install recipe in requirements-dev.txt carries a version and a
# digest; a stale digest there would send a maintainer to the wrong binary.
mirror "actionlint" "$(comment_value v)" \
  "$(workflow_env ACTIONLINT_VERSION "$STATIC_CHECKS")" "$STATIC_CHECKS"
mirror "actionlint digest" "$(comment_value sha)" \
  "$(workflow_env ACTIONLINT_LINUX_AMD64_SHA256 "$STATIC_CHECKS")" "$STATIC_CHECKS"

# --- Pinned-version guard -----------------------------------------------------
# A local run with different tool versions can disagree with CI in both
# directions, so a mismatch is a failure rather than a warning.
bold "Pinned tool versions"

check_version() { # check_version <label> <expected> <actual>
  if [[ -z "$3" ]]; then
    record fail "$1 is not installed (expected $2; see tools/requirements-dev.txt)"
  elif [[ "$3" == "$2" ]]; then
    record pass "$1 $3"
  else
    record fail "$1 $3 does not match the pinned $2"
  fi
}

# Each tool prints its version differently, and `ruff --version` in particular
# reports whichever binary PATH resolves first, which may not be the pinned one.
check_version ruff "$(pin ruff)" \
  "$(ruff --version 2>/dev/null | head -1 | awk '{print $2}')"
check_version mypy "$(pin mypy)" \
  "$(python3 -m mypy --version 2>/dev/null | head -1 | awk '{print $2}')"
check_version coverage "$(pin 'coverage\[toml\]')" \
  "$(python3 -m coverage --version 2>/dev/null | head -1 | awk '{print $3}')"
# check_portfolio_drift.py refuses to run on any other PyYAML, so a mismatch here
# fails the portfolio suites with a message that looks unrelated.
check_version PyYAML "$(pin PyYAML)" \
  "$(python3 -c 'import yaml; print(yaml.__version__)' 2>/dev/null)"

if command -v actionlint >/dev/null 2>&1; then
  record pass "actionlint $(actionlint --version 2>/dev/null | head -1) present"
  HAVE_ACTIONLINT=1
else
  record skip "actionlint absent; hosted CI pins and verifies 1.7.12"
  HAVE_ACTIONLINT=0
fi

# --- Python quality baseline --------------------------------------------------
bold "Python quality baseline"
run "Ruff lint" ruff check tools
run "mypy type check" python3 -m mypy

# --- Authoritative workflow validation ---------------------------------------
bold "Workflow validation"
if [[ "$HAVE_ACTIONLINT" == "1" ]]; then
  run "Authoritative workflow fixtures" \
    python3 tools/test_workflow_validation.py --actionlint "$(command -v actionlint)"
else
  record skip "Authoritative workflow fixtures (needs actionlint)"
fi

# --- Release, documentation and focused gates ---------------------------------
bold "Release and documentation contracts"
run "Release-integrity self-test"  python3 tools/check_release.py --root . --self-test
run "Release provenance fixtures"  python3 tools/test_release_provenance.py
run "Excel evidence fixtures"      python3 tools/test_excel_evidence.py
run "Documentation fixtures"       python3 tools/test_documentation.py
run "Documentation drift"          python3 tools/check_documentation.py --root .
run "Release-semantics self-test"  python3 tools/check_release_semantics.py --root . --self-test
run "Release semantics"            python3 tools/check_release_semantics.py --root .

bold "Focused hardening gates"
run "Committed-whitespace self-test" python3 tools/check_committed_whitespace.py --root . --self-test
run "Working-tree whitespace"        python3 tools/check_committed_whitespace.py --root . --mode working-tree
run "VBA jump self-test"             python3 tools/check_vba_jumps.py --root . --self-test
run "VBA jumps"                      python3 tools/check_vba_jumps.py --root .
run "VBA conditional self-test"      python3 tools/check_vba_conditionals.py --root . --self-test
run "VBA conditional compilation"    python3 tools/check_vba_conditionals.py --root .
run "VBA public API self-test"       python3 tools/check_vba_public_api.py --root . --self-test
run "VBA public API"                 python3 tools/check_vba_public_api.py --root .
run "Local-Action self-test"         python3 tools/check_local_actions.py --root . --self-test
run "Repository-local Actions"       python3 tools/check_local_actions.py --root .
run "Template-contract self-test"    python3 tools/check_template_contract.py --root . --self-test
run "Template contract"              python3 tools/check_template_contract.py --root .

bold "Canonical repository gate"
run "Checker self-test"    python3 tools/check_repo.py --root . --self-test
run "Repository integrity" python3 tools/check_repo.py --root .

# --- Template-maintainer contracts -------------------------------------------
# Removed from generated projects by initialization; only meaningful here.
bold "Template-maintainer contracts"
run "Checker-development contract" python3 tools/checker_development.py --root . --self-test
run "Wiki contract"                python3 tools/check_wiki.py --root .
run "Wiki fixtures"                python3 tools/test_wiki.py
run "Portfolio drift fixtures"     python3 tools/test_portfolio_drift.py
run "Portfolio quality fixtures"   python3 tools/test_portfolio_quality.py
run "Provisioning fixtures"        python3 tools/test_provision_repository.py
run "Verification-depth fixtures"  python3 tools/test_verification_depth.py

if [[ "$FAST" == "1" ]]; then
  record skip "Initializer self-test (FAST=1)"
  record skip "Semantic policy coverage (FAST=1)"
else
  run "Initializer self-test"   python3 tools/initialize_repository.py --root . --self-test
  run "Semantic policy coverage" python3 tools/check_policy_coverage.py --root . --self-test
fi

# --- Verdict ------------------------------------------------------------------
echo
bold "Local verdict"
printf '  passed %d   failed %d   skipped %d\n' "$PASS" "$FAIL" "$SKIP"
if ((FAIL > 0)); then
  echo
  echo "  Failing:"
  printf '    - %s\n' "${FAILED_NAMES[@]}"
  echo
  echo "  Hosted CI will fail too. Fix these before pushing."
  exit 1
fi
echo
echo "  Local checks clean. Skipped checks are still hosted-CI responsibilities:"
echo "  actionlint (unless installed), statement-coverage floor, Excel evidence,"
echo "  live GitHub state, and the terminal CI verdict."
exit 0
