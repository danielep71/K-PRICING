#!/usr/bin/env python3
"""Preview, check or apply the repository settings baseline through an authenticated GitHub CLI (github.com).

The baseline is docs/setup/repository-settings.json: an optional visibility,
the general repository settings, the security features, the GitHub Actions
settings (reported, never written) and the repository rulesets. Preview (the
default) and --check only read and print the plan with its SHA-256; --apply
writes that approved plan, journaling before the first write and after each,
then reads everything back. --check --token-scope workflow compares only what a
workflow GITHUB_TOKEN with contents: read can read, lists the rest as not
checked and plans nothing.
"""
# Postponed annotations let the version check in main() run on older Python versions.
from __future__ import annotations

import argparse
import hashlib
import json
import os
import re
import subprocess
import sys
import tempfile
from collections.abc import Callable
from datetime import datetime, timezone
from pathlib import Path
from typing import TYPE_CHECKING, Any
from urllib.parse import quote

ROOT = Path(__file__).resolve().parents[1]
CONFIG = ROOT / 'docs/setup/repository-settings.json'
SCHEMA_VERSION = 4
EVIDENCE_SCHEMA_VERSION = 1
TOP_KEYS = ('schema_version', 'visibility', 'repository', 'security', 'actions', 'rulesets')
RULE_KEYS = ('name', 'target', 'enforcement', 'conditions', 'rules', 'bypass_actors')
# The merge methods are written before the rulesets, which may allow only enabled ones.
MERGE_KEYS = ('allow_merge_commit', 'allow_squash_merge', 'allow_rebase_merge')
GENERAL_KEYS = MERGE_KEYS + (
    'allow_auto_merge', 'allow_update_branch', 'delete_branch_on_merge',
    'web_commit_signoff_required', 'has_issues', 'has_projects', 'has_wiki',
    'has_discussions', 'has_downloads',
)
MERGE_METHODS = {'merge': 'allow_merge_commit', 'squash': 'allow_squash_merge',
                 'rebase': 'allow_rebase_merge'}
# Security features in write order (alerts before the security updates that need
# them), and whether GitHub offers them only on public repositories (private
# repositories need a paid licence for secret scanning).
SECURITY_PUBLIC_ONLY = {
    'private_vulnerability_reporting': True,
    'vulnerability_alerts': False,
    'automated_security_fixes': False,
    'secret_scanning': True,
    'secret_scanning_push_protection': True,
}
ANALYSIS_KEYS = ('secret_scanning', 'secret_scanning_push_protection')
# Each Actions setting: the endpoint under actions/permissions that reports it,
# its field there, and whether it applies only to public repositories.
ACTIONS_SOURCES = {
    'sha_pinning_required': ('', 'sha_pinning_required', False),
    'default_workflow_permissions': ('/workflow', 'default_workflow_permissions', False),
    'can_approve_pull_request_reviews': ('/workflow', 'can_approve_pull_request_reviews', False),
    'fork_pr_approval_policy': ('/fork-pr-contributor-approval', 'approval_policy', True),
    'retention_days': ('/artifact-and-log-retention', 'days', False),
}
FORK_POLICIES = ('first_time_contributors_new_to_github', 'first_time_contributors',
                 'all_external_contributors')
REPOSITORY_PATTERN = re.compile(r'[A-Za-z0-9][A-Za-z0-9-]*/[A-Za-z0-9_.-]+')
DIGEST_PATTERN = re.compile(r'[0-9a-fA-F]{64}')
# Ruleset statuses that need no write and are not drift.
SATISFIED = ('matches', 'covered')
NOT_APPLIED = 'not applied (public repositories only)'
DISTRIBUTION = 'docs/repo_governance/DISTRIBUTION.md'
# What the token reads: everything that an administrator's token can ('full'), or
# only what a workflow GITHUB_TOKEN with contents: read can ('workflow').
TOKEN_SCOPES = ('full', 'workflow')
# Endpoints that the workflow scope never reads: they need Administration: read,
# which no workflow permission grants.
WORKFLOW_UNREAD = 'not read: needs Administration: read, which a workflow token cannot have'
# The HTTP statuses with which GitHub refuses a token that may not read an
# endpoint. Any other failure is a failed request.
ACCESS_LIMITS = ('401', '403', '404')

# GitHub API values are JSON; a call takes a method, an endpoint and an optional payload.
# The aliases appear only in annotations, which are never evaluated, so Python 3.8
# imports the module and reaches the version check.
Json = Any
if TYPE_CHECKING:
    Call = Callable[..., Json]
    Plan = list[tuple[str, str, Json]]
    Row = dict[str, Json]


class SetupError(Exception):
    """A refusal or failure; the command exits with status 2."""


class ApiError(SetupError):
    """A GitHub API call that failed, with the HTTP status gh reported, if any."""

    def __init__(self, method: str, endpoint: str, status: str) -> None:
        super().__init__(f'{method} {endpoint} failed (HTTP {status}). Check gh authentication, '
                         'repository access, Administration permissions and plan availability. '
                         'No automatic rollback is performed; rerun preview to inspect partial changes.')
        self.status = status


def api(method: str, endpoint: str, payload: Json = None) -> Json:
    command = ['gh', 'api', '--hostname', 'github.com', '--method', method,
               '-H', 'Accept: application/vnd.github+json',
               '-H', 'X-GitHub-Api-Version: 2022-11-28', endpoint]
    if payload is not None:
        command += ['--input', '-']
    # Never inherit HTTP debug logging: it can expose authentication headers.
    env = dict(os.environ)
    env.pop('GH_DEBUG', None)
    # gh writes UTF-8; never decode with the locale code page (cp1252 on Windows).
    result = subprocess.run(command, input=json.dumps(payload) if payload is not None else None,
                            text=True, encoding='utf-8', capture_output=True, env=env, check=False)
    if result.returncode:
        # Do not print CLI stderr (may contain sensitive server or credential detail).
        match = re.search(r'HTTP (\d{3})', result.stderr)
        raise ApiError(method, endpoint, match.group(1) if match else 'unknown')
    return json.loads(result.stdout) if result.stdout.strip() else None


def read(call: Call, endpoint: str, tolerated: tuple[str, ...] | None = None) -> tuple[Json, str | None]:
    """GET an endpoint: its value, or None and the HTTP status of the failure.

    With tolerated, only a failure with one of those statuses is returned; any
    other, such as a server error or no answer, is raised.
    """
    try:
        return call('GET', endpoint), None
    except ApiError as error:
        if tolerated is not None and error.status not in tolerated:
            raise
        return None, error.status


def pages(call: Call, endpoint: str, key: str | None = None) -> list[Json]:
    result: list[Json] = []
    page = 1
    while True:
        separator = '&' if '?' in endpoint else '?'
        value = call('GET', f'{endpoint}{separator}per_page=100&page={page}')
        batch = value[key] if key else value
        result.extend(batch)
        if len(batch) < 100:
            return result
        page += 1


def canonical(value: Json) -> Json:
    """Rule arrays are sets for this baseline; ignore API ordering, not values."""
    if isinstance(value, dict):
        return {k: canonical(v) for k, v in sorted(value.items())}
    if isinstance(value, list):
        return sorted((canonical(x) for x in value), key=lambda x: json.dumps(x, sort_keys=True))
    return value


def rule_spec(rule: dict[str, Json]) -> Json:
    # Ignore response metadata, but preserve unknown rules/parameters in comparison.
    return canonical({k: rule.get(k, [] if k == 'bypass_actors' else None) for k in RULE_KEYS})


# Configuration


def require_booleans(value: Json, keys: tuple[str, ...], what: str) -> None:
    if not isinstance(value, dict) or set(value) != set(keys) or any(type(v) is not bool for v in value.values()):
        raise SetupError(f'Configuration must contain exactly the {len(keys)} documented boolean '
                         f'{what}: {", ".join(keys)}.')


def check_actions(actions: Json) -> None:
    if not isinstance(actions, dict) or set(actions) != set(ACTIONS_SOURCES):
        raise SetupError('Configuration must contain exactly the documented Actions settings: '
                         + ', '.join(ACTIONS_SOURCES) + '.')
    days = actions['retention_days']
    if (type(actions['sha_pinning_required']) is not bool
            or actions['default_workflow_permissions'] not in ('read', 'write')
            or type(actions['can_approve_pull_request_reviews']) is not bool
            or actions['fork_pr_approval_policy'] not in FORK_POLICIES
            or type(days) is not int or days < 1):
        raise SetupError('Invalid Actions settings: sha_pinning_required and '
                         'can_approve_pull_request_reviews are booleans, default_workflow_permissions '
                         'is read or write, fork_pr_approval_policy is one of '
                         + ', '.join(FORK_POLICIES) + ', and retention_days is a positive integer.')


def check_rule(name: str, rule: Json, settings: dict[str, bool]) -> None:
    """Refuse a rule that the configured repository settings would make GitHub reject."""
    parameters = rule.get('parameters') or {}
    for method in parameters.get('allowed_merge_methods', []):
        setting = MERGE_METHODS.get(method)
        if setting is None:
            raise SetupError(f'Ruleset "{name}" allows an unknown merge method: {method}.')
        if not settings[setting]:
            raise SetupError(f'Ruleset "{name}" allows {method} merges, but the configuration '
                             f'sets {setting} to false.')
    for check in parameters.get('required_status_checks', []):
        if (not isinstance(check, dict) or not isinstance(check.get('context'), str)
                or type(check.get('integration_id')) is not int):
            raise SetupError(f'Ruleset "{name}" has a required status check without a string '
                             'context and an integer integration_id.')


def check_rulesets(rulesets: Json, settings: dict[str, bool]) -> None:
    if not isinstance(rulesets, list) or not rulesets:
        raise SetupError('Configuration must contain at least one ruleset.')
    names: list[str] = []
    for ruleset in rulesets:
        if not isinstance(ruleset, dict) or set(ruleset) != set(RULE_KEYS):
            raise SetupError('Each configured ruleset must have exactly: ' + ', '.join(RULE_KEYS) + '.')
        name = ruleset['name']
        if ruleset['target'] not in ('branch', 'tag') or ruleset['enforcement'] != 'active':
            raise SetupError(f'Ruleset "{name}" must be an active branch or tag ruleset.')
        if ruleset['bypass_actors'] or ruleset['conditions']['ref_name']['exclude']:
            raise SetupError(f'Ruleset "{name}" must not contain bypass actors or excluded refs.')
        types = [rule['type'] for rule in ruleset['rules']]
        if len(types) != len(set(types)):
            raise SetupError(f'Ruleset "{name}" repeats a rule type.')
        for rule in ruleset['rules']:
            check_rule(name, rule, settings)
        names.append(name)
    duplicates = sorted({name for name in names if names.count(name) > 1})
    if duplicates:
        raise SetupError('Configured ruleset names must be unique: ' + ', '.join(duplicates) + '.')


def load_config(path: Path) -> dict[str, Json]:
    value = json.loads(path.read_text(encoding='utf-8'))
    if not isinstance(value, dict) or set(value) != set(TOP_KEYS) or value['schema_version'] != SCHEMA_VERSION:
        raise SetupError('Unsupported configuration schema.')
    if value['visibility'] not in (None, 'public', 'private'):
        raise SetupError('The configured visibility must be null, "public" or "private".')
    require_booleans(value['repository'], GENERAL_KEYS, 'repository settings')
    require_booleans(value['security'], tuple(SECURITY_PUBLIC_ONLY), 'security features')
    security = value['security']
    if security['secret_scanning_push_protection'] and not security['secret_scanning']:
        raise SetupError('Push protection requires secret scanning.')
    if security['automated_security_fixes'] and not security['vulnerability_alerts']:
        raise SetupError('Dependabot security updates require vulnerability alerts.')
    check_actions(value['actions'])
    check_rulesets(value['rulesets'], value['repository'])
    return dict(value)


# Observation: read-only, and tolerant of what a token without administrator access cannot read.


def dependabot_states(call: Call, root: str) -> tuple[Row, Row]:
    """Dependabot alerts and security updates as observed: True, False or None when GitHub does not say."""
    value, failure = read(call, root + '/automated-security-fixes')
    if failure is None:
        value = value or {}
        paused = bool(value.get('enabled')) and bool(value.get('paused'))
        fixes: Row = {'observed': bool(value.get('enabled')), 'paused': paused,
                      'source': 'GET automated-security-fixes enabled' + (', paused' if paused else '')}
    elif failure == '404':
        fixes = {'observed': False, 'paused': False, 'source': 'GET automated-security-fixes (HTTP 404: off)'}
    else:
        fixes = {'observed': None, 'paused': False,
                 'source': f'GET automated-security-fixes (HTTP {failure}: unreadable)'}

    # The endpoint answers without content when the alerts are on and 404 when
    # they are off. Enabled security updates prove them on: GitHub requires them.
    _, failure = read(call, root + '/vulnerability-alerts')
    if failure is None:
        alerts: Row = {'observed': True, 'source': 'GET vulnerability-alerts (on)'}
    elif failure == '404':
        alerts = {'observed': False, 'source': 'GET vulnerability-alerts (HTTP 404: off)'}
    elif fixes['observed'] is True:
        alerts = {'observed': True, 'source': f'GET vulnerability-alerts (HTTP {failure}: unreadable); on, '
                                              'because Dependabot security updates are enabled and require them'}
    else:
        alerts = {'observed': None, 'source': f'GET vulnerability-alerts (HTTP {failure}: unreadable)'}
    return alerts, fixes


def security_states(call: Call, root: str, metadata: dict[str, Json], public: bool,
                    scope: str = 'full') -> dict[str, Row]:
    """Each security feature as observed: True, False or None when GitHub does not say."""
    states: dict[str, Row] = {}
    if public:
        # The workflow scope reports only a refused read as not checked; a failed
        # request is an API failure, never a value it may skip.
        value, failure = read(call, root + '/private-vulnerability-reporting',
                              ACCESS_LIMITS if scope == 'workflow' else None)
        states['private_vulnerability_reporting'] = (
            {'observed': bool((value or {}).get('enabled')), 'source': 'GET private-vulnerability-reporting enabled'}
            if failure is None else
            {'observed': None, 'source': f'GET private-vulnerability-reporting (HTTP {failure}: unreadable)'})
    if scope == 'workflow':
        # Never read: what GitHub answers a token without the permission could look like "off".
        states['vulnerability_alerts'] = {'observed': None,
                                          'source': f'GET vulnerability-alerts ({WORKFLOW_UNREAD})'}
        states['automated_security_fixes'] = {'observed': None, 'paused': False,
                                              'source': f'GET automated-security-fixes ({WORKFLOW_UNREAD})'}
    else:
        states['vulnerability_alerts'], states['automated_security_fixes'] = dependabot_states(call, root)

    # GitHub reports this block only to administrators of the repository.
    analysis = metadata.get('security_and_analysis')
    for key in ANALYSIS_KEYS if public else ():
        if not isinstance(analysis, dict):
            states[key] = {'observed': None,
                           'source': f'security_and_analysis.{key} (not reported without administrator access)'}
        else:
            states[key] = {'observed': (analysis.get(key) or {}).get('status') == 'enabled',
                           'source': f'security_and_analysis.{key}'}
    for feature, only in SECURITY_PUBLIC_ONLY.items():
        if only and not public:
            states[feature] = {'observed': None, 'source': NOT_APPLIED, 'applies': False}
    return {feature: states[feature] for feature in SECURITY_PUBLIC_ONLY}


def actions_states(call: Call, root: str, public: bool, scope: str = 'full') -> dict[str, Row]:
    """Each Actions setting as observed, or None with the reason it is not known."""
    answers: dict[str, tuple[Json, str | None]] = {}
    states: dict[str, Row] = {}
    for setting, (suffix, field, public_only) in ACTIONS_SOURCES.items():
        if public_only and not public:
            states[setting] = {'observed': None, 'source': NOT_APPLIED, 'applies': False}
            continue
        endpoint = 'actions/permissions' + suffix
        if scope == 'workflow':
            states[setting] = {'observed': None, 'source': f'GET {endpoint} ({WORKFLOW_UNREAD})'}
            continue
        if endpoint not in answers:
            answers[endpoint] = read(call, f'{root}/{endpoint}')
        value, failure = answers[endpoint]
        if failure is not None:
            states[setting] = {'observed': None, 'source': f'GET {endpoint} (HTTP {failure}: unreadable)'}
        elif not isinstance(value, dict) or field not in value:
            states[setting] = {'observed': None, 'source': f'GET {endpoint} (no {field} field)'}
        else:
            states[setting] = {'observed': value[field], 'source': f'GET {endpoint} {field}'}
    return states


def observe(call: Call, repo: str, config: dict[str, Json], scope: str = 'full') -> dict[str, Json]:
    """Read every value the configuration manages, within the token scope. Read-only."""
    root = f'repos/{repo}'
    metadata = call('GET', root)
    public = metadata.get('visibility') == 'public'
    listed = pages(call, root + '/rulesets?includes_parents=true')
    own = [r for r in listed if r.get('source_type', 'Repository') == 'Repository']
    managed = {rule['name'] for rule in config['rulesets']}
    names = [r['name'] for r in own if r['name'] in managed]
    duplicates = sorted({name for name in names if names.count(name) > 1})
    if duplicates:
        raise SetupError('Ambiguous duplicate ruleset name: ' + ', '.join(duplicates))
    return {
        'metadata': metadata,
        'public': public,
        'rulesets': [call('GET', f'{root}/rulesets/{r["id"]}') for r in own],
        'inherited': [r['name'] for r in listed if r.get('source_type', 'Repository') != 'Repository'],
        'security': security_states(call, root, metadata, public, scope),
        'actions': actions_states(call, root, public, scope),
    }


# Comparison


_ABSENT = '<absent>'


def keyed(spec: Json) -> Json:
    """A ruleset spec with its rules keyed by type, for readable difference paths."""
    result = dict(spec)
    rules = spec.get('rules') or []
    types = [rule.get('type') for rule in rules]
    if len(types) == len(set(types)):
        result['rules'] = {rule.get('type'): rule.get('parameters', {}) for rule in rules}
    return result


def differences(observed: Json, configured: Json, path: str = '') -> list[str]:
    """Each path where an observed value departs from the configured one."""
    if isinstance(observed, dict) and isinstance(configured, dict):
        found: list[str] = []
        for key in sorted(set(observed) | set(configured), key=str):
            found += differences(observed.get(key, _ABSENT), configured.get(key, _ABSENT),
                                 f'{path}.{key}' if path else str(key))
        return found
    if canonical(observed) == canonical(configured):
        return []
    return [f'{path}: observed {json.dumps(canonical(observed))}, '
            f'configured {json.dumps(canonical(configured))}']


def compare(key: str, name: str, configured: Json, state: Row) -> Row:
    """A row of the record; matches is None for a value that does not apply here."""
    row: Row = {key: name, 'configured': configured, 'observed': state['observed'], 'source': state['source']}
    if 'paused' in state:
        row['paused'] = state['paused']
    if not state.get('applies', True):
        row['matches'] = None
    elif row.get('paused') and configured is True:
        # Paused updates open no pull requests: drift when they should be on, but still
        # enabled, so turning the feature off needs the DELETE.
        row['matches'] = False
    else:
        row['matches'] = state['observed'] is not None and state['observed'] == configured
    return row


# Coverage: live rulesets that enforce a configured one at least as strictly,
# under any names. Only an active ruleset of the same target with a reported and
# empty bypass list and no excluded refs can count: GitHub applies every active
# ruleset, so its rules hold whatever the configured name's ruleset does.


def parameters(rule: Json) -> dict[str, Json]:
    return dict(rule.get('parameters') or {})


def include_covers(candidate: Json, desired: dict[str, Json], default_branch: str) -> bool:
    """Whether a live ruleset's ref conditions reach every ref the configured one names."""
    conditions = (candidate.get('conditions') or {}).get('ref_name') or {}
    if conditions.get('exclude'):
        return False
    include = set(conditions.get('include') or [])
    if '~ALL' in include:
        return True
    if desired['target'] == 'branch' and '~DEFAULT_BRANCH' in include:
        include.add(f'refs/heads/{default_branch}')
    return set(desired['conditions']['ref_name']['include']) <= include


def candidates(desired: dict[str, Json], rulesets: list[Json], default_branch: str) -> list[Json]:
    """The live rulesets whose scope can stand in for the configured one, the named one first."""
    found = [r for r in rulesets
             if r.get('enforcement') == 'active' and r.get('target') == desired['target']
             and r.get('bypass_actors') == [] and include_covers(r, desired, default_branch)]
    return sorted(found, key=lambda r: r.get('name') != desired['name'])


def covers_pull_request(configured: dict[str, Json], actual: dict[str, Json]) -> bool:
    """More approvals, a true option, fewer merge methods or more reviewers are stricter."""
    for key, value in configured.items():
        if key == 'allowed_merge_methods':
            if not set(actual.get(key, tuple(MERGE_METHODS))) <= set(value):
                return False
        elif key == 'required_approving_review_count':
            if actual.get(key, 0) < value:
                return False
        elif type(value) is bool:
            if actual.get(key, False) not in (True, value):
                return False
        elif isinstance(value, list):
            dumped = {json.dumps(canonical(item), sort_keys=True) for item in actual.get(key) or []}
            if not {json.dumps(canonical(item), sort_keys=True) for item in value} <= dumped:
                return False
        elif actual.get(key) != value:
            return False
    return True


def covers_status_checks(configured: dict[str, Json], actual: dict[str, Json]) -> bool:
    """A strict policy, more checks or enforcement on creation are stricter."""
    def checks(found: dict[str, Json]) -> set[tuple[Json, Json]]:
        return {(c.get('context'), c.get('integration_id')) for c in found.get('required_status_checks') or []}
    strict, on_create = 'strict_required_status_checks_policy', 'do_not_enforce_on_create'
    return (actual.get(strict, False) in (True, configured.get(strict, False))
            and checks(configured) <= checks(actual)
            and actual.get(on_create, False) in (False, configured.get(on_create, False)))


def covers_rule(rule: Json, candidate: Json) -> bool:
    """Whether a live rule of the configured rule's type is at least as strict as it."""
    configured, actual = parameters(rule), parameters(candidate)
    kind = rule['type']
    if kind in ('deletion', 'non_fast_forward'):
        return True
    if kind == 'update':
        allows = 'update_allows_fetch_and_merge'
        return actual.get(allows, False) in (False, configured.get(allows, False))
    if kind == 'required_status_checks':
        return covers_status_checks(configured, actual)
    if kind == 'pull_request':
        return covers_pull_request(configured, actual)
    return bool(canonical(actual) == canonical(configured))


def covered_by(desired: dict[str, Json], rulesets: list[Json], default_branch: str) -> list[str]:
    """The names of the live rulesets that together enforce every configured rule, or none."""
    possible = candidates(desired, rulesets, default_branch)
    names: list[str] = []
    for rule in desired['rules']:
        covering = next((c for c in possible if any(r.get('type') == rule['type'] and covers_rule(rule, r)
                                                    for r in c.get('rules') or [])), None)
        if covering is None:
            return []
        names.append(covering['name'])
    return sorted(set(names))


def assess_ruleset(desired: dict[str, Json], state: dict[str, Json]) -> Row:
    """The status of one configured ruleset: matches, covered, differs, missing or unknown."""
    actual = next((r for r in state['rulesets'] if r.get('name') == desired['name']), None)
    row: Row = {'name': desired['name'], 'id': None if actual is None else actual.get('id'),
                'status': 'missing', 'differences': [], 'covered_by': []}
    if actual is not None and 'bypass_actors' not in actual:
        # GitHub omits the bypass list unless the token can write the ruleset; an
        # omitted list is not an empty one, so the ruleset cannot be verified.
        visible = dict(actual, bypass_actors=desired['bypass_actors'])
        found = differences(keyed(rule_spec(visible)), keyed(rule_spec(desired)))
        found.append('bypass_actors: not reported (needs write access to the ruleset)')
        row.update(status='unknown', differences=found)
        return row
    if actual is not None:
        found = differences(keyed(rule_spec(actual)), keyed(rule_spec(desired)))
        if rule_spec(actual) != rule_spec(desired) and not found:
            found = ['rules: observed and configured differ']
        row.update(status='differs' if found else 'matches', differences=found)
    if row['status'] == 'differs' and actual is not None:
        # The named ruleset counts for itself: stronger parameters cover the configured ones.
        # A weaker one is replaced, whatever other rulesets enforce; a missing one is judged
        # in assess() against the rulesets as the plan leaves them.
        names = covered_by(desired, [actual], str(state['metadata'].get('default_branch', '')))
        if names:
            row.update(status='covered', covered_by=names)
    return row


def effective_rulesets(config: dict[str, Json], state: dict[str, Json], rows: list[Row]) -> list[Json]:
    """The live rulesets as the plan leaves them.

    A named ruleset that the plan replaces counts with its configured definition, not
    the live one: what the live definition enforces beyond the configuration is gone
    once it is replaced, so it cannot cover a missing ruleset.
    """
    replaced = {row['name'] for row in rows if row['status'] == 'differs'}
    desired = {r['name']: r for r in config['rulesets']}
    return [dict(desired[live['name']], id=live.get('id')) if live.get('name') in replaced else live
            for live in state['rulesets']]


def with_unreported_bypass(state: dict[str, Json]) -> tuple[dict[str, Json], list[str]]:
    """The state with each unreported bypass list taken as empty, and the names of those rulesets.

    GitHub never shows a workflow token a ruleset's bypass list. The workflow scope
    compares the rest of the ruleset and lists the bypass list as not checked.
    """
    unreported = sorted({str(r.get('name')) for r in state['rulesets'] if 'bypass_actors' not in r})
    rulesets = [r if 'bypass_actors' in r else dict(r, bypass_actors=[]) for r in state['rulesets']]
    return dict(state, rulesets=rulesets), unreported


def not_checked(report: dict[str, Json], unreported: list[str]) -> list[str]:
    """Each managed value that the workflow scope could not read; its row is neither drift nor a match."""
    found: list[str] = []
    for row in report['settings']:
        if row['observed'] is None:
            row['matches'] = None
            found.append(f'setting {row["setting"]}: not reported to this token')
    for group, key in (('security', 'feature'), ('actions', 'setting')):
        for row in report[group]:
            if row['matches'] is not None and row['observed'] is None:
                row['matches'] = None
                found.append(f'{group} {row[key]}: {row["source"]}')
    # Only the bypass lists of the named rulesets and of those that cover one matter.
    relevant = {row['name'] for row in report['rulesets'] if row['id'] is not None}
    relevant |= {name for row in report['rulesets'] for name in row['covered_by']}
    found += [f'ruleset "{name}" bypass_actors: not reported (needs write access to the ruleset)'
              for name in unreported if name in relevant]
    return found


def assess(config: dict[str, Json], state: dict[str, Json], scope: str = 'full') -> dict[str, Json]:
    """Compare every managed value with what GitHub reports.

    In the workflow scope a value the token cannot read is listed in not_checked
    and is neither drift nor a match; the result is PARTIAL when everything read
    matches and something is not checked.
    """
    unreported: list[str] = []
    if scope == 'workflow':
        state, unreported = with_unreported_bypass(state)
    metadata = state['metadata']
    configured_visibility = config['visibility']
    visibility = {'configured': configured_visibility, 'observed': metadata.get('visibility'),
                  'matches': configured_visibility in (None, metadata.get('visibility'))}
    # Rows in the documented order, whatever the order of the configuration file;
    # security features in write order.
    settings = [{'setting': key, 'configured': config['repository'][key], 'observed': metadata.get(key),
                 'matches': metadata.get(key) == config['repository'][key]}
                for key in GENERAL_KEYS]
    security = [compare('feature', key, config['security'][key], state['security'][key])
                for key in SECURITY_PUBLIC_ONLY]
    actions = [compare('setting', key, config['actions'][key], state['actions'][key])
               for key in ACTIONS_SOURCES]
    rulesets = [assess_ruleset(desired, state) for desired in config['rulesets']]
    effective = effective_rulesets(config, state, rulesets)
    for row, desired in zip(rulesets, config['rulesets']):
        if row['status'] == 'missing':
            names = covered_by(desired, effective, str(metadata.get('default_branch', '')))
            if names:
                row.update(status='covered', covered_by=names)
    managed = {r['name'] for r in config['rulesets']}
    report: dict[str, Json] = {
        'result': None, 'visibility': visibility,
        'settings': settings, 'security': security, 'actions': actions, 'rulesets': rulesets,
        'unmanaged_rulesets': [r['name'] for r in state['rulesets'] if r.get('name') not in managed],
        'inherited_rulesets': sorted(state['inherited'])}
    if scope == 'workflow':
        report['not_checked'] = not_checked(report, unreported)
    drift = (not visibility['matches']
             or any(row['matches'] is False for row in settings + security + actions)
             or any(row['status'] not in SATISFIED for row in rulesets))
    report['result'] = 'DRIFT' if drift else 'PARTIAL' if report.get('not_checked') else 'VERIFIED'
    return report


def plan_digest(repo: str, config_sha256: str, report: dict[str, Json], steps: Plan,
                rulesets: list[Json] | None = None) -> str:
    """SHA-256 of the plan: repository, configuration, observed managed values and exact writes.

    The observed definitions of the live rulesets that cover a managed one are
    included, so an approval goes stale when the covering protection changes;
    with no covering ruleset the document, and so the digest, is as without them.
    Sources and timestamps are left out, so the same unchanged repository and
    configuration give the same digest, and any change to either changes it.
    """
    def stable(rows: list[Row]) -> list[Row]:
        return [{key: value for key, value in row.items() if key != 'source'} for row in rows]
    observed: dict[str, Json] = {
        'visibility': report['visibility'], 'settings': report['settings'],
        'security': stable(report['security']), 'actions': stable(report['actions']),
        'rulesets': report['rulesets'],
        'unmanaged_rulesets': sorted(report['unmanaged_rulesets']),
        'inherited_rulesets': sorted(report['inherited_rulesets'])}
    covering = sorted({name for row in report['rulesets'] for name in row['covered_by']})
    if covering:
        observed['covering_rulesets'] = {name: [rule_spec(r) for r in rulesets or [] if r.get('name') == name]
                                         for name in covering}
    document = {
        'repository': repo, 'configuration_sha256': config_sha256, 'observed': observed,
        'writes': [{'method': method, 'endpoint': endpoint, 'payload': payload}
                   for method, endpoint, payload in steps],
    }
    text = json.dumps(document, sort_keys=True, separators=(',', ':'))
    return hashlib.sha256(text.encode('utf-8')).hexdigest()


def step_record(method: str, endpoint: str, payload: Json) -> dict[str, Json]:
    """A write as the records show it: method and endpoint, never the payload.

    A ruleset write also names its ruleset, so that the journal and the evidence
    tell apart writes that share the endpoint ``repos/OWNER/REPO/rulesets``.
    """
    record: dict[str, Json] = {'method': method, 'endpoint': endpoint}
    if endpoint.split('/')[3:4] == ['rulesets'] and isinstance(payload, dict) and payload.get('name'):
        record['ruleset'] = payload['name']
    return record


def plan(config: dict[str, Json], report: dict[str, Json], repo: str) -> Plan:
    """The writes that would make the repository match. Visibility and Actions are never written.

    Order matters. The merge methods come first, because GitHub may reject a ruleset
    that allows a merge method the repository disables. Missing rulesets are created
    before differing ones are replaced. The other general settings
    come last, after the rulesets and security features: automatic deletion of merged
    branches must not start before the deletion rulesets protect the retained branches.
    """
    root = f'repos/{repo}'
    steps: Plan = []
    differing = {row['setting']: row['configured'] for row in report['settings'] if not row['matches']}
    merge = {key: value for key, value in differing.items() if key in MERGE_KEYS}
    if merge:
        steps.append(('PATCH', root, merge))
    desired = {r['name']: r for r in config['rulesets']}
    # Missing rulesets are created before differing ones are replaced, so that a
    # replacement never narrows protection that only the replaced definition gave.
    for row in report['rulesets']:
        if row['status'] == 'missing':
            steps.append(('POST', root + '/rulesets', desired[row['name']]))
    for row in report['rulesets']:
        if row['status'] == 'differs':
            steps.append(('PUT', f'{root}/rulesets/{row["id"]}', desired[row['name']]))
    analysis: dict[str, dict[str, str]] = {}
    for row in report['security']:
        if row['matches'] is not False:
            continue
        if row['feature'] in ANALYSIS_KEYS:
            analysis[row['feature']] = {'status': 'enabled' if row['configured'] else 'disabled'}
        else:
            endpoint = root + '/' + row['feature'].replace('_', '-')
            steps.append(('PUT' if row['configured'] else 'DELETE', endpoint, None))
    if analysis:
        steps.append(('PATCH', root, {'security_and_analysis': analysis}))
    others = {key: value for key, value in differing.items() if key not in MERGE_KEYS}
    if others:
        steps.append(('PATCH', root, others))
    return steps


# Output


def shown(row: Row) -> str:
    return json.dumps(row['observed']) + (' (paused)' if row.get('paused') else '')


def print_notes(config: dict[str, Json], state: dict[str, Json], report: dict[str, Json]) -> None:
    metadata = state['metadata']
    print(f'Destination: {metadata["full_name"]}; visibility: {metadata["visibility"]}; '
          f'default branch: {metadata["default_branch"]}')
    if state['inherited']:
        print('NOTE: inherited rules also apply; they are not modified or certified by this script.')
    for detail in state['rulesets']:
        if detail.get('name') not in report['unmanaged_rulesets']:
            continue
        refs = ', '.join(((detail.get('conditions') or {}).get('ref_name') or {}).get('include') or [])
        print(f'NOTE: unmanaged ruleset "{detail["name"]}" '
              f'({detail.get("target", "unknown target")}: {refs or "no ref conditions"}; '
              f'{detail.get("enforcement", "unknown enforcement")}) stays active. This script '
              'never changes or deletes it; delete it in the repository settings if a managed '
              'ruleset replaces it.')
    skipped = [row['feature'] for row in report['security']
               if row['source'] == NOT_APPLIED and row['configured']]
    if skipped:
        print(f'NOTE: {", ".join(skipped)} apply only to public repositories here; they are not '
              f'applied while the destination is {metadata.get("visibility")}. Rerun the check '
              'after making the repository public.')
    if not state['public']:
        print(f'NOTE: fork_pr_approval_policy applies only to public repositories; it is not '
              f'checked while the destination is {metadata.get("visibility")}.')
    fixes = state['security']['automated_security_fixes']
    if fixes.get('paused') and config['security']['automated_security_fixes']:
        print('NOTE: Dependabot security updates are enabled but paused, so they open no pull '
              'requests. GitHub pauses them when Dependabot pull requests go unanswered; they '
              'count as drift, and an apply that does not resume them fails its read-back.')


def difference_lines(report: dict[str, Json]) -> list[str]:
    """One line per managed value that differs from the configuration, and one per covered ruleset."""
    lines: list[str] = []
    visibility = report['visibility']
    if not visibility['matches']:
        lines.append(f'DIFF visibility: observed {visibility["observed"]}, configured {visibility["configured"]} '
                     f'(never changed by this tool; see {DISTRIBUTION})')
    for row in report['settings']:
        if row['matches'] is False:
            lines.append(f'DIFF setting {row["setting"]}: observed {json.dumps(row["observed"])}, '
                         f'configured {json.dumps(row["configured"])}')
    for row in report['security']:
        if row['matches'] is False:
            lines.append(f'DIFF security {row["feature"]}: observed {shown(row)}, '
                         f'configured {json.dumps(row["configured"])} [{row["source"]}]')
    for row in report['actions']:
        if row['matches'] is False:
            lines.append(f'DIFF actions {row["setting"]}: observed {shown(row)}, '
                         f'configured {json.dumps(row["configured"])} [{row["source"]}]')
    for row in report['rulesets']:
        if row['status'] == 'covered':
            names = ', '.join(f'"{name}"' for name in row['covered_by'])
            lines.append(f'NOTE: ruleset "{row["name"]}" is covered by {names}: every configured rule '
                         'is enforced at least as strictly there, so no write is planned for it.')
            continue
        if row['status'] == 'missing':
            lines.append(f'DIFF ruleset "{row["name"]}": missing')
        lines += [f'DIFF ruleset "{row["name"]}" ({row["id"]}) {line}' for line in row['differences']]
    return lines


def print_differences(report: dict[str, Json]) -> None:
    for line in difference_lines(report):
        print(line)
    for line in report.get('not_checked', []):
        print(f'NOT CHECKED {line}')


def print_plan(steps: Plan) -> None:
    for method, endpoint, payload in steps:
        print(f'{method} {endpoint}')
        if payload is not None:
            print(json.dumps(payload, indent=2))


# Records


def markdown_record(state: dict[str, Json], report: dict[str, Json], plan_sha256: str, mode: str,
                    now: datetime) -> str:
    """A dated Markdown record of the managed settings as GitHub reports them."""
    metadata = state['metadata']

    def cell(value: Json) -> str:
        return 'on' if value is True else 'off' if value is False else 'unknown' if value is None else str(value)

    def observed(row: Row) -> str:
        if row['matches'] is None:
            return NOT_APPLIED
        return cell(row['observed']) + (' (paused)' if row.get('paused') else '')

    def status(row: Row) -> str:
        if row['status'] == 'covered':
            return 'covered by ' + ', '.join(f'"{name}"' for name in row['covered_by'])
        return str(row['status'])

    lines = ['# Repository settings evidence', '',
             f'- Checked: **{now.strftime("%Y-%m-%dT%H:%M:%SZ")}** with `tools/setup_repository.py` ({mode})',
             f'- Destination: `{metadata["full_name"]}`; visibility: {metadata["visibility"]} '
             f'(configured: {report["visibility"]["configured"] or "any"}); '
             f'default branch: `{metadata["default_branch"]}`',
             f'- Result: **{report["result"]}**',
             f'- Plan SHA-256: `{plan_sha256}`', '',
             '## General settings', '', '| Setting | Configured | Observed |', '| --- | --- | --- |']
    lines += [f'| `{row["setting"]}` | {cell(row["configured"])} | {cell(row["observed"])} |'
              for row in report['settings']]
    lines += ['', '## Security features', '', '| Feature | Configured | Observed |', '| --- | --- | --- |']
    lines += [f'| `{row["feature"]}` | {cell(row["configured"])} | {observed(row)} |'
              for row in report['security']]
    lines += ['', '## Actions settings', '', '| Setting | Configured | Observed |', '| --- | --- | --- |']
    lines += [f'| `{row["setting"]}` | {cell(row["configured"])} | {observed(row)} |'
              for row in report['actions']]
    lines += ['', '## Managed rulesets', '', '| Ruleset | ID | Status |', '| --- | --- | --- |']
    lines += [f'| {row["name"]} | {"—" if row["id"] is None else row["id"]} | {status(row)} |'
              for row in report['rulesets']]
    lines += ['', '## Other rules', '',
              f'- Unmanaged repository rulesets: {", ".join(report["unmanaged_rulesets"]) or "none"}',
              f'- Inherited rulesets: {", ".join(report["inherited_rulesets"]) or "none"}',
              '', 'Unmanaged settings are outside this record.']
    return '\n'.join(lines) + '\n'


def evidence_record(config_path: Path, config_sha256: str, repo: str, state: dict[str, Json],
                    report: dict[str, Json], steps: Plan, plan_sha256: str, mode: str,
                    now: datetime) -> dict[str, Json]:
    """A dated JSON record of the managed settings as GitHub reports them."""
    try:
        shown_path = config_path.resolve().relative_to(ROOT).as_posix()
    except ValueError:
        shown_path = config_path.name
    metadata = state['metadata']
    return {
        'schema_version': EVIDENCE_SCHEMA_VERSION,
        'record': 'repository-settings-evidence',
        'checked_at': now.strftime('%Y-%m-%dT%H:%M:%SZ'),
        'tool': 'tools/setup_repository.py',
        'mode': mode,
        'repository': metadata.get('full_name', repo),
        'default_branch': metadata.get('default_branch'),
        'configuration': {'path': shown_path, 'sha256': config_sha256},
        'result': report['result'],
        'visibility': report['visibility'],
        'settings': report['settings'],
        'security': report['security'],
        'actions': report['actions'],
        'rulesets': report['rulesets'],
        'unmanaged_rulesets': report['unmanaged_rulesets'],
        'inherited_rulesets': report['inherited_rulesets'],
        'planned_changes': [step_record(*step) for step in steps],
        'plan_sha256': plan_sha256,
        'scope': 'Only the settings that the configuration manages are recorded; '
                 'everything else is outside this record.',
    }


def markdown_summary(state: dict[str, Json], report: dict[str, Json], scope: str) -> str:
    """A short Markdown summary of the comparison, for a workflow's job summary."""
    metadata = state['metadata']
    lines = ['## Repository settings drift', '',
             f'- Destination: `{metadata.get("full_name")}`; default branch: `{metadata.get("default_branch")}`',
             '- Token scope: ' + ('workflow: what a workflow token with contents: read can read; '
                                  'the values it cannot read are not checked'
                                  if scope == 'workflow' else 'full: every unreadable value counts as drift'),
             f'- Result: **{report["result"]}**', '']
    found = difference_lines(report)
    sections = [('Drift', [line for line in found if line.startswith('DIFF ')]),
                ('Notes', [line for line in found if not line.startswith('DIFF ')])]
    if scope == 'workflow':
        sections.append(('Not checked', report['not_checked']))
    for title, items in sections:
        if items or title == 'Drift':
            lines += [f'### {title}', ''] + ([f'- {item}' for item in items] or ['None.']) + ['']
    if report.get('not_checked'):
        lines.append("Only `tools/setup_repository.py --check` with an administrator's token compares them.")
    return '\n'.join(lines).rstrip('\n') + '\n'


def append_text(path: Path, text: str) -> None:
    path.parent.mkdir(parents=True, exist_ok=True)
    with path.open('a', encoding='utf-8', newline='\n') as handle:
        handle.write(text)


def write_text(path: Path, text: str, what: str) -> None:
    path.parent.mkdir(parents=True, exist_ok=True)
    path.write_text(text, encoding='utf-8', newline='\n')
    print(f'{what} written to {path}')


def write_json_atomically(path: Path, record: dict[str, Json]) -> None:
    """Write the record to a temporary file beside the path, then replace the path with it."""
    path.parent.mkdir(parents=True, exist_ok=True)
    descriptor, temporary = tempfile.mkstemp(dir=path.parent, prefix=path.name + '.', suffix='.tmp')
    try:
        with os.fdopen(descriptor, 'w', encoding='utf-8', newline='\n') as handle:
            handle.write(json.dumps(record, indent=2) + '\n')
        os.replace(temporary, path)
    finally:
        if os.path.exists(temporary):
            os.unlink(temporary)


# Apply


def check_prerequisites(state: dict[str, Json], report: dict[str, Json]) -> None:
    """Refuse to apply anything that cannot be written safely, before any write."""
    metadata = state['metadata']
    if metadata.get('archived') or metadata.get('disabled'):
        raise SetupError('Destination is archived or disabled.')
    if not (metadata.get('permissions') or {}).get('admin'):
        raise SetupError('Administrative access to the destination is required to apply the baseline.')
    visibility = report['visibility']
    if not visibility['matches']:
        raise SetupError(f'The destination is {visibility["observed"]}, but the configuration requires '
                         f'{visibility["configured"]}. This tool never changes visibility; follow '
                         f'{DISTRIBUTION} before changing it, then rerun.')
    unknown = [row['setting'] for row in report['settings'] if row['observed'] is None]
    unknown += [row['feature'] for row in report['security']
                if row['matches'] is not None and row['observed'] is None]
    # Actions settings are never written, but an unread one would only fail the
    # read-back after every other write; refuse before any write instead.
    unknown += [row['setting'] for row in report['actions']
                if row['matches'] is not None and row['observed'] is None]
    unknown += [f'ruleset "{row["name"]}"' for row in report['rulesets'] if row['status'] == 'unknown']
    if unknown:
        raise SetupError('Cannot apply while these managed values are unknown: ' + ', '.join(unknown)
                         + '. Rerun preview to see why GitHub did not report them.')


def check_required_checks(call: Call, repo: str, config: dict[str, Json], metadata: dict[str, Json]) -> None:
    root = f'repos/{repo}'
    branch = call('GET', root + '/branches/' + quote(metadata['default_branch'], safe=''))
    checks = pages(call, root + '/commits/' + branch['commit']['sha'] + '/check-runs?filter=latest', 'check_runs')
    required = [check for r in config['rulesets'] for rule in r['rules']
                if rule['type'] == 'required_status_checks'
                for check in rule['parameters']['required_status_checks']]
    for required_check in required:
        if not any(c['name'] == required_check['context'] and
                   c.get('app', {}).get('id') == required_check['integration_id'] and
                   c.get('conclusion') == 'success' and c.get('status') == 'completed' for c in checks):
            raise SetupError('Required GitHub Actions check has not succeeded on the current default-branch '
                             'HEAD. Enable Actions and run Repository integrity before applying protection.')


def readback_error(report: dict[str, Json]) -> SetupError:
    """The read-back failure, naming each Actions setting to set by hand."""
    actions = [row for row in report['actions'] if row['matches'] is False]
    rest = report['result'] == 'DRIFT' and (
        not report['visibility']['matches']
        or any(row['matches'] is False for row in report['settings'] + report['security'])
        or any(row['status'] not in SATISFIED for row in report['rulesets']))
    message = 'Read-back verification failed'
    message += ('; the differences above remain. Rerun preview to inspect them.' if rest
                else ': the other settings match, but this tool never writes Actions settings.')
    if actions:
        message += (' Set ' + ', '.join(f'{row["setting"]} to {json.dumps(row["configured"])} '
                                        f'(observed {shown(row)})' for row in actions)
                    + " in the repository's Settings > Actions > General, then rerun --check.")
    return SetupError(message)


def check_approval(args: argparse.Namespace, digest: str) -> None:
    """Refuse an apply that is not bound to the plan just printed, before any write."""
    approved, journal = getattr(args, 'approve_plan', None), getattr(args, 'journal', None)
    if not approved or journal is None:
        raise SetupError('--apply requires --approve-plan SHA256 and --journal PATH.')
    if approved.lower() != digest:
        raise SetupError('The plan differs from the approved one: the repository or the configuration '
                         'changed since the reviewed plan, or the digest is not the one preview printed. '
                         'Nothing was written. Run preview again, review the plan above and approve '
                         'its PLAN SHA-256.')
    if Path(journal).exists():
        raise SetupError(f'Journal {journal} exists; choose a new path. Nothing was written.')


def apply(call: Call, args: argparse.Namespace, config: dict[str, Json], state: dict[str, Json],
          report: dict[str, Json], steps: Plan, digest: str, config_sha256: str,
          started: datetime) -> tuple[dict[str, Json], dict[str, Json], Plan]:
    """Write the approved plan, journaling before the first write and after each result.

    A failing write stops the sequence; the journal records it and the error
    propagates. Nothing is retried or rolled back. The read-back state is returned.
    """
    if any(method == 'PUT' and '/rulesets/' in endpoint for method, endpoint, _ in steps) \
            and not args.update_existing:
        raise SetupError('A named ruleset differs. Review the preview, then use --update-existing '
                         'to replace that entire named ruleset. Additional rules in it will be removed.')
    if steps:
        check_required_checks(call, args.repo, config, state['metadata'])
    journal: dict[str, Json] = {
        'record': 'repository-settings-journal',
        'tool': 'tools/setup_repository.py',
        'plan_sha256': digest,
        'repository': state['metadata'].get('full_name', args.repo),
        'configuration_sha256': config_sha256,
        'started_at': started.strftime('%Y-%m-%dT%H:%M:%SZ'),
        'writes': [{**step_record(*step), 'status': 'pending'} for step in steps],
        'readback': None,
    }
    write_json_atomically(args.journal, journal)
    for entry, (method, endpoint, payload) in zip(journal['writes'], steps):
        try:
            response = call(method, endpoint, payload)
        except Exception as error:
            entry.update(status='failed', error=str(error))
            write_json_atomically(args.journal, journal)
            raise
        entry['status'] = 'done'
        # A ruleset write records the ID GitHub answers with, so that a created
        # ruleset can be found again even when a later write fails.
        if 'ruleset' in entry and isinstance(response, dict) and isinstance(response.get('id'), int):
            entry['ruleset_id'] = response['id']
        write_json_atomically(args.journal, journal)
        print(f'Applied {method} {endpoint}')
    # Read back after every apply, an empty one included, so that the journal
    # records a fresh observation, never the one the plan was built from.
    state = observe(call, args.repo, config)
    report = assess(config, state)
    steps = plan(config, report, args.repo)
    journal['readback'] = report['result']
    write_json_atomically(args.journal, journal)
    print(f'Journal written to {args.journal}')
    return state, report, steps


def write_summary(args: argparse.Namespace, state: dict[str, Json], report: dict[str, Json],
                  scope: str) -> None:
    if getattr(args, 'summary', None):
        append_text(args.summary, markdown_summary(state, report, scope))


def workflow_result(report: dict[str, Json]) -> int:
    """The exit status of a workflow-scope check: 1 for drift in what it read, else 0."""
    if report['result'] == 'DRIFT':
        print('DRIFT: documented settings that this token can read do not match.')
        return 1
    if report['result'] == 'PARTIAL':
        print(f'PARTIAL: every documented value that this token can read matches; '
              f'{len(report["not_checked"])} value(s) listed as NOT CHECKED were not compared. '
              'Only --check with an administrator\'s token verifies them.')
        return 0
    print('VERIFIED: all documented settings match. Unmanaged settings are outside this verification.')
    return 0


def run(args: argparse.Namespace, call: Call = api, now: datetime | None = None) -> int:
    if not REPOSITORY_PATTERN.fullmatch(args.repo) or args.repo.split('/')[1] in ('.', '..'):
        raise SetupError('Destination must be an explicit owner/repository on github.com.')
    config = load_config(args.config)
    config_sha256 = hashlib.sha256(args.config.read_bytes()).hexdigest()
    mode = 'apply' if args.apply else 'check' if args.check else 'preview'
    scope = getattr(args, 'token_scope', 'full')
    if scope == 'workflow' and (mode != 'check' or getattr(args, 'report', None)
                                or getattr(args, 'evidence', None)):
        raise SetupError('--token-scope workflow works only with --check, without --report or --evidence.')
    started = now or datetime.now(timezone.utc)
    state = observe(call, args.repo, config, scope)
    report = assess(config, state, scope)
    print_notes(config, state, report)
    print_differences(report)
    if scope == 'workflow':
        write_summary(args, state, report, scope)
        return workflow_result(report)
    steps = plan(config, report, args.repo)
    digest = plan_digest(state['metadata'].get('full_name', args.repo), config_sha256, report, steps,
                         state['rulesets'])
    print_plan(steps)
    print(f'PLAN SHA-256: {digest}')
    if args.apply:
        check_approval(args, digest)
        check_prerequisites(state, report)
        state, report, steps = apply(call, args, config, state, report, steps, digest, config_sha256, started)
        digest = plan_digest(state['metadata'].get('full_name', args.repo), config_sha256, report, steps,
                             state['rulesets'])
    stamp = now or datetime.now(timezone.utc)
    if getattr(args, 'report', None):
        write_text(args.report, markdown_record(state, report, digest, mode, stamp), 'Record')
    if getattr(args, 'evidence', None):
        target = Path(str(args.evidence).replace('{date}', stamp.strftime('%Y-%m-%d')))
        record = evidence_record(args.config, config_sha256, args.repo, state, report, steps, digest, mode, stamp)
        write_text(target, json.dumps(record, indent=2) + '\n', 'Evidence')
    write_summary(args, state, report, scope)
    if report['result'] == 'VERIFIED':
        print('VERIFIED: all documented settings match' + (' after apply' if args.apply else '')
              + '. Unmanaged settings are outside this verification.')
        return 0
    if args.apply:
        print_differences(report)
        raise readback_error(report)
    if args.check:
        print('DRIFT: documented settings do not match.')
        return 1
    print('PREVIEW ONLY: no changes made. Use --apply to execute this plan.')
    return 0


def main() -> int:
    if sys.version_info < (3, 11):
        print('ERROR: Python 3.11 or later is required.', file=sys.stderr)
        return 2
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument('--repo', required=True, help='Explicit destination owner/repository')
    parser.add_argument('--config', type=Path, default=CONFIG)
    mode = parser.add_mutually_exclusive_group()
    mode.add_argument('--apply', action='store_true',
                      help='Apply the approved plan, then verify (needs --approve-plan and --journal)')
    mode.add_argument('--check', action='store_true', help='Read-only verification; exit 1 for drift')
    parser.add_argument('--approve-plan', metavar='SHA256',
                        help='The PLAN SHA-256 that preview printed for the plan to apply')
    parser.add_argument('--journal', type=Path,
                        help='Write a JSON journal of the apply to this new path, before the first '
                             'write and after each')
    parser.add_argument('--update-existing', action='store_true', help='Allow replacement of differing named rulesets')
    parser.add_argument('--report', type=Path, help='Write a dated Markdown evidence record of the managed settings')
    parser.add_argument('--evidence', type=Path,
                        help='Write a JSON evidence record of the managed settings ({date} becomes YYYY-MM-DD, UTC)')
    parser.add_argument('--summary', type=Path, help='Append a Markdown summary of the comparison to this file')
    parser.add_argument('--token-scope', choices=TOKEN_SCOPES, default='full',
                        help='workflow: with --check, compare only what a workflow GITHUB_TOKEN with '
                             'contents: read can read and list the rest as not checked')
    args = parser.parse_args()
    for option in ('approve_plan', 'journal', 'update_existing'):
        if getattr(args, option) and not args.apply:
            parser.error(f'--{option.replace("_", "-")} requires --apply')
    if args.token_scope == 'workflow' and (not args.check or args.report or args.evidence):
        parser.error('--token-scope workflow requires --check and takes no --report or --evidence')
    if args.apply and (args.approve_plan is None or args.journal is None):
        parser.error('--apply requires --approve-plan SHA256 and --journal PATH')
    if args.approve_plan is not None and not DIGEST_PATTERN.fullmatch(args.approve_plan):
        parser.error('--approve-plan takes the 64 hexadecimal digits of the PLAN SHA-256 line that preview printed')
    try:
        return run(args)
    except (SetupError, OSError, ValueError, KeyError, TypeError) as exc:
        print(f'ERROR: {exc}', file=sys.stderr)
        return 2


if __name__ == '__main__':
    sys.exit(main())
