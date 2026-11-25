"""Governance invariants of the workflows and Dependabot, read with a real YAML parser.

The parser is pinned in .github/requirements.txt; install it with
``python -m pip install -r .github/requirements.txt``. Without it the tests
are skipped locally and fail in GitHub Actions.
"""

# mypy checks the bodies of these tests but does not require them to be annotated.
# mypy: disallow-untyped-defs=False, disallow-incomplete-defs=False, disallow-untyped-calls=False

import contextlib
import io
import json
import os
import re
import shutil
import subprocess
import sys
import tempfile
import unittest
import unittest.mock
from email.message import Message
from pathlib import Path
from urllib.error import HTTPError

ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(ROOT / 'tests'))
sys.path.insert(0, str(ROOT / 'tools'))

import check_repository  # noqa: E402

try:
    import strict_yaml
    import yaml
except ImportError:  # pragma: no cover - exercised only without the parser
    yaml = None

WORKFLOWS = ROOT / '.github/workflows'
REQUIRED_JOB = ('repository-integrity.yml', 'repository-integrity')
REQUIRED_CHECK = 'Repository integrity'
MINIMUM_JOB = ('minimum-python.yml', 'minimum-python')
# The only jobs that may write, and what each may write.
WRITE_PERMISSIONS = {
    ('labels-sync.yml', 'reconcile'): {'issues'},
    ('traffic-history.yml', 'collect'): {'contents'},
    ('traffic-history.yml', 'alert'): {'issues'},
    ('pr-milestones.yml', 'reconcile'): {'pull-requests'},
    ('codeql.yml', 'analyze'): {'security-events'},
    ('scorecard.yml', 'analysis'): {'id-token', 'security-events'},
    ('release.yml', 'attest'): {'id-token', 'attestations'},
    ('release.yml', 'draft'): {'contents'},
}
PINNED = re.compile(r'^[\w.-]+/[\w.-]+(/[\w./-]+)?@[0-9a-f]{40}$')
DOCKER_PINNED = re.compile(r'^docker://[^@\s]+@sha256:[0-9a-f]{64}$')
# The version a pin stands for, after it on the same line: "# v4.39.1".
VERSION_COMMENT = r'[ \t]+#[ \t]*v?\d+(\.\d+){0,3}\b'
ACTION_FILES = ('action.yml', 'action.yaml')


def load(path):
    document = strict_yaml.load(path.read_text(encoding='utf-8'))
    # YAML 1.1 reads the bare key "on" as True.
    if True in document:
        document['on'] = document.pop(True)
    return document


def workflow_paths(root=ROOT):
    folder = root / '.github/workflows'
    return sorted([*folder.glob('*.yml'), *folder.glob('*.yaml')])


def workflows():
    return {path.name: load(path) for path in workflow_paths()}


class Source:
    """The raw text of one workflow or action file, with the uses: lines read so far.

    YAML drops comments, so each reference's version comment is read from its own
    line of the raw text, occurrence by occurrence: the n-th time a reference is
    parsed, the n-th line that writes it must carry the comment. A reference
    written in a form this does not match, such as a flow mapping, has fewer such
    lines than occurrences and fails the check, so the rule fails closed."""

    def __init__(self, text):
        self.text = text
        self.seen = {}

    def has_version_comment(self, reference, key='uses'):
        quoted = rf'(?P<q>[\'"]?){re.escape(reference)}(?P=q)'
        lines = re.finditer(rf'(?m)^[ \t-]*["\']?{key}["\']?:[ \t]*{quoted}(?P<rest>[^\n]*)$',
                            self.text)
        commented = [re.match(VERSION_COMMENT, line['rest']) is not None for line in lines]
        index = self.seen.get((key, reference), 0)
        self.seen[(key, reference)] = index + 1
        return index < len(commented) and commented[index]


def reference_problems(where, reference, source, root, seen, key='uses'):
    """The problems of one reference, following a local action into its metadata."""
    if not isinstance(reference, str) or not reference:
        return [f'{where}: {key}: is not a reference']
    if reference.startswith('./'):
        return local_problems(where, reference, root, seen)
    if reference.startswith('docker://'):
        problems = [] if DOCKER_PINNED.match(reference) else [
            f'{where}: {reference} is not pinned to an @sha256: digest']
    else:
        problems = [] if PINNED.match(reference) else [
            f'{where}: {reference} is not pinned to a full commit SHA']
    if not source.has_version_comment(reference, key):
        problems.append(f'{where}: {reference} has no version comment')
    return problems


def local_problems(where, reference, root, seen):
    """A local reusable workflow must exist; a local action is checked step by step."""
    target = (root / reference).resolve()
    # ./ is an action at the repository root, which is inside it.
    if target != root.resolve() and root.resolve() not in target.parents:
        return [f'{where}: {reference} is outside the repository']
    workflows_folder = (root / '.github/workflows').resolve()
    if target.suffix in ('.yml', '.yaml'):
        if target.parent != workflows_folder:
            return [f'{where}: {reference} is not in .github/workflows']
        # It is scanned as a workflow in its own right.
        return [] if target.is_file() else [f'{where}: {reference} does not exist']
    metadata = next((target / name for name in ACTION_FILES if (target / name).is_file()), None)
    if metadata is None:
        return [f'{where}: {reference} has no action.yml or action.yaml']
    if metadata in seen:
        return []  # Already checked, or a cycle: either way, nothing new to read.
    seen.add(metadata)
    source = Source(metadata.read_text(encoding='utf-8'))
    label = metadata.relative_to(root.resolve()).as_posix()
    runs = (load_text(source.text) or {}).get('runs') or {}
    problems = []
    for index, step in enumerate(runs.get('steps') or []):
        if isinstance(step, dict) and 'uses' in step:
            problems += reference_problems(f'{label} step {index + 1}', step['uses'], source,
                                           root, seen)
    image = runs.get('image')
    if isinstance(image, str) and image.startswith('docker://'):
        problems += reference_problems(f'{label} image', image, source, root, seen, 'image')
    return problems


def load_text(text):
    document = strict_yaml.load(text)
    if isinstance(document, dict) and True in document:
        document['on'] = document.pop(True)
    return document


def pin_problems(root=ROOT):
    """Every uses: reference reachable from the workflows of ``root`` that is not immutable.

    A remote action or reusable workflow is pinned to a full commit SHA and a
    Docker image to an @sha256: digest, each with a version comment; a local
    reusable workflow must exist in .github/workflows; a local action is followed
    into its action.yml, recursively and once."""
    problems = []
    seen: set[Path] = set()
    for path in workflow_paths(root):
        source = Source(path.read_text(encoding='utf-8'))
        document = load_text(source.text) or {}
        for job_id, job in (document.get('jobs') or {}).items():
            where = f'{path.name} job {job_id}'
            if 'uses' in job:
                problems += reference_problems(where, job['uses'], source, root, seen)
            for index, step in enumerate(job.get('steps') or []):
                if isinstance(step, dict) and 'uses' in step:
                    problems += reference_problems(f'{where} step {index + 1}', step['uses'],
                                                   source, root, seen)
    return problems


@unittest.skipIf(yaml is None and not os.environ.get('GITHUB_ACTIONS'),
                 'PyYAML is not installed: python -m pip install -r .github/requirements.txt')
class WorkflowControlTests(unittest.TestCase):
    def test_parser_is_available(self):
        self.assertIsNotNone(yaml, 'install .github/requirements.txt before running the tests')

    def test_duplicate_keys_are_rejected(self):
        self.assertEqual(strict_yaml.load('a: 1\nb: {c: 2}\n'), {'a': 1, 'b': {'c': 2}})
        for text in ('permissions: {}\npermissions:\n  contents: write\n',
                     'jobs:\n  build:\n    runs-on: a\n    runs-on: b\n'):
            with self.subTest(text=text), self.assertRaisesRegex(yaml.YAMLError, 'duplicate key'):
                strict_yaml.load(text)

    def test_repository_yaml_has_no_duplicate_keys(self):
        files = sorted([*ROOT.glob('.github/**/*.yml'), *ROOT.glob('.github/**/*.yaml')])
        self.assertTrue(any(path.parent.name == 'ISSUE_TEMPLATE' for path in files))
        for path in files:
            with self.subTest(path=path.relative_to(ROOT).as_posix()):
                strict_yaml.load(path.read_text(encoding='utf-8'))

    def test_actions_are_pinned_and_checkouts_drop_credentials(self):
        self.assertEqual(pin_problems(), [])
        for name, document in workflows().items():
            for job_id, job in document['jobs'].items():
                for step in job.get('steps', []):
                    if 'uses' not in step:
                        continue
                    with self.subTest(workflow=name, job=job_id, uses=step['uses']):
                        if step['uses'].startswith('actions/checkout@'):
                            self.assertIs((step.get('with') or {}).get('persist-credentials'), False)

    def test_pins_cover_every_reachable_reference(self):
        sha, digest = 'a' * 40, 'b' * 64
        pinned = f'actions/checkout@{sha} # v7.0.1'
        cases = {
            'pinned step, job-level call, local reusable workflow and local action': (
                {'.github/workflows/ci.yml': f'jobs:\n  build:\n    steps:\n      - uses: {pinned}\n'
                                             '      - uses: ./.github/actions/setup\n'
                                             f'  call:\n    uses: owner/repo/.github/workflows/x.yml@{sha} # v1\n'
                                             '  local:\n    uses: ./.github/workflows/called.yml\n',
                 '.github/workflows/called.yml': f'jobs:\n  run:\n    steps:\n      - uses: {pinned}\n',
                 '.github/actions/setup/action.yml': f'runs:\n  using: composite\n  steps:\n    - uses: {pinned}\n'},
                []),
            'unpinned job-level call': (
                {'.github/workflows/ci.yml': 'jobs:\n  call:\n    uses: owner/repo/.github/workflows/x.yml@main # v1\n'},
                ['ci.yml job call: owner/repo/.github/workflows/x.yml@main is not pinned to a full commit SHA']),
            'pin without a version comment': (
                {'.github/workflows/ci.yml': f'jobs:\n  build:\n    steps:\n      - uses: actions/checkout@{sha}\n'},
                [f'ci.yml job build step 1: actions/checkout@{sha} has no version comment']),
            'a flow mapping cannot show its comment, so it fails closed': (
                {'.github/workflows/ci.yml': f'jobs:\n  build:\n    steps:\n      - {{uses: "actions/checkout@{sha}"}} # v7\n'},
                [f'ci.yml job build step 1: actions/checkout@{sha} has no version comment']),
            'a repeated reference needs its own comment, on its own line': (
                {'.github/workflows/ci.yml': f'jobs:\n  build:\n    steps:\n      - uses: {pinned}\n'
                                             f'      - uses: actions/checkout@{sha}\n'
                                             f'      - uses: actions/checkout@{sha}\n'
                                             '        # v7.0.1\n'},
                [f'ci.yml job build step 2: actions/checkout@{sha} has no version comment',
                 f'ci.yml job build step 3: actions/checkout@{sha} has no version comment']),
            'an action at the repository root': (
                {'.github/workflows/ci.yml': 'jobs:\n  build:\n    steps:\n      - uses: ./\n',
                 'action.yml': 'runs:\n  using: composite\n  steps:\n    - uses: actions/checkout@v7 # v7\n'},
                ['action.yml step 1: actions/checkout@v7 is not pinned to a full commit SHA']),
            'unpinned step inside a local action': (
                {'.github/workflows/ci.yml': 'jobs:\n  build:\n    steps:\n      - uses: ./tools/setup\n',
                 'tools/setup/action.yaml': 'runs:\n  using: composite\n  steps:\n    - uses: actions/checkout@v7 # v7\n'},
                ['tools/setup/action.yaml step 1: actions/checkout@v7 is not pinned to a full commit SHA']),
            'Docker tag without a digest, in a step and in an action image': (
                {'.github/workflows/ci.yml': 'jobs:\n  build:\n    steps:\n      - uses: docker://alpine:3.20 # v3.20\n'
                                             f'      - uses: docker://alpine@sha256:{digest} # v3.20\n'
                                             '      - uses: ./box\n',
                 'box/action.yml': 'runs:\n  using: docker\n  image: docker://alpine:3.20 # v3.20\n'},
                ['ci.yml job build step 1: docker://alpine:3.20 is not pinned to an @sha256: digest',
                 'box/action.yml image: docker://alpine:3.20 is not pinned to an @sha256: digest']),
            'local actions that use each other are read once': (
                {'.github/workflows/ci.yml': 'jobs:\n  build:\n    steps:\n      - uses: ./a\n',
                 'a/action.yml': 'runs:\n  using: composite\n  steps:\n    - uses: ./b\n',
                 'b/action.yml': 'runs:\n  using: composite\n  steps:\n    - uses: ./a\n'},
                []),
            'missing or misplaced local references': (
                {'.github/workflows/ci.yml': 'jobs:\n  a:\n    uses: ./.github/workflows/missing.yml\n'
                                             '  b:\n    uses: ./other/x.yml\n'
                                             '  c:\n    steps:\n      - uses: ./no-action\n'
                                             '      - uses: ./../outside\n'},
                ['ci.yml job a: ./.github/workflows/missing.yml does not exist',
                 'ci.yml job b: ./other/x.yml is not in .github/workflows',
                 'ci.yml job c step 1: ./no-action has no action.yml or action.yaml',
                 'ci.yml job c step 2: ./../outside is outside the repository']),
        }
        for label, (files, expected) in cases.items():
            with self.subTest(case=label), tempfile.TemporaryDirectory() as temporary:
                root = Path(temporary)
                for name, text in files.items():
                    (root / name).parent.mkdir(parents=True, exist_ok=True)
                    (root / name).write_text(text, encoding='utf-8')
                self.assertEqual(pin_problems(root), expected)

    def test_every_job_has_a_timeout_and_permissions_are_explicit(self):
        for name, document in workflows().items():
            with self.subTest(workflow=name):
                self.assertIn('permissions', document)
            for job_id, job in document['jobs'].items():
                # A job that calls a reusable workflow cannot set a timeout; the
                # jobs of a local one are checked in its own file.
                if 'uses' in job:
                    continue
                with self.subTest(workflow=name, job=job_id):
                    self.assertIsInstance(job.get('timeout-minutes'), int)

    def test_only_the_listed_jobs_write(self):
        def writes(permissions):
            if permissions in ('write-all', 'read-all'):
                return {permissions}
            return {scope for scope, level in (permissions or {}).items() if level == 'write'}

        for name, document in workflows().items():
            with self.subTest(workflow=name, level='workflow'):
                self.assertEqual(writes(document['permissions']), set())
            for job_id, job in document['jobs'].items():
                with self.subTest(workflow=name, job=job_id):
                    self.assertEqual(writes(job.get('permissions')),
                                     WRITE_PERMISSIONS.get((name, job_id), set()))

    def test_required_check_runs_unconditionally_on_every_pull_request(self):
        name, job_id = REQUIRED_JOB
        document = workflows()[name]
        # A pull request to any base gets the required check, or it could never
        # merge. Pushes are filtered to the protected branches (next test).
        self.assertIn('pull_request', document['on'])
        self.assertFalse(document['on']['pull_request'],
                         'pull_request must have no branch or path filters')
        job = document['jobs'][job_id]
        self.assertNotIn('if', job)
        self.assertNotIn('strategy', job)
        # The rulesets require this name; tests/test_setup_repository.py checks the
        # settings file against it while the setup tooling is kept.
        self.assertEqual(job['name'], REQUIRED_CHECK)

    def test_every_workflow_declares_concurrency(self):
        for name, document in workflows().items():
            with self.subTest(workflow=name):
                self.assertTrue(document.get('concurrency', {}).get('group'))
                self.assertIn('cancel-in-progress', document['concurrency'])

    def test_integrity_runs_on_pull_requests_and_on_protected_pushes_only(self):
        triggers = workflows()[REQUIRED_JOB[0]]['on']
        # A task branch's head is checked by its pull request run; a second
        # run on the push would only duplicate it.
        self.assertIsNone(triggers['pull_request'])
        # The default branch, by its name in this repository, and the release
        # branches. codeql_default_branch() in tools/check_repository.py checks
        # that the name is the default branch's, so a renamed one passes here.
        self.assertEqual(set(triggers['push']), {'branches'})
        branches = triggers['push']['branches']
        self.assertEqual(len(branches), 2, branches)
        self.assertIn('release/**', branches)
        default = next(name for name in branches if name != 'release/**')
        self.assertNotIn('*', default)
        self.assertIn('workflow_dispatch', triggers)

    def test_integrity_cancels_only_a_superseded_pull_request_run(self):
        concurrency = workflows()[REQUIRED_JOB[0]]['concurrency']
        # A pull request's runs share a group; every other run, a push to a
        # protected branch included, has its own and is never cancelled.
        self.assertIn("github.event_name == 'pull_request'", concurrency['group'])
        self.assertIn('github.event.pull_request.number', concurrency['group'])
        self.assertIn('github.run_id', concurrency['group'])
        self.assertNotIn('github.ref', concurrency['group'])
        self.assertEqual(concurrency['cancel-in-progress'],
                         "${{ github.event_name == 'pull_request' }}")

    def test_live_label_reconciliation_runs_only_on_the_default_branch(self):
        condition = workflows()['labels-sync.yml']['jobs']['reconcile']['if']
        self.assertEqual(condition,
                         "github.ref == format('refs/heads/{0}', "
                         "github.event.repository.default_branch)")

    def test_label_workflows_leave_pull_requests_to_repository_integrity(self):
        documents = workflows()
        # No pull request runs these workflows: Repository integrity runs their
        # offline commands on every pull request instead.
        expected = {'labels-sync.yml': ({'push', 'workflow_dispatch'}, ['reconcile']),
                    'labels-drift.yml': ({'schedule', 'workflow_dispatch'}, ['detect'])}
        for name, (triggers, jobs) in expected.items():
            with self.subTest(workflow=name):
                document = documents[name]
                self.assertEqual(set(document['on']), triggers)
                self.assertNotIn('pull_request', document['on'])
                self.assertNotIn('pull_request_target', document['on'])
                self.assertEqual(list(document['jobs']), jobs)
                self.assertNotIn('pull_request', str(document['jobs']))
        integrity = {' '.join(['node', *check.command[1:]])
                     for check in check_repository.checks(check_repository.Context(ROOT))
                     if check.command and len(check.command) > 1
                     and check.command[1].startswith('.github/scripts/')}
        offline = {'node .github/scripts/labels-sync.mjs --self-test',
                   'node .github/scripts/labels-sync.mjs --mode validate',
                   'node .github/scripts/labels-drift.mjs --self-test'}
        self.assertLessEqual(offline, integrity)
        # The live jobs still exercise the offline paths before reading GitHub.
        for name, job_id, command in (
                ('labels-sync.yml', 'reconcile', 'node .github/scripts/labels-sync.mjs --self-test'),
                ('labels-drift.yml', 'detect', 'node .github/scripts/labels-drift.mjs --self-test')):
            with self.subTest(workflow=name, command=command):
                runs = [step.get('run') for step in documents[name]['jobs'][job_id]['steps']]
                self.assertIn(command, runs)

    def test_scans_run_weekly_off_the_hour(self):
        # A weekly cron names one weekday; a daily one has * there.
        documents = workflows()
        for name in ('labels-drift.yml', 'pr-milestones.yml'):
            with self.subTest(workflow=name):
                crons = [entry['cron'] for entry in documents[name]['on']['schedule']]
                self.assertEqual(len(crons), 1, crons)
                minute, hour, day, month, weekday = crons[0].split()
                self.assertNotIn(minute, ('0', '00'), 'schedule off the hour')
                self.assertRegex(minute, r'^\d{1,2}$')
                self.assertRegex(hour, r'^\d{1,2}$')
                self.assertEqual((day, month), ('*', '*'))
                self.assertRegex(weekday, r'^[0-6]$', f'{name} must run weekly: {crons[0]}')

    def test_label_pruning_sees_pull_requests(self):
        # Without pull-requests: read the issues API omits pull requests, so a
        # label only they carry would look unused and be deleted.
        permissions = workflows()['labels-sync.yml']['jobs']['reconcile']['permissions']
        self.assertEqual(permissions.get('pull-requests'), 'read')

    @unittest.skipUnless((WORKFLOWS / 'traffic-history.yml').is_file(),
                         '.github/workflows/traffic-history.yml does not exist')
    def test_traffic_app_credentials_stay_in_their_steps(self):
        jobs = workflows()['traffic-history.yml']['jobs']
        steps = jobs['collect']['steps']
        with_key = [step for step in steps if 'TRAFFIC_APP_PRIVATE_KEY' in str(step)]
        self.assertEqual([step['name'] for step in with_key],
                         ['Check the traffic App configuration', 'Create the traffic App token'])
        # The configuration check sees only whether the key is set.
        self.assertEqual(with_key[0]['env']['HAS_KEY'],
                         "${{ secrets.TRAFFIC_APP_PRIVATE_KEY != '' }}")
        token_step = with_key[1]
        self.assertTrue(token_step['uses'].startswith('actions/create-github-app-token@'))
        self.assertEqual(token_step['with'], {
            'client-id': '${{ vars.TRAFFIC_APP_CLIENT_ID }}',
            'private-key': '${{ secrets.TRAFFIC_APP_PRIVATE_KEY }}',
            'permission-administration': 'read',
        })
        with_token = [step for step in steps if 'steps.app-token.outputs.token' in str(step)]
        self.assertEqual([step['name'] for step in with_token], ['Fetch traffic from the GitHub API'])
        self.assertNotRegex(with_token[0]['run'], r'\bpython|\.py\b|source/')
        alert = jobs['alert']
        self.assertNotIn('environment', alert)
        self.assertNotIn('secrets.', str(alert))

    @unittest.skipUnless((WORKFLOWS / 'traffic-history.yml').is_file(),
                         '.github/workflows/traffic-history.yml does not exist')
    def test_traffic_alert_reads_the_commit_its_collection_wrote(self):
        jobs = workflows()['traffic-history.yml']['jobs']
        collect, alert = jobs['collect'], jobs['alert']
        self.assertEqual(collect['outputs'], {'history-sha': '${{ steps.push.outputs.sha }}'})
        push = next(step for step in collect['steps'] if step.get('id') == 'push')
        # Both successful exits, pushed and nothing to commit, record the commit.
        self.assertIn('echo "sha=$(git rev-parse HEAD)" >> "$GITHUB_OUTPUT"', push['run'])
        self.assertEqual(push['run'].count('exit 0'), 1, 'only recorded() exits successfully')
        self.assertEqual(push['run'].count('recorded\n'), 2)
        self.assertEqual(alert['needs'], 'collect')
        history = [step for step in alert['steps']
                   if step.get('uses', '').startswith('actions/checkout@') and step['with'].get('path') == 'history']
        self.assertEqual(len(history), 1)
        self.assertEqual(history[0]['with']['ref'], '${{ needs.collect.outputs.history-sha }}')
        # An empty output must stop the job, not check out the default branch.
        guard = alert['steps'][alert['steps'].index(history[0]) - 1]
        self.assertEqual(guard['env'], {'HISTORY_SHA': '${{ needs.collect.outputs.history-sha }}'})
        self.assertIn('^[0-9a-f]{40}$', guard['run'])
        self.assertIn('exit 1', guard['run'])

    def test_milestone_workflow_never_runs_pull_request_code(self):
        document = workflows()['pr-milestones.yml']
        self.assertIn('github.event.pull_request.number', document['concurrency']['group'])
        self.assertEqual(set(document['on']), {'push', 'pull_request_target', 'schedule', 'workflow_dispatch'})
        self.assertEqual(document['on']['push'], {'paths': ['.github/workflows/pr-milestones.yml']})
        self.assertEqual(document['on']['pull_request_target'],
                         {'types': ['opened', 'reopened', 'edited', 'milestoned']})
        self.assertEqual(document['permissions'], {})
        job = document['jobs']['reconcile']
        self.assertEqual(job['permissions'], {'issues': 'read', 'pull-requests': 'write'})
        self.assertIn('github.event.repository.default_branch', job['if'])
        self.assertEqual(len(job['steps']), 1)
        step = job['steps'][0]
        self.assertEqual(set(step), {'name', 'env', 'run'})
        self.assertEqual(step['env'], {'GH_TOKEN': '${{ github.token }}'})
        self.assertNotIn('${{', step['run'])

    def test_closing_keywords_read_the_pull_request_from_the_event_file(self):
        document = workflows()['closing-keywords.yml']
        self.assertEqual(document['on'], {'pull_request': {
            'types': ['opened', 'synchronize', 'reopened', 'edited']}})
        self.assertIn('github.event.pull_request.number', document['concurrency']['group'])
        self.assertEqual(document['permissions'], {'contents': 'read'})
        self.assertEqual(list(document['jobs']), ['check'])
        job = document['jobs']['check']
        # Read-only, and no condition: every pull request reports a result.
        self.assertEqual(job['permissions'],
                         {'contents': 'read', 'issues': 'read', 'pull-requests': 'read'})
        self.assertNotIn('if', job)
        checkout = next(step for step in job['steps']
                        if str(step.get('uses', '')).startswith('actions/checkout@'))
        self.assertEqual(checkout['with']['fetch-depth'], 0)
        step = next(step for step in job['steps']
                    if 'check_closing_keywords.py' in (step.get('run') or ''))
        self.assertEqual(step['env'], {'GITHUB_TOKEN': '${{ secrets.GITHUB_TOKEN }}'})
        self.assertNotIn('continue-on-error', step)
        # The title and body are untrusted: the tool reads them from the event file.
        self.assertNotIn('${{', step['run'])
        self.assertIn('--event "$GITHUB_EVENT_PATH"', step['run'])

    def test_settings_drift_checks_on_a_schedule_with_the_read_only_token(self):
        document = workflows()['settings-drift.yml']
        self.assertEqual(set(document['on']), {'schedule', 'workflow_dispatch'})
        crons = [entry['cron'] for entry in document['on']['schedule']]
        self.assertEqual(len(crons), 1)
        self.assertNotEqual(crons[0].split()[0], '0', 'schedule off the hour')
        # Contents read only: no administrator token, no secret and no write.
        self.assertEqual(document['permissions'], {'contents': 'read'})
        self.assertNotIn('secrets.', str(document))
        jobs = document['jobs']
        self.assertEqual(list(jobs), ['detect'])
        job = jobs['detect']
        self.assertEqual(job['permissions'], {'contents': 'read'})
        self.assertNotIn('if', job)
        steps = job['steps']
        step = next(step for step in steps if 'setup_repository.py' in (step.get('run') or ''))
        self.assertEqual(step['env']['GH_TOKEN'], '${{ github.token }}')
        self.assertNotIn('continue-on-error', step)
        for argument in ('python tools/setup_repository.py', '--repo "$GITHUB_REPOSITORY"', '--check',
                         '--token-scope workflow', '--summary "$GITHUB_STEP_SUMMARY"'):
            with self.subTest(argument=argument):
                self.assertIn(argument, step['run'])
        for option in ('--apply', '--report', '--evidence'):
            with self.subTest(option=option):
                self.assertNotIn(option, step['run'])
        setup = [i for i, s in enumerate(steps) if str(s.get('uses', '')).startswith('actions/setup-python@')]
        self.assertTrue(setup and setup[0] < steps.index(step), 'set up Python before the check')

    def test_scorecard_publishes_only_from_the_default_branch_of_a_public_repository(self):
        document = workflows()['scorecard.yml']
        self.assertNotIn('pull_request', document['on'])
        self.assertIn('schedule', document['on'])
        # Scheduled events carry no repository in their payload: the weekly
        # run must be exempt from the branch test, and the visibility is read
        # from the API in a job of its own.
        gate = document['jobs']['visibility']
        self.assertIn("github.event_name == 'schedule' ||", gate['if'])
        self.assertIn('github.ref_name == github.event.repository.default_branch', gate['if'])
        self.assertNotIn('github.event.repository.private', str(document))
        self.assertIn('.visibility', gate['steps'][0]['run'])
        self.assertEqual(gate['permissions'], {'contents': 'read'})
        job = document['jobs']['analysis']
        self.assertEqual(job['needs'], 'visibility')
        self.assertEqual(job['if'], "needs.visibility.outputs.public == 'true'")
        # The Scorecard service accepts published results only from a job made
        # of approved actions, with no run steps.
        self.assertTrue(all('uses' in step and 'run' not in step for step in job['steps']))
        self.assertEqual([step['uses'].split('@')[0] for step in job['steps']],
                         ['actions/checkout', 'ossf/scorecard-action', 'actions/upload-artifact',
                          'github/codeql-action/upload-sarif'])
        self.assertIs(job['steps'][1]['with']['publish_results'], True)

    def test_code_scanning_covers_each_language_and_skips_when_unavailable(self):
        document = workflows()['codeql.yml']
        triggers = document['on']
        self.assertEqual(set(triggers), {'push', 'pull_request', 'schedule', 'workflow_dispatch'})
        # The default branch is listed by name and may be renamed; Repository
        # integrity checks that it is present (see the next test). Pushes are
        # analyzed on the default branch only, not on the release branches.
        self.assertEqual(set(triggers['push']), {'branches'})
        branches = triggers['push']['branches']
        self.assertEqual(len(branches), 1, branches)
        self.assertNotIn('*', branches[0])
        self.assertNotIn('release/**', branches)
        # Pull requests into the default and release branches are analyzed,
        # unless they change only Markdown files and files under docs/.
        self.assertEqual(set(triggers['pull_request']), {'branches', 'paths-ignore'})
        self.assertEqual(triggers['pull_request']['branches'], [branches[0], 'release/**'])
        self.assertEqual(triggers['pull_request']['paths-ignore'], ['**/*.md', 'docs/**'])
        jobs = document['jobs']
        gate = jobs['availability']['steps'][0]
        self.assertIn('.visibility', gate['run'])
        self.assertEqual(gate['env']['CODE_SCANNING'], '${{ vars.CODE_SCANNING }}')
        expected = {
            'analyze': ("needs.availability.outputs.enabled == 'true' && "
                        "github.event_name != 'pull_request'",
                        ['actions/checkout', 'github/codeql-action/init',
                         'github/codeql-action/analyze']),
            'analyze-pr': ("needs.availability.outputs.enabled == 'true' && "
                           "github.event_name == 'pull_request'",
                           ['actions/checkout', 'github/codeql-action/init',
                            'github/codeql-action/analyze', 'actions/upload-artifact']),
        }
        for job_id, (condition, sequence) in expected.items():
            with self.subTest(job=job_id):
                job = jobs[job_id]
                self.assertEqual(job['needs'], 'availability')
                self.assertEqual(job['if'], condition)
                self.assertEqual(sorted(job['strategy']['matrix']['language']),
                                 ['actions', 'javascript-typescript', 'python'])
                steps = [step for step in job['steps'] if 'uses' in step]
                self.assertEqual([step['uses'].split('@')[0] for step in steps], sequence)
                self.assertEqual(steps[1]['with']['queries'], 'security-extended')
        # Pull request analysis writes nothing to code scanning.
        self.assertEqual(jobs['analyze-pr']['permissions'], {'contents': 'read'})
        pr_analyze = jobs['analyze-pr']['steps'][2]['with']
        self.assertEqual(pr_analyze['upload'], 'never')
        self.assertIs(pr_analyze['upload-database'], False)
        artifact = jobs['analyze-pr']['steps'][3]['with']
        self.assertEqual(artifact['path'], pr_analyze['output'])
        self.assertEqual(artifact['if-no-files-found'], 'error')

    def test_integrity_runs_every_check_with_its_context(self):
        name, job_id = REQUIRED_JOB
        steps = workflows()[name]['jobs'][job_id]['steps']
        step = next(step for step in steps if step.get('name') == 'Run the repository checks')
        self.assertEqual(step['env'], {
            'IS_TEMPLATE': '${{ github.event.repository.is_template }}',
            'DEFAULT_BRANCH': '${{ github.event.repository.default_branch }}',
            'PR_HEAD_SHA': '${{ github.event.pull_request.head.sha }}'})
        # The default branch lets the run check that CodeQL covers it.
        # --ci fails a working tree that an earlier step or a check changed.
        for argument in ('tools/check_repository.py --strict', '--ci',
                         '--repository "$GITHUB_REPOSITORY"',
                         '--default-branch "$DEFAULT_BRANCH"', '--template "$IS_TEMPLATE"',
                         '--summary "$GITHUB_STEP_SUMMARY"'):
            self.assertIn(argument, step['run'])
        # The parser, the quality tools and actionlint are installed by earlier steps.
        position = steps.index(step)
        for earlier in ('Install pinned configuration parser',
                        'Install hash-locked Python quality tools',
                        'Install checksum-verified actionlint'):
            with self.subTest(step=earlier):
                self.assertLess(next(i for i, s in enumerate(steps) if s.get('name') == earlier),
                                position)
        # The reports are kept whatever the outcome.
        upload = steps[position + 1]
        self.assertTrue(upload['uses'].startswith('actions/upload-artifact@'))
        self.assertEqual(upload['if'], 'always()')
        self.assertEqual(upload['with']['path'], 'test-results/integrity/')
        self.assertEqual(upload['with']['if-no-files-found'], 'error')
        self.assertLessEqual(int(upload['with']['retention-days']), 30)
        self.assertIn('${{ github.sha }}', upload['with']['name'])

    def test_every_pip_install_is_hash_locked(self):
        installs = 0
        for name, document in workflows().items():
            for job_id, job in document['jobs'].items():
                for step in job.get('steps', []):
                    run = step.get('run') or ''
                    if 'pip install' not in run:
                        continue
                    installs += 1
                    with self.subTest(workflow=name, job=job_id, step=step.get('name')):
                        self.assertIn('--require-hashes', run)
                        self.assertIn('--only-binary=:all:', run)
                        self.assertRegex(run, r'-r \S+requirements[\w.-]*\.txt')
        self.assertGreater(installs, 0)
        for path in ('.github/requirements.txt', 'tools/requirements-quality-ci.txt'):
            text = (ROOT / path).read_text(encoding='utf-8')
            for requirement in re.split(r'\n(?=\S)', re.sub(r'(?m)^#.*\n?', '', text).strip()):
                with self.subTest(file=path, requirement=requirement.split()[0]):
                    self.assertRegex(requirement.split()[0], r'^[\w.-]+==[\w.]+$')
                    self.assertRegex(requirement, r'--hash=sha256:[0-9a-f]{64}')

    def test_jobs_that_run_node_set_it_up_first(self):
        # The repository tests and the issue checker run the label scripts too.
        uses_node = re.compile(r'\bnode\s|unittest|check_issues\.py|check_repository\.py')
        versions, jobs = set(), 0
        for name, document in workflows().items():
            for job_id, job in document['jobs'].items():
                steps = job.get('steps', [])
                runs = [i for i, step in enumerate(steps) if uses_node.search(step.get('run') or '')]
                if not runs:
                    continue
                jobs += 1
                setups = [i for i, step in enumerate(steps)
                          if str(step.get('uses', '')).startswith('actions/setup-node@')]
                with self.subTest(workflow=name, job=job_id):
                    self.assertTrue(setups and setups[0] < runs[0], 'set up Node.js before using it')
                    versions.add(str(steps[setups[0]]['with']['node-version']))
        self.assertGreater(jobs, 0)
        self.assertEqual(len(versions), 1, f'one Node.js version everywhere, found {versions}')

    def test_toolchain_versions_are_exact_and_shared(self):
        # Floating versions let a new patch release change CI without a commit.
        exact = re.compile(r'^\d+\.\d+\.\d+$')
        minimum_job = MINIMUM_JOB
        floor = re.search(r'^python_version = "(\d+\.\d+)"$',
                          (ROOT / 'pyproject.toml').read_text(encoding='utf-8'), re.M)
        assert floor, 'pyproject.toml sets no python_version'
        minimum = floor.group(1)
        python, node = set(), set()
        for name, document in workflows().items():
            for job_id, job in document['jobs'].items():
                for step in job.get('steps', []):
                    uses, settings = str(step.get('uses', '')), step.get('with') or {}
                    with self.subTest(workflow=name, job=job_id, uses=uses):
                        if uses.startswith('actions/setup-python@'):
                            version = str(settings['python-version'])
                            self.assertRegex(version, exact)
                            if (name, job_id) == minimum_job:
                                self.assertTrue(version.startswith(minimum + '.'),
                                                f'the minimum-version job runs {minimum}.x')
                            else:
                                python.add(version)
                        if uses.startswith('actions/setup-node@'):
                            self.assertRegex(str(settings['node-version']), exact)
                            self.assertIs(settings.get('package-manager-cache'), False)
                            node.add(str(settings['node-version']))
        self.assertEqual(len(python), 1, f'one Python version outside the minimum job, found {python}')
        self.assertEqual(len(node), 1, f'one Node.js version everywhere, found {node}')

    def test_minimum_python_runs_apart_from_the_required_check(self):
        documents = workflows()
        # Repository integrity holds the required job alone.
        self.assertEqual(list(documents[REQUIRED_JOB[0]]['jobs']), [REQUIRED_JOB[1]])
        name, job_id = MINIMUM_JOB
        document = documents[name]
        triggers = document['on']
        self.assertEqual(set(triggers), {'pull_request', 'push', 'workflow_dispatch'})
        # Only a pull request that changes what Python runs, or its requirements.
        self.assertEqual(triggers['pull_request'], {'paths': [
            '**/*.py', 'tests/**', 'tools/**', '.github/scripts/**', 'pyproject.toml',
            '.github/requirements.txt', '.github/workflows/minimum-python.yml']})
        # Pushes to the default branch only, listed by name; Repository integrity
        # checks that it is present.
        self.assertEqual(set(triggers['push']), {'branches'})
        branches = triggers['push']['branches']
        self.assertEqual(len(branches), 1, branches)
        self.assertNotIn('*', branches[0])
        self.assertEqual(document['permissions'], {'contents': 'read'})
        self.assertEqual(list(document['jobs']), [job_id])
        job = document['jobs'][job_id]
        self.assertEqual(job['name'], 'Minimum Python (3.11)')
        self.assertNotEqual(job['name'], REQUIRED_CHECK)
        self.assertNotIn('if', job)
        self.assertNotIn('permissions', job)
        self.assertEqual(job['timeout-minutes'], 5)
        self.assertEqual(job['steps'][-1]['run'], 'python -m unittest discover -s tests')
        concurrency = document['concurrency']
        self.assertIn('github.event.pull_request.number', concurrency['group'])
        self.assertIn('github.run_id', concurrency['group'])
        self.assertEqual(concurrency['cancel-in-progress'],
                         "${{ github.event_name == 'pull_request' }}")

    def test_dependabot_covers_each_pinned_dependency_once(self):
        document = strict_yaml.load((ROOT / '.github/dependabot.yml').read_text(encoding='utf-8'))
        self.assertEqual(document['version'], 2)
        updates = document['updates']
        # One block per ecosystem: a week with updates in both pip directories
        # still opens a single pull request.
        self.assertEqual(sorted(u['package-ecosystem'] for u in updates), ['github-actions', 'pip'])
        covered = sorted(directory for u in updates
                         for directory in u.get('directories', [u.get('directory')]))
        self.assertEqual(covered, ['/', '/.github', '/tools'])
        # Every tracked requirements file sits in a directory a pip block covers.
        tracked = subprocess.run(['git', 'ls-files', '*requirements*.txt'], cwd=ROOT, check=True,
                                 capture_output=True, text=True).stdout.split()
        self.assertTrue(tracked, 'no tracked requirements file found')
        for name in tracked:
            self.assertIn('/' + str(Path(name).parent), covered, name)
        for update in updates:
            with self.subTest(ecosystem=update['package-ecosystem']):
                self.assertNotIn('milestone', update)
                self.assertEqual([label for label in update['labels'] if label in {'P1', 'P2', 'P3'}], ['P3'])
                # One weekly pull request per ecosystem, at a stated local time.
                schedule = update['schedule']
                self.assertEqual(schedule['interval'], 'weekly')
                self.assertTrue({'day', 'time', 'timezone'} <= set(schedule), schedule)
                self.assertIsInstance(update.get('open-pull-requests-limit'), int)
                self.assertTrue(update.get('commit-message', {}).get('prefix'))
                # A version update waits at least 14 days after its release, so a
                # package that is withdrawn within days is never proposed.
                self.assertGreaterEqual(update.get('cooldown', {}).get('default-days', 0), 14)
                # A catch-all group keeps related pins, such as CodeQL's init
                # and analyze steps, in a single pull request.
                self.assertTrue(any('*' in group.get('patterns', [])
                                    for group in (update.get('groups') or {}).values()),
                                'needs a group with pattern "*"')

    def test_codeql_action_steps_share_one_pin(self):
        pins = {step['uses'].split('@', 1)[1]
                for document in workflows().values()
                for job in document['jobs'].values()
                for step in job.get('steps', [])
                if str(step.get('uses', '')).startswith('github/codeql-action/')}
        self.assertTrue(pins, 'no github/codeql-action steps found')
        self.assertEqual(len(pins), 1, f'CodeQL action steps use different pins: {pins}')


def milestone_script():
    run = workflows()['pr-milestones.yml']['jobs']['reconcile']['steps'][0]['run']
    return run.split("<<'PY'\n", 1)[1].rsplit('PY', 1)[0]


class FakeGitHub:
    def __init__(self, pulls, issues=(), *, read_only=False, sticky=False, unpersisted=False):
        self.items = {pull['number']: dict(pull, pull_request={}) for pull in pulls}
        self.items.update({issue['number']: dict(issue) for issue in issues})
        self.patches = []
        self.requests = []
        self.read_only = read_only  # every PATCH is refused with HTTP 403
        self.sticky = sticky  # a PATCH succeeds but the milestone stays
        # A PATCH whose response reports the change, which a later read does not show.
        self.unpersisted = unpersisted

    def __call__(self, method, path, data=None):
        self.requests.append((method, path))
        if path.startswith('/repos/o/r/pulls?'):
            match = re.search(r'[?&]page=(\d+)', path)
            assert match, path
            page = int(match.group(1))
            pulls = [item for item in self.items.values() if 'pull_request' in item]
            return pulls[(page - 1) * 100:page * 100]
        number = int(path.rsplit('/', 1)[1])
        if method == 'PATCH':
            if self.read_only:
                raise HTTPError(path, 403, 'Resource not accessible by integration', Message(), None)
            self.patches.append(number)
            if self.unpersisted:
                return dict(self.items[number], milestone=data['milestone'])
            if not self.sticky:
                self.items[number]['milestone'] = data['milestone']
        return dict(self.items[number])


@unittest.skipIf(yaml is None and not os.environ.get('GITHUB_ACTIONS'),
                 'PyYAML is not installed: python -m pip install -r .github/requirements.txt')
class MilestoneScriptTests(unittest.TestCase):
    def script(self):
        namespace = {'__name__': 'milestones'}
        exec(milestone_script(), namespace)  # trusted: the script under test, from this repository
        return namespace

    def reconcile(self):
        """The script's reconcile, with the line it prints for each pull request
        captured, so that the test log stays readable."""
        reconcile = self.script()['reconcile']

        def quiet(*args):
            with contextlib.redirect_stdout(io.StringIO()):
                return reconcile(*args)
        return quiet

    def run_job(self, api, event):
        """Run the job's logic; return the job summary and what it printed."""
        published: list[str] = []
        with unittest.mock.patch('sys.stdout', new_callable=io.StringIO) as out:
            self.script()['run'](api, 'o/r', event, published.append)
        return ''.join(published), out.getvalue()

    def test_pull_request_event_clears_only_that_pull_request(self):
        api = FakeGitHub([{'number': 1, 'milestone': {'title': 'v1'}},
                          {'number': 2, 'milestone': {'title': 'v1'}}])
        self.reconcile()(api, 'o/r', {'number': 1, 'pull_request': {}})
        self.assertEqual(api.patches, [1])
        self.assertIsNotNone(api.items[2]['milestone'])

    def test_scan_clears_every_pull_request_and_never_touches_issues(self):
        pulls = [{'number': n, 'milestone': {'title': 'v1'} if n % 2 else None} for n in range(1, 151)]
        api = FakeGitHub(pulls, issues=[{'number': 500, 'milestone': {'title': 'v1'}}])
        self.reconcile()(api, 'o/r', {'schedule': '17 3 * * *'})
        self.assertEqual(api.patches, list(range(1, 151, 2)))
        self.assertIsNotNone(api.items[500]['milestone'])

    def test_an_issue_number_is_refused(self):
        api = FakeGitHub([], issues=[{'number': 7, 'milestone': {'title': 'v1'}}])
        with self.assertRaisesRegex(ValueError, 'Refusing to change an issue'):
            self.reconcile()(api, 'o/r', {'number': 7, 'pull_request': {}})
        self.assertEqual(api.patches, [])

    def test_a_pull_request_without_a_milestone_is_left_alone(self):
        api = FakeGitHub([{'number': 3, 'milestone': None}])
        self.assertEqual(self.reconcile()(api, 'o/r', {'number': 3, 'pull_request': {}}), [])
        self.assertEqual(api.patches, [])

    def test_a_milestone_that_survives_the_change_fails(self):
        api = FakeGitHub([{'number': 4, 'milestone': {'title': 'v1'}}], sticky=True)
        with self.assertRaisesRegex(ValueError, 'PR #4 still has a milestone'):
            self.reconcile()(api, 'o/r', {'number': 4, 'pull_request': {}})

    def test_a_removal_is_read_back_after_the_change(self):
        api = FakeGitHub([{'number': 8, 'milestone': {'title': 'v1'}}])
        self.assertEqual(self.reconcile()(api, 'o/r', {'number': 8, 'pull_request': {}}), [(8, 'v1')])
        self.assertEqual(api.requests, [('GET', '/repos/o/r/issues/8'), ('PATCH', '/repos/o/r/issues/8'),
                                        ('GET', '/repos/o/r/issues/8')])

    def test_a_change_the_response_reports_but_github_does_not_keep_fails(self):
        api = FakeGitHub([{'number': 9, 'milestone': {'title': 'v1'}}], unpersisted=True)
        published: list[str] = []
        with self.assertRaisesRegex(ValueError, 'PR #9 still has a milestone after the change'):
            self.script()['run'](api, 'o/r', {'number': 9, 'pull_request': {}}, published.append)
        self.assertEqual(published, [], 'no removal may be recorded')

    def test_the_summary_lists_each_removal_with_its_scope(self):
        api = FakeGitHub([{'number': 1, 'milestone': {'title': 'v1 | first\nrelease'}},
                          {'number': 2, 'milestone': None},
                          {'number': 3, 'milestone': {'title': 'v2'}}])
        summary, printed = self.run_job(api, {'schedule': '17 3 * * *'})
        self.assertEqual(api.patches, [1, 3])
        self.assertIn('Scope: every pull request.', summary)
        self.assertIn('| #1 | v1 \\| first release |', summary)
        self.assertIn('| #3 | v2 |', summary)
        self.assertIn('Removed 2 pull request milestone(s); scope: every pull request', printed)

    def test_a_run_with_nothing_to_remove_says_so(self):
        summary, _ = self.run_job(FakeGitHub([{'number': 5, 'milestone': None}]),
                                  {'number': 5, 'pull_request': {}})
        self.assertIn('Scope: pull request #5.', summary)
        self.assertIn('nothing was removed', summary)

    def test_a_read_only_token_on_a_pull_request_event_warns(self):
        api = FakeGitHub([{'number': 6, 'milestone': {'title': 'v1'}}], read_only=True)
        summary, printed = self.run_job(api, {'number': 6, 'pull_request': {}})
        self.assertIn('::warning title=Milestone not removed::', printed)
        self.assertIn('HTTP 403', summary)
        self.assertIn('weekly scan or a manual run', summary)
        self.assertIsNotNone(api.items[6]['milestone'])

    def test_a_read_only_token_fails_a_scan(self):
        api = FakeGitHub([{'number': 6, 'milestone': {'title': 'v1'}}], read_only=True)
        with self.assertRaises(HTTPError):
            self.run_job(api, {'schedule': '37 3 * * 5'})


def traffic_alert_script():
    steps = workflows()['traffic-history.yml']['jobs']['alert']['steps']
    return next(step for step in steps if step['name'] == 'Open the alert issue, if one is due')['run']


# A gh that answers the milestone lookup with real jq over a fixed list, as
# gh --jq does, finds no existing issue, and records every call.
FAKE_GH = """#!/usr/bin/env bash
set -euo pipefail
printf '%s\\n' "$*" >> "$GH_CALLS"
if [ "$1" = api ]; then
  while [ "$1" != --jq ]; do shift; done
  jq -r "$2" "$MILESTONES"
elif [ "$1 $2" = "issue list" ]; then
  echo 0
fi
"""
SNAPSHOT_HEADER = ('snapshot_date,collected_at,observed_at,first_date,last_date,views_14d,'
                   'unique_views_14d,clones_14d,unique_clones_14d,stars,forks,watchers,open_issues\n')


@unittest.skipUnless((WORKFLOWS / 'traffic-history.yml').is_file(),
                     '.github/workflows/traffic-history.yml does not exist')
@unittest.skipIf(yaml is None and not os.environ.get('GITHUB_ACTIONS'),
                 'PyYAML is not installed: python -m pip install -r .github/requirements.txt')
@unittest.skipIf((shutil.which('bash') is None or shutil.which('jq') is None)
                 and not os.environ.get('GITHUB_ACTIONS'), 'bash and jq are needed')
class TrafficAlertScriptTests(unittest.TestCase):
    """The alert step, run with bash against a fake gh and a collected history."""

    def run_step(self, latest_views, milestones):
        """Two snapshots, 2026-10-13 and 2026-10-14; return the result and the gh calls."""
        with tempfile.TemporaryDirectory() as directory:
            root = Path(directory)
            (root / 'source').symlink_to(ROOT, target_is_directory=True)
            data = root / 'history/data'
            data.mkdir(parents=True)
            rows = ''.join(f'{day},{day}T06:00:00Z,{day}T06:00:00Z,{day},{day},{views},1,1,1,1,0,1,4\n'
                           for day, views in (('2026-10-13', 50), ('2026-10-14', latest_views)))
            (data / 'snapshots.csv').write_text(SNAPSHOT_HEADER + rows, encoding='utf-8')
            bin_dir = root / 'bin'
            bin_dir.mkdir()
            (bin_dir / 'gh').write_text(FAKE_GH, encoding='utf-8')
            (bin_dir / 'gh').chmod(0o755)
            (bin_dir / 'python').symlink_to(sys.executable)
            (root / 'milestones.json').write_text(json.dumps(milestones), encoding='utf-8')
            (root / 'temp').mkdir()
            (root / 'step.sh').write_text(traffic_alert_script(), encoding='utf-8')
            env = dict(os.environ, PATH=f'{bin_dir}{os.pathsep}{os.environ["PATH"]}',
                       RUNNER_TEMP=str(root / 'temp'), GH_TOKEN='test-token', GH_REPO='o/r',
                       GH_CALLS=str(root / 'calls.txt'), MILESTONES=str(root / 'milestones.json'),
                       PYTHONDONTWRITEBYTECODE='1')
            result = subprocess.run(['bash', '--noprofile', '--norc', '-eo', 'pipefail', 'step.sh'],
                                    cwd=root, env=env, capture_output=True, text=True,
                                    encoding='utf-8', timeout=60, check=False)
            calls_file = root / 'calls.txt'
            calls = calls_file.read_text(encoding='utf-8').splitlines() if calls_file.exists() else []
        return result, calls

    def test_a_run_with_no_alert_due_needs_no_milestone(self):
        result, calls = self.run_step(50, [])
        self.assertEqual(result.returncode, 0, result.stdout + result.stderr)
        self.assertIn('No alert', result.stdout)
        self.assertEqual(calls, [], 'no milestone lookup and no issue without an alert')

    def test_the_milestone_is_due_after_the_snapshot_date_not_the_job_date(self):
        # Both milestones are overdue by the job's clock; only the snapshot date counts.
        milestones = [{'title': 'Same day', 'due_on': '2026-10-14T07:00:00Z'},
                      {'title': 'Next day', 'due_on': '2026-10-15T07:00:00Z'}]
        result, calls = self.run_step(500, milestones)
        self.assertEqual(result.returncode, 0, result.stdout + result.stderr)
        self.assertTrue(calls[0].startswith('api repos/o/r/milestones?state=open'), calls)
        created = [call for call in calls if call.startswith('issue create')]
        self.assertEqual(len(created), 1, calls)
        self.assertIn('--title Traffic alert: 2026-10-14', created[0])
        self.assertTrue(created[0].endswith('--milestone Next day'), created[0])

    def test_an_alert_with_no_milestone_after_the_snapshot_date_fails(self):
        result, calls = self.run_step(500, [{'title': 'Same day', 'due_on': '2026-10-14T07:00:00Z'},
                                            {'title': 'Undated', 'due_on': None}])
        self.assertEqual(result.returncode, 1)
        self.assertIn('::error::No open milestone due after the collection date 2026-10-14',
                      result.stdout)
        self.assertFalse([call for call in calls if call.startswith('issue')], calls)


if __name__ == '__main__':
    unittest.main()
