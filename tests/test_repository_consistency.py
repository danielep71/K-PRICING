"""Cross-file contracts: label catalogue, editor and Git byte rules, Dependabot, Markdown, API
versions and template-only names."""

# mypy checks the bodies of these tests but does not require them to be annotated.
# mypy: disallow-untyped-defs=False, disallow-incomplete-defs=False, disallow-untyped-calls=False

import json
import re
import subprocess
import unittest
from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]


def catalogue_labels():
    """The resolved label catalogue: the core plus the selected overlays, sorted by name."""
    manifest = json.loads((ROOT / '.github/labels.json').read_text(encoding='utf-8'))
    labels = list(manifest['core'])
    for name in manifest['selected_overlays']:
        labels += manifest['overlays'][name]
    return sorted(labels, key=lambda label: (label['name'].lower(), label['name']))


def tracked(*patterns):
    """Files Git tracks, so ignored directories such as .venv/ are never scanned."""
    try:
        names = subprocess.run(['git', '-C', str(ROOT), 'ls-files', '-z', '--', *patterns],
                               capture_output=True, text=True, encoding='utf-8',
                               check=True).stdout.split('\0')
    except (OSError, subprocess.CalledProcessError):
        names = [path.relative_to(ROOT).as_posix() for pattern in patterns
                 for path in ROOT.glob(f'**/{pattern}') if '.git' not in path.parts]
    return sorted(ROOT / name for name in names if name and (ROOT / name).is_file())


class CatalogueDocumentationTests(unittest.TestCase):
    def test_guide_matches_label_names_colors_and_descriptions(self):
        guide = (ROOT / 'docs/repo_governance/ISSUE_LABELS.md').read_text(encoding='utf-8')
        rows = re.findall(r'^\| `([^`]+)` \| `#([0-9A-F]+)` \| (.+) \|$', guide, re.M)
        expected = [(label['name'], label['color'], label['description'])
                    for label in catalogue_labels()]
        self.assertEqual(rows, expected)

    def test_usage_guide_covers_every_label_once(self):
        guide = (ROOT / 'docs/repo_governance/ISSUE_LABELS.md').read_text(encoding='utf-8')
        section = re.search(r'^## 🧭 Using the labels$(.*?)^## ', guide, re.M | re.S)
        assert section, 'ISSUE_LABELS.md has no "Using the labels" section'
        names = re.findall(r'^\| `([^`]+)` \|', section.group(1), re.M)
        self.assertEqual(sorted(names, key=str.lower),
                         sorted((label['name'] for label in catalogue_labels()), key=str.lower))

    def test_documented_label_counts_match_the_catalogue(self):
        expected = str(len(catalogue_labels()))
        pattern = re.compile(r'labels-(\d+)-|\b(\d+) (?:issue )?labels\b', re.I)
        counts = []
        for path in tracked('*.md'):
            text = path.read_text(encoding='utf-8')
            counts += [(path.relative_to(ROOT).as_posix(), badge or prose)
                       for badge, prose in pattern.findall(text)]
        self.assertTrue(counts, 'no documented label counts found')
        self.assertEqual([item for item in counts if item[1] != expected], [])


def yaml_list(text, key):
    """Return the block-list items under ``key:`` in a simple YAML file."""
    items, indent = [], None
    for line in text.splitlines():
        stripped = line.strip()
        if indent is None:
            if stripped == f'{key}:':
                indent = len(line) - len(line.lstrip())
            continue
        if stripped.startswith('- ') and len(line) - len(line.lstrip()) > indent:
            items.append(stripped[2:].strip().strip('"\''))
        elif stripped:
            indent = None
    return items


class IgnoreRuleTests(unittest.TestCase):
    def test_regulatory_source_copies_are_never_ignored_for_their_office_type(self):
        text = (ROOT / '.gitignore').read_text(encoding='utf-8')
        section = text[text.index('# GENERATED OFFICE PACKAGES'):
                       text.index('# Reserve one exact path')]
        types = [line.strip()[1:] for line in section.splitlines() if line.startswith('*.')]
        self.assertIn('.xltx', types)
        for suffix in [*types, '.log']:
            with self.subTest(suffix=suffix):
                def ignored(path):
                    return subprocess.run(['git', '-C', str(ROOT), 'check-ignore', '-q',
                                           '--no-index', path], check=False).returncode == 0
                self.assertTrue(ignored(f'build/source{suffix}'))
                self.assertFalse(ignored(f'docs/methodology/sources/eba/source{suffix}'))


def editorconfig_sections():
    """The settings of each .editorconfig section, by section glob."""
    sections: dict[str, list[str]] = {}
    current = None
    for line in (ROOT / '.editorconfig').read_text(encoding='utf-8').splitlines():
        line = line.strip()
        if line.startswith('[') and line.endswith(']'):
            current = sections.setdefault(line[1:-1], [])
        elif current is not None and line and not line.startswith('#'):
            current.append(line)
    return sections


def git_attributes(path, *names):
    """The values .gitattributes gives ``path`` for ``names``."""
    printed = subprocess.run(['git', '-C', str(ROOT), 'check-attr', *names, '--', path],
                             capture_output=True, text=True, check=True).stdout
    return dict(line.split(': ')[1:] for line in printed.splitlines())


class EditorAndGitByteRuleTests(unittest.TestCase):
    # git diff --check, run by Repository integrity, reports what editors keep.
    def test_editors_trim_the_trailing_white_space_git_reports_in_markdown(self):
        sections = editorconfig_sections()
        for section, path in (('*.md', 'docs/notes.md'),
                              ('docs/methodology/sources/README.md',
                               'docs/methodology/sources/README.md')):
            with self.subTest(section=section):
                self.assertIn('trim_trailing_whitespace = true', sections[section])
                self.assertEqual(git_attributes(path, 'whitespace'),
                                 {'whitespace': 'unspecified'})

    def test_byte_preserved_files_are_left_alone_by_editors_and_git(self):
        sections = editorconfig_sections()
        unset = ['charset = unset', 'end_of_line = unset', 'insert_final_newline = unset',
                 'trim_trailing_whitespace = false']
        self.assertEqual(sections['*.reg'],
                         [*unset, 'indent_style = unset', 'indent_size = unset',
                          'tab_width = unset'])
        for section, path in (('docs/methodology/sources/**',
                               'docs/methodology/sources/SRC-01_rules.csv'),
                              ('tests/fixtures/*/*/**', 'tests/fixtures/a-case/a-variant/a.csv'),
                              ('*.reg', 'tools/settings.reg')):
            with self.subTest(section=section):
                self.assertEqual(sections[section][:4], unset)
                attributes = git_attributes(path, 'text', 'whitespace', 'diff')
                self.assertEqual(attributes['text'], 'unset')
                # A binary file has no text diff, so git diff --check skips it.
                self.assertIn('unset', (attributes['whitespace'], attributes['diff']))


class DependabotTests(unittest.TestCase):
    def test_labels_exist_and_assignee_matches_the_issue_forms(self):
        text = (ROOT / '.github/dependabot.yml').read_text(encoding='utf-8')
        catalogue = {label['name'] for label in catalogue_labels()}
        labels = yaml_list(text, 'labels')
        self.assertTrue(labels)
        self.assertEqual([label for label in labels if label not in catalogue], [])
        form = (ROOT / '.github/ISSUE_TEMPLATE/bug.yml').read_text(encoding='utf-8')
        assignee = re.search(r'^assignees:\s*\[\s*"?([^",\]]+)"?\s*\]', form, re.M)
        assert assignee, 'bug.yml has no assignees'
        self.assertEqual(set(yaml_list(text, 'assignees')), {assignee.group(1)})


class MarkdownTests(unittest.TestCase):
    # tools/check_links.py checks link targets and anchors (tests/test_check_links.py).
    def test_code_fences_close_on_a_line_of_their_own(self):
        # A closing fence followed by text does not close the block on GitHub.
        problems = []
        for path in tracked('*.md'):
            relative = path.relative_to(ROOT)
            opened = None
            for number, line in enumerate(path.read_text(encoding='utf-8').splitlines(), 1):
                if not line.lstrip().startswith('```'):
                    continue
                if opened is None:
                    opened = number
                elif line.strip() == '```':
                    opened = None
                else:
                    problems.append((relative.as_posix(), number))
            if opened is not None:
                problems.append((relative.as_posix(), opened))
        self.assertEqual(problems, [])

    def test_every_verify_published_command_names_the_tag_and_repository(self):
        # The tag is a required positional argument and --repository a
        # required option; a shorthand without the python prefix escapes
        # tools/check_doc_commands.py, so both are checked here.
        command = re.compile(r'release\.py verify-published (?:v|--)[^`\n|]*(?:\\\n[^`\n|]*)?')
        tag = re.compile(r'release\.py verify-published v(?:X\.Y\.Z|\d+\.\d+\.\d+)\s')
        problems = []
        for path in tracked('*.md'):
            if path.name == 'CHANGELOG.md':
                continue
            for found in command.finditer(path.read_text(encoding='utf-8')):
                if '--repository' not in found.group(0) or not tag.match(found.group(0)):
                    problems.append(f'{path.relative_to(ROOT).as_posix()}: {found.group(0)}')
        self.assertEqual(problems, [])


class ApiVersionTests(unittest.TestCase):
    def test_every_github_api_call_uses_the_same_version(self):
        found: dict[str, list[str]] = {}
        for path in [*ROOT.glob('tools/*.py'), *ROOT.glob('.github/workflows/*.yml'),
                     *ROOT.glob('.github/scripts/*.mjs')]:
            text = path.read_text(encoding='utf-8')
            for version in re.findall(r'(?:X-GitHub-Api-Version|API_VERSION)\W+(\d{4}-\d{2}-\d{2})', text):
                found.setdefault(version, []).append(path.relative_to(ROOT).as_posix())
        self.assertEqual(len(found), 1, found)


NOTICE_FILE = re.compile(r'^([.\w][\w./-]*\.[\w]+|\.[\w]+)$', re.M)
LICENSES = {'Mozilla Public License 2.0': 'MPL-2.0', 'MIT License': 'MIT'}


def notice_entries():
    """(repository URL, license, file) for every file of every notice section."""
    notice = (ROOT / 'docs/licenses/THIRD_PARTY_NOTICES.txt').read_text(encoding='utf-8')
    entries = set()
    for section in re.split(r'\n(?=The (?:following|target))', notice)[1:]:
        match = re.search(r'\((https://github\.com/[\w.-]+/[\w.-]+)\)', section)
        assert match, f'a notice section names no repository: {section[:60]!r}'
        url = match.group(1)
        license_id = next(code for name, code in LICENSES.items()
                          if name in ' '.join(section.split()))
        entries |= {(url, license_id, path) for path in NOTICE_FILE.findall(section)}
    return entries


class LicenseNoticeTests(unittest.TestCase):
    def test_every_file_listed_in_the_notice_exists(self):
        notice = (ROOT / 'docs/licenses/THIRD_PARTY_NOTICES.txt').read_text(encoding='utf-8')
        listed = NOTICE_FILE.findall(notice)
        self.assertTrue(listed, 'the notice lists no files')
        self.assertEqual([path for path in listed if not (ROOT / path).is_file()], [])

    def test_provenance_register_agrees_with_the_notice(self):
        register = json.loads((ROOT / 'docs/licenses/provenance.json').read_text(encoding='utf-8'))
        self.assertEqual(register['schema_version'], 1)
        entries = set()
        for source in register['sources']:
            with self.subTest(repository=source['repository']):
                self.assertEqual(set(source), {'repository', 'license', 'revision',
                                               'revision_note', 'files'})
                self.assertRegex(source['repository'], r'^https://github\.com/[\w.-]+/[\w.-]+$')
                revision = source['revision']
                self.assertTrue(revision is None or re.fullmatch(r'[0-9a-f]{40}', revision),
                                f'revision {revision!r} is not a full commit SHA or null')
                self.assertTrue(str(source['revision_note']).strip(),
                                'explain what the revision is, or why it is not recorded')
                self.assertEqual(len(source['files']), len(set(source['files'])))
                entries |= {(source['repository'], source['license'], path)
                            for path in source['files']}
        self.assertEqual(entries, notice_entries())


class TemplateOnlyContentTests(unittest.TestCase):
    def test_reference_repositories_are_named_only_where_allowed(self):
        agents = (ROOT / 'AGENTS.md').read_text(encoding='utf-8')
        block = re.search(r'<!-- template-only:start -->(.*?)<!-- template-only:end -->',
                          agents, re.S)
        if block is None:
            self.skipTest('no template-only section (not the template)')
        names = set(re.findall(r'`[\w.-]+/([\w.-]+)`', block.group(1)))
        self.assertTrue(names, 'the template-only section names no reference repository')
        outside = agents.replace(block.group(0), '')
        found = [('AGENTS.md', name) for name in sorted(names)
                 if re.search(re.escape(name), outside, re.I)]
        for path in tracked():
            relative = path.relative_to(ROOT)
            if relative.parts[:2] == ('docs', 'licenses') or relative.as_posix() == 'AGENTS.md':
                continue
            text = path.read_text(encoding='utf-8', errors='ignore')
            found += [(relative.as_posix(), name) for name in sorted(names)
                      if re.search(re.escape(name), text, re.I)]
        self.assertEqual(found, [])
