#!/usr/bin/env python3
"""Check reviewed wiki sources, render a pinned bundle, or compare a wiki checkout."""
from __future__ import annotations

import argparse
import hashlib
import json
import re
from pathlib import Path, PurePosixPath
from typing import Any

from _gatelib import git_text, run_gate, tracked_files
from initialize_repository import _build_changes, _fixture_arguments

WIKI = 'docs/wiki'
CATALOGUE = f'{WIKI}/catalogue.json'
INVENTORY = 'File-and-directory-reference.md'
NOTICE = '> **Guide, not policy:**'


def require(condition: Any, message: str) -> None:
    if not condition:
        raise ValueError(message)


def sha256(data: bytes) -> str:
    return hashlib.sha256(data).hexdigest()


def read_json(path: Path) -> dict[str, Any]:
    def unique(pairs: list[tuple[str, Any]]) -> dict[str, Any]:
        result: dict[str, Any] = {}
        for key, value in pairs:
            require(key not in result, f'duplicate JSON key: {key}')
            result[key] = value
        return result
    value = json.loads(path.read_text(encoding='utf-8'), object_pairs_hook=unique)
    require(isinstance(value, dict), 'JSON root must be an object')
    return value


def directories(files: set[str]) -> set[str]:
    return {str(parent) for file in files for parent in PurePosixPath(file).parents
            if str(parent) != '.'}


def lifecycle(root: Path, files: set[str]) -> tuple[dict[str, list[str]], dict[str, list[str]]]:
    """Use the actual dry-run renderer, not a second implementation of initialization."""
    matrix: dict[str, list[str]] = {p: [] for p in files}
    generated: dict[str, list[str]] = {}
    for profile in ('application', 'library', 'ui-component'):
        scalars, repeats = _fixture_arguments(profile)
        changes, _ = _build_changes(root, profile, scalars, repeats)
        for path in files:
            state = 'R' if path not in changes else 'X' if changes[path] is None else 'T'
            matrix[path].append(state)
        for path, content in changes.items():
            if path not in files and content is not None:
                generated.setdefault(path, []).append(profile)
    return matrix, generated


def validate_catalogue(root: Path, data: dict[str, Any], files: set[str]) -> None:
    require(data.get('schema_version') == 1, 'unsupported wiki catalogue')
    require(set(data['files']) == files, 'file inventory differs from git ls-files: '
            + ', '.join(sorted(set(data['files']) ^ files)))
    require(set(data['directories']) == directories(files), 'directory inventory differs from tracked tree')
    pages = data['pages']
    require(isinstance(pages, list) and pages, 'page order is required')
    names = [page['name'] for page in pages]
    require(len(set(names)) == len(names) and names[0] == 'Home', 'duplicate pages or missing Home first')
    require(all(re.fullmatch(r'[A-Z][A-Za-z0-9-]*', name) for name in names), 'unsafe page name')
    expected = {f'{WIKI}/{name}.md' for name in names} | {f'{WIKI}/_Sidebar.md'}
    actual = {p for p in files if p.startswith(WIKI + '/') and p.endswith('.md')}
    require(expected == actual, 'page registry differs from source Markdown set')
    for page in pages:
        require(page['authority'] in files, f'missing page authority: {page["name"]}')
        text = (root / WIKI / (page['name'] + '.md')).read_text(encoding='utf-8')
        require(NOTICE in text, f'missing authority notice: {page["name"]}')
        notice = next(line for line in text.splitlines() if NOTICE in line)
        targets = re.findall(r'\]\(([^\s)]+)\)', notice)
        require(any((root / WIKI / target.split('#')[0]).resolve()
                    == (root / page['authority']).resolve() for target in targets),
                f'authority notice does not link its registered contract: {page["name"]}')
    for path, row in {**data['files'], **data['directories']}.items():
        require(row['kind'] in data['kinds'], f'unknown kind: {path}')
        require(row['authority'] in files, f'missing inventory authority: {path}')
        require(isinstance(row['purpose'], str) and row['purpose'].strip(), f'missing purpose: {path}')
    for kind, row in data['kinds'].items():
        require(all(isinstance(row.get(key), str) and row[key].strip()
                    for key in ('owner', 'edit', 'guard', 'validation')), f'incomplete kind: {kind}')


def sidebar(data: dict[str, Any]) -> str:
    return ('# 🧭 Maintainer journey\n\n'
            + '\n'.join(f'- [{i:02d} · {p["title"]}]({p["name"]}.md)'
                        for i, p in enumerate(data['pages']))
            + '\n\n' + NOTICE + ' [Repository authorities](../../docs/README.md) govern.\n')


def inventory(root: Path, data: dict[str, Any], files: set[str]) -> str:
    matrix, generated = lifecycle(root, files)
    out = ['# 🗂️ File and directory reference', '',
           NOTICE + ' [Repository Structure](../REPOSITORY_STRUCTURE.md) and the linked authorities govern.', '',
           'This inventory is generated from the reviewed catalogue and checked against `git ls-files`.',
           'Each tracked file and each ancestor directory appears once. Group rules apply to every row',
           'of that kind; the row supplies its specific purpose and authoritative document.', '',
           '**Lifecycle order: application / library / UI component.** R = retained byte-for-byte;',
           'T = transformed; X = removed. Values come from the actual initializer dry run with its',
           'standard synthetic values. Supplying optional/repeatable values changes rendered text.',
           'The optional `assets/social-preview.png` is retained when selected; otherwise removed.',
           'An untracked directory is not a Git artifact. Directory retention follows its children.', '',
           '## Editing and validation rules by kind', '',
           '| Kind / owner | When to edit | Invariants and unsafe edits | Validation |',
           '| --- | --- | --- | --- |']
    for kind, row in sorted(data['kinds'].items()):
        out.append(f'| {kind} / {row["owner"]} | {row["edit"]} | {row["guard"]} | {row["validation"]} |')
    out += ['', '## Tracked files', '', '| Path / kind | Purpose | A / L / UI | Authority |',
            '| --- | --- | --- | --- |']
    for path, row in sorted(data['files'].items()):
        status = ' / '.join(matrix[path])
        out.append(f'| `{path}` / {row["kind"]} | {row["purpose"]} | {status} | '
                   f'[{row["authority"]}](../../{row["authority"]}) |')
    out += ['', '## Tracked directories', '', '| Directory / kind | Purpose | A / L / UI | Authority |',
            '| --- | --- | --- | --- |']
    for path, row in sorted(data['directories'].items()):
        states = ['R' if any(matrix[p][i] != 'X' for p in files if p.startswith(path + '/')) else 'X'
                  for i in range(3)]
        out.append(f'| `{path}/` / {row["kind"]} | {row["purpose"]} | {" / ".join(states)} | '
                   f'[{row["authority"]}](../../{row["authority"]}) |')
    out += ['', '## Files created only during initialization', '',
            '| Generated path | Profiles | Purpose and editing rule |', '| --- | --- | --- |']
    for path, profiles in sorted(generated.items()):
        purpose = ('Immutable initialization inputs; do not hand-edit to conceal a mismatch.'
                   if path.endswith('initialization.json') else
                   'Directory guidance; replace with applicable exported source when that role is implemented.')
        out.append(f'| `{path}` | {", ".join(profiles)} | {purpose} |')
    return '\n'.join(out) + '\n'


def render_page(text: str, page: str, names: set[str], repository: str, sha: str) -> str:
    """Translate reviewed relative Markdown links; keep remote links and anchors."""
    def link(match: re.Match[str]) -> str:
        target = match[1]
        if target.startswith(('https://', 'http://', 'mailto:', '#')):
            return match[0]
        path, separator, fragment = target.partition('#')
        if path.endswith('.md') and path[:-3] in names:
            return '](' + path[:-3] + (separator + fragment if separator else '') + ')'
        resolved = PurePosixPath(WIKI) / path
        parts: list[str] = []
        for part in resolved.parts:
            if part == '..':
                require(bool(parts), 'link escapes repository')
                parts.pop()
            elif part != '.':
                parts.append(part)
        return f'](https://github.com/{repository}/blob/{sha}/{"/".join(parts)}' + (separator + fragment if separator else '') + ')'
    rendered = re.sub(r'\]\(([^\s)]+)\)', link, text)
    return rendered.rstrip() + f'\n\n---\nPublished from [{sha}](https://github.com/{repository}/blob/{sha}/{WIKI}/{page}).\n'


def bundle(root: Path, data: dict[str, Any], sha: str) -> dict[str, bytes]:
    require(re.fullmatch(r'[0-9a-f]{40}', sha), 'source SHA must be a full commit')
    repository = read_json(root / '.github/repository-profile.json')['repository']
    require(re.fullmatch(r'[A-Za-z0-9_.-]+/[A-Za-z0-9_.-]+', repository), 'invalid repository')
    names = {page['name'] for page in data['pages']}
    result = {}
    for name in sorted(names | {'_Sidebar'}):
        page = name + '.md'
        source = (root / WIKI / page).read_text(encoding='utf-8')
        result[page] = render_page(source, page, names, repository, sha).encode('utf-8')
    record = {'schema_version': 1, 'repository': repository, 'source_sha': sha,
              'pages': {p: sha256(content) for p, content in result.items()}}
    result['Wiki-Source.json'] = (json.dumps(record, indent=2, sort_keys=True) + '\n').encode()
    return result


def compare_bundle(expected: dict[str, bytes], destination: Path) -> list[str]:
    require(destination.is_dir() and not destination.is_symlink(), 'published directory unavailable')
    paths = [p for p in destination.rglob('*') if '.git' not in p.relative_to(destination).parts]
    require(not any(p.is_symlink() for p in paths), 'published directory contains a symlink')
    actual = {p.relative_to(destination).as_posix() for p in paths if p.is_file()}
    differences = [f'published path set differs: {p}' for p in sorted(actual ^ set(expected))]
    differences += [f'published bytes differ: {p}' for p in sorted(actual & set(expected))
                    if (destination / p).read_bytes() != expected[p]]
    return differences


def write_bundle(expected: dict[str, bytes], destination: Path, root: Path) -> None:
    require(not destination.exists(), 'export destination must not exist; no files are overwritten')
    require(not destination.resolve().is_relative_to(root.resolve()), 'export must be outside source repository')
    destination.mkdir(parents=True)
    for path, content in expected.items():
        (destination / path).write_bytes(content)


def build_report(args: argparse.Namespace) -> dict[str, Any]:
    root = args.root.resolve()
    findings: list[str] = []
    try:
        files = tracked_files(root)
        data = read_json(root / CATALOGUE)
        validate_catalogue(root, data, files)
        targets = {INVENTORY: inventory(root, data, files), '_Sidebar.md': sidebar(data)}
        for path, expected in targets.items():
            target = root / WIKI / path
            if args.write_reference:
                target.write_text(expected, encoding='utf-8', newline='\n')
            elif target.read_text(encoding='utf-8') != expected:
                findings.append(f'{path}: stale generated reference; use --write-reference')
        if args.export_dir or args.published_dir:
            require(not findings, 'fix deterministic source findings before publication comparison')
            sha = git_text(root, 'rev-parse', 'HEAD', check=True).stdout.strip()
            require(not git_text(root, 'status', '--porcelain', check=True).stdout,
                    'publication requires a clean committed source tree')
            expected_bundle = bundle(root, data, sha)
            if args.export_dir:
                write_bundle(expected_bundle, args.export_dir, root)
            if args.published_dir:
                findings.extend(compare_bundle(expected_bundle, args.published_dir))
    except (ValueError, OSError, KeyError, TypeError, RuntimeError) as error:
        findings.append(str(error))
    return {'status': 'fail' if findings else 'pass', 'findings': sorted(findings),
            'scope_note': 'Offline source/inventory/navigation and exact-source wiki-byte checks. '
                          'No GitHub writes, network availability claim, Excel execution or release certification.'}


def markdown(report: dict[str, Any]) -> str:
    return '# Wiki contract\n\nResult: ' + report['status'].upper() + '\n\n' + '\n'.join(
        '- ' + f for f in report['findings']) + '\n\n' + report['scope_note'] + '\n'


def main() -> int:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument('--root', type=Path, default=Path.cwd())
    parser.add_argument('--output', type=Path)
    parser.add_argument('--summary', type=Path)
    mode = parser.add_mutually_exclusive_group()
    mode.add_argument('--write-reference', action='store_true')
    mode.add_argument('--export-dir', type=Path)
    mode.add_argument('--published-dir', type=Path)
    args = parser.parse_args()
    return run_gate(args, build=lambda: build_report(args), markdown=markdown, errors=(OSError,))


if __name__ == '__main__':
    raise SystemExit(main())
