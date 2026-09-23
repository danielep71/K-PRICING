#!/usr/bin/env python3
"""Offline failure fixtures for wiki publication; no GitHub or Excel execution."""
from __future__ import annotations

import json
import tempfile
import unittest
from pathlib import Path

from check_wiki import (
    compare_bundle,
    directories,
    read_json,
    render_page,
    sidebar,
    validate_catalogue,
    write_bundle,
)


def catalogue_fixture(root: Path) -> tuple[dict, set[str]]:
    (root / 'docs/wiki').mkdir(parents=True)
    (root / 'docs/Authority.md').write_text('Authority')
    (root / 'docs/wiki/Home.md').write_text('> **Guide, not policy:** [Authority](../Authority.md) governs.\n')
    files = {'docs/Authority.md', 'docs/wiki/Home.md', 'docs/wiki/_Sidebar.md', 'docs/wiki/catalogue.json'}
    row = {'kind': 'document', 'purpose': 'Explains a maintained contract', 'authority': 'docs/Authority.md'}
    data = {'schema_version': 1, 'pages': [{'name': 'Home', 'title': 'Start', 'authority': 'docs/Authority.md'}],
            'files': {p: dict(row) for p in files},
            'directories': {p: dict(row) for p in directories(files)},
            'kinds': {'document': {'owner': 'Maintainer', 'edit': 'Contract changes',
                                   'guard': 'Keep one authority', 'validation': 'Source checks'}}}
    return data, files


class WikiTests(unittest.TestCase):
    def test_complete_catalogue(self) -> None:
        with tempfile.TemporaryDirectory() as tmp:
            root = Path(tmp)
            data, files = catalogue_fixture(root)
            validate_catalogue(root, data, files)

    def test_catalogue_omissions_and_wrong_authority(self) -> None:
        with tempfile.TemporaryDirectory() as tmp:
            root = Path(tmp)
            data, files = catalogue_fixture(root)
            for mutation in ('missing-file', 'extra-file', 'missing-directory', 'duplicate-page',
                             'missing-authority', 'empty-purpose', 'empty-edit-rule', 'unknown-kind'):
                with self.subTest(mutation=mutation):
                    variant = json.loads(json.dumps(data))
                    if mutation == 'missing-file':
                        del variant['files']['docs/Authority.md']
                    elif mutation == 'extra-file':
                        variant['files']['extra'] = variant['files']['docs/Authority.md']
                    elif mutation == 'missing-directory':
                        del variant['directories']['docs']
                    elif mutation == 'duplicate-page':
                        variant['pages'].append(dict(variant['pages'][0]))
                    elif mutation == 'missing-authority':
                        variant['pages'][0]['authority'] = 'missing'
                    elif mutation == 'empty-purpose':
                        variant['files']['docs/Authority.md']['purpose'] = ''
                    elif mutation == 'empty-edit-rule':
                        variant['kinds']['document']['edit'] = ''
                    else:
                        variant['files']['docs/Authority.md']['kind'] = 'unknown'
                    with self.assertRaises(ValueError):
                        validate_catalogue(root, variant, files)

    def test_notice_requires_actual_authority_link(self) -> None:
        with tempfile.TemporaryDirectory() as tmp:
            root = Path(tmp)
            data, files = catalogue_fixture(root)
            for text in ('No notice', '> **Guide, not policy:** [Wrong](../Other.md)'):
                (root / 'docs/wiki/Home.md').write_text(text)
                with self.assertRaises(ValueError):
                    validate_catalogue(root, data, files)

    def test_navigation_and_pinned_authorities(self) -> None:
        text = '[Next](Next.md) [Rule](../INITIALIZATION.md#apply) [Web](https://example.com/a)'
        result = render_page(text, 'Home.md', {'Home', 'Next'}, 'owner/repo', 'a' * 40)
        self.assertIn('[Next](Next)', result)
        self.assertIn('/blob/' + 'a' * 40 + '/docs/INITIALIZATION.md#apply', result)
        self.assertIn('[Web](https://example.com/a)', result)
        self.assertIn('/docs/wiki/Home.md)', result)

    def test_root_authority_and_anchor(self) -> None:
        result = render_page('[Rule](../../CONTRIBUTING.md) [Here](#local)', 'Home.md', {'Home'}, 'o/r', 'a' * 40)
        self.assertIn('/CONTRIBUTING.md)', result)
        self.assertIn('[Here](#local)', result)

    def test_escape_refused(self) -> None:
        with self.assertRaisesRegex(ValueError, 'escapes'):
            render_page('[Bad](../../../outside)', 'Home.md', {'Home'}, 'o/r', 'a' * 40)

    def test_all_directory_ancestors(self) -> None:
        self.assertEqual(directories({'README.md', 'src/core/a.bas', '.github/x.json'}),
                         {'src', 'src/core', '.github'})

    def test_sidebar_order(self) -> None:
        data = {'pages': [{'name': 'Home', 'title': 'Start'}, {'name': 'Next', 'title': 'Continue'}]}
        result = sidebar(data)
        self.assertLess(result.index('Home.md'), result.index('Next.md'))

    def test_duplicate_catalogue_keys(self) -> None:
        with tempfile.TemporaryDirectory() as tmp:
            path = Path(tmp) / 'catalogue.json'
            path.write_text('{"files": {}, "files": {}}')
            with self.assertRaisesRegex(ValueError, 'duplicate'):
                read_json(path)

    def test_matching_bundle_and_git_metadata(self) -> None:
        with tempfile.TemporaryDirectory() as tmp:
            root = Path(tmp)
            (root / 'Home.md').write_bytes(b'page\n')
            (root / '.git').mkdir()
            (root / '.git/config').write_text('metadata')
            self.assertEqual(compare_bundle({'Home.md': b'page\n'}, root), [])

    def test_missing_extra_and_modified_pages(self) -> None:
        with tempfile.TemporaryDirectory() as tmp:
            root = Path(tmp)
            (root / 'Home.md').write_bytes(b'changed')
            (root / 'Extra.md').write_bytes(b'extra')
            findings = compare_bundle({'Home.md': b'page', 'Missing.md': b'missing'}, root)
            self.assertEqual(len(findings), 3)

    def test_manifest_source_change_is_drift(self) -> None:
        with tempfile.TemporaryDirectory() as tmp:
            root = Path(tmp)
            (root / 'Wiki-Source.json').write_text(json.dumps({'source_sha': 'a' * 40}))
            self.assertTrue(compare_bundle({'Wiki-Source.json': json.dumps({'source_sha': 'b' * 40}).encode()}, root))

    def test_symlink_refused(self) -> None:
        with tempfile.TemporaryDirectory() as tmp:
            root = Path(tmp)
            (root / 'target').write_bytes(b'page')
            (root / 'Home.md').symlink_to(root / 'target')
            with self.assertRaisesRegex(ValueError, 'symlink'):
                compare_bundle({'Home.md': b'page'}, root)

    def test_missing_readback_is_not_pass(self) -> None:
        with tempfile.TemporaryDirectory() as tmp:
            with self.assertRaisesRegex(ValueError, 'unavailable'):
                compare_bundle({}, Path(tmp) / 'missing')

    def test_export_does_not_overwrite_or_write_in_source(self) -> None:
        with tempfile.TemporaryDirectory() as tmp:
            root = Path(tmp) / 'source'
            root.mkdir()
            with self.assertRaisesRegex(ValueError, 'outside'):
                write_bundle({'Home.md': b'page'}, root / 'export', root)
            dest = Path(tmp) / 'export'
            write_bundle({'Home.md': b'page'}, dest, root)
            with self.assertRaisesRegex(ValueError, 'must not exist'):
                write_bundle({'Home.md': b'changed'}, dest, root)
            self.assertEqual((dest / 'Home.md').read_bytes(), b'page')


if __name__ == '__main__':
    unittest.main()
