"""Prove that the Ruff complexity ceiling, mypy strict mode with warn_unreachable,
and actionlint are active.

Ruff and mypy are installed from tools/requirements-quality-ci.txt, and
actionlint as a checksum-verified release. The Repository integrity job sets
REQUIRE_QUALITY_TOOLS=1, so these tests fail there if a tool is missing;
elsewhere they are skipped without the tools.
"""

# mypy checks the bodies of these tests but does not require them to be annotated.
# mypy: disallow-untyped-defs=False, disallow-incomplete-defs=False, disallow-untyped-calls=False

import functools
import os
import re
import shutil
import subprocess
import sys
import tempfile
import unittest
from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(ROOT / 'tools'))

import check_repository  # noqa: E402

REQUIRED = os.environ.get('REQUIRE_QUALITY_TOOLS') == '1'


def run(command, stdin=None):
    result = subprocess.run(command, cwd=ROOT, input=stdin, text=True,
                            capture_output=True, check=False)
    return result.returncode, result.stdout + result.stderr


INSTALL = {
    'ruff': 'python -m pip install --only-binary=:all: --require-hashes '
            '-r tools/requirements-quality-ci.txt',
    'mypy': 'python -m pip install --only-binary=:all: --require-hashes '
            '-r tools/requirements-quality-ci.txt',
    'actionlint': 'install the checksum-verified release that '
                  '.github/workflows/repository-integrity.yml downloads',
}


def tool(name):
    return unittest.skipIf(shutil.which(name) is None and not REQUIRED,
                           f'{name} is not installed: {INSTALL[name]}')


NARROWED = ('def probe(value: int | str) -> str:\n'
            '    if isinstance(value, int):\n        return str(value)\n'
            '    if isinstance(value, str):\n        return value\n')
# For each mypy test, a probe that mypy must pass and one it must reject.
MYPY_PROBES = {
    'strict_passes.py': 'def probe(value: int) -> int:\n    return value\n',
    'strict_fails.py': 'def probe(value):\n    return value\n',
    'unreachable_passes.py': NARROWED,
    'unreachable_fails.py': NARROWED + "    return 'never'\n",
}


@functools.cache
def mypy_probes():
    """mypy's exit code, its output and the errors it reports in each probe, from one
    run over all of them with this repository's configuration and a cache of its own."""
    with tempfile.TemporaryDirectory() as directory:
        paths = []
        for name, source in MYPY_PROBES.items():
            path = Path(directory) / name
            path.write_text(source, encoding='utf-8')
            paths.append(str(path))
        code, log = run(['mypy', '--no-pretty', '--no-incremental', '--cache-dir',
                         str(Path(directory) / 'cache'), *paths])
    errors = {name: [line for line in log.splitlines()
                     if re.search(rf'[\\/]{re.escape(name)}:\d+: error:', line)]
              for name in MYPY_PROBES}
    return code, log, errors


class QualityGateTests(unittest.TestCase):
    @tool('ruff')
    def test_complexity_15_passes_and_16_fails(self):
        def source(branches):
            body = ''.join(f'    if x == {i}:\n        return {i}\n' for i in range(branches - 1))
            return f'def probe(x: int) -> int:\n{body}    return -1\n'

        command = ['ruff', 'check', '--no-cache', '--select', 'C901',
                   '--stdin-filename', 'tools/complexity_probe.py', '-']
        ok_code, ok_log = run(command, source(15))
        bad_code, bad_log = run(command, source(16))
        self.assertEqual(ok_code, 0, ok_log)
        self.assertNotEqual(bad_code, 0, bad_log)
        self.assertIn('is too complex (16 > 15)', bad_log)

    @tool('mypy')
    def test_strict_mode_rejects_an_unannotated_function(self):
        self.assertMypyProbes('strict', '[no-untyped-def]')

    @tool('mypy')
    def test_unreachable_code_is_reported(self):
        self.assertMypyProbes('unreachable', '[unreachable]')

    def assertMypyProbes(self, name, code):
        """The passing probe of ``name`` has no error and its failing probe ``code``."""
        exit_code, log, errors = mypy_probes()
        self.assertIn(f'checked {len(MYPY_PROBES)} source files', log)
        self.assertNotEqual(exit_code, 0, log)
        self.assertEqual(errors[f'{name}_passes.py'], [], log)
        self.assertTrue(errors[f'{name}_fails.py'], log)
        self.assertTrue(all(code in line for line in errors[f'{name}_fails.py']), log)

    @tool('actionlint')
    def test_actionlint_rejects_an_unknown_key(self):
        workflow = ('name: Probe\non: push\npermissions: {}\njobs:\n  probe:\n'
                    '    runs-on: ubuntu-24.04\n    steps:\n      - run: echo ok\n')
        command = ['actionlint', '-no-color', '-shellcheck=', '-pyflakes=', '-']
        ok_code, ok_log = run(command, workflow)
        bad_code, bad_log = run(command, workflow + '        bogus-key: true\n')
        self.assertEqual(ok_code, 0, ok_log)
        self.assertNotEqual(bad_code, 0, bad_log)
        self.assertIn('unexpected key "bogus-key"', bad_log)

    @tool('actionlint')
    def test_actionlint_rejects_the_probe_of_the_actionlint_check(self):
        # The probe lies outside .github/workflows/, which GitHub runs and actionlint lints.
        probe = check_repository.ACTIONLINT_PROBE
        self.assertFalse(probe.startswith('.github/workflows/'))
        code, log = run(['actionlint', '-no-color', '-shellcheck=', '-pyflakes=', probe])
        self.assertNotEqual(code, 0, log)
        self.assertIn(check_repository.ACTIONLINT_PROBE_FINDING, log)


if __name__ == '__main__':
    unittest.main()
