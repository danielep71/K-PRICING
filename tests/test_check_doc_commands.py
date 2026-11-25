"""Tests for the check of documented commands (tools/check_doc_commands.py)."""

from __future__ import annotations

import contextlib
import io
import re
import subprocess
import sys
import tempfile
import textwrap
import unittest
from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]
SETUP_GUIDE = ROOT / "docs/setup/REPOSITORY_SETUP.md"
sys.path.insert(0, str(ROOT / "tools"))

import check_doc_commands  # noqa: E402

TOOL = textwrap.dedent("""\
    import argparse
    import sys


    def main(argv=None):
        parser = argparse.ArgumentParser()
        parser.add_argument("folder")
        parser.add_argument("--commit", required=True)
        parser.add_argument("-v", "--verbose", action="store_true")
        parser.add_argument("-q", action="count")
        mode = parser.add_mutually_exclusive_group()
        mode.add_argument("--check", action="store_true")
        mode.add_argument("--write", action="store_true")
        return parser.parse_args(argv)


    if __name__ == "__main__":
        sys.exit(main())
    """)

SUBCOMMANDS = textwrap.dedent("""\
    from argparse import ArgumentParser


    def common(parser):
        parser.add_argument("--data", required=True)


    def parse(argv):
        parser = ArgumentParser()
        parser.add_argument("--root")
        commands = parser.add_subparsers(dest="command", required=True)
        merge = commands.add_parser("merge", aliases=["combine"])
        common(merge)
        merge.add_argument("--raw", nargs="+")
        common(commands.add_parser("init"))
        return parser.parse_args(argv)


    def main():
        parse(None)


    if __name__ == "__main__":
        raise SystemExit(main())
    """)

UNREADABLE = textwrap.dedent("""\
    import argparse

    import shared

    if __name__ == "__main__":
        parser = argparse.ArgumentParser()
        shared.add_options(parser)
        parser.parse_args()
    """)


def run_check(files: dict[str, str], track: bool = True) -> list[str]:
    with tempfile.TemporaryDirectory() as directory:
        root = Path(directory)
        for name, text in files.items():
            path = root / name
            path.parent.mkdir(parents=True, exist_ok=True)
            path.write_text(text, encoding="utf-8")
        if track:
            subprocess.run(["git", "init", "-q", str(root)], check=True)
            subprocess.run(["git", "-C", str(root), "add", "."], check=True)
        findings, _, _ = check_doc_commands.check(root)
        return [str(finding) for finding in findings]


def guide(text: str, **tools: str) -> dict[str, str]:
    files = {"GUIDE.md": textwrap.dedent(text)}
    files.update({f"tools/{name}.py": source for name, source in tools.items()})
    return files


class CommandTests(unittest.TestCase):
    def test_matching_commands_pass(self) -> None:
        self.assertEqual(run_check(guide("""\
            Run `python tools/run.py out --commit <commit> --check` first, then
            `py -3 tools/run.py out --commit=X.Y.Z -vq`.

            ```sh
            python3 tools/run.py "$folder" --commit abc --write  # comment --nowhere
            python tools/run.py out --commit abc | tee log --nowhere
            ```
            """, run=TOOL)), [])

    def test_unknown_options_fail(self) -> None:
        findings = run_check(guide("""\
            Run `python tools/run.py out --commit abc --require-binding -x`.
            """, run=TOOL))
        self.assertEqual(findings, [
            "GUIDE.md:1: `python tools/run.py out --commit abc --require-binding -x`: "
            "unknown option --require-binding",
            "GUIDE.md:1: `python tools/run.py out --commit abc --require-binding -x`: "
            "unknown option -x",
        ])
        # A placeholder or variable as the attached value does not hide the option's name.
        attached = run_check(guide("""\
            `python tools/run.py out --commit=<commit> --require-bindng=${BINDING}`
            `python tools/run.py out --commit abc --require-bindng=<binding> -x<value>`
            `python tools/run.py out --commit=$COMMIT -q -v --check`
            """, run=TOOL))
        self.assertEqual([finding.split(": ", 2)[-1] for finding in attached],
                         ["unknown option --require-bindng", "unknown option --require-bindng",
                          "unknown option -x<value>"])
        # Only a name that is itself a placeholder is skipped.
        self.assertEqual(run_check(guide("`python tools/run.py out --commit a --<option> "
                                         "<options>`\n", run=TOOL)), [])
        # An abbreviation argparse would accept is reported: it is not the option's name.
        self.assertEqual(len(run_check(guide("`python tools/run.py --comm abc`\n", run=TOOL))), 1)

    def test_missing_or_untracked_script_fails(self) -> None:
        self.assertEqual(run_check(guide("`python tools/gone.py --check`\n")), [
            "GUIDE.md:1: `python tools/gone.py --check`: tools/gone.py is not a tracked file"])
        with tempfile.TemporaryDirectory() as directory:
            root = Path(directory)
            (root / "tools").mkdir()
            (root / "GUIDE.md").write_text("`python tools/run.py --check`\n", encoding="utf-8")
            subprocess.run(["git", "init", "-q", str(root)], check=True)
            subprocess.run(["git", "-C", str(root), "add", "GUIDE.md"], check=True)
            (root / "tools/run.py").write_text(TOOL, encoding="utf-8")
            findings, _, _ = check_doc_commands.check(root)
        self.assertEqual([finding.message for finding in findings],
                         ["`python tools/run.py --check`: tools/run.py is not a tracked file"])

    def test_a_tool_removed_once_run_is_skipped_only_once_deleted(self) -> None:
        cited = "`python tools/initialize_repository.py --gone`\n"
        self.assertEqual(run_check({"GUIDE.md": cited}), [])
        self.assertEqual(run_check({"GUIDE.md": cited, "tools/initialize_repository.py": TOOL}), [
            "GUIDE.md:1: `python tools/initialize_repository.py --gone`: unknown option --gone"])

    def test_subcommands_and_their_options(self) -> None:
        tools = {"traffic": SUBCOMMANDS}
        self.assertEqual(run_check(guide("""\
            `python tools/traffic.py --root . merge --data d --raw a b`,
            `python tools/traffic.py combine --data d`, `python tools/traffic.py init --data d`
            and `python tools/traffic.py <command> --anything`.
            """, **tools)), [])
        self.assertEqual(run_check(guide("""\
            `python tools/traffic.py init --raw r`
            `python tools/traffic.py merge --root .`
            `python tools/traffic.py prune --data d`
            """, **tools)), [
            "GUIDE.md:1: `python tools/traffic.py init --raw r`: unknown option --raw",
            # The parent's options are not accepted after a subcommand.
            "GUIDE.md:2: `python tools/traffic.py merge --root .`: unknown option --root",
            "GUIDE.md:3: `python tools/traffic.py prune --data d`: unknown subcommand prune "
            "(defined: combine, init, merge)",
        ])

    def test_continued_lines_are_one_command(self) -> None:
        text = """\
            ```sh
            python tools/run.py out \\
              --commit abc --nowhere
            ```

            ```powershell
            python tools\\run.py out `
              --commit abc `
              --elsewhere
            ```

            A span may wrap: `python tools/run.py out --commit abc
            --wrapped`.
            """
        self.assertEqual([finding.split(": ")[-1] for finding in run_check(guide(text, run=TOOL))],
                         ["unknown option --nowhere", "unknown option --elsewhere",
                          "unknown option --wrapped"])
        self.assertEqual([finding.split(":")[1] for finding in run_check(guide(text, run=TOOL))],
                         ["2", "7", "12"])

    def test_shell_operators_end_a_command_without_spaces(self) -> None:
        text = """\
            ```sh
            python tools/run.py out --commit abc --check; echo --done
            python tools/run.py out --commit abc|grep --count x
            python tools/run.py out --commit abc&&echo --ok
            python tools/run.py out --commit abc 2>&1 >log.txt
            python tools/run.py <folder> --commit=<sha> --write>out.txt
            python tools/run.py out --commit abc --gone;
            python tools/run.py ";" --commit "&&" --quoted-semicolon
            python tools/run.py '|' --commit '>' --quoted-pipe
            python tools/run.py \\; --commit abc --escaped
            ```
            """
        self.assertEqual(run_check(guide(text, run=TOOL)), [
            "GUIDE.md:7: `python tools/run.py out --commit abc --gone`: unknown option --gone",
            "GUIDE.md:8: `python tools/run.py ; --commit && --quoted-semicolon`: "
            "unknown option --quoted-semicolon",
            "GUIDE.md:9: `python tools/run.py | --commit > --quoted-pipe`: "
            "unknown option --quoted-pipe",
            "GUIDE.md:10: `python tools/run.py \\; --commit abc --escaped`: "
            "unknown option --escaped"])

    def test_a_hash_is_a_comment_only_where_a_word_starts(self) -> None:
        text = """\
            ```sh
            python tools/run.py "a # b" --commit abc --quoted
            python tools/run.py 'a # b' --commit abc --single
            python tools/run.py a#b --commit abc --inside
            python tools/run.py out --commit abc # --ignored
            python tools/run.py out --commit abc;# --ignored
            python tools/run.py out --commit abc --next-line
            ```
            """
        self.assertEqual([finding.split(": ")[-1] for finding in run_check(guide(text, run=TOOL))],
                         ["unknown option --quoted", "unknown option --single",
                          "unknown option --inside", "unknown option --next-line"])

    def test_other_languages_prose_and_released_history_are_skipped(self) -> None:
        self.assertEqual(run_check({
            **guide("""\
                ```text
                python tools/gone.py --anything
                ```

                ```python
                subprocess.run(["python", "tools/gone.py"])
                ```

                `python tools/<name>.py --flag`, `python -m unittest --whatever` and
                `python -I tests/script.py --x` are not commands of a tool.
                """),
            "CHANGELOG.md": "## [Unreleased]\n\n- `python tools/gone.py`\n\n"
                            "## [1.0.0] - 2026-01-01\n\n- `python tools/old.py --gone`\n",
        }), ["CHANGELOG.md:3: `python tools/gone.py`: tools/gone.py is not a tracked file"])

    def test_repository_modules(self) -> None:
        self.assertEqual(run_check({
            **guide("`python -m tools.run out --commit a --nope` and `python -m tools.gone`\n",
                    run=TOOL),
        }), ["GUIDE.md:1: `python -m tools.run out --commit a --nope`: unknown option --nope",
             "GUIDE.md:1: `python -m tools.gone`: no tracked file runs module tools.gone"])

    def test_unreadable_parsers_are_reported(self) -> None:
        self.assertEqual(run_check(guide("`python tools/shared_run.py`\n",
                                         shared_run=UNREADABLE)), [
            "GUIDE.md:1: `python tools/shared_run.py`: the parser of tools/shared_run.py "
            "cannot be read statically: line 7 passes the parser to shared.add_options, "
            "which the script does not define"])
        loop = TOOL.replace('parser.add_argument("folder")',
                            'for name in ("a", "b"):\n        parser.add_argument(f"--{name}")')
        self.assertIn("add_argument with names that are not string literals",
                      run_check(guide("`python tools/run.py`\n", run=loop))[0])
        parents = TOOL.replace("ArgumentParser()", "ArgumentParser(parents=[])")
        self.assertIn("sets parents=", run_check(guide("`python tools/run.py`\n", run=parents))[0])

    def test_script_without_a_parser_accepts_only_positional_arguments(self) -> None:
        tool = "import sys\n\nif __name__ == '__main__':\n    print(sys.argv[1:])\n"
        self.assertEqual(run_check(guide("`python tools/plain.py file.xlsm`\n", plain=tool)), [])
        self.assertEqual(run_check(guide("`python tools/plain.py --fast`\n", plain=tool)), [
            "GUIDE.md:1: `python tools/plain.py --fast`: --fast: the script defines no "
            "argparse option"])

    def test_without_git_every_file_counts(self) -> None:
        self.assertEqual(run_check(guide("`python tools/run.py x --commit a`\n", run=TOOL),
                                   track=False), [])


class RepositoryTests(unittest.TestCase):
    def test_repository_guides_pass(self) -> None:
        output = io.StringIO()
        with contextlib.redirect_stdout(output):
            code = check_doc_commands.main(["--root", str(ROOT)])
        self.assertEqual(code, 0, output.getvalue())
        self.assertIn("documented-commands: PASS", output.getvalue())

    # The repository may retire the setup guide (its *Optional: retire the
    # setup tooling* steps); the list has nothing to match then.
    @unittest.skipUnless(SETUP_GUIDE.is_file(), f"{SETUP_GUIDE.relative_to(ROOT)} is retired")
    def test_removed_once_run_matches_the_setup_guide(self) -> None:
        guide_text = SETUP_GUIDE.read_text(encoding="utf-8")
        rows = [row.split("|")[1] for row in guide_text.splitlines()
                if row.startswith("|") and "| **Removes itself once run** |" in row]
        tools = {name for row in rows for name in re.findall(r"`(tools/[^`]+\.py)`", row)}
        self.assertEqual(tools, set(check_doc_commands.REMOVED_ONCE_RUN))

    def test_a_retired_setup_guide_skips_the_comparison(self) -> None:
        # Run this module in a copy without the guide, as after the retirement steps.
        name = "test_removed_once_run_matches_the_setup_guide"
        with tempfile.TemporaryDirectory() as directory:
            copy = Path(directory)
            for path in ("tests/test_check_doc_commands.py", "tools/check_doc_commands.py",
                         "tools/gitio.py"):
                (copy / path).parent.mkdir(parents=True, exist_ok=True)
                (copy / path).write_bytes((ROOT / path).read_bytes())
            result = subprocess.run(
                [sys.executable, "-m", "unittest", "-v", f"tests.test_check_doc_commands.RepositoryTests.{name}"],
                cwd=copy, capture_output=True, text=True, check=False)
        self.assertEqual(result.returncode, 0, result.stderr)
        self.assertIn(f"{name} (", result.stderr)
        self.assertIn("skipped 'docs/setup/REPOSITORY_SETUP.md is retired'", result.stderr)

    def test_every_tool_parser_is_readable(self) -> None:
        for path in sorted((ROOT / "tools").rglob("*.py")):
            with self.subTest(path=path.relative_to(ROOT).as_posix()):
                self.assertIsNone(check_doc_commands.read_script(path).problem)


if __name__ == "__main__":
    unittest.main()
