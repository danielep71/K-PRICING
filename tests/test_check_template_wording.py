"""Tests for tools/check_template_wording.py: text that describes the template."""

from __future__ import annotations

import contextlib
import io
import subprocess
import sys
import tempfile
import unittest
from collections.abc import Mapping
from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(ROOT / "tools"))

import check_template_wording as C  # noqa: E402

# Sentences that a repository created from the template had to rewrite by hand.
OLD_SENTENCES = {
    "docs/REPOSITORY_STRUCTURE.md": ("| Location | Holds | In the template |\n"
                                     "| --- | --- | --- |\n\n"
                                     "`src/`, in two tiers; in the template both are empty.\n"),
    "INSTALLATION.md": ("contracts and their reference values do; the template's illustrative\n"
                        "contract validates none.\n"),
    "RELEASING.md": "Use `--template true` in the template repository itself.\n",
    ".github/dependabot.yml": ("# Update target-branch when a new release branch opens; in a "
                               "generated\n# repository, set it to that repository's release "
                               "branch.\nversion: 2\n"),
    ".editorconfig": "# POLICY\n#   These defaults also cover file types used by generated\n"
                     "#   repositories.\nroot = true\n",
}


def repository(root: Path, files: Mapping[str, str | bytes],
               untracked: Mapping[str, str] | None = None) -> Path:
    for path, content in files.items():
        (root / path).parent.mkdir(parents=True, exist_ok=True)
        if isinstance(content, bytes):
            (root / path).write_bytes(content)
        else:
            (root / path).write_text(content, encoding="utf-8")
    subprocess.run(["git", "init", "-q", str(root)], check=True)
    subprocess.run(["git", "-C", str(root), "add", "-A"], check=True)
    for path, text in (untracked or {}).items():
        (root / path).write_text(text, encoding="utf-8")
    return root


class TemplateWording(unittest.TestCase):
    def findings(self, files: Mapping[str, str | bytes], allowed: tuple[C.Allowed, ...] = (),
                 scope: dict[str, str] | None = None,
                 untracked: Mapping[str, str] | None = None) -> list[str]:
        with tempfile.TemporaryDirectory() as folder:
            root = repository(Path(folder), files, untracked)
            return C.findings(root, {} if scope is None else scope, allowed)[0]

    def located(self, files: Mapping[str, str | bytes]) -> list[str]:
        """Each finding as path:line: phrase."""
        return [finding.split(" describes", 1)[0] for finding in self.findings(files)]

    def test_the_repository_passes(self) -> None:
        problems, read = C.findings(ROOT)
        self.assertEqual(problems, [])
        self.assertGreater(read, 0)

    def test_the_old_sentences_fail_with_file_line_and_phrase(self) -> None:
        found = self.findings(OLD_SENTENCES)
        self.assertEqual([finding.split(" describes", 1)[0] for finding in found], [
            ".editorconfig:2: 'generated repository'",
            ".github/dependabot.yml:1: 'generated repository'",
            "INSTALLATION.md:1: \"the template's\"",
            "RELEASING.md:1: 'in the template'",
            "RELEASING.md:1: 'template repository'",
            "docs/REPOSITORY_STRUCTURE.md:1: 'in the template'",
            "docs/REPOSITORY_STRUCTURE.md:4: 'in the template'",
        ])
        for finding in found:
            self.assertIn("describes the template, not the repository the file is in; "
                          "reword it", finding)
            self.assertTrue(finding.endswith("with a reason"), finding)

    def test_case_line_breaks_comment_marks_and_apostrophes(self) -> None:
        self.assertEqual(self.located({
            "a.md": "IN THE TEMPLATE\n",
            "b.md": "As the **template’s** own record.\n",
            "c.py": "# kept in the\n# template\n",
            "d.bas": "' a Generated\r\n' Repositories list\r\n",
            "e.js": "// the\n// template's own\n",
            "f.md": "> stays in the\n> `template`.\n",
        }), ["a.md:1: 'in the template'", "b.md:1: \"the template's\"",
             "c.py:1: 'in the template'", "d.bas:1: 'generated repository'",
             "e.js:1: \"the template's\"", "f.md:1: 'in the template'"])

    def test_longer_words_and_other_phrases_pass(self) -> None:
        text = ("Delete the template-only section. Keep in the templates folder. Run the "
                "templated build, a regenerated repository, the templates' rules, a "
                "template’d form, the workbook template's sheets as copied from a template "
                "and created from the template.\n")
        self.assertEqual(self.findings({"GUIDE.md": text}), [])

    def test_only_the_template_only_sections_of_agents_and_readme_are_left_out(self) -> None:
        section = ("# Rules\n\n<!-- template-only:start -->\nIn the template, the template's "
                   "rules.\n<!-- template-only:end -->\n\nOutside: in the template.\n")
        self.assertEqual(self.located({"AGENTS.md": section, "README.md": section,
                                       "OTHER.md": section}), [
            "AGENTS.md:7: 'in the template'",
            "OTHER.md:4: 'in the template'",
            "OTHER.md:4: \"the template's\"",
            "OTHER.md:7: 'in the template'",
            "README.md:7: 'in the template'",
        ])

    def test_excluded_binary_and_untracked_files_are_out_of_scope(self) -> None:
        scope = {"CHANGELOG.md": "history", "docs/licenses/": "third-party texts"}
        found = self.findings({"CHANGELOG.md": "in the template\n",
                               "docs/licenses/NOTICE.txt": "in the template\n",
                               "docs/licensesX.md": "in the template\n",
                               "image.png": b"\x89PNG\0in the template",
                               "latin.txt": "in the template \xe9".encode("latin-1")},
                              scope=scope, untracked={"NOTES.md": "in the template\n"})
        self.assertEqual([finding.split(" describes", 1)[0] for finding in found],
                         ["docs/licensesX.md:1: 'in the template'"])

    def test_every_exclusion_and_allowed_use_has_a_reason(self) -> None:
        for path, reason in C.EXCLUDED.items():
            self.assertTrue(reason.strip(), path)
        for entry in C.ALLOWED:
            self.assertIn(entry.phrase, C.PHRASES)
            self.assertTrue(entry.reason.strip(), entry)
            self.assertTrue((ROOT / entry.path).is_file(), entry.path)

    def test_allowed_uses(self) -> None:
        files: dict[str, str | bytes] = {"GUIDE.md": "The template's sheets.\n"}
        kept = C.Allowed("GUIDE.md", "the template's", "The workbook template.")
        self.assertEqual(self.findings(files, (kept,)), [])
        # An allowed use covers its phrase in its file only.
        found = self.findings({**files, "OTHER.md": "The template's sheets.\n"}, (kept,))
        self.assertEqual(len(found), 1, found)
        self.assertTrue(found[0].startswith("OTHER.md:1: \"the template's\""), found[0])
        elsewhere = C.Allowed("OTHER.md", "the template's", "Another guide.")
        found = self.findings(files, (elsewhere,))
        self.assertEqual(len(found), 2, found)
        self.assertTrue(found[0].startswith("GUIDE.md:1:"), found[0])
        self.assertEqual(found[1], "allowed use \"the template's\" in OTHER.md matches nothing; "
                                   "remove it from ALLOWED in tools/check_template_wording.py")
        unreasoned = C.Allowed("GUIDE.md", "the template's", " ")
        self.assertEqual(self.findings(files, (unreasoned,)),
                         ["allowed use \"the template's\" in GUIDE.md has no reason"])
        unknown = C.Allowed("GUIDE.md", "template", "Not a phrase.")
        found = self.findings({"GUIDE.md": "clean\n"}, (unknown,))
        self.assertEqual(len(found), 1, found)
        self.assertIn("allowed use 'template' in GUIDE.md is not one of the phrases", found[0])

    def test_tracked_files_that_cannot_be_listed_fail(self) -> None:
        with tempfile.TemporaryDirectory() as folder:
            problems, read = C.findings(Path(folder) / "missing", {}, ())
        self.assertEqual(read, 0)
        self.assertEqual(len(problems), 1, problems)
        self.assertIn("the tracked files cannot be listed", problems[0])

    def test_main_reports_and_exits(self) -> None:
        with tempfile.TemporaryDirectory() as folder:
            repository(Path(folder), {"GUIDE.md": "Empty in the template.\n"})
            output = io.StringIO()
            with contextlib.redirect_stdout(output):
                code = C.main(["--root", folder])
            self.assertEqual(code, 1)
            self.assertIn("FAIL: GUIDE.md:1: 'in the template'", output.getvalue())
            self.assertIn("template-wording: FAIL (1 file(s) read", output.getvalue())
        output = io.StringIO()
        with contextlib.redirect_stdout(output):
            self.assertEqual(C.main(["--root", str(ROOT)]), 0)
        self.assertIn("template-wording: PASS (", output.getvalue())


if __name__ == "__main__":
    unittest.main()
