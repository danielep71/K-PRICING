"""Tests for the Markdown link checker (tools/check_links.py)."""

from __future__ import annotations

import contextlib
import io
import json
import subprocess
import sys
import tempfile
import unittest
import unittest.mock
import warnings
from datetime import date
from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(ROOT / "tools"))

import check_links  # noqa: E402


def run_check(files: dict[str, str]) -> tuple[list[str], set[str], int]:
    with tempfile.TemporaryDirectory() as directory:
        root = Path(directory)
        for name, text in files.items():
            path = root / name
            path.parent.mkdir(parents=True, exist_ok=True)
            path.write_text(text, encoding="utf-8")
        # Not a Git checkout: the checker falls back to every *.md file.
        return check_links.check(root)


class SlugTests(unittest.TestCase):
    def test_github_slugs(self) -> None:
        self.assertEqual(check_links.github_slug("Branches and publication"), "branches-and-publication")
        self.assertEqual(check_links.github_slug("🧭 Before you start"), "-before-you-start")
        # U+FE0F after an emoji is a combining mark that GitHub keeps; "①" is not a word character.
        self.assertEqual(check_links.github_slug("▶\ufe0f Launch sequence"), "\ufe0f-launch-sequence")
        self.assertEqual(check_links.github_slug("① Preview the changes"), "-preview-the-changes")
        self.assertEqual(check_links.github_slug("Café 2026_v1"), "café-2026_v1")
        self.assertEqual(check_links.github_slug("Run `check.py` [now](x.md)"), "run-checkpy-now")
        self.assertEqual(check_links.github_slug("Rock &amp; Roll"), "rock--roll")

    def test_duplicate_and_setext_headings(self) -> None:
        _, anchors = check_links.markdown_links("# Notes\n\n## Notes\n\nTitle\n=====\n")
        self.assertEqual(anchors, {"notes", "notes-1", "title"})
        _, anchors = check_links.markdown_links("# Foo\n# Foo-1\n# Foo\n")
        self.assertEqual(anchors, {"foo", "foo-1", "foo-2"})

    def test_code_is_ignored_and_explicit_anchors_count(self) -> None:
        text = ('<a id="custom"></a>\n\n```sh\n[x](missing.md)\n# not a heading\n```\n'
                "`[y](also-missing.md)` [z](real.md#part)\n")
        links, anchors = check_links.markdown_links(text)
        self.assertEqual(links, ["real.md#part"])
        self.assertEqual(anchors, {"custom"})


class CheckTests(unittest.TestCase):
    def test_valid_links_pass_and_external_links_are_collected(self) -> None:
        errors, external, count = run_check({
            "README.md": "[Guide](docs/guide.md#-install) · [site](https://example.com/page#top)\n"
                         "[self](#readme) <https://example.org>\n\n# README\n",
            "docs/guide.md": "# 🚀 Install\n\n[Back](../README.md)\n",
        })
        self.assertEqual(errors, [])
        self.assertEqual(external, {"https://example.com/page", "https://example.org"})
        self.assertEqual(count, 5)

    def test_missing_files_anchors_and_paths_are_errors(self) -> None:
        errors, _, _ = run_check({
            "README.md": "[a](docs/missing.md) [b](docs/guide.md#nowhere) [c](/abs.md)\n"
                         "[d](../outside.md) [e](empty.md) [f](ftp://example.com)\n",
            "docs/guide.md": "# Guide\n",
            "empty.md": "",
        })
        self.assertEqual(errors, [
            "README.md: missing target: docs/missing.md",
            "README.md: missing anchor: docs/guide.md#nowhere",
            "README.md: root-relative link: /abs.md",
            "README.md: path outside the repository: ../outside.md",
            "README.md: empty target: empty.md",
            "README.md: unsupported URL scheme: ftp://example.com",
        ])

    def test_plain_http_links_are_errors_outside_code(self) -> None:
        errors, external, _ = run_check({
            "README.md": "[a](http://a.example/x) <http://b.example> [c](HTTP://c.example)\n"
                         "[d]: http://d.example/\n<a href=\"http://e.example\">e</a>\n"
                         "`[f](http://f.example)` and `<http://g.example>` are code\n\n"
                         "```text\n[h](http://h.example)\n```\n[ok](https://ok.example)\n",
        })
        self.assertEqual(errors, [
            "README.md: plain HTTP link, use https: http://a.example/x",
            "README.md: plain HTTP link, use https: HTTP://c.example",
            "README.md: plain HTTP link, use https: http://b.example",
            "README.md: plain HTTP link, use https: http://d.example/",
            "README.md: plain HTTP link, use https: http://e.example",
        ])
        self.assertEqual(external, {"https://ok.example"})

    def test_tracked_file_missing_from_the_working_tree_is_reported(self) -> None:
        with tempfile.TemporaryDirectory() as directory:
            root = Path(directory)
            (root / "README.md").write_text("[guide](guide.md)\n", encoding="utf-8")
            (root / "guide.md").write_text("# Guide\n", encoding="utf-8")
            subprocess.run(["git", "init", "-q", str(root)], check=True)
            subprocess.run(["git", "-C", str(root), "add", "."], check=True)
            (root / "guide.md").unlink()
            errors, _, _ = check_links.check(root)
        self.assertEqual(errors, ["guide.md: tracked but missing from the working tree",
                                  "README.md: missing target: guide.md"])

    def test_repository_links_pass(self) -> None:
        errors, _, count = check_links.check(ROOT)
        self.assertGreater(count, 0)
        self.assertEqual(errors, [])


class ExternalTests(unittest.TestCase):
    def test_own_repository_urls_are_skipped(self) -> None:
        own = check_links.own_repository
        self.assertTrue(own("https://github.com/Owner/Repo/blob/HEAD/x.md", "owner/repo"))
        self.assertTrue(own("https://github.com/owner/repo", "owner/repo"))
        self.assertTrue(own("https://raw.githubusercontent.com/owner/repo/main/x", "owner/repo"))
        self.assertFalse(own("https://github.com/owner/repo-other", "owner/repo"))
        self.assertFalse(own("https://github.com/owner/repo", ""))

    def test_only_errors_fail_and_own_urls_are_not_requested(self) -> None:
        results = {"https://a.example/": ("PASS", "https://a.example/"),
                   "https://b.example/": ("WARN", "https://b.example/: unavailable"),
                   "https://c.example/": ("ERROR", "https://c.example/: HTTP 404")}
        with unittest.mock.patch.object(check_links, "check_external",
                                        side_effect=lambda url: results[url]) as fake, \
                contextlib.redirect_stdout(io.StringIO()):
            errors = check_links.external_errors(
                [*results, "https://github.com/owner/repo/issues"], "owner/repo")
        self.assertEqual(errors, ["https://c.example/: HTTP 404"])
        self.assertEqual(sorted(call.args[0] for call in fake.call_args_list), sorted(results))

    def test_a_local_run_skips_the_origin_repository(self) -> None:
        # Without --repository or GITHUB_REPOSITORY, the origin remote names it.
        with tempfile.TemporaryDirectory() as directory:
            root = Path(directory)
            (root / "README.md").write_text(
                "[own](https://github.com/Owner/Repo/blob/HEAD/x.md) [other](https://a.example/)\n",
                encoding="utf-8")
            subprocess.run(["git", "init", "-q", str(root)], check=True)
            subprocess.run(["git", "-C", str(root), "remote", "add", "origin",
                            "git@github.com:owner/repo.git"], check=True)
            subprocess.run(["git", "-C", str(root), "add", "README.md"], check=True)
            self.assertEqual(check_links.detect_repository(root), "owner/repo")
            with unittest.mock.patch.dict("os.environ", {"GITHUB_REPOSITORY": ""}), \
                    unittest.mock.patch.object(check_links, "check_external",
                                               return_value=("PASS", "ok")) as fake, \
                    contextlib.redirect_stdout(io.StringIO()):
                self.assertEqual(check_links.main(["--root", str(root), "--external"]), 0)
        self.assertEqual([call.args[0] for call in fake.call_args_list], ["https://a.example/"])
        with tempfile.TemporaryDirectory() as directory:
            self.assertEqual(check_links.detect_repository(Path(directory)), "")

    def test_http_codes_are_classified(self) -> None:
        def curl(code: str, returncode: int = 0) -> unittest.mock.Mock:
            return unittest.mock.Mock(stdout=code, returncode=returncode)

        for code, returncode, status in (("200", 0, "PASS"), ("404", 0, "ERROR"),
                                         ("429", 0, "WARN"), ("503", 0, "WARN"),
                                         ("", 6, "WARN"), ("301", 1, "ERROR")):
            with self.subTest(code=code), unittest.mock.patch(
                    "check_links.subprocess.run", return_value=curl(code, returncode)):
                self.assertEqual(check_links.check_external("https://x.example/")[0], status)

    def test_requests_use_https_only(self) -> None:
        with unittest.mock.patch("check_links.subprocess.run",
                                 return_value=unittest.mock.Mock(stdout="200", returncode=0)) as run:
            check_links.check_external("https://x.example/")
        command = run.call_args.args[0]
        self.assertEqual(command[command.index("--proto") + 1], "=https")
        self.assertEqual(command[command.index("--proto-redir") + 1], "=https")


class GitHubReadingTests(unittest.TestCase):
    """Markdown read the way GitHub renders it; one test per behaviour."""

    def test_probe_cases_read_as_github_renders_them(self) -> None:
        links_cases: list[tuple[str, list[str]]] = [
            ("\\[text](missing.md)\n", []),
            ("[x](a\\_b.md)\n", ["a_b.md"]),
            ('[x](notes&amp;examples.md) <a href="notes&amp;examples.md">x</a>\n',
             ["notes&examples.md", "notes&examples.md"]),
            ("para\n\n    [x](code.md)\n", []),
            ("A `span\n[x](span.md)` here\n", []),
            ("<!-- [x](comment.md) -->\n", []),
            ("<!--\n[x](comment.md)\n-->\n", []),
            ("> ```\n> [x](fenced.md)\n> ```\n", []),
            ("- ```\n  [x](listfence.md)\n  ```\n", []),
        ]
        for text, expected in links_cases:
            with self.subTest(text=text):
                self.assertEqual(check_links.markdown_links(text)[0], expected)
        for text, anchors in (("> # Quoted Title\n", {"quoted-title"}),
                              ("- # Item Title\n", {"item-title"}),
                              ("# __API__ notes\n", {"api-notes", "__api__-notes"})):
            with self.subTest(text=text):
                self.assertEqual(check_links.markdown_links(text)[1], anchors)

    def test_link_needs_an_unescaped_opener(self) -> None:
        errors, _, _ = run_check({
            "README.md": "\\[example](missing.md) !\\[alt](missing.png) [x\\](missing.md)\n"
                         "\\\\[real](missing-r.md) ![alt](missing-img.png)\n"
                         "[![badge](badge.svg)](missing-n.md) [text\nmore](missing-m.md)\n",
            "badge.svg": "<svg/>\n",
        })
        self.assertEqual(errors, ["README.md: missing target: missing-r.md",
                                  "README.md: missing target: missing-img.png",
                                  "README.md: missing target: missing-n.md",
                                  "README.md: missing target: missing-m.md"])

    def test_backslash_escapes_in_destinations_are_removed(self) -> None:
        errors, _, count = run_check({
            "README.md": "[ok](notes\\).md) [ok](notes\\(draft.md) [ok](a(b).md#top)\n"
                         "[ok](a\\_b.md) [ok](<a\\_b.md>)\n[ref]: a\\_b.md\n[bad](gone\\).md)\n",
            "notes).md": "# Notes\n",
            "notes(draft.md": "# Draft\n",
            "a(b).md": "# Top\n",
            "a_b.md": "# Underscore\n",
        })
        self.assertEqual(count, 7)
        self.assertEqual(errors, ["README.md: missing target: gone).md"])

    def test_character_references_are_decoded(self) -> None:
        errors, external, _ = run_check({
            "README.md": '<a href="notes&amp;examples.md#top">x</a> '
                         "[y](notes&amp;examples.md) [z](notes\\&examples.md)\n"
                         '<img src="a&#32;b.svg"> <a href="a&ampb.md">literal</a>\n'
                         '<a href="missing&amp;file.md">w</a> [r](#r-and-d)\n'
                         "[q](https://example.com/?a=1&amp;b=2)\n\n"
                         "[def]: notes&amp;examples.md\n"
                         '<a id="r&amp;d"></a> <a name="r-and-d"></a>\n',
            "notes&examples.md": "# Top\n",
            "a b.svg": "<svg/>\n",
            "a&ampb.md": "# Literal\n",
            "other.md": "[a](README.md#r&d)\n",
        })
        self.assertEqual(errors, ["README.md: missing target: missing&file.md"])
        self.assertEqual(external, {"https://example.com/?a=1&b=2"})

    def test_code_span_across_lines_is_ignored(self) -> None:
        # A code span may continue over line breaks within its paragraph.
        errors, _, count = run_check({
            "README.md": "Use this literal example: `\n[example](missing.md)\n`\n\n"
                         "Or ``this `one`\n[x](missing.md#a) ``.\n\n"
                         "> Quoted `\nlazy [y](missing.md)\n> `\n",
        })
        self.assertEqual((errors, count), ([], 0))

    def test_unclosed_and_escaped_backticks_do_not_hide_links(self) -> None:
        # An unmatched run is literal, and a span never crosses a blank line.
        errors, _, _ = run_check({
            "README.md": "A lone ` backtick\n[a](missing-a.md)\n\n"
                         "Open `here\n\n[b](missing-b.md) and `\n\n"
                         "Escaped \\`[c](missing-c.md)\\`\n\n"
                         "Not escaped \\\\`[d](missing-d.md)`\n",
        })
        self.assertEqual(errors, ["README.md: missing target: missing-a.md",
                                  "README.md: missing target: missing-b.md",
                                  "README.md: missing target: missing-c.md"])

    def test_backticks_inside_comments_do_not_open_code_spans(self) -> None:
        # Comments and code spans are read left to right: a backtick inside a
        # comment is part of the comment, and "<!--" inside a code span is code.
        errors, _, _ = run_check({
            "README.md": "Text <!-- ` -->\n[x](missing.md)\nText <!-- ` -->\n\n"
                         "Text <!-- ` -->\n[h](http://h.example)\nText <!-- ` -->\n\n"
                         "A `<!--` span [y](missing-y.md) `-->` and <!-- [z](missing.md) -->\n",
        })
        self.assertEqual(errors, ["README.md: missing target: missing.md",
                                  "README.md: plain HTTP link, use https: http://h.example",
                                  "README.md: missing target: missing-y.md"])

    def test_indented_code_blocks_are_ignored(self) -> None:
        errors, _, count = run_check({
            "README.md": "Text\n\n    [a](missing.md)\n\n    [b](missing.md)\n\n"
                         "\t[c](missing.md)\n\n# Heading\n    [d](missing.md)\n\n"
                         "- item\n\n      [e](missing.md)\n\n"
                         "1. step\n\n    ```text\n    [f](missing.md)\n    ```\n",
        })
        self.assertEqual((errors, count), ([], 0))

    def test_indented_lines_that_are_not_code_are_checked(self) -> None:
        # Indented code cannot interrupt a paragraph, and list content is
        # indented relative to its item, not to the margin.
        errors, _, _ = run_check({
            "README.md": "Paragraph\n    [a](missing-a.md)\n\n"
                         "- item\n\n    [b](missing-b.md)\n\n"
                         "  - nested\n\n    [c](missing-c.md)\n\n"
                         "1. step\n   continued [d](missing-d.md)\n",
        })
        self.assertEqual(errors, [f"README.md: missing target: missing-{name}.md"
                                  for name in "abcd"])

    def test_html_comments_are_ignored(self) -> None:
        errors, _, _ = run_check({
            "README.md": "<!--\n[example](missing.md)\n# Hidden\n\n<a id=\"ghost\"></a>\n-->\n"
                         "Text <!-- [a](missing.md) --> more <!-- [b](missing.md)\n"
                         "[c](missing.md) --> end\n\n"
                         "> <!-- [d](missing.md) -->\n\n"
                         "<!-- one line --> [e](missing.md)\n\n"
                         "[f](missing-f.md) [h](#hidden) [g](#ghost)\n",
        })
        self.assertEqual(errors, ["README.md: missing target: missing-f.md",
                                  "README.md: missing anchor: #hidden",
                                  "README.md: missing anchor: #ghost"])

    def test_fences_and_headings_inside_containers(self) -> None:
        errors, _, _ = run_check({
            "README.md": "> ~~~md\n> [example](missing.md)\n> ~~~\n> [b](missing-b.md)\n\n"
                         "- ~~~md\n  [example](missing.md)\n  ~~~\n\n"
                         "> # Quoted heading\n\n- ## Item heading\n1. # Numbered\n\n"
                         "> ```\n> code\n\n[a](missing-a.md)\n\n"
                         "[q](#quoted-heading) [i](#item-heading) [n](#numbered)\n",
        })
        self.assertEqual(errors, ["README.md: missing target: missing-b.md",
                                  "README.md: missing target: missing-a.md"])

    def test_setext_headings_and_breaks_inside_containers(self) -> None:
        _, anchors = check_links.markdown_links(
            "- Item title\n  ----------\n\n> Quoted title\n> ===\n\n"
            "> Intro\n> ***\n> After break\n> ===\n\n"
            "- First\n  * * *\n  Second\n  ===\n")
        self.assertEqual(anchors, {"item-title", "quoted-title", "after-break", "second"})

    def test_lazy_lines_continue_a_quoted_paragraph_only(self) -> None:
        errors, _, _ = run_check({
            "README.md": "> Quoted `\nlazy [x](missing.md)\n> ` text\n\n"
                         "> Quoted\n```\n[y](missing.md)\n```\n[z](missing-z.md)\n",
        })
        self.assertEqual(errors, ["README.md: missing target: missing-z.md"])

    def test_closing_fence_has_nothing_after_its_run(self) -> None:
        errors, _, _ = run_check({
            "README.md": "```\n``` not a closing fence\n[a](missing.md)\n```\n[b](missing-b.md)\n\n"
                         "1.  ```\n    [c](missing.md)\n    ```\n\n[d](missing-d.md)\n\n"
                         "- ```\n  [e](missing.md)\n[f](missing-f.md)\n\n"
                         "```\n> ```\n[g](missing.md)\n```\n",
        })
        # A fence in a list item ends with the item, at the line left of its content;
        # a quote marker inside a fence is code.
        self.assertEqual(errors, ["README.md: missing target: missing-b.md",
                                  "README.md: missing target: missing-d.md",
                                  "README.md: missing target: missing-f.md"])

    def test_heading_emphasis_is_not_part_of_the_slug(self) -> None:
        slug = check_links.github_slug
        self.assertEqual(slug("__API__"), "api")
        self.assertEqual(slug("__API__ and _x_"), "api-and-x")
        self.assertEqual(slug("foo_bar_ baz"), "foo_bar_-baz")
        self.assertEqual(slug("`__init__` and __init__"), "__init__-and-init")
        self.assertEqual(slug("__API__", emphasis=False), "__api__")
        errors, _, _ = run_check({
            "README.md": "[a](guide.md#api) [b](guide.md#snake_case-names) "
                         "[c](guide.md#__init__-method) [d](guide.md#old-new-bold-text)\n"
                         "[e](guide.md#_private-helper) [f](guide.md#api-1) [g](guide.md#__api__)\n"
                         "[h](guide.md#nowhere) [i](guide.md#api-2)\n",
            "guide.md": "# __API__\n## snake_case names\n## `__init__` method\n"
                        "## ~~old~~ *new* **bold** _text_\n## _private helper\n## API\n",
        })
        self.assertEqual(errors, ["README.md: missing anchor: guide.md#nowhere",
                                  "README.md: missing anchor: guide.md#api-2"])

    def test_escaped_underscores_stay_in_the_slug(self) -> None:
        self.assertEqual(check_links.github_slug("\\_foo\\_"), "_foo_")
        self.assertEqual(check_links.github_slug("\\_foo_ and _bar_"), "_foo_-and-bar")
        errors, _, _ = run_check({
            "README.md": "# \\_foo\\_\n# \\_foo\\_\n\n[a](#_foo_) [b](#_foo_-1) [c](#foo)\n",
        })
        self.assertEqual(errors, ["README.md: missing anchor: #foo"])

    def test_docstring_has_no_invalid_escape(self) -> None:
        source = (ROOT / "tools" / "check_links.py").read_text(encoding="utf-8")
        with warnings.catch_warnings():
            warnings.simplefilter("error")
            compile(source, "check_links.py", "exec")


def write_exclusions(files: dict[str, str], *entries: object) -> dict[str, str]:
    return {**files, check_links.EXCLUSIONS.as_posix(): json.dumps(
        {"schema_version": 1, "exclusions": list(entries)})}


class ExclusionTests(unittest.TestCase):
    NAME = ".github/link-check-exclusions.json"

    def test_absent_file_means_no_exclusions(self) -> None:
        with tempfile.TemporaryDirectory() as directory:
            self.assertEqual(check_links.load_exclusions(Path(directory)), ({}, []))
        errors, _, _ = run_check({"README.md": "[a](missing.md)\n"})
        self.assertEqual(errors, ["README.md: missing target: missing.md"])

    def test_listed_report_is_not_an_error_and_unused_entry_is(self) -> None:
        errors, _, _ = run_check(write_exclusions(
            {"README.md": "[a](missing.md#part) [b](still-missing.md) [c](ftp://x.example)\n"},
            {"path": "README.md", "link": "missing.md#part", "reason": "Rendered as text"},
            {"path": "README.md", "link": "ftp://x.example", "reason": "Rendered as text"},
            {"path": "README.md", "link": "gone.md", "reason": "Fixed since"},
            {"path": "docs/other.md", "link": "missing.md#part", "reason": "Wrong file"}))
        self.assertEqual(errors, [
            "README.md: missing target: still-missing.md",
            f"{self.NAME}: README.md: gone.md is no longer reported; remove its exclusion",
            f"{self.NAME}: docs/other.md: missing.md#part is no longer reported; "
            "remove its exclusion",
        ])

    def test_malformed_entries_fail(self) -> None:
        entry = {"path": "README.md", "link": "missing.md", "reason": "Rendered as text"}
        cases: list[tuple[str, str]] = [
            ("{", f"{self.NAME}: not valid UTF-8 JSON"),
            (json.dumps({"exclusions": []}),
             f"{self.NAME}: needs exactly schema_version 1 and an exclusions list"),
            (json.dumps({"schema_version": 1, "exclusions": {}}),
             f"{self.NAME}: needs exactly schema_version 1 and an exclusions list"),
        ]
        entry_cases: list[tuple[object, str]] = [
            ("x", "needs exactly path, link and reason"),
            ({"link": "missing.md", "reason": "r"}, "needs exactly path, link and reason"),
            ({**entry, "expires": "2027-01-01"}, "needs exactly path, link and reason"),
            ({**entry, "path": ""}, "needs a non-empty path"),
            ({**entry, "link": " "}, "needs a non-empty link"),
            ({**entry, "reason": ""}, "needs a non-empty reason"),
            ({**entry, "reason": 1}, "needs a non-empty reason"),
            ({**entry, "link": "http://a.example/"},
             "a plain HTTP link cannot be excluded: http://a.example/"),
        ]
        cases += [(json.dumps({"schema_version": 1, "exclusions": [data]}),
                   f"{self.NAME}: exclusion 1: {problem}") for data, problem in entry_cases]
        cases.append((json.dumps({"schema_version": 1, "exclusions": [entry, entry]}),
                      f"{self.NAME}: exclusion 2: README.md: missing.md is listed twice"))
        for data, expected in cases:
            with self.subTest(data=data):
                errors, _, _ = run_check({"README.md": "[a](missing.md)\n", self.NAME: data})
                self.assertTrue(any(error.startswith(expected) for error in errors), errors)

    def test_plain_http_link_cannot_be_excluded(self) -> None:
        errors, _, _ = run_check(write_exclusions(
            {"README.md": "[a](http://a.example/)\n"},
            {"path": "README.md", "link": "http://a.example/", "reason": "Old site"}))
        self.assertEqual(errors, [
            "README.md: plain HTTP link, use https: http://a.example/",
            f"{self.NAME}: exclusion 1: a plain HTTP link cannot be excluded: http://a.example/",
        ])


class WikiTests(unittest.TestCase):
    PUBLIC = {"https://github.com/owner/repo": ("PASS", "https://github.com/owner/repo")}

    def test_wiki_pages_are_requested_as_raw_markdown(self) -> None:
        page = check_links.wiki_page
        self.assertEqual(page("https://github.com/owner/repo/wiki/API-Reference", "owner/repo"),
                         "https://raw.githubusercontent.com/wiki/owner/repo/API-Reference.md")
        self.assertEqual(page("https://github.com/Owner/Repo/wiki", "owner/repo"),
                         "https://raw.githubusercontent.com/wiki/owner/repo/Home.md")
        self.assertEqual(page("https://github.com/owner/repo/wiki/", "owner/repo"),
                         "https://raw.githubusercontent.com/wiki/owner/repo/Home.md")
        for url in ("https://github.com/owner/repo/issues/11",
                    "https://github.com/owner/repo/wiki/_pages",
                    "https://github.com/owner/repo/wiki/Page/_history",
                    "https://github.com/owner/repo/wikis",
                    "https://github.com/owner/repo-other/wiki/Page",
                    "https://example.com/wiki/Page"):
            self.assertIsNone(page(url, "owner/repo"), url)
        self.assertIsNone(page("https://github.com/owner/repo/wiki/Page", ""))

    def run_external(self, results: dict[str, tuple[str, str]],
                     urls: list[str]) -> tuple[list[str], list[str], str]:
        output = io.StringIO()
        with unittest.mock.patch.object(check_links, "check_external",
                                        side_effect=lambda url: results[url]) as fake, \
                contextlib.redirect_stdout(output):
            errors = check_links.external_errors(urls, "owner/repo")
        return errors, [call.args[0] for call in fake.call_args_list], output.getvalue()

    def test_public_repository_wiki_pages_are_checked(self) -> None:
        raw = "https://raw.githubusercontent.com/wiki/owner/repo/"
        results = {**self.PUBLIC, f"{raw}Home.md": ("PASS", f"{raw}Home.md"),
                   f"{raw}Gone.md": ("ERROR", f"{raw}Gone.md: HTTP 404")}
        errors, requested, output = self.run_external(results, [
            "https://github.com/owner/repo/wiki", "https://github.com/owner/repo/wiki/Gone",
            "https://github.com/owner/repo/issues/1"])
        self.assertEqual(errors, [f"https://github.com/owner/repo/wiki/Gone (requested as "
                                  f"{raw}Gone.md): HTTP 404"])
        self.assertEqual(sorted(requested), sorted(results))
        self.assertIn("PASS: https://github.com/owner/repo/wiki (requested as "
                      f"{raw}Home.md)", output)
        self.assertIn("SKIP: https://github.com/owner/repo/issues/1:", output)

    def test_private_or_unreachable_repository_wiki_is_not_requested(self) -> None:
        wiki = "https://github.com/owner/repo/wiki/Page"
        for answer, line in ((("ERROR", "https://github.com/owner/repo: HTTP 404"), "SKIP: "),
                             (("WARN", "https://github.com/owner/repo: unavailable for "
                                       "verification (HTTP 503)"), "WARN: ")):
            with self.subTest(answer=answer):
                errors, requested, output = self.run_external(
                    {"https://github.com/owner/repo": answer}, [wiki])
                self.assertEqual(errors, [])
                self.assertEqual(requested, ["https://github.com/owner/repo"])
                self.assertIn(line + wiki + ":", output)

    def test_repository_is_not_requested_without_wiki_links(self) -> None:
        errors, requested, _ = self.run_external(
            {"https://a.example/": ("PASS", "https://a.example/")},
            ["https://a.example/", "https://github.com/owner/repo/blob/HEAD/x.md"])
        self.assertEqual((errors, requested), ([], ["https://a.example/"]))


def write_exceptions(root: Path, data: object) -> None:
    path = root / check_links.EXCEPTIONS
    path.parent.mkdir(parents=True, exist_ok=True)
    path.write_text(data if isinstance(data, str) else json.dumps(data), encoding="utf-8")


def exceptions_file(*entries: object) -> dict[str, object]:
    return {"schema_version": 1, "exceptions": list(entries)}


ENTRY: dict[str, object] = {"url": "https://slow.example/page",
                            "reason": "Answers 403 to automated requests",
                            "expires": "2027-01-31"}


class ExceptionTests(unittest.TestCase):
    def test_absent_file_means_no_exceptions(self) -> None:
        with tempfile.TemporaryDirectory() as directory:
            self.assertEqual(check_links.load_exceptions(Path(directory)), ({}, []))

    def test_valid_file_is_loaded_and_its_urls_must_be_linked(self) -> None:
        with tempfile.TemporaryDirectory() as directory:
            root = Path(directory)
            write_exceptions(root, exceptions_file(ENTRY, {**ENTRY, "url": "https://gone.example/"}))
            (root / "README.md").write_text("[slow](https://slow.example/page#part)\n",
                                            encoding="utf-8")
            exceptions, problems = check_links.load_exceptions(root)
            errors, _, _ = check_links.check(root)
        self.assertEqual(problems, [])
        self.assertEqual(exceptions["https://slow.example/page"], check_links.LinkException(
            "Answers 403 to automated requests", date(2027, 1, 31)))
        self.assertEqual(errors, [".github/link-exceptions.json: https://gone.example/ is not "
                                  "linked from any Markdown file; remove its exception"])

    def test_malformed_files_fail(self) -> None:
        name = ".github/link-exceptions.json"
        cases: list[tuple[object, str]] = [
            ("{", f"{name}: not valid UTF-8 JSON"),
            ({"exceptions": []}, f"{name}: needs exactly schema_version 1 and an exceptions list"),
            ({"schema_version": True, "exceptions": []}, f"{name}: needs exactly schema_version 1"),
            ({"schema_version": 1, "exceptions": {}}, f"{name}: needs exactly schema_version 1"),
            ({**exceptions_file(), "extra": 1}, f"{name}: needs exactly schema_version 1"),
            (exceptions_file("x"), f"{name}: exception 1: needs exactly url, reason and expires"),
            (exceptions_file({**ENTRY, "note": ""}), f"{name}: exception 1: needs exactly url"),
            (exceptions_file({**ENTRY, "url": "http://slow.example/page"}),
             f"{name}: exception 1: url must be an https URL without a fragment"),
            (exceptions_file({**ENTRY, "url": "https://slow.example/page#x"}),
             f"{name}: exception 1: url must be an https URL without a fragment"),
            (exceptions_file(ENTRY, ENTRY),
             f"{name}: exception 2: https://slow.example/page is listed twice"),
            (exceptions_file({**ENTRY, "reason": " "}),
             f"{name}: exception 1: https://slow.example/page needs a reason"),
            (exceptions_file({**ENTRY, "expires": "31/01/2027"}),
             f"{name}: exception 1: https://slow.example/page needs expires as YYYY-MM-DD"),
            (exceptions_file({**ENTRY, "expires": 20270131}),
             f"{name}: exception 1: https://slow.example/page needs expires as YYYY-MM-DD"),
            (exceptions_file({**ENTRY, "expires": "2027-02-30"}),
             f"{name}: exception 1: https://slow.example/page has an expires date that does not "
             "exist"),
        ]
        for data, expected in cases:
            with self.subTest(data=data), tempfile.TemporaryDirectory() as directory:
                root = Path(directory)
                write_exceptions(root, data)
                (root / "README.md").write_text("[slow](https://slow.example/page)\n",
                                                encoding="utf-8")
                errors, _, _ = check_links.check(root)
                self.assertEqual(len(errors), 1, errors)
                self.assertTrue(errors[0].startswith(expected), errors[0])

    def test_tolerated_answer_fails_once_its_exception_expires(self) -> None:
        exception = check_links.LinkException("Answers 403", date(2027, 1, 31))
        url = "https://slow.example/page"
        warn = ("WARN", f"{url}: unavailable for verification (HTTP 403)")
        for today, listed, expected in ((date(2027, 1, 30), True, []),
                                        (date(2027, 1, 31), True, []),
                                        (date(2027, 2, 1), True, [
                                            f"{url}: unavailable for verification (HTTP 403); its "
                                            "exception in .github/link-exceptions.json expired "
                                            "on 2027-01-31 (Answers 403)"]),
                                        (date(2099, 1, 1), False, [])):
            output = io.StringIO()
            with self.subTest(today=today, listed=listed), unittest.mock.patch.object(
                    check_links, "check_external", return_value=warn), \
                    contextlib.redirect_stdout(output):
                errors = check_links.external_errors(
                    [url], "", {url: exception} if listed else {}, today)
                self.assertEqual(errors, expected)
                if listed and not expected:
                    self.assertIn("tolerated until 2027-01-31 (Answers 403)", output.getvalue())

    def test_exception_does_not_hide_an_error(self) -> None:
        url = "https://gone.example/"
        with unittest.mock.patch.object(check_links, "check_external",
                                        return_value=("ERROR", f"{url}: HTTP 404")), \
                contextlib.redirect_stdout(io.StringIO()):
            errors = check_links.external_errors(
                [url], "", {url: check_links.LinkException("Gone", date(2099, 1, 1))},
                date(2027, 1, 1))
        self.assertEqual(errors, [f"{url}: HTTP 404"])


if __name__ == "__main__":
    unittest.main()
