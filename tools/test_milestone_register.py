#!/usr/bin/env python3
"""Frozen-snapshot tests for tools/milestone_register.py: rendering, bounded replacement and results."""

from __future__ import annotations

import contextlib
import copy
import email.message
import io
import json
import re
import tempfile
import unittest
import urllib.error
from collections.abc import Callable
from pathlib import Path
from typing import Any
from unittest.mock import patch

import milestone_register as mr

FIXTURES = mr.ROOT / "tests" / "fixtures" / "milestone_register"
SNAPSHOT: dict[str, Any] = json.loads((FIXTURES / "snapshot.json").read_text(encoding="utf-8"))
EXPECTED = (FIXTURES / "expected.md").read_text(encoding="utf-8")
REPOSITORY = "example/repo"
WORKFLOWS = mr.ROOT / ".github" / "workflows"


def document(region: str = "") -> str:
    return f"# Register\n\nIntroduction kept.\n\n{mr.BEGIN}\n{region}\n{mr.END}\n\nClosing text kept.\n"


class RenderTests(unittest.TestCase):
    def test_render_is_byte_identical_to_the_frozen_output(self) -> None:
        title, records = mr.fetch_register(mr.snapshot_fetch(SNAPSHOT), REPOSITORY, "v0.0.4")
        self.assertEqual(mr.render(title, records), EXPECTED)
        self.assertEqual(mr.render(title, records), mr.render(*mr.fetch_register(
            mr.snapshot_fetch(copy.deepcopy(SNAPSHOT)), REPOSITORY, "v0.0.4")))

    def test_issues_ascend_and_pull_requests_are_left_out(self) -> None:
        _, records = mr.fetch_register(mr.snapshot_fetch(SNAPSHOT), REPOSITORY, "v0.0.4")
        self.assertEqual([record["number"] for record in records], [46, 48, 52])
        self.assertNotIn("Pull request in the milestone", EXPECTED)

    def test_body_is_complete_with_normalized_line_ends_and_a_longer_fence(self) -> None:
        self.assertNotIn("\r", EXPECTED)
        self.assertFalse([line for line in EXPECTED.split("\n") if line != line.rstrip()])
        self.assertIn("`````text\nBody with ````four```` backticks\n````\ninside\n````\n`````", EXPECTED)
        self.assertIn("### #52 Certify \\| publish", EXPECTED)


class RegionTests(unittest.TestCase):
    def run_tool(self, *argv: str) -> tuple[int, str]:
        output = io.StringIO()
        with contextlib.redirect_stdout(output), contextlib.redirect_stderr(output):
            code = mr.main(list(argv))
        return code, output.getvalue()

    def setUp(self) -> None:
        temp = tempfile.TemporaryDirectory()
        self.addCleanup(temp.cleanup)
        self.folder = Path(temp.name)
        self.document = self.folder / "REGISTER.md"
        self.document.write_text(document("stale"), encoding="utf-8")
        self.snapshot = self.folder / "snapshot.json"
        self.save(SNAPSHOT)

    def save(self, data: dict[str, Any]) -> None:
        self.snapshot.write_text(json.dumps(data), encoding="utf-8")

    def tool(self, mode: str) -> tuple[int, str]:
        return self.run_tool(mode, "--document", str(self.document), "--repository", REPOSITORY,
                             "--snapshot", str(self.snapshot))

    def test_write_replaces_only_the_bounded_region_and_is_idempotent(self) -> None:
        self.assertEqual(self.tool("--write")[0], mr.CURRENT)
        written = self.document.read_text(encoding="utf-8")
        self.assertEqual(written, document(EXPECTED))
        self.assertTrue(written.startswith("# Register\n\nIntroduction kept.\n\n"))
        self.assertTrue(written.endswith("\n\nClosing text kept.\n"))
        self.assertEqual(self.tool("--write")[0], mr.CURRENT)
        self.assertEqual(self.document.read_text(encoding="utf-8"), written)
        self.assertEqual(self.tool("--check"), (mr.CURRENT, self.tool("--check")[1]))

    def test_a_live_edit_is_drift_until_regeneration(self) -> None:
        self.tool("--write")
        cases: tuple[tuple[str, Callable[[list[dict[str, Any]]], None]], ...] = (
            ("body", lambda issues: issues[0].update(body="## Exit gate\n- [x] Tag")),
            ("title", lambda issues: issues[0].update(title="Certify and publish")),
            ("labels", lambda issues: issues[0]["labels"].append({"name": "ci"})),
            ("state", lambda issues: issues[3].update(state="closed")),
        )
        for field, edit in cases:
            with self.subTest(field=field):
                changed = copy.deepcopy(SNAPSHOT)
                edit(changed["issues"])
                self.save(changed)
                code, output = self.tool("--check")
                self.assertEqual(code, mr.DRIFT)
                self.assertIn(f"field {field} differs", output)
                self.assertIn(mr.REGENERATE, output)
                self.assertRegex(output, r"DRIFT: (issue #\d+|the milestone summary)")
        added = copy.deepcopy(SNAPSHOT)
        added["issues"].append(dict(added["issues"][3], number=99, title="Added"))
        self.save(added)
        code, output = self.tool("--check")
        self.assertEqual(code, mr.DRIFT)
        self.assertEqual(self.tool("--write")[0], mr.CURRENT)
        self.assertEqual(self.tool("--check")[0], mr.CURRENT)

    def test_body_text_that_looks_like_a_table_is_still_the_body(self) -> None:
        self.tool("--write")
        changed = copy.deepcopy(SNAPSHOT)
        changed["issues"][0]["body"] = "| State | closed |\n" + changed["issues"][0]["body"]
        self.save(changed)
        code, output = self.tool("--check")
        self.assertEqual(code, mr.DRIFT)
        self.assertIn("issue #52: field body differs", output)

    def test_a_missing_or_doubled_region_is_unverified(self) -> None:
        for text in ("# No markers\n", document() + mr.BEGIN + "\n", f"{mr.END}\n{mr.BEGIN}\n"):
            with self.subTest(text=text[:20]):
                self.document.write_text(text, encoding="utf-8")
                self.assertEqual(self.tool("--check")[0], mr.UNVERIFIED)
                self.assertEqual(self.tool("--write")[0], mr.UNVERIFIED)
                self.assertEqual(self.document.read_text(encoding="utf-8"), text)

    def test_a_marker_line_in_a_body_is_refused_before_writing(self) -> None:
        self.tool("--write")
        written = self.document.read_text(encoding="utf-8")
        for marker in (mr.BEGIN, mr.END):
            with self.subTest(marker=marker[:20]):
                changed = copy.deepcopy(SNAPSHOT)
                changed["issues"][0]["body"] = f"Format:\r\n{marker}\r\nend"
                self.save(changed)
                for mode in ("--write", "--check"):
                    code, output = self.tool(mode)
                    self.assertEqual(code, mr.UNVERIFIED)
                    self.assertIn("issue #52: field body has a line equal to a generated-region marker", output)
                self.assertEqual(self.document.read_text(encoding="utf-8"), written)
        self.save(SNAPSHOT)
        self.assertEqual(self.tool("--check")[0], mr.CURRENT)

    def test_summary_names_the_result(self) -> None:
        self.tool("--write")
        summary = self.folder / "summary.md"
        self.run_tool("--check", "--document", str(self.document), "--repository", REPOSITORY,
                      "--snapshot", str(self.snapshot), "--summary", str(summary))
        self.assertIn("Result: **CURRENT**", summary.read_text(encoding="utf-8"))
        self.assertIn("makes no claim", summary.read_text(encoding="utf-8"))
        for code, word in ((mr.DRIFT, "DRIFT"), (mr.UNVERIFIED, "UNVERIFIED")):
            self.assertIn(f"Result: **{word}**", mr.summary(code, "detail"))


class UnverifiedTests(unittest.TestCase):
    def check(self, fetch: mr.Fetch) -> int:
        with tempfile.TemporaryDirectory() as temp:
            path = Path(temp) / "REGISTER.md"
            path.write_text(document(EXPECTED), encoding="utf-8")
            return mr.run("check", path, fetch, REPOSITORY, "v0.0.4")[0]

    def http_error(self, code: int, remaining: str | None = None) -> urllib.error.HTTPError:
        headers = email.message.Message()
        if remaining is not None:
            headers["x-ratelimit-remaining"] = remaining
        return urllib.error.HTTPError("https://api.github.com/x", code, "error", headers, None)

    def test_api_authentication_rate_limit_and_network_failures_are_unverified(self) -> None:
        failures: list[BaseException] = [
            self.http_error(401), self.http_error(403, "0"), self.http_error(404), self.http_error(429),
            self.http_error(500), urllib.error.URLError("no route"), TimeoutError("slow"), OSError("reset"),
        ]
        for failure in failures:
            with self.subTest(failure=repr(failure)):
                with patch("urllib.request.urlopen", side_effect=failure):
                    self.assertEqual(self.check(mr.github_fetch("token")), mr.UNVERIFIED)

    def test_malformed_responses_are_unverified(self) -> None:
        class Response(io.BytesIO):
            def __enter__(self) -> Response:
                return self

        with patch("urllib.request.urlopen", return_value=Response(b"<html>not json</html>")):
            self.assertEqual(self.check(mr.github_fetch(None)), mr.UNVERIFIED)
        malformed: list[dict[str, Any]] = [
            {"milestones": {}, "issues": []},
            {"milestones": [], "issues": []},
            {"milestones": SNAPSHOT["milestones"], "issues": [{"number": "52"}]},
            {"milestones": SNAPSHOT["milestones"], "issues": [dict(SNAPSHOT["issues"][0], labels="P2")]},
            {"milestones": SNAPSHOT["milestones"], "issues": [dict(SNAPSHOT["issues"][0], milestone=None)]},
            {"milestones": SNAPSHOT["milestones"], "issues": [dict(SNAPSHOT["issues"][0], body=7)]},
            {"milestones": SNAPSHOT["milestones"], "issues": ["not an object"]},
        ]
        for data in malformed:
            with self.subTest(data=str(data)[:60]):
                self.assertEqual(self.check(mr.snapshot_fetch(data)), mr.UNVERIFIED)


class WorkflowBoundaryTests(unittest.TestCase):
    def test_the_deterministic_static_gate_never_reads_live_issues(self) -> None:
        static = (WORKFLOWS / "static-checks.yml").read_text(encoding="utf-8")
        self.assertNotRegex(static, r"(?<!test_)milestone_register\.py")
        self.assertIn("tools/test_milestone_register.py", static, "the frozen tests run in the gate")
        self.assertNotRegex(static, r"issues:\s*(read|write)")
        self.assertNotIn("api.github.com", static)

    def test_the_live_workflow_is_separate_read_only_and_pinned(self) -> None:
        text = (WORKFLOWS / "milestone-register.yml").read_text(encoding="utf-8")
        triggers = text.split("\non:\n", 1)[1].split("\npermissions:", 1)[0]
        self.assertNotIn("pull_request", triggers)
        for event in ("push:", "issues:", "workflow_dispatch:"):
            self.assertIn(event, triggers)
        self.assertIn("permissions: {}", text)
        self.assertEqual(re.findall(r"^      (\w[\w-]*): (\w+)$", text, re.M),
                         [("contents", "read"), ("issues", "read")])
        uses = re.findall(r"uses: (\S+)", text)
        self.assertTrue(uses)
        self.assertTrue(all(re.fullmatch(r"[\w./-]+@[0-9a-f]{40}", use) for use in uses), uses)
        self.assertIn("ref: ${{ github.event.repository.default_branch }}", text)
        self.assertIn("--check --summary", text)


if __name__ == "__main__":
    unittest.main()
