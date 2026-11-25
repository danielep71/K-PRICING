"""End-to-end tests of .github/scripts/labels-sync.mjs against a local fake GitHub API.

The script's self-test covers its functions; these tests run its command line,
so that the order of its steps is covered too. They need Node.js.
"""

from __future__ import annotations

import functools
import json
import os
import shutil
import subprocess
import sys
import tempfile
import threading
import unittest
from http.server import BaseHTTPRequestHandler, ThreadingHTTPServer
from pathlib import Path
from typing import Any
from urllib.parse import parse_qs, urlsplit

ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(ROOT / "tools"))

import check_repository  # noqa: E402

SCRIPT = ROOT / ".github/scripts/labels-sync.mjs"
MANIFEST = ROOT / ".github/labels.json"
NODE = shutil.which("node")


class FakeGitHub:
    """The repository's labels, the issues that carry them, and every request made.

    issue_labels maps a label name to the numbers of the issues and pull
    requests that carry it.
    """

    def __init__(self, labels: list[dict[str, Any]], issue_labels: dict[str, list[int]]) -> None:
        self.labels = labels
        self.issue_labels = issue_labels
        self.requests: list[tuple[str, str]] = []

    def issues(self) -> list[dict[str, Any]]:
        """Every issue and pull request, newest first, as GitHub lists them by default."""
        numbers = sorted({number for carried in self.issue_labels.values() for number in carried}, reverse=True)
        return [{"number": number,
                 "labels": [{"name": name} for name, carried in self.issue_labels.items() if number in carried]}
                for number in numbers]

    def answer(self, method: str, target: str) -> tuple[int, Any]:
        self.requests.append((method, target))
        url = urlsplit(target)
        query = parse_qs(url.query)
        if method == "GET" and url.path.endswith("/labels"):
            page = int(query.get("page", ["1"])[0])
            return 200, self.labels if page == 1 else []
        if method == "GET" and url.path.endswith("/issues"):
            items = self.issues()
            wanted = query.get("labels", [None])[0]
            if wanted is not None:
                items = [item for item in items
                         if any(label["name"].lower() == wanted.lower() for label in item["labels"])]
            if query.get("direction", ["desc"])[0] == "asc":
                items.reverse()
            per_page = int(query.get("per_page", ["30"])[0])
            page = int(query.get("page", ["1"])[0])
            return 200, items[(page - 1) * per_page:page * per_page]
        return 500, {"message": f"unexpected {method} {target}"}

    @property
    def writes(self) -> list[tuple[str, str]]:
        return [request for request in self.requests if request[0] != "GET"]


def serve(api: FakeGitHub) -> ThreadingHTTPServer:
    class Handler(BaseHTTPRequestHandler):
        def respond(self) -> None:
            status, body = api.answer(self.command, self.path)
            data = json.dumps(body).encode()
            self.send_response(status)
            self.send_header("Content-Type", "application/json")
            self.send_header("Content-Length", str(len(data)))
            self.end_headers()
            self.wfile.write(data)

        do_GET = do_POST = do_PATCH = do_DELETE = respond

        def log_message(self, *args: object) -> None:
            pass

    server = ThreadingHTTPServer(("127.0.0.1", 0), Handler)
    # A short poll, so that server.shutdown() returns at once rather than after
    # the default half second.
    threading.Thread(target=functools.partial(server.serve_forever, poll_interval=0.05),
                     daemon=True).start()
    return server


@unittest.skipIf(NODE is None and not os.environ.get("GITHUB_ACTIONS"), "Node.js is not on PATH")
class ReconcileOrderTests(unittest.TestCase):
    def setUp(self) -> None:
        manifest = json.loads(MANIFEST.read_text(encoding="utf-8"))
        self.assertTrue(manifest["prune"])
        # The resolved catalogue: the core plus the labels of every selected overlay.
        catalogue = list(manifest["core"])
        for name in manifest["selected_overlays"]:
            catalogue += manifest["overlays"][name]
        # Every catalogue label, plus one outside it that an issue still carries.
        live = [dict(label) for label in catalogue] + [
            {"name": "legacy", "color": "ABCDEF", "description": "outside the catalogue"}]
        self.api = FakeGitHub(live, issue_labels={"legacy": [1]})
        self.server = serve(self.api)
        self.addCleanup(self.server.server_close)
        self.addCleanup(self.server.shutdown)

    def run_script(self, *extra: str, manifest: Path = MANIFEST) -> subprocess.CompletedProcess[str]:
        with tempfile.TemporaryDirectory() as directory:
            env = {key: value for key, value in os.environ.items()
                   if key not in {"GITHUB_STEP_SUMMARY", "GITHUB_REPOSITORY"}}
            env.update(GITHUB_API_URL=f"http://127.0.0.1:{self.server.server_port}",
                       GITHUB_TOKEN="test-token", HOME=directory)
            return subprocess.run(
                [NODE or "node", str(SCRIPT), "--manifest", str(manifest), "--mode", "reconcile",
                 "--repository", "example/project", *extra],
                capture_output=True, text=True, encoding="utf-8", env=env, timeout=60, check=False)

    def test_a_selected_overlay_is_reconciled_with_the_core(self) -> None:
        manifest = json.loads(MANIFEST.read_text(encoding="utf-8"))
        manifest["overlays"] = {"regime": [
            {"name": "zz-regime", "color": "ABCDEF", "description": "overlay label"}]}
        with tempfile.TemporaryDirectory() as directory:
            path = Path(directory) / "labels.json"
            manifest["selected_overlays"] = ["regime"]
            path.write_text(json.dumps(manifest), encoding="utf-8")
            selected = self.run_script("--dry-run", manifest=path)
            manifest["selected_overlays"] = []
            path.write_text(json.dumps(manifest), encoding="utf-8")
            unselected = self.run_script("--dry-run", manifest=path)
        self.assertEqual(selected.returncode, 0, selected.stderr)
        self.assertIn("| create | `zz-regime` | missing live label |", selected.stdout)
        core = len(manifest["core"])
        self.assertIn(f"- Core / overlays defined / selected / resolved: **{core} / 1 / 1 / {core + 1}**",
                      selected.stdout)
        self.assertIn("- Selected overlays (1): `regime`", selected.stdout)
        self.assertEqual(unselected.returncode, 0, unselected.stderr)
        self.assertNotIn("zz-regime", unselected.stdout)
        self.assertIn(f"- Core / overlays defined / selected / resolved: **{core} / 1 / 0 / {core}**",
                      unselected.stdout)
        self.assertEqual(self.api.writes, [])

    def test_a_label_of_an_unselected_overlay_is_outside_the_catalogue(self) -> None:
        # The label is defined in the manifest, but the catalogue holds only the
        # core and the selected overlays: the remedy is to select the overlay.
        manifest = json.loads(MANIFEST.read_text(encoding="utf-8"))
        label = {"name": "zz-regime", "color": "ABCDEF", "description": "overlay label"}
        manifest["overlays"] = {"regime": [label]}
        manifest["selected_overlays"] = []
        self.api.labels = [entry for entry in self.api.labels if entry["name"] != "legacy"] + [label]
        self.api.issue_labels = {"zz-regime": [4]}
        with tempfile.TemporaryDirectory() as directory:
            path = Path(directory) / "labels.json"
            path.write_text(json.dumps(manifest), encoding="utf-8")
            result = self.run_script("--dry-run", manifest=path)
        self.assertEqual(result.returncode, 0, result.stderr)
        self.assertIn("| delete | `zz-regime` | outside the catalogue; still in use, a real run would refuse it; "
                      "used by #4 |", result.stdout)
        self.assertIn("Remove each label from its issues and pull requests, or add it to the catalogue "
                      "or select its overlay.", result.stdout)
        self.assertEqual(self.api.writes, [])

    def test_a_dry_run_previews_the_deletion_a_real_run_would_refuse(self) -> None:
        result = self.run_script("--dry-run")
        self.assertEqual(result.returncode, 0, result.stderr)
        self.assertIn("(dry run)", result.stdout)
        self.assertIn("- Deletions still in use: **1**", result.stdout)
        self.assertIn("| delete | `legacy` | outside the catalogue; still in use, a real run would refuse it; "
                      "used by #1 |", result.stdout)
        self.assertIn('A real run would refuse every change. Refusing to prune labels still used by an issue '
                      'or pull request: "legacy" (used by #1). No change was applied.', result.stdout)
        self.assertEqual(self.api.writes, [])
        self.assertEqual([target for _, target in self.api.requests if "/issues" in target],
                         ["/repos/example/project/issues?state=all&direction=asc&per_page=11&labels=legacy"],
                         "a dry run reads the usage of each planned deletion, once")

    def test_a_dry_run_names_up_to_ten_items_that_use_a_refused_deletion(self) -> None:
        # The fake lists newest first unless asked otherwise; the summary names
        # the lowest numbers, ascending, and "and others" only past ten.
        odd = [3, 5, 7, 9, 11, 13, 15, 17, 19, 21]
        ten = "#3, #5, #7, #9, #11, #13, #15, #17, #19, #21"
        for numbers, sample in (([21, 4], "#4, #21"), (odd, ten), ([*odd, 23], f"{ten} and others"),
                                ([*odd, 23, 40, 2000], f"{ten} and others")):
            with self.subTest(uses=len(numbers)):
                self.api.issue_labels = {"legacy": numbers, "P2": [1, 2, 22]}
                self.api.requests.clear()
                result = self.run_script("--dry-run")
                self.assertEqual(result.returncode, 0, result.stderr)
                self.assertIn("| delete | `legacy` | outside the catalogue; still in use, a real run would "
                              f"refuse it; used by {sample} |", result.stdout)
                self.assertIn(f'request: "legacy" (used by {sample}). No change was applied.', result.stdout)
                self.assertEqual(result.stdout.count("and others"), 2 if len(numbers) > 10 else 0)
                self.assertEqual(self.api.writes, [])
                self.assertEqual(len([target for _, target in self.api.requests if "/issues" in target]), 1,
                                 "a plain name costs one request")

    def test_a_dry_run_collects_the_items_of_a_comma_name_from_every_page(self) -> None:
        # 150 items, read in two pages of 100: the comma name is on both.
        self.api.labels = [label for label in self.api.labels if label["name"] != "legacy"] + [
            {"name": "a,b", "color": "ABCDEF", "description": "outside the catalogue"}]
        self.api.issue_labels = {"P2": list(range(1, 151)), "A,B": [140, 9, 120]}
        result = self.run_script("--dry-run")
        self.assertEqual(result.returncode, 0, result.stderr)
        self.assertIn("| delete | `a,b` | outside the catalogue; still in use, a real run would refuse it; "
                      "used by #9, #120, #140 |", result.stdout)
        self.assertEqual([target for _, target in self.api.requests if "/issues" in target],
                         ["/repos/example/project/issues?state=all&per_page=100&page=1",
                          "/repos/example/project/issues?state=all&per_page=100&page=2"])
        self.assertEqual(self.api.writes, [])

    def test_a_dry_run_without_a_deletion_reads_no_usage(self) -> None:
        self.api.labels = [label for label in self.api.labels if label["name"] != "legacy"]
        result = self.run_script("--dry-run")
        self.assertEqual(result.returncode, 0, result.stderr)
        self.assertIn("- Deletions still in use: **0**", result.stdout)
        self.assertNotIn("would refuse", result.stdout)
        self.assertEqual(self.api.writes, [])
        self.assertFalse([target for _, target in self.api.requests if "/issues" in target])

    def test_reconcile_refuses_to_prune_a_label_in_use_before_any_write(self) -> None:
        result = self.run_script()
        self.assertEqual(result.returncode, 1)
        self.assertIn('Refusing to prune labels still used by an issue or pull request: "legacy" (used by #1).',
                      result.stderr)
        self.assertIn('"legacy" (used by #1)', result.stdout, "the failure summary names the items too")
        self.assertEqual(self.api.writes, [])


@unittest.skipIf(NODE is None and not os.environ.get("GITHUB_ACTIONS"), "Node.js is not on PATH")
class IntegrityCommandTests(unittest.TestCase):
    """The label checks of Repository integrity fail on a defective script or catalogue.

    Each runs the check's own command in a copy of .github/ in which one
    thing is broken.
    """

    @staticmethod
    def command(name: str) -> list[str]:
        context = check_repository.Context(ROOT)
        check = next(check for check in check_repository.checks(context) if check.name == name)
        assert check.command is not None
        return [NODE or "node", *check.command[1:]]

    def run_check(self, name: str, path: str = "", find: str = "",
                  replace: str = "") -> subprocess.CompletedProcess[str]:
        """Run a check in a copy of .github/ where the text find in path is replaced."""
        with tempfile.TemporaryDirectory() as directory:
            root = Path(directory)
            shutil.copytree(ROOT / ".github/scripts", root / ".github/scripts")
            shutil.copy(MANIFEST, root / ".github/labels.json")
            if path:
                target = root / path
                text = target.read_text(encoding="utf-8")
                self.assertEqual(text.count(find), 1, f"{find!r} in {path}")
                target.write_text(text.replace(find, replace), encoding="utf-8")
            env = {key: value for key, value in os.environ.items()
                   if key not in {"GITHUB_STEP_SUMMARY", "GITHUB_REPOSITORY", "GITHUB_TOKEN"}}
            return subprocess.run(self.command(name), cwd=root, capture_output=True, text=True,
                                  encoding="utf-8", env=env, timeout=120, check=False)

    def test_the_catalogue_check_rejects_an_invalid_label(self) -> None:
        self.assertEqual(self.run_check("labels-catalogue").returncode, 0)
        manifest = json.loads(MANIFEST.read_text(encoding="utf-8"))
        color = manifest["core"][0]["color"]
        result = self.run_check("labels-catalogue", ".github/labels.json",
                                f'"color": "{color}"', f'"color": "#{color.lower()}"')
        self.assertEqual(result.returncode, 1, result.stdout)
        self.assertIn("core[0].color must be six uppercase hexadecimal characters without #",
                      result.stderr)

    def test_the_sync_self_test_fails_when_validation_accepts_a_bad_colour(self) -> None:
        self.assertEqual(self.run_check("labels-sync-self-test").returncode, 0)
        result = self.run_check("labels-sync-self-test", ".github/scripts/labels-sync.mjs",
                                'typeof label.color !== "string" || '
                                '!/^[0-9A-F]{6}$/.test(label.color)',
                                'typeof label.color !== "string"')
        self.assertEqual(result.returncode, 1, result.stdout)
        self.assertNotIn("SELF-TEST PASS", result.stdout)

    def test_the_drift_self_test_fails_when_a_disagreement_is_ignored(self) -> None:
        self.assertEqual(self.run_check("labels-drift-self-test").returncode, 0)
        result = self.run_check("labels-drift-self-test", ".github/scripts/labels-drift.mjs",
                                "const agree = JSON.stringify(reported) === "
                                "JSON.stringify(planned);", "const agree = true;")
        self.assertNotEqual(result.returncode, 0, result.stdout)
        self.assertNotIn("SELF-TEST PASS", result.stdout)


if __name__ == "__main__":
    unittest.main()
