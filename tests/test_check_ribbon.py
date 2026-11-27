"""Tests for tools/check_ribbon.py.

The repository's Ribbon must pass. Each negative test copies the XML and the two VBA modules
into a temporary directory, breaks one rule and expects the check to fail and say why.
"""

from __future__ import annotations

import contextlib
import io
import shutil
import sys
import tempfile
import unittest
from collections.abc import Iterator
from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(ROOT / "tools"))

import check_ribbon  # noqa: E402

FILES = (check_ribbon.CUSTOM_UI, check_ribbon.CALLBACKS, check_ribbon.ACTIONS)


@contextlib.contextmanager
def copy_root(edits: dict[str, tuple[str, str]] | None = None) -> Iterator[Path]:
    """A directory with the Ribbon's three files, each edit replacing one text in one file."""
    with tempfile.TemporaryDirectory(prefix="ribbon-") as directory:
        root = Path(directory)
        for name in FILES:
            (root / name).parent.mkdir(parents=True, exist_ok=True)
            shutil.copy(ROOT / name, root / name)
        for name, (old, new) in (edits or {}).items():
            text = (root / name).read_bytes().decode("ascii")
            if old not in text:
                raise AssertionError(f"{name} has no {old!r}")
            (root / name).write_bytes(text.replace(old, new, 1).encode("ascii"))
        yield root


def run(root: Path) -> tuple[int, str]:
    with contextlib.redirect_stdout(io.StringIO()) as output:
        status = check_ribbon.main(["--root", str(root)])
    return status, output.getvalue()


class RepositoryTest(unittest.TestCase):
    def test_the_repository_ribbon_passes(self) -> None:
        status, output = run(ROOT)
        self.assertEqual(status, 0, output)
        self.assertIn("ribbon: PASS", output)

    def test_every_static_action_is_an_action_text(self) -> None:
        _, uses, tags = check_ribbon.xml_findings(check_ribbon.load_xml(ROOT)[0])  # type: ignore[arg-type]
        self.assertEqual(sorted(tags), ["dev:tests", "setup:clear", "setup:register"])
        self.assertEqual({use.name for use in uses},
                         {"RibbonGetFunctions", "RibbonOnAction", "RibbonGetDevelopmentVisible"})

    def test_no_ribbon_passes_and_says_so(self) -> None:
        with tempfile.TemporaryDirectory() as directory:
            status, output = run(Path(directory))
        self.assertEqual(status, 0)
        self.assertIn("nothing to check", output)


class RefusalTest(unittest.TestCase):
    def assert_fails(self, edits: dict[str, tuple[str, str]], message: str) -> None:
        with copy_root(edits) as root:
            status, output = run(root)
        self.assertEqual(status, 1, output)
        self.assertIn(message, output)

    def test_a_missing_callback_fails(self) -> None:
        self.assert_fails({check_ribbon.CUSTOM_UI: ('getContent="RibbonGetFunctions"',
                                                    'getContent="RibbonGetMissing"')},
                          "names RibbonGetMissing, which is no procedure")

    def test_a_callback_with_the_wrong_signature_fails(self) -> None:
        self.assert_fails({check_ribbon.CALLBACKS: (
            "Public Sub RibbonGetFunctions(ByVal Control As IRibbonControl, ByRef Content As Variant)",
            "Public Sub RibbonGetFunctions(ByVal Control As IRibbonControl, ByVal Content As Variant)")},
            "must be ByRef")

    def test_a_private_callback_fails(self) -> None:
        self.assert_fails({check_ribbon.CALLBACKS: ("Public Sub RibbonOnAction(",
                                                    "Private Sub RibbonOnAction(")},
                          "must be a Public Sub")

    def test_an_unused_ribbon_sub_fails(self) -> None:
        self.assert_fails({check_ribbon.CUSTOM_UI: (
            'getVisible="RibbonGetDevelopmentVisible"', "")},
            "RibbonGetDevelopmentVisible takes a Ribbon parameter, but no attribute")

    def test_a_menu_callback_that_does_not_exist_fails(self) -> None:
        self.assert_fails({check_ribbon.CALLBACKS: ('ON_ACTION             As String = "RibbonOnAction"',
                                                    'ON_ACTION             As String = "RibbonClick"')},
                          "names RibbonClick, which is no procedure")

    def test_an_unknown_action_tag_fails(self) -> None:
        self.assert_fails({check_ribbon.CUSTOM_UI: ('tag="setup:clear"', 'tag="setup:wipe"')},
                          "tag 'setup:wipe' is no action text")

    def test_an_action_without_a_tag_fails(self) -> None:
        self.assert_fails({check_ribbon.CUSTOM_UI: ('tag="dev:tests"', "")},
                          "has an onAction but no tag")

    def test_a_duplicate_or_foreign_id_fails(self) -> None:
        self.assert_fails({check_ribbon.CUSTOM_UI: ('id="KPR_ClearHelp"', 'id="KPR_RegisterHelp"')},
                          "control id 'KPR_RegisterHelp' is used 2 times")
        self.assert_fails({check_ribbon.CUSTOM_UI: ('id="KPR_RunTests"', 'id="RunTests"')},
                          "id 'RunTests' does not match")

    def test_a_custom_image_fails(self) -> None:
        self.assert_fails({check_ribbon.CUSTOM_UI: ('imageMso="Delete"', 'image="delete.png"')},
                          "uses image; use a built-in imageMso icon")

    def test_another_namespace_or_broken_xml_fails(self) -> None:
        self.assert_fails({check_ribbon.CUSTOM_UI: ("office/2009/07/customui", "office/2006/01/customui")},
                          "not customUI of")
        self.assert_fails({check_ribbon.CUSTOM_UI: ("</customUI>", "")}, "is not well-formed XML")
        self.assert_fails({check_ribbon.CUSTOM_UI: ("<customUI ", '<!DOCTYPE x [<!ENTITY e "y">]><customUI ')},
                          "declares a DOCTYPE or an entity")


if __name__ == "__main__":
    unittest.main()
