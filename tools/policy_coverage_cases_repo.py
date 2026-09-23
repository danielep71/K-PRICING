from __future__ import annotations

from pathlib import Path
from types import ModuleType
from typing import Callable

from policy_coverage_core import add_force, mutate_config

Case = tuple[str, str, str | None, Callable[[Path], None]]


def _register(cases: list[Case], name: str, rule: str, pattern: str | None = None):
    def decorator(function: Callable[[Path], None]) -> Callable[[Path], None]:
        cases.append((name, rule, pattern, function))
        return function

    return decorator


def _case(cases: list[Case]):
    def register(name: str, rule: str, pattern: str | None = None):
        return _register(cases, name, rule, pattern)

    return register


def _required_and_placeholder_cases(module: ModuleType) -> list[Case]:
    cases: list[Case] = []

    def template_mode(document: dict) -> None:
        document.update(
            mode="template",
            profile=None,
            repository="example/TEMPLATE-IDENTITY",
        )

    case = _case(cases)

    @case("required-file-not-tracked", "required-paths", "Required file is not tracked")
    def _(root: Path) -> None:
        def mutation(document: dict) -> None:
            document["required_paths"] = sorted(
                [*document["required_paths"], "missing.txt"], key=str.casefold
            )

        mutate_config(module, root, mutation)

    @case("required-directory-absent", "required-paths", "Required directory is absent")
    def _(root: Path) -> None:
        def mutation(document: dict) -> None:
            document["required_directories"] = sorted(
                [*document["required_directories"], "missing-dir"], key=str.casefold
            )

        mutate_config(module, root, mutation)

    @case("required-directory-empty", "required-paths", "has no tracked")
    def _(root: Path) -> None:
        (root / "empty-dir").mkdir()

        def mutation(document: dict) -> None:
            document["required_directories"] = sorted(
                [*document["required_directories"], "empty-dir"], key=str.casefold
            )

        mutate_config(module, root, mutation)

    @case("placeholder-executable", "placeholders", "prohibited in executable")
    def _(root: Path) -> None:
        token = "{{" + "REQUIRED_NOTE}}"
        module._write_fixture(root / "tools/check_repo.py", f"# {token}\n")

    @case("placeholder-template-unregistered", "placeholders", "not registered")
    def _(root: Path) -> None:
        mutate_config(module, root, template_mode)
        token = "{{" + "UNKNOWN_NOTE}}"
        module._write_fixture(root / "README.md", f"# Fixture\n\n{token}\n")

    @case("placeholder-template-unused", "placeholders", "unused")
    def _(root: Path) -> None:
        mutate_config(module, root, template_mode)

    return cases


def _dotfile_and_structured_cases(module: ModuleType) -> list[Case]:
    cases: list[Case] = []
    case = _case(cases)

    @case("dotfile-editor-unreadable", "dotfile-policy", "Cannot read policy")
    def _(root: Path) -> None:
        (root / ".editorconfig").unlink()

    @case("dotfile-vba-eol", "dotfile-policy", "VBA component section")
    def _(root: Path) -> None:
        text = (root / ".editorconfig").read_text(encoding="utf-8").replace(
            "end_of_line = crlf", "end_of_line = lf"
        )
        module._write_fixture(root / ".editorconfig", text)

    @case("dotfile-vba-final-newline", "dotfile-policy", "final newline")
    def _(root: Path) -> None:
        text = (root / ".editorconfig").read_text(encoding="utf-8")
        original = (
            "[*.{bas,cls,frm}]\ncharset = latin1\nend_of_line = crlf\n"
            "insert_final_newline = true"
        )
        replacement = (
            "[*.{bas,cls,frm}]\ncharset = latin1\nend_of_line = crlf\n"
            "insert_final_newline = false"
        )
        module._write_fixture(root / ".editorconfig", text.replace(original, replacement))

    @case("dotfile-gitattributes-vba", "dotfile-policy", "eol=crlf")
    def _(root: Path) -> None:
        text = (root / ".gitattributes").read_text(encoding="utf-8").replace(
            "*.bas text eol=crlf", "*.bas text eol=lf"
        )
        module._write_fixture(root / ".gitattributes", text)

    @case("dotfile-gitattributes-lf", "dotfile-policy", "eol=lf")
    def _(root: Path) -> None:
        text = (root / ".gitattributes").read_text(encoding="utf-8").replace(
            "*.json text eol=lf", "*.json text eol=crlf"
        )
        module._write_fixture(root / ".gitattributes", text)

    @case("dotfile-office-binary", "dotfile-policy", "text=unset")
    def _(root: Path) -> None:
        text = (root / ".gitattributes").read_text(encoding="utf-8").replace(
            "*.xlsm binary", "*.xlsm text"
        )
        module._write_fixture(root / ".gitattributes", text)

    @case("dotfile-ignore-probe", "dotfile-policy", "not ignored")
    def _(root: Path) -> None:
        text = (root / ".gitignore").read_text(encoding="utf-8").replace(".env\n", "")
        module._write_fixture(root / ".gitignore", text)

    @case("dotfile-env-example", "dotfile-policy", "env.example")
    def _(root: Path) -> None:
        module._write_fixture(
            root / ".gitignore",
            ".env*\n__pycache__/\n*.pem\n*.xlsm\ntest-results/\n~$*\n",
        )

    @case("structured-yaml-invalid-encoding", "structured-data", "Cannot decode YAML")
    def _(root: Path) -> None:
        path = root / ".github/workflows/bad-encoding.yml"
        path.write_bytes(b"\xff\n")
        add_force(module, root, ".github/workflows/bad-encoding.yml")

    return cases


def _markdown_and_text_cases(module: ModuleType) -> list[Case]:
    cases: list[Case] = []
    case = _case(cases)

    @case("markdown-escape", "markdown-links", "escapes the repository")
    def _(root: Path) -> None:
        module._write_fixture(root / "README.md", "# Fixture\n\n[Escape](../outside.md)\n")

    @case("markdown-untracked-target", "markdown-links", "not tracked")
    def _(root: Path) -> None:
        module._write_fixture(root / "README.md", "# Fixture\n\n[Untracked](docs/UNTRACKED.md)\n")
        module._write_fixture(root / "docs/UNTRACKED.md", "# Untracked\n")

    @case("markdown-empty-directory", "markdown-links", "no tracked content")
    def _(root: Path) -> None:
        (root / "docs/empty").mkdir(parents=True)
        module._write_fixture(root / "README.md", "# Fixture\n\n[Empty](docs/empty/)\n")

    @case("markdown-missing-anchor", "markdown-links", "heading does not exist")
    def _(root: Path) -> None:
        module._write_fixture(
            root / "README.md", "# Fixture\n\n[Bad anchor](docs/DETAILS.md#missing)\n"
        )

    @case("text-nul", "text-integrity", "NUL byte")
    def _(root: Path) -> None:
        (root / "README.md").write_bytes(b"# Fixture\n\x00\n")

    @case("text-invalid-encoding", "text-integrity", "invalid encoding")
    def _(root: Path) -> None:
        (root / "README.md").write_bytes(b"# Fixture\n\xff\n")

    @case("text-private-key", "text-integrity", "private-key material")
    def _(root: Path) -> None:
        marker = "-----BEGIN " + "PRIVATE KEY-----"
        module._write_fixture(root / "README.md", f"# Fixture\n\n{marker}\n")

    @case("text-github-token", "text-integrity", "GitHub token material")
    def _(root: Path) -> None:
        marker = "gh" + "p_" + "fixturetoken"
        module._write_fixture(root / "README.md", f"# Fixture\n\n{marker}\n")

    @case("text-aws-key", "text-integrity", "AWS access key")
    def _(root: Path) -> None:
        marker = "AK" + "IA" + "ABCDEFGHIJKLMNOP"
        module._write_fixture(root / "README.md", f"# Fixture\n\n{marker}\n")

    return cases


def _artifact_and_line_cases(module: ModuleType) -> list[Case]:
    cases: list[Case] = []
    case = _case(cases)

    @case("artifact-office-lock", "forbidden-artifacts", "Office lock file")
    def _(root: Path) -> None:
        module._write_fixture(root / "~$fixture.xlsx", b"lock")
        add_force(module, root, "~$fixture.xlsx")

    @case("artifact-office-binary", "forbidden-artifacts", "Office binary is not permitted")
    def _(root: Path) -> None:
        module._write_fixture(root / "fixture.xlsx", b"office")
        add_force(module, root, "fixture.xlsx")

    @case("artifact-env", "forbidden-artifacts", "Local environment or secret file")
    def _(root: Path) -> None:
        module._write_fixture(root / ".env", "SECRET=fixture\n")
        add_force(module, root, ".env")

    @case("artifact-private-directory", "forbidden-artifacts", "Private review material")
    def _(root: Path) -> None:
        module._write_fixture(root / "private/note.txt", "private\n")
        add_force(module, root, "private/note.txt")

    @case("artifact-private-token", "forbidden-artifacts", "Private review material")
    def _(root: Path) -> None:
        module._write_fixture(root / "docs/confidential-review.txt", "private\n")
        add_force(module, root, "docs/confidential-review.txt")

    @case("line-bom", "line-endings", "BOM")
    def _(root: Path) -> None:
        (root / "README.md").write_bytes(b"\xef\xbb\xbf# Fixture\n")

    @case("line-missing-final-newline", "line-endings", "end with a newline")
    def _(root: Path) -> None:
        (root / "README.md").write_bytes(b"# Fixture")

    @case("line-cross-platform-crlf", "line-endings", "Cross-platform text must use LF")
    def _(root: Path) -> None:
        module._write_fixture(root / "README.md", "# Fixture\n\nCRLF\n", crlf=True)

    @case("line-vba-lf", "line-endings", "Windows/VBA source must use CRLF")
    def _(root: Path) -> None:
        path = root / "src/modules/Quality.bas"
        text = path.read_bytes().decode("cp1252").replace("\r\n", "\n")
        path.write_bytes(text.encode("cp1252"))

    return cases


def repository_cases(module: ModuleType) -> list[Case]:
    return [
        *_required_and_placeholder_cases(module),
        *_dotfile_and_structured_cases(module),
        *_markdown_and_text_cases(module),
        *_artifact_and_line_cases(module),
    ]
