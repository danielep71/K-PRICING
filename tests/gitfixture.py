"""Git repositories for the tests, isolated from the user's Git configuration.

Every command runs with ``GIT_CONFIG_GLOBAL`` set to the null device and
``GIT_CONFIG_NOSYSTEM``, so that no global or system setting (a signing key, a
hook path, autocrlf, a default branch name) changes what a test sees, and with
a fixed author, no signing and no line-end conversion. :func:`init` writes the
same settings into the new repository, so that a tool a test runs in it commits
and tags as the fixture does.
"""

from __future__ import annotations

import os
import subprocess
from pathlib import Path

NAME = "Fixture"
EMAIL = "fixture@example.invalid"
SETTINGS = (("user", "name", NAME), ("user", "email", EMAIL), ("commit", "gpgsign", "false"),
            ("tag", "gpgsign", "false"), ("core", "autocrlf", "false"))
OPTIONS = tuple(item for section, key, value in SETTINGS
                for item in ("-c", f"{section}.{key}={value}"))


def environment() -> dict[str, str]:
    """The current environment, without the global and system Git configuration."""
    return {**os.environ, "GIT_CONFIG_GLOBAL": os.devnull, "GIT_CONFIG_NOSYSTEM": "1"}


def git(root: Path, *args: str, input: str | None = None) -> str:
    """The standard output of ``git -C root args`` with the fixture's settings;
    CalledProcessError when it fails."""
    return subprocess.run(["git", "-C", str(root), *OPTIONS, *args], input=input,
                          env=environment(), capture_output=True, text=True,
                          check=True).stdout


def init(root: Path, *options: str, branch: str = "main") -> Path:
    """A new repository at ``root``, created if missing, on ``branch``, with the
    fixture's settings in its own configuration. ``options`` go to git init,
    such as --bare."""
    root.mkdir(parents=True, exist_ok=True)
    git(root, "init", "-q", "-b", branch, *options)
    config = root / "config" if "--bare" in options else root / ".git" / "config"
    with config.open("a", encoding="utf-8") as handle:
        handle.writelines(f"[{section}]\n\t{key} = {value}\n" for section, key, value in SETTINGS)
    return root


def commit(root: Path, message: str = "Fixture", *options: str) -> None:
    """Stage every change and commit it; ``options`` go to git commit."""
    git(root, "add", "-A")
    git(root, "commit", "-q", "-m", message, *options)


def head(root: Path) -> str:
    """The SHA of HEAD."""
    return git(root, "rev-parse", "HEAD").strip()
