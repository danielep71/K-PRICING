#!/usr/bin/env python3
"""Validate retained KPR migration parity evidence without executing Excel."""

from __future__ import annotations

import argparse
import hashlib
import json
import re
import subprocess
import sys
import tempfile
from pathlib import Path
from typing import Any

SCHEMA_VERSION = 1
FROZEN_SOURCE_SHA = "f26450d1fa7b11261162e901dedba062f21c99a7"
REQUIRED_LOG_IDS = {
    "source-exact-compile",
    "destination-compile",
    "source-observations",
    "destination-observations",
    "destination-host-record",
}
REQUIRED_CLEANUP = {"host", "shape", "array"}
REQUIRED_OBSERVATIONS = {
    "direct/days-in-month",
    "direct/add-days",
    "direct/strict-date-error",
    "host/1900/diagnostic",
    "host/1900/value",
    "host/1900/propagated-na",
    "host/1904/diagnostic",
    "host/1904/value",
    "host/1904/propagated-na",
    "host/1900-again/diagnostic",
    "host/1900-again/value",
    "host/1900-again/propagated-na",
    "shape/row",
    "shape/column",
    "shape/rectangle",
    "shape/single",
    "shape/multi-area",
    "shape/beyond-usedrange",
    "shape/blank-b2",
    "shape/error-c3",
    "shape/value-b3",
    "shape/state-application",
    "shape/state-selection",
    "array/api",
    "array/1900-spill",
    "array/1900-row1",
    "array/1900-row2",
    "array/1900-row3",
    "array/1904-call",
    "array/1904-spill",
    "array/1904-neighbour-c2",
}
SHA40 = re.compile(r"^[0-9a-f]{40}$")
SHA256 = re.compile(r"^[0-9a-f]{64}$")


def require(condition: bool, message: str) -> None:
    if not condition:
        raise ValueError(message)


def exact_keys(value: Any, keys: set[str], label: str) -> dict[str, Any]:
    require(isinstance(value, dict), f"{label} must be an object")
    actual = set(value)
    require(actual == keys, f"{label} keys differ: expected {sorted(keys)}, got {sorted(actual)}")
    return value


def safe_file(base: Path, relative: str) -> Path:
    require(isinstance(relative, str) and bool(relative) and not Path(relative).is_absolute(), "unsafe evidence path")
    path = base / relative
    require(".." not in Path(relative).parts, "evidence path traversal is forbidden")
    require(path.resolve().is_relative_to(base.resolve()), "evidence path escapes bundle")
    require(path.is_file(), f"missing evidence file: {relative}")
    require(not any(part.is_symlink() for part in (path, *path.parents)), f"symlinked evidence path: {relative}")
    return path


def read_bound(base: Path, entry: dict[str, Any]) -> str:
    exact_keys(entry, {"id", "path", "sha256"}, "log entry")
    require(isinstance(entry["id"], str) and bool(entry["id"]), "log id must be nonempty")
    require(isinstance(entry["sha256"], str) and SHA256.fullmatch(entry["sha256"]) is not None,
            f"invalid SHA-256 for log {entry['id']}")
    path = safe_file(base, entry["path"])
    raw = path.read_bytes()
    require(hashlib.sha256(raw).hexdigest() == entry["sha256"], f"log digest mismatch: {entry['id']}")
    return raw.decode("utf-8-sig").replace("\r\n", "\n")


def parse_observations(
    text: str,
    side: str,
) -> tuple[list[tuple[str, str]], dict[str, str], dict[str, tuple[int, int]]]:
    observations: list[tuple[str, str]] = []
    seen: set[str] = set()
    cleanup: dict[str, str] = {}
    runner_summaries: dict[str, tuple[int, int]] = {}
    summary_pattern = re.compile(
        r"^KPR (host|shape|array) regression\s+checks:\s*(\d+)\s+failures:\s*(\d+)\s*$"
    )

    for raw in text.splitlines():
        if raw.startswith("OBS\t"):
            fields = raw.split("\t", 2)
            require(len(fields) == 3 and bool(fields[1]) and bool(fields[2]), f"{side}: malformed OBS record")
            require(fields[1] not in seen, f"{side}: duplicate observation id {fields[1]}")
            seen.add(fields[1])
            observations.append((fields[1], fields[2]))
        elif raw.startswith("CLEANUP\t"):
            fields = raw.split("\t", 2)
            require(len(fields) == 3 and bool(fields[1]) and bool(fields[2]), f"{side}: malformed CLEANUP record")
            require(fields[1] not in cleanup, f"{side}: duplicate cleanup record {fields[1]}")
            cleanup[fields[1]] = fields[2]
        else:
            summary = summary_pattern.fullmatch(raw.strip())
            if summary is not None:
                runner = summary.group(1)
                require(runner not in runner_summaries, f"{side}: duplicate {runner} runner summary")
                runner_summaries[runner] = (int(summary.group(2)), int(summary.group(3)))

    missing = REQUIRED_OBSERVATIONS - seen
    require(not missing,
            f"{side}: missing required v0.0.2 observation ids: {sorted(missing)}")
    payloads = dict(observations)
    require(payloads["array/api"] == "TEXT:SUPPORTED",
            f"{side}: dynamic-array API is not supported")
    require(payloads["shape/state-application"] == "UNCHANGED",
            f"{side}: application state changed during shape observations")
    require(payloads["shape/state-selection"] == "UNCHANGED",
            f"{side}: selection state changed during shape observations")
    require(set(cleanup) == REQUIRED_CLEANUP, f"{side}: cleanup runners differ from required set")
    for runner, status in cleanup.items():
        require(status == "PASS", f"{side}: cleanup {runner} is {status}")

    require(set(runner_summaries) == REQUIRED_CLEANUP,
            f"{side}: stateful runner summaries differ from required set")
    for runner, (checks, failures) in runner_summaries.items():
        require(checks > 0, f"{side}: {runner} runner reported no checks")
        require(failures == 0, f"{side}: {runner} runner failures={failures}")

    return observations, cleanup, runner_summaries


def checked_out_sha(root: Path) -> str:
    completed = subprocess.run(
        ["git", "-C", str(root), "rev-parse", "HEAD"],
        check=True,
        capture_output=True,
        text=True,
    )
    sha = completed.stdout.strip()
    require(SHA40.fullmatch(sha) is not None, "checked-out candidate SHA is invalid")

    status = subprocess.run(
        ["git", "-C", str(root), "status", "--porcelain", "--untracked-files=no"],
        check=True,
        capture_output=True,
        text=True,
    )
    require(not status.stdout.strip(), "checked-out candidate has modified tracked files")
    return sha


def require_compile_log(text: str, label: str, require_native_run: bool) -> None:
    records: dict[str, str] = {}
    for raw in text.splitlines():
        if "=" not in raw:
            continue
        key, value = raw.split("=", 1)
        key = key.strip().upper()
        value = value.strip().upper()
        if key in {"IMPORT", "COMPILE", "NATIVE_RUN"}:
            require(key not in records, f"{label}: duplicate {key} record")
            records[key] = value

    require(records.get("IMPORT") == "PASS", f"{label}: import did not pass")
    require(records.get("COMPILE") == "PASS", f"{label}: compile did not pass")
    if require_native_run:
        require(records.get("NATIVE_RUN") == "PASS", f"{label}: native run did not pass")


def validate_manifest(
    root: Path,
    manifest_path: Path,
    expected_destination_sha: str,
) -> dict[str, Any]:
    manifest = json.loads(manifest_path.read_text(encoding="utf-8"))
    exact_keys(
        manifest,
        {"schema_version", "source", "destination", "environment", "instrumentation", "logs", "known_differences"},
        "migration manifest",
    )
    require(manifest["schema_version"] == SCHEMA_VERSION, "unsupported migration manifest schema")

    source = exact_keys(manifest["source"], {"repository", "sha"}, "source")
    destination = exact_keys(manifest["destination"], {"repository", "sha"}, "destination")
    require(source["repository"] == "danielep71/KPR", "unexpected source repository")
    require(destination["repository"] == "danielep71/K-PRICING", "unexpected destination repository")
    for label, record in (("source", source), ("destination", destination)):
        require(isinstance(record["sha"], str) and SHA40.fullmatch(record["sha"]) is not None,
                f"invalid {label} SHA")
    require(source["sha"] == FROZEN_SOURCE_SHA, "source SHA differs from frozen KPR baseline")
    require(destination["sha"] == expected_destination_sha,
            "destination SHA differs from checked-out candidate")

    environment = exact_keys(
        manifest["environment"],
        {"excel_version", "excel_build", "office_bitness", "windows_build", "locale", "references", "macro_policy"},
        "environment",
    )
    require(environment["office_bitness"] == "64-bit",
            "exact frozen-source parity requires the documented 64-bit Office baseline")
    require(all(isinstance(environment[key], str) and environment[key].strip()
                for key in ("excel_version", "excel_build", "windows_build", "locale", "macro_policy")),
            "environment text fields must be nonempty")
    require(isinstance(environment["references"], list) and all(isinstance(x, str) and x for x in environment["references"]),
            "environment references must be a nonempty string list")

    differences = manifest["known_differences"]
    require(isinstance(differences, list), "known_differences must be an array")
    difference_issues: set[str] = set()
    for item in differences:
        exact_keys(
            item,
            {"issue", "scope", "source_behavior", "destination_behavior", "disposition"},
            "known difference",
        )
        require(all(isinstance(item[key], str) and item[key].strip() for key in item),
                "known-difference fields must be nonempty strings")
        require(item["issue"].startswith("https://github.com/danielep71/K-PRICING/issues/"),
                "known difference must bind a K-PRICING issue")
        require(item["issue"] not in difference_issues, "duplicate known-difference issue")
        difference_issues.add(item["issue"])
    require(
        "https://github.com/danielep71/K-PRICING/issues/32" in difference_issues,
        "known pillar range correction #32 must be explicit in migration evidence",
    )

    instrumentation = exact_keys(manifest["instrumentation"], {"path", "sha256"}, "instrumentation")
    require(instrumentation["path"] == "tests/modules/KPR_REGRESSION_TESTS.bas",
            "unexpected migration instrumentation path")
    require(isinstance(instrumentation["sha256"], str) and SHA256.fullmatch(instrumentation["sha256"]) is not None,
            "invalid instrumentation SHA-256")
    candidate = root / instrumentation["path"]
    require(candidate.is_file(), "instrumentation module is missing from candidate")
    require(hashlib.sha256(candidate.read_bytes()).hexdigest() == instrumentation["sha256"],
            "instrumentation digest differs from candidate")

    logs = manifest["logs"]
    require(isinstance(logs, list), "logs must be an array")
    ids = [entry.get("id") if isinstance(entry, dict) else None for entry in logs]
    require(all(isinstance(x, str) and x for x in ids), "every log needs an id")
    require(len(ids) == len(set(ids)), "duplicate log id")
    require(set(ids) == REQUIRED_LOG_IDS, f"log ids differ from required set: {sorted(REQUIRED_LOG_IDS)}")

    base = manifest_path.parent
    text_by_id = {entry["id"]: read_bound(base, entry) for entry in logs}

    require_compile_log(text_by_id["source-exact-compile"], "source exact compile", True)
    require_compile_log(text_by_id["destination-compile"], "destination compile", False)

    source_obs, _, source_summaries = parse_observations(text_by_id["source-observations"], "source")
    destination_obs, _, destination_summaries = parse_observations(
        text_by_id["destination-observations"], "destination"
    )
    require(source_obs == destination_obs, "source/destination serialized observations differ")
    require(source_summaries == destination_summaries,
            "source/destination stateful runner summaries differ")

    return {
        "schema_version": SCHEMA_VERSION,
        "status": "pass",
        "source_sha": source["sha"],
        "destination_sha": destination["sha"],
        "observations": len(source_obs),
        "known_differences": len(differences),
        "bound_logs": sorted(REQUIRED_LOG_IDS),
        "scope_note": (
            "Validates manifest identity, retained-file digests, observation parity and cleanup records; "
            "does not execute Excel or authenticate the human/runner that produced the files."
        ),
    }


def self_test() -> None:
    with tempfile.TemporaryDirectory(prefix="migration-evidence-") as tmp:
        base = Path(tmp)
        root = base / "candidate"
        bundle = base / "bundle"
        (root / "tests/modules").mkdir(parents=True)
        bundle.mkdir()
        instrumentation = b"synthetic migration observer\n"
        (root / "tests/modules/KPR_REGRESSION_TESTS.bas").write_bytes(instrumentation)

        payloads = {name: "TEXT:synthetic" for name in REQUIRED_OBSERVATIONS}
        payloads["array/api"] = "TEXT:SUPPORTED"
        payloads["shape/state-application"] = "UNCHANGED"
        payloads["shape/state-selection"] = "UNCHANGED"
        obs = "\n".join(
            [f"OBS\t{name}\t{payloads[name]}" for name in sorted(REQUIRED_OBSERVATIONS)]
            + [
                "CLEANUP\thost\tPASS",
                "KPR host regression  checks: 9  failures: 0",
                "CLEANUP\tshape\tPASS",
                "KPR shape regression  checks: 12  failures: 0",
                "CLEANUP\tarray\tPASS",
                "KPR array regression  checks: 7  failures: 0",
                "",
            ]
        )
        files = {
            "source-exact-compile": "IMPORT=PASS\nCOMPILE=PASS\nNATIVE_RUN=PASS\n",
            "destination-compile": "IMPORT=PASS\nCOMPILE=PASS\n",
            "source-observations": obs,
            "destination-observations": obs,
            "destination-host-record": '{"synthetic":"host record"}\n',
        }
        log_entries: list[dict[str, str]] = []
        for ident, content in files.items():
            log_path = f"{ident}.log"
            (bundle / log_path).write_text(content, encoding="utf-8")
            log_entries.append(
                {"id": ident, "path": log_path, "sha256": hashlib.sha256(content.encode()).hexdigest()}
            )

        manifest: dict[str, Any] = {
            "schema_version": 1,
            "source": {"repository": "danielep71/KPR", "sha": FROZEN_SOURCE_SHA},
            "destination": {"repository": "danielep71/K-PRICING", "sha": "b" * 40},
            "environment": {
                "excel_version": "16.0",
                "excel_build": "synthetic",
                "office_bitness": "64-bit",
                "windows_build": "synthetic",
                "locale": "en-US",
                "references": ["VBA", "Excel"],
                "macro_policy": "synthetic",
            },
            "instrumentation": {
                "path": "tests/modules/KPR_REGRESSION_TESTS.bas",
                "sha256": hashlib.sha256(instrumentation).hexdigest(),
            },
            "logs": log_entries,
            "known_differences": [
                {
                    "issue": "https://github.com/danielep71/K-PRICING/issues/32",
                    "scope": "pillar numerical-range classification",
                    "source_behavior": "oversized valid quantity can fall through as PILLAR_TOKEN_MALFORMED",
                    "destination_behavior": "PILLAR_AGGREGATE_RANGE",
                    "disposition": "accepted destination correctness fix; verify separately from equal-observation parity",
                }
            ],
        }
        manifest_file = bundle / "migration.json"
        manifest_file.write_text(json.dumps(manifest), encoding="utf-8")
        report = validate_manifest(root, manifest_file, "b" * 40)
        require(report["status"] == "pass", "positive self-test did not pass")

        degraded: dict[str, Any] = json.loads(manifest_file.read_text())
        degraded["source"]["sha"] = "a" * 40
        manifest_file.write_text(json.dumps(degraded), encoding="utf-8")
        try:
            validate_manifest(root, manifest_file, "b" * 40)
        except ValueError as error:
            require("frozen KPR baseline" in str(error),
                    "degraded source-identity case failed for wrong reason")
        else:
            raise RuntimeError("degraded source-identity self-test unexpectedly passed")

        degraded = json.loads(json.dumps(manifest))
        degraded["destination"]["sha"] = "c" * 40
        manifest_file.write_text(json.dumps(degraded), encoding="utf-8")
        try:
            validate_manifest(root, manifest_file, "b" * 40)
        except ValueError as error:
            require("checked-out candidate" in str(error),
                    "degraded destination-identity case failed for wrong reason")
        else:
            raise RuntimeError("degraded destination-identity self-test unexpectedly passed")

        degraded = json.loads(json.dumps(manifest))
        dest = next(x for x in degraded["logs"] if x["id"] == "destination-observations")
        changed = obs.replace(
            "OBS\tdirect/days-in-month\tTEXT:synthetic",
            "OBS\tdirect/days-in-month\tLONG:999",
        )
        (bundle / dest["path"]).write_text(changed, encoding="utf-8")
        dest["sha256"] = hashlib.sha256(changed.encode()).hexdigest()
        manifest_file.write_text(json.dumps(degraded), encoding="utf-8")
        try:
            validate_manifest(root, manifest_file, "b" * 40)
        except ValueError as error:
            require("observations differ" in str(error), "degraded parity case failed for wrong reason")
        else:
            raise RuntimeError("degraded parity self-test unexpectedly passed")

        degraded = json.loads(json.dumps(manifest))
        dest = next(x for x in degraded["logs"] if x["id"] == "destination-observations")
        failed_runner = obs.replace(
            "KPR shape regression  checks: 12  failures: 0",
            "KPR shape regression  checks: 12  failures: 1",
        )
        (bundle / dest["path"]).write_text(failed_runner, encoding="utf-8")
        dest["sha256"] = hashlib.sha256(failed_runner.encode()).hexdigest()
        manifest_file.write_text(json.dumps(degraded), encoding="utf-8")
        try:
            validate_manifest(root, manifest_file, "b" * 40)
        except ValueError as error:
            require("shape runner failures=1" in str(error),
                    "degraded runner-failure case failed for wrong reason")
        else:
            raise RuntimeError("degraded runner-failure self-test unexpectedly passed")

        # Restore the destination observation log for the digest-binding case.
        (bundle / "destination-observations.log").write_text(obs, encoding="utf-8")

        degraded = manifest.copy()
        original_logs = manifest["logs"]
        require(isinstance(original_logs, list), "self-test logs must be a list")
        degraded["logs"] = [dict(x) for x in original_logs if isinstance(x, dict)]
        host = next(x for x in degraded["logs"] if x["id"] == "destination-host-record")
        host["sha256"] = "0" * 64
        manifest_file.write_text(json.dumps(degraded), encoding="utf-8")
        try:
            validate_manifest(root, manifest_file, "b" * 40)
        except ValueError as error:
            require("digest mismatch" in str(error), "degraded digest case failed for wrong reason")
        else:
            raise RuntimeError("degraded digest self-test unexpectedly passed")

    print(
        "PASS migration-evidence self-test: positive, source identity, destination identity, "
        "parity mismatch, runner failure, digest mismatch"
    )


def parse_args(argv: list[str]) -> argparse.Namespace:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--root", type=Path, default=Path.cwd())
    parser.add_argument("--manifest", type=Path)
    parser.add_argument("--output", type=Path)
    parser.add_argument("--self-test", action="store_true")
    return parser.parse_args(argv)


def main(argv: list[str] | None = None) -> int:
    args = parse_args(sys.argv[1:] if argv is None else argv)
    try:
        if args.self_test:
            self_test()
            return 0
        require(args.manifest is not None, "--manifest is required unless --self-test is used")
        root = args.root.resolve()
        report = validate_manifest(
            root,
            args.manifest.resolve(),
            checked_out_sha(root),
        )
        rendered = json.dumps(report, indent=2) + "\n"
        if args.output is not None:
            args.output.parent.mkdir(parents=True, exist_ok=True)
            args.output.write_text(rendered, encoding="utf-8")
        print(rendered, end="")
        return 0
    except (
        OSError,
        ValueError,
        TypeError,
        KeyError,
        RuntimeError,
        json.JSONDecodeError,
        subprocess.SubprocessError,
    ) as error:
        print(f"ERROR: {error}", file=sys.stderr)
        return 1


if __name__ == "__main__":
    raise SystemExit(main())
