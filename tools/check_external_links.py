#!/usr/bin/env python3
"""Bounded anonymous link observations; independent of the local-link gate."""

from __future__ import annotations

import argparse
import hashlib
import http.client
import ipaddress
import re
import socket
import ssl
import time
from concurrent.futures import ThreadPoolExecutor
from datetime import date
from pathlib import Path
from typing import Any
from urllib.parse import quote, urljoin, urlsplit

from _gatelib import run_gate, tracked_files
from check_documentation import load_policy
from check_repo import _markdown_destinations
from release_provenance import nonempty, require

CLASSIFICATION_STATUSES = {
    "restricted-historical": "RESTRICTED_HISTORICAL",
    "pending-publication": "PENDING_PUBLICATION",
}
PASSING_STATUSES = frozenset({"OK", "NOT_APPLICABLE", "EXCEPTED"})
COUNT_LABELS = {
    "deterministic_public_defects": "Deterministic public defects",
    "restricted_historical": "restricted historical",
    "pending_publication": "pending publication",
    "access_restricted": "access restricted",
    "transient_failures": "transient",
}


def _validate_expiring_id(item: Any, expected: set[str], as_of: date, seen: set[str]) -> None:
    require(isinstance(item, dict) and set(item) == expected, "invalid external-link policy item")
    identifier = item["id"]
    require(
        isinstance(identifier, str)
        and re.fullmatch(r"[0-9a-f]{64}", identifier)
        and identifier not in seen,
        "invalid or duplicate external-link policy identity",
    )
    require(nonempty(item["reason"]), "external-link policy item requires a reason")
    require(
        isinstance(item["expires"], str) and date.fromisoformat(item["expires"]) >= as_of,
        "expired external-link policy item",
    )
    seen.add(identifier)


def validate_policy(policy: Any, as_of: date) -> None:
    require(isinstance(policy, dict), "network policy must be an object")
    for key, limit in {
        "attempts": 3,
        "timeout_seconds": 15,
        "concurrency": 8,
        "redirects": 5,
        "max_links": 500,
    }.items():
        require(
            type(policy.get(key)) is int and 1 <= policy[key] <= limit,
            f"invalid network limit: {key}",
        )
    domains = policy.get("domains")
    require(isinstance(domains, dict) and bool(domains), "approved domains are required")
    for host, reason in domains.items():
        require(
            bool(re.fullmatch(r"[a-z0-9](?:[a-z0-9.-]*[a-z0-9])?", host))
            and "." in host
            and nonempty(reason),
            "domains require exact names and reasons",
        )

    exceptions = policy.get("exceptions")
    classifications = policy.get("classifications")
    require(isinstance(exceptions, list), "exceptions must be an array")
    require(isinstance(classifications, list), "classifications must be an array")
    seen: set[str] = set()
    for item in exceptions:
        _validate_expiring_id(item, {"id", "reason", "expires"}, as_of, seen)
    for item in classifications:
        _validate_expiring_id(
            item,
            {"id", "kind", "reason", "expires"},
            as_of,
            seen,
        )
        require(
            item["kind"] in CLASSIFICATION_STATUSES,
            "unsupported external-link classification",
        )


def destinations(text: str) -> list[tuple[int, str]]:
    result = list(_markdown_destinations(text))
    fenced = False
    for number, line in enumerate(text.splitlines(), 1):
        if line.lstrip().startswith(("```", "~~~")):
            fenced = not fenced
        elif not fenced:
            result.extend((number, value) for value in re.findall(r"<(https?://[^\s<>]+)>", line))
            result.extend(
                (number, value)
                for value in re.findall(r"(?:href|src)=[\"']([^\"']+)[\"']", line)
            )
    return result


def url_status(url: str, policy: dict[str, Any]) -> str | None:
    try:
        parsed = urlsplit(url)
        if parsed.scheme in ("mailto", "tel") or "{{" in url:
            return "NOT_APPLICABLE"
        if parsed.scheme != "https" or parsed.port not in (None, 443):
            return "POLICY_BLOCKED"
        if parsed.username is not None or parsed.password is not None or parsed.query:
            return "ACCESS_RESTRICTED"
        if parsed.hostname not in policy["domains"] or any(ord(char) < 32 for char in url):
            return "POLICY_BLOCKED"
    except ValueError:
        return "POLICY_BLOCKED"
    return None


class PinnedHTTPS(http.client.HTTPSConnection):
    def __init__(self, host: str, address: str, timeout: int):
        self.tls_context = ssl.create_default_context()
        super().__init__(host, timeout=timeout, context=self.tls_context)
        self.address = address

    def connect(self) -> None:
        # Connect to the already checked address, retaining the hostname for TLS.
        # No second DNS lookup, environment proxy, cookie jar or auth handler.
        raw = socket.create_connection((self.address, 443), self.timeout)
        try:
            self.sock = self.tls_context.wrap_socket(raw, server_hostname=self.host)
        except BaseException:
            raw.close()
            raise


def request(url: str, timeout: int) -> tuple[int, str | None]:
    parsed = urlsplit(url)
    assert parsed.hostname is not None
    addresses = socket.getaddrinfo(parsed.hostname, 443, type=socket.SOCK_STREAM)
    require(
        bool(addresses)
        and all(ipaddress.ip_address(item[4][0]).is_global for item in addresses),
        "non-public DNS address",
    )
    connection = PinnedHTTPS(parsed.hostname, str(addresses[0][4][0]), timeout)
    try:
        path = quote(parsed.path or "/", safe="/%:@-._~!$&'()*+,;=")
        connection.request(
            "GET",
            path,
            headers={"User-Agent": "Template-Link-Check/1.0", "Accept": "*/*"},
        )
        response = connection.getresponse()
        return response.status, response.getheader("Location")
    finally:
        connection.close()


def attempt(url: str, policy: dict[str, Any], transport=request) -> tuple[str, int | None]:
    seen = set()
    for _ in range(policy["redirects"] + 1):
        blocked = url_status(url, policy)
        if blocked:
            return blocked, None
        if url in seen:
            return "REDIRECT_FAILURE", None
        seen.add(url)
        try:
            code, location = transport(url, policy["timeout_seconds"])
        except ValueError:
            return "POLICY_BLOCKED", None
        except (OSError, http.client.HTTPException, UnicodeError):
            return "TRANSIENT_FAILURE", None
        if code in (301, 302, 303, 307, 308):
            if not location:
                return "REDIRECT_FAILURE", code
            url = urljoin(url, location)
            continue
        if 200 <= code < 300:
            return "OK", code
        if code in (401, 403, 407, 451):
            return "ACCESS_RESTRICTED", code
        if code in (408, 425, 429) or code >= 500:
            return "TRANSIENT_FAILURE", code
        return "PERMANENT_FAILURE", code
    return "REDIRECT_FAILURE", None


def probe(
    url: str,
    policy: dict[str, Any],
    transport=request,
    pause=time.sleep,
) -> dict[str, Any]:
    statuses = []
    codes = []
    for number in range(policy["attempts"]):
        status, code = attempt(url, policy, transport)
        statuses.append(status)
        codes.append(code)
        if status not in ("PERMANENT_FAILURE", "TRANSIENT_FAILURE"):
            break
        if number + 1 < policy["attempts"]:
            pause(number + 1)
    if statuses[-1] == "OK":
        outcome = "OK"
    elif all(status == "PERMANENT_FAILURE" for status in statuses):
        outcome = "PERMANENT_FAILURE"
    elif "TRANSIENT_FAILURE" in statuses:
        outcome = "TRANSIENT_FAILURE"
    else:
        outcome = statuses[-1]
    return {"status": outcome, "attempts": len(statuses), "http_codes": codes}


def collect(root: Path) -> dict[str, dict[str, Any]]:
    links: dict[str, dict[str, Any]] = {}
    for path in sorted(tracked_files(root)):
        if not path.endswith(".md"):
            continue
        for line, raw in destinations((root / path).read_text(encoding="utf-8")):
            if not raw.strip():
                continue
            url = raw.strip().strip("<>").split()[0]
            if not re.match(r"[A-Za-z][A-Za-z0-9+.-]*:", url):
                continue
            # Fragments are not sent to servers and are not checked by this job.
            url = url.split("#", 1)[0]
            digest = hashlib.sha256(url.encode()).hexdigest()
            links.setdefault(digest, {"url": url, "locations": []})["locations"].append(
                f"{path}:{line}"
            )
    return links


def _counts(rows: list[dict[str, Any]]) -> dict[str, int]:
    return {
        "deterministic_public_defects": sum(
            row["status"] == "PERMANENT_FAILURE" for row in rows
        ),
        "restricted_historical": sum(
            row["status"] == "RESTRICTED_HISTORICAL" for row in rows
        ),
        "pending_publication": sum(row["status"] == "PENDING_PUBLICATION" for row in rows),
        "access_restricted": sum(row["status"] == "ACCESS_RESTRICTED" for row in rows),
        "transient_failures": sum(row["status"] == "TRANSIENT_FAILURE" for row in rows),
    }


def build_report(
    root: Path,
    as_of: date,
    transport=request,
    pause=time.sleep,
) -> dict[str, Any]:
    policy = load_policy(root)["network"]
    validate_policy(policy, as_of)
    links = collect(root)
    exceptions = {item["id"] for item in policy["exceptions"]}
    classifications = {item["id"]: item for item in policy["classifications"]}

    def check(item: tuple[str, dict[str, Any]]) -> dict[str, Any]:
        digest, value = item
        url = value["url"]
        try:
            domain = urlsplit(url).hostname or "<none>"
        except ValueError:
            domain = "<invalid>"
        classification = classifications.get(digest)
        policy_status = url_status(url, policy)
        if classification is not None and policy_status in {
            "POLICY_BLOCKED",
            "ACCESS_RESTRICTED",
        }:
            result = {
                "status": policy_status,
                "attempts": 0,
                "http_codes": [],
                "classification": classification["kind"],
            }
        elif classification is not None:
            result = {
                "status": CLASSIFICATION_STATUSES[classification["kind"]],
                "attempts": 0,
                "http_codes": [],
                "classification": classification["kind"],
            }
        elif digest in exceptions:
            result = {"status": "EXCEPTED", "attempts": 0, "http_codes": []}
        else:
            result = probe(url, policy, transport, pause)
        return {
            "id": digest,
            "domain": domain,
            "locations": value["locations"],
            **result,
        }

    selected = sorted(links.items())[: policy["max_links"]]
    with ThreadPoolExecutor(max_workers=policy["concurrency"]) as pool:
        rows = list(pool.map(check, selected))
    limited = len(links) > policy["max_links"]
    counts = _counts(rows)
    status = (
        "fail"
        if limited or any(row["status"] not in PASSING_STATUSES for row in rows)
        else "pass"
    )
    return {
        "status": status,
        "as_of": as_of.isoformat(),
        "limit_exceeded": limited,
        "discovered": len(links),
        "counts": counts,
        "links": rows,
        "scope_note": (
            "Anonymous HTTP observations only; access restrictions, declared historical "
            "restrictions, pending-publication references and transients remain non-green "
            "without being counted as deterministic public-page defects. Queries, credentials "
            "and fragments are not probed; URLs are represented by domain and SHA-256 ID."
        ),
    }


def markdown(report: dict[str, Any]) -> str:
    counts = report["counts"]
    rendered_counts = "; ".join(
        f"{COUNT_LABELS[key]}: {value}" for key, value in counts.items()
    )
    lines = [
        "# External documentation links",
        "",
        f"Result: {report['status'].upper()}",
        "",
        f"Discovered: {report['discovered']}; limit exceeded: {report['limit_exceeded']}",
        rendered_counts,
        "",
        "| Location | Domain | Outcome | Attempts | ID |",
        "| --- | --- | --- | --- | --- |",
    ]
    for row in report["links"]:
        lines.append(
            f"| {row['locations'][0]} | {row['domain']} | {row['status']} | "
            f"{row['attempts']} | {row['id']} |"
        )
    return "\n".join(lines) + "\n\n" + report["scope_note"] + "\n"


def main() -> int:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--root", type=Path, default=Path.cwd())
    parser.add_argument("--as-of", type=date.fromisoformat, required=True)
    parser.add_argument("--output", type=Path)
    parser.add_argument("--summary", type=Path)
    args = parser.parse_args()
    return run_gate(
        args,
        build=lambda: build_report(args.root, args.as_of),
        markdown=markdown,
        errors=(ValueError, OSError, RuntimeError),
    )


if __name__ == "__main__":
    raise SystemExit(main())
