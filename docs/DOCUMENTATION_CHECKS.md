# 🔗 Documentation Checks

[![Local checks: blocking](https://img.shields.io/badge/local%20checks-blocking-217346)](#-local-and-offline-checks)
[![Network: separate](https://img.shields.io/badge/network-separate-1D76DB)](#-external-observations)
[![Policy: versioned](https://img.shields.io/badge/policy-versioned-6f42c1)](#-policy-and-exceptions)

`.github/documentation-policy.json` owns maintained-document references and
external-link reliability policy. Local defects and external availability have
separate verdicts, so a temporary service outage cannot weaken local validation.

## 📚 Local and Offline Checks

`tools/check_repo.py` retains its existing deterministic, blocking local-link
and anchor validation. `tools/check_documentation.py` additionally checks:

- literal Python invocations of tracked scripts under `tools/` in maintained Markdown,
  including inline and continued examples, against tracked filenames;
- documented long options against literal `add_argument` declarations and
  shared argument helpers, using Python AST parsing without executing code;
- registered filename references against tracked targets and document text;
- registered literal workflow names and named job contexts against their YAML;
- registered JSON policy values against their JSON-pointer targets.

The command scanner is not a shell interpreter: it does not execute examples,
verify argument values, resolve arbitrary dynamic CLI generation or test Office.
Reference coverage is explicit in the policy. It cannot infer every prose claim
or find every unregistered inline filename. Register new cross-file claims when
they would otherwise drift silently. Complex/dynamic workflow names need a
deliberate checker extension; the current registered-name syntax is literal.

```bash
python3 tools/check_documentation.py --root . \
  --output test-results/documentation.json --summary test-results/documentation.md
python3 tools/test_documentation.py -v
```

Both run under `Repository integrity`. A removed script, renamed flag/context,
missing registered file, or changed registered policy value fails the gate.
Fix the implementation or documentation together; do not remove the reference
merely to hide an inconsistency. Historical document exclusions have versioned
reasons and apply only to command scanning; their local links still get checked.

## 🌐 External Observations

The separate **External documentation links** workflow runs weekly on Monday
at 06:17 UTC and through `workflow_dispatch`. It is not part of the required
PR status check. GitHub [scheduled workflows run from the default branch](https://docs.github.com/en/actions/reference/workflows-and-actions/events-that-trigger-workflows#schedule),
so a workflow added only on a release branch has not started its weekly schedule.
Dispatch it after publication, or run the same command locally:

```bash
python3 tools/check_external_links.py --root . --as-of 2026-09-09 \
  --output test-results/external-links.json --summary test-results/external-links.md
```

Use the actual UTC observation date for `--as-of`; expiry is evaluated against
it. The workflow supplies that date automatically. The report contains
observations, not a permanent certification of remote content.

The checker discovers Markdown destinations/reference definitions, HTML
`href`/`src` values and angle-bracket HTTP links outside fenced examples. It
does not crawl linked pages or validate remote fragment anchors. Duplicate URLs
share one probe; all referring file/line locations are retained in JSON.

Default limits are 200 unique links, four workers, three attempts, an eight-second
socket timeout, three redirects and one-/two-second retry backoff. A consistently
missing response is reported only after all attempts. Mixed missing/transient
responses remain transient. Socket timeouts do not impose a DNS resolver deadline;
the workflow's ten-minute job timeout bounds the whole hosted operation. Exceeding
the link cap is explicitly non-passing, never a complete scan.

## 🛡️ Policy and Exceptions

Only HTTPS on port 443 is probed, and only exact approved domain names. Every
domain has a versioned rationale; subdomains and redirect destinations need
their own entries. Every redirect is checked before connection. All resolved
addresses must be public; the connection uses a checked IP while retaining the
original hostname for TLS verification, preventing a second DNS lookup from
changing the destination. No response body is downloaded intentionally.

There are no authentication headers, cookies, environment proxies, credential
files or private-site login attempts. URLs with user information or query
parameters are not sent, including harmless-looking badge query strings.
`mailto`, `tel` and unexpanded template URLs are reported as not applicable.
HTTP and other unapproved schemes are policy-blocked.

Reports show the domain, source locations and SHA-256 URL identifier rather
than the raw URL. Query strings, user information, redirect destinations,
classification rationales and raw exception messages are not printed. Locate a
finding through its referring document; do not paste credential-bearing URLs
into issue or workflow logs.

Temporary exceptions have exactly `id`, `reason` and `expires` fields. The ID
is SHA-256 of the discovered URL with its fragment removed, as shown in the
report. An example shape is:

```json
{"id": "REPLACE_WITH_64_CHARACTER_REPORT_ID", "reason": "Reviewed temporary service migration; owner assigned", "expires": "2026-09-16"}
```

An exception bypasses the probe with an explicit `EXCEPTED` result, not an
availability claim. Missing rationale, duplicate identifiers and expired entries
fail policy validation. Review renewal or remove the exception after correction.
There are no shipped temporary exceptions. Domain approvals are durable
public-service decisions; temporary per-link exceptions always expire.

### Restricted and pre-publication classifications

Known non-public observations use the separate `classifications` array. Each
entry has exactly `id`, `kind`, `reason` and `expires`. Supported kinds are:

- `access-restricted` for an exact current private target verified through an
  authenticated repository review;
- `restricted-historical` for a preserved evidence URL known through prior
  authenticated review to belong to a private/restricted historical target; and
- `pending-publication` for a reviewed candidate URL whose target does not exist
  anonymously until the associated tag/release is published.

These classifications are **not exceptions**. Before a classification is applied,
the checker evaluates the existing local URL policy without performing network
I/O. `POLICY_BLOCKED` and `ACCESS_RESTRICTED` therefore take precedence over
`RESTRICTED_HISTORICAL` and `PENDING_PUBLICATION`: a classification may explain
an observed link state, but it never exempts an invalid scheme/domain/port or a
credential/query-bearing URL from local policy. A policy-clean classified URL
still skips the anonymous request that cannot answer the intended question, and
the resulting classification status remains non-green.

`ACCESS_RESTRICTED`, `RESTRICTED_HISTORICAL` and `PENDING_PUBLICATION` never become `PASS`
and never substitute for authenticated evidence or the later post-publication
check.

Classifications are exact URL-hash assertions, not domain or path wildcards.
They require a rationale and an expiry/review date; unsupported kinds, duplicate
IDs (including an ID also present in `exceptions`) and expired entries fail
closed. K-PRICING records its current private targets with the
`access-restricted` kind. Unused inherited pilot classifications were removed
during setup. IDs derive from the exact maintained URLs; no credentials are
stored. Review them when visibility changes or their review date expires.

A release candidate may temporarily classify an exact tag-dependent URL as
`pending-publication` only when its rationale and expiry are reviewed with the
candidate. After publication, remove that classification and require the normal
anonymous check. Do not use this mechanism for a public page that should already
exist.

## 🚦 Interpreting Results

| Outcome | Meaning and response |
| --- | --- |
| `OK` | Anonymous HTTP success, possibly after retry; not content or anchor validation |
| `PERMANENT_FAILURE` | Repeated non-transient public HTTP errors; counted as a deterministic public-documentation defect unless a policy-clean URL is explicitly classified before observation |
| `TRANSIENT_FAILURE` | Timeout, transport failure, rate limit or server error; retry later |
| `ACCESS_RESTRICTED` | Authentication/access response, user information/query-bearing URL, or explicitly classified current private target; no private login attempted; local URL policy takes precedence |
| `RESTRICTED_HISTORICAL` | Exact reviewed private/restricted historical target; no anonymous probe attempted; remains non-green |
| `PENDING_PUBLICATION` | Exact reviewed target depends on publication/tag creation; no probe attempted; remains non-green until publication |
| `POLICY_BLOCKED` | Unapproved scheme/domain/port or non-public DNS destination; review policy and URL; local policy takes precedence over classification |
| `REDIRECT_FAILURE` | Missing redirect target, redirect loop or redirect limit reached |
| `EXCEPTED` | Active reviewed temporary exception; availability was not checked |
| `NOT_APPLICABLE` | Non-HTTP contact link or unresolved template destination |

The JSON and Markdown summary report the same five aggregates:
`deterministic_public_defects`, `restricted_historical`, `pending_publication`,
`access_restricted`, and `transient_failures`. This lets release review
distinguish a genuinely missing public page from preserved restricted evidence
without reconstructing raw URLs from hashes. A classification changes the
diagnosis only after local policy passes; it never weakens the terminal verdict.

An undeclared anonymous 404 remains `PERMANENT_FAILURE` after bounded retries.
A 404 must never be inferred to mean “private” merely because GitHub can hide
private repositories that way. Restricted status is accepted only through the
explicit, expiring classification above.

The network command exits 1 for any unresolved failure/restriction,
`RESTRICTED_HISTORICAL`, `PENDING_PUBLICATION`, or exceeded limit, and 2 for
invalid policy or operational failure. The workflow preserves its own non-green
result without changing `Repository integrity`. JSON/Markdown observations
upload even after a check fails, with 14-day retention; a job killed before
report creation fails artifact publication rather than fabricating a report.

## 🧪 Validation Scope

Offline fixtures cover command/context/policy renames, repeated missing public
pages, restricted historical targets, pending-publication targets, classification
expiry/conflicts, classification-versus-local-policy precedence, aggregate
JSON/Markdown parity, transient recovery, restricted access, redirects, request
limits, temporary exceptions, DNS/IP containment, concurrency and report
redaction. They simulate HTTP results and do not establish current availability
of every referenced website. The separate network workflow supplies dated
observations when actually executed.
