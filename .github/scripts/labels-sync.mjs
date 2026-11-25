#!/usr/bin/env node

import assert from "node:assert/strict";
import { appendFile, mkdir, mkdtemp, readFile, rm, writeFile } from "node:fs/promises";
import { tmpdir } from "node:os";
import { dirname, join } from "node:path";
import process from "node:process";
import { pathToFileURL } from "node:url";

const EXPECTED_SCHEMA_VERSION = 3;
const ROOT_KEYS = ["schema_version", "prune", "core", "overlays", "selected_overlays"];
const DEFAULT_MANIFEST = ".github/labels.json";
const API_VERSION = "2022-11-28";
const LABELS_PER_PAGE = 100;
const MAX_LABEL_PAGES = 100;
const ISSUES_PER_PAGE = 100;
const MAX_ISSUE_PAGES = 100;
// The issues and pull requests a summary names for a label still in use.
export const USAGE_SAMPLE = 10;
const MAX_RATE_LIMIT_RETRIES = 2;
export const MAX_RETRY_DELAY_MS = 5000;

export function compareNames(left, right) {
  const leftKey = left.toLowerCase();
  const rightKey = right.toLowerCase();
  if (leftKey < rightKey) return -1;
  if (leftKey > rightKey) return 1;
  return left < right ? -1 : left > right ? 1 : 0;
}

function sameKeys(value, expected) {
  const actual = Object.keys(value).sort();
  return JSON.stringify(actual) === JSON.stringify([...expected].sort());
}

export function normalizedLiveLabel(label) {
  return {
    name: String(label.name ?? ""),
    color: String(label.color ?? "").toUpperCase(),
    description: label.description == null ? "" : String(label.description)
  };
}

function validateLabelArray(labels, location, seen) {
  const errors = [];
  if (!Array.isArray(labels)) return [`${location} must be an array`];

  labels.forEach((label, index) => {
    const item = `${location}[${index}]`;
    if (!label || typeof label !== "object" || Array.isArray(label)) {
      errors.push(`${item} must be an object`);
      return;
    }
    if (!sameKeys(label, ["name", "color", "description"])) {
      errors.push(`${item} must contain exactly name, color, and description`);
    }

    if (typeof label.name !== "string") {
      errors.push(`${item}.name must be a string`);
    } else {
      if (label.name.length === 0 || label.name !== label.name.trim()) {
        errors.push(`${item}.name must be non-empty with no surrounding whitespace`);
      }
      if (label.name.length > 50) errors.push(`${item}.name exceeds GitHub's 50-character limit`);
      if (/\r|\n/.test(label.name)) errors.push(`${item}.name must be single-line`);
      const key = label.name.toLowerCase();
      if (seen.has(key)) {
        errors.push(`${item}.name duplicates ${seen.get(key)} case-insensitively`);
      } else {
        seen.set(key, `${item}.name`);
      }
    }

    if (typeof label.color !== "string" || !/^[0-9A-F]{6}$/.test(label.color)) {
      errors.push(`${item}.color must be six uppercase hexadecimal characters without #`);
    }
    if (typeof label.description !== "string") {
      errors.push(`${item}.description must be a string`);
    } else {
      if (label.description.length === 0) errors.push(`${item}.description must not be empty`);
      if (label.description.length > 100) errors.push(`${item}.description exceeds GitHub's 100-character limit`);
      if (/\r|\n/.test(label.description)) errors.push(`${item}.description must be single-line`);
    }
  });

  for (let index = 1; index < labels.length; index += 1) {
    const previous = labels[index - 1]?.name;
    const current = labels[index]?.name;
    if (typeof previous === "string" && typeof current === "string" && compareNames(previous, current) >= 0) {
      errors.push(`${location} must be sorted case-insensitively by name: ${previous} precedes ${current}`);
      break;
    }
  }
  return errors;
}

function validateSortedNames(names, location) {
  for (let index = 1; index < names.length; index += 1) {
    if (compareNames(names[index - 1], names[index]) >= 0) {
      return [`${location} must be sorted case-insensitively by name: ${names[index - 1]} precedes ${names[index]}`];
    }
  }
  return [];
}

// Overlays are named label sets besides the core: the repository defines
// its own, such as one per regulatory regime, and selects them.
function validateOverlays(overlays, seen) {
  if (!overlays || typeof overlays !== "object" || Array.isArray(overlays)) {
    return ["overlays must be an object mapping each overlay name to its label array"];
  }
  const errors = [];
  const names = Object.keys(overlays);
  const seenOverlays = new Map();
  for (const name of names) {
    const item = `overlays[${JSON.stringify(name)}]`;
    if (name.length === 0 || name !== name.trim() || /\r|\n/.test(name)) {
      errors.push(`${item} name must be non-empty, single-line, with no surrounding whitespace`);
    }
    const key = name.toLowerCase();
    if (seenOverlays.has(key)) {
      errors.push(`${item} duplicates overlay ${JSON.stringify(seenOverlays.get(key))} case-insensitively`);
    } else {
      seenOverlays.set(key, name);
    }
    errors.push(...validateLabelArray(overlays[name], item, seen));
    if (Array.isArray(overlays[name]) && overlays[name].length === 0) errors.push(`${item} must not be empty`);
  }
  errors.push(...validateSortedNames(names, "overlays"));
  return errors;
}

function validateSelection(selected, overlays) {
  if (!Array.isArray(selected)) return ["selected_overlays must be an array of overlay names"];
  const errors = [];
  const defined = overlays && typeof overlays === "object" && !Array.isArray(overlays) ? overlays : {};
  const chosen = new Set();
  selected.forEach((name, index) => {
    const item = `selected_overlays[${index}]`;
    if (typeof name !== "string") {
      errors.push(`${item} must be a string`);
      return;
    }
    if (!Object.hasOwn(defined, name)) {
      errors.push(`${item} selects unknown overlay ${JSON.stringify(name)}`);
    }
    if (chosen.has(name.toLowerCase())) {
      errors.push(`${item} selects overlay ${JSON.stringify(name)} again`);
    }
    chosen.add(name.toLowerCase());
  });
  if (selected.every(name => typeof name === "string")) {
    errors.push(...validateSortedNames(selected, "selected_overlays"));
  }
  return errors;
}

export function validateManifest(document) {
  const errors = [];
  if (!document || typeof document !== "object" || Array.isArray(document)) {
    throw new Error("Manifest must be a JSON object.");
  }
  if (document.schema_version === 2 && Object.hasOwn(document, "labels")) {
    throw new Error(
      "Invalid label manifest:\n- schema_version 2 with a labels list is no longer supported: "
      + `set schema_version ${EXPECTED_SCHEMA_VERSION}, rename labels to core, and add overlays {} and selected_overlays []`
    );
  }
  if (!sameKeys(document, ROOT_KEYS)) {
    errors.push(`root must contain exactly ${ROOT_KEYS.slice(0, -1).join(", ")}, and ${ROOT_KEYS.at(-1)}`);
  }
  if (document.schema_version !== EXPECTED_SCHEMA_VERSION) {
    errors.push(`schema_version must be ${EXPECTED_SCHEMA_VERSION}`);
  }
  if (typeof document.prune !== "boolean") errors.push("prune must be a boolean");

  // One map across the core and every overlay: a name is unique in the whole catalogue.
  const seen = new Map();
  errors.push(...validateLabelArray(document.core, "core", seen));
  if (Array.isArray(document.core) && document.core.length === 0) errors.push("core must not be empty");
  errors.push(...validateOverlays(document.overlays, seen));
  errors.push(...validateSelection(document.selected_overlays, document.overlays));

  if (errors.length > 0) throw new Error(`Invalid label manifest:\n- ${errors.join("\n- ")}`);
  return structuredClone(document);
}

export function catalogueCounts(manifest) {
  if (!manifest) return null;
  return {
    core: manifest.core.length,
    overlays: Object.keys(manifest.overlays).length,
    selected: manifest.selected_overlays.length,
    resolved: resolveDesired(manifest).length
  };
}

// The core plus the labels of every selected overlay, sorted by name.
export function resolveDesired(document) {
  const manifest = validateManifest(document);
  const labels = [...manifest.core];
  for (const name of manifest.selected_overlays) labels.push(...manifest.overlays[name]);
  labels.sort((left, right) => compareNames(left.name, right.name));
  return labels;
}

export function planChanges(desiredLabels, liveLabels, { prune = true } = {}) {
  const desired = new Map(desiredLabels.map(label => [label.name.toLowerCase(), label]));
  const live = new Map(liveLabels.map(label => {
    const normalized = normalizedLiveLabel(label);
    return [normalized.name.toLowerCase(), normalized];
  }));
  const changes = [];

  for (const [key, target] of desired) {
    const current = live.get(key);
    if (!current) {
      changes.push({ action: "create", name: target.name, target });
      continue;
    }
    const fields = [];
    if (current.name !== target.name) fields.push("name");
    if (current.color !== target.color) fields.push("color");
    if (current.description !== target.description) fields.push("description");
    if (fields.length > 0) {
      changes.push({ action: "update", name: target.name, currentName: current.name, fields, target });
    }
  }

  if (prune) {
    for (const [key, current] of live) {
      if (!desired.has(key)) changes.push({ action: "delete", name: current.name, currentName: current.name });
    }
  }

  const order = { update: 0, create: 1, delete: 2 };
  changes.sort((left, right) => order[left.action] - order[right.action] || compareNames(left.name, right.name));
  return changes;
}

function parseArguments(argv) {
  const parsed = {
    manifest: DEFAULT_MANIFEST,
    mode: null,
    live: null,
    repository: process.env.GITHUB_REPOSITORY || null,
    planOutput: null,
    dryRun: false,
    selfTest: false
  };
  for (let index = 0; index < argv.length; index += 1) {
    const argument = argv[index];
    if (argument === "--dry-run") {
      parsed.dryRun = true;
    } else if (argument === "--self-test") {
      parsed.selfTest = true;
    } else if (["--manifest", "--mode", "--live", "--repository", "--plan-output"].includes(argument)) {
      const value = argv[index + 1];
      if (!value || value.startsWith("--")) throw new Error(`${argument} requires a value`);
      parsed[argument === "--plan-output" ? "planOutput" : argument.slice(2)] = value;
      index += 1;
    } else {
      throw new Error(`Unknown argument: ${argument}`);
    }
  }
  if (parsed.planOutput !== null && parsed.mode !== "plan") throw new Error("--plan-output requires --mode plan");
  return parsed;
}

// The planned changes as JSON, so that another tool can compare its own plan
// with the reconciler's on the same live snapshot.
export async function writePlan(path, changes) {
  const directory = dirname(path);
  if (directory !== ".") await mkdir(directory, { recursive: true });
  await writeFile(path, `${JSON.stringify({ changes }, null, 2)}\n`, "utf8");
}

export async function loadJson(path) {
  let text;
  try {
    text = await readFile(path, "utf8");
  } catch (error) {
    throw new Error(`Cannot read ${path}: ${error.message}`);
  }
  try {
    return JSON.parse(text);
  } catch (error) {
    throw new Error(`Invalid JSON in ${path}: ${error.message}`);
  }
}

function splitRepository(repository) {
  const parts = String(repository ?? "").split("/");
  if (parts.length !== 2 || parts.some(part => part.length === 0)) {
    throw new Error("repository must be in owner/name form");
  }
  return parts.map(encodeURIComponent);
}

function apiUrl(repository, suffix) {
  const [owner, repo] = splitRepository(repository);
  // The runner sets GITHUB_API_URL, which differs from api.github.com on GitHub Enterprise.
  const base = (process.env.GITHUB_API_URL || "https://api.github.com").replace(/\/+$/, "");
  return `${base}/repos/${owner}/${repo}${suffix}`;
}

function sleep(milliseconds) {
  return new Promise(resolve => setTimeout(resolve, milliseconds));
}

function isRateLimited(response) {
  return response.status === 429
    || (
      response.status === 403
      && (
        response.headers.get("x-ratelimit-remaining") === "0"
        || response.headers.has("retry-after")
      )
    );
}

export function retryDelayMs(response, attempt) {
  const header = response.headers.get("retry-after");
  const retryAfter = header !== null && header.trim() !== "" ? Number(header) : NaN;
  if (Number.isFinite(retryAfter) && retryAfter >= 0) {
    return Math.min(retryAfter * 1000, MAX_RETRY_DELAY_MS);
  }
  return Math.min(1000 * (2 ** attempt), MAX_RETRY_DELAY_MS);
}

async function githubRequest(repository, token, method, suffix, body) {
  for (let attempt = 0; ; attempt += 1) {
    const response = await fetch(apiUrl(repository, suffix), {
      method,
      headers: {
        Accept: "application/vnd.github+json",
        Authorization: `Bearer ${token}`,
        "X-GitHub-Api-Version": API_VERSION,
        "User-Agent": "repository-labels-sync"
      },
      body: body === undefined ? undefined : JSON.stringify(body)
    });
    if (response.ok) {
      if (response.status === 204) return null;
      return response.json();
    }

    const detail = await response.text();
    const retryable = method === "GET" && isRateLimited(response);
    if (!retryable || attempt >= MAX_RATE_LIMIT_RETRIES) {
      throw new Error(`${method} ${suffix} failed with HTTP ${response.status}: ${detail.slice(0, 500)}`);
    }
    await sleep(retryDelayMs(response, attempt));
  }
}

export async function listLiveLabels(repository, token) {
  const labels = [];
  for (let page = 1; page <= MAX_LABEL_PAGES; page += 1) {
    const batch = await githubRequest(
      repository,
      token,
      "GET",
      `/labels?per_page=${LABELS_PER_PAGE}&page=${page}`
    );
    if (!Array.isArray(batch)) throw new Error("GitHub labels response is not an array");
    labels.push(...batch.map(normalizedLiveLabel));
    if (batch.length < LABELS_PER_PAGE) return labels;
  }
  throw new Error(`GitHub label pagination exceeded ${MAX_LABEL_PAGES} pages`);
}

// The number of an issue or pull request, a positive integer, or an error:
// a summary never names an item it cannot read.
function itemNumber(item) {
  const number = item?.number;
  if (!Number.isSafeInteger(number) || number < 1) {
    throw new Error("GitHub issues response has an item without a valid number");
  }
  return number;
}

// The usage of a label from the numbers of the items that carry it: null when
// none does, otherwise the lowest USAGE_SAMPLE numbers, ascending, and whether
// more items carry it.
export function usageSample(numbers) {
  const sorted = [...new Set(numbers)].sort((left, right) => left - right);
  if (sorted.length === 0) return null;
  return { items: sorted.slice(0, USAGE_SAMPLE), more: sorted.length > USAGE_SAMPLE };
}

// The issues and pull requests that carry a label: null when none does,
// otherwise `{ items, more }` as usageSample returns it. The issues endpoint
// covers pull requests; state=all adds closed items.
export async function labelUsage(repository, token, name, request = githubRequest) {
  if (!name.includes(",")) {
    // One page, oldest first, with one item more than the sample: the extra
    // item says whether more items carry the label, without another request.
    const batch = await request(
      repository,
      token,
      "GET",
      `/issues?state=all&direction=asc&per_page=${USAGE_SAMPLE + 1}&labels=${encodeURIComponent(name)}`
    );
    if (!Array.isArray(batch)) throw new Error("GitHub issues response is not an array");
    return usageSample(batch.map(itemNumber));
  }
  // The labels filter is a comma-separated list, so it cannot query this name:
  // read the labels of every issue and pull request instead. A malformed page
  // or more pages than the limit throws, so nothing is deleted on a guess.
  const wanted = name.toLowerCase();
  const numbers = [];
  for (let page = 1; page <= MAX_ISSUE_PAGES; page += 1) {
    const batch = await request(
      repository,
      token,
      "GET",
      `/issues?state=all&per_page=${ISSUES_PER_PAGE}&page=${page}`
    );
    if (!Array.isArray(batch)) throw new Error("GitHub issues response is not an array");
    for (const item of batch) {
      if (!Array.isArray(item?.labels) || item.labels.some(label => typeof label?.name !== "string")) {
        throw new Error("GitHub issues response has malformed labels");
      }
      const number = itemNumber(item);
      if (item.labels.some(label => label.name.toLowerCase() === wanted)) numbers.push(number);
    }
    if (batch.length < ISSUES_PER_PAGE) return usageSample(numbers);
  }
  throw new Error(`GitHub issue pagination exceeded ${MAX_ISSUE_PAGES} pages while checking ${JSON.stringify(name)}`);
}

// The planned deletions that an issue or pull request still carries, each as
// `{ name, items, more }`, read without any change: a dry run previews them,
// a real run refuses on them. `usageOf(name)` returns what labelUsage returns.
export async function deletionsInUse(changes, usageOf) {
  const inUse = [];
  for (const change of changes) {
    if (change.action !== "delete") continue;
    const usage = await usageOf(change.currentName);
    if (usage !== null) inUse.push({ name: change.currentName, items: usage.items, more: usage.more });
  }
  return inUse;
}

// Where a label is used, such as "#3, #7 and others": "and others" only when
// more items carry it than the sample names.
export function renderUsage(entry) {
  return `${entry.items.map(number => `#${number}`).join(", ")}${entry.more ? " and others" : ""}`;
}

export function deletionRefusal(inUse) {
  const labels = inUse.map(entry => `${JSON.stringify(entry.name)} (used by ${renderUsage(entry)})`);
  return `Refusing to prune labels still used by an issue or pull request: ${labels.join(", ")}. `
    + "No change was applied. Remove each label from its issues and pull requests, or add it to the catalogue or select its overlay.";
}

// Deleting a label removes it from every issue and pull request, closed ones
// included, and cannot be undone: refuse before any change is applied.
export async function assertSafeDeletions(changes, usageOf) {
  const inUse = await deletionsInUse(changes, usageOf);
  if (inUse.length > 0) throw new Error(deletionRefusal(inUse));
}

async function applyChanges(repository, token, changes) {
  for (const change of changes) {
    if (change.action === "update") {
      await githubRequest(repository, token, "PATCH", `/labels/${encodeURIComponent(change.currentName)}`, {
        new_name: change.target.name,
        color: change.target.color,
        description: change.target.description
      });
    } else if (change.action === "create") {
      await githubRequest(repository, token, "POST", "/labels", change.target);
    } else if (change.action === "delete") {
      await githubRequest(repository, token, "DELETE", `/labels/${encodeURIComponent(change.currentName)}`);
    }
  }
}

function singleLine(value) {
  return String(value).replace(/\r\n?|\n/g, " ");
}

// Plain text in a table cell: one line, with pipes escaped.
export function markdownEscape(value) {
  return singleLine(value).replaceAll("|", "\\|");
}

// A code span that the value cannot close: its delimiter is longer than any
// run of backticks in the value. In a table cell, pipes are escaped as well.
export function inlineCode(value, { table = false } = {}) {
  let text = singleLine(value);
  if (table) text = text.replaceAll("|", "\\|");
  const longest = Math.max(0, ...(text.match(/`+/g) ?? []).map(run => run.length));
  const fence = "`".repeat(longest + 1);
  const padding = /^[` ]|[` ]$/.test(text) ? " " : "";
  return `${fence}${padding}${text}${padding}${fence}`;
}

function renderLabelNames(labels) {
  if (labels.length === 0) return "none";
  return labels.map(label => inlineCode(label.name)).join(", ");
}

function renderSummary({
  mode,
  manifest,
  manifestPath,
  desiredLabels,
  changes = [],
  verified = null,
  dryRun = false,
  inUse = null,
  error = null,
  partial = false
}) {
  const resolvedLabels = desiredLabels ?? [];
  const planned = changes !== null;
  const counts = Object.fromEntries(["create", "update", "delete"].map(action => [
    action,
    planned ? changes.filter(change => change.action === action).length : "n/a"
  ]));
  const catalogue = catalogueCounts(manifest);
  const lines = [
    "# Repository label synchronization",
    "",
    `- Mode: ${inlineCode(mode)}${dryRun ? " (dry run)" : ""}`,
    `- Label manifest: ${inlineCode(manifestPath ?? "unavailable")}`,
    `- Schema: **${manifest?.schema_version ?? "unavailable"}**`,
    `- Labels: **${resolvedLabels.length}**`,
    `- Core / overlays defined / selected / resolved: **${catalogue
      ? `${catalogue.core} / ${catalogue.overlays} / ${catalogue.selected} / ${catalogue.resolved}`
      : "unavailable"}**`,
    `- Prune: **${manifest ? (manifest.prune ? "enabled" : "disabled") : "unavailable"}**`,
    `- Planned changes: **${planned ? changes.length : "not computed (live labels not read)"}**`,
    `- Create / update / delete: **${counts.create} / ${counts.update} / ${counts.delete}**`,
    `- Deletions still in use: **${inUse === null ? "not checked (no repository read)" : inUse.length}**`
  ];
  if (verified !== null) lines.push(`- Post-run exact match: **${verified ? "yes" : "no"}**`);
  lines.push(
    "",
    "## Label catalogue",
    "",
    `- Labels (${resolvedLabels.length}): ${renderLabelNames(resolvedLabels)}`,
    `- Selected overlays (${catalogue?.selected ?? 0}): ${manifest && manifest.selected_overlays.length > 0
      ? manifest.selected_overlays.map(name => inlineCode(name)).join(", ")
      : "none"}`
  );
  if (error) {
    lines.push("", "## Failure", "", inlineCode(error));
    if (partial) {
      lines.push("", "The failure happened during or after reconciliation: some changes may already be applied. Compare the live labels with the catalogue.");
    }
  } else if (!planned) {
    lines.push("", "Live labels were not read; no changes were planned.");
  } else if (changes.length === 0) {
    lines.push("", "No label changes are required.");
  } else {
    lines.push("", "| Action | Label | Detail |", "|---|---|---|");
    const refused = new Map((inUse ?? []).map(entry => [entry.name.toLowerCase(), entry]));
    for (const change of changes) {
      const entry = change.action === "delete" ? refused.get(change.currentName.toLowerCase()) : undefined;
      const detail = change.action === "update" ? change.fields.join(", ")
        : change.action === "delete" ? (entry
          ? `outside the catalogue; still in use, a real run would refuse it; used by ${renderUsage(entry)}`
          : "outside the catalogue")
          : "missing live label";
      lines.push(`| ${change.action} | ${inlineCode(change.name, { table: true })} | ${markdownEscape(detail)} |`);
    }
    if (dryRun && inUse !== null && inUse.length > 0) {
      lines.push("", `A real run would refuse every change. ${markdownEscape(deletionRefusal(inUse))}`);
    }
  }
  return `${lines.join("\n")}\n`;
}

async function publishSummary(summary) {
  process.stdout.write(summary);
  if (process.env.GITHUB_STEP_SUMMARY) await appendFile(process.env.GITHUB_STEP_SUMMARY, summary, "utf8");
}

function expectFailure(operation, pattern) {
  assert.throws(operation, pattern);
}

async function runSelfTest(manifestPath) {
  for (const [header, attempt, expected] of [
    [null, 0, 1000], [null, 2, 4000], ["", 1, 2000],
    ["invalid", 1, 2000], ["-1", 1, 2000], ["0", 2, 0],
    ["2", 0, 2000], ["999999", 0, MAX_RETRY_DELAY_MS],
    [null, 20, MAX_RETRY_DELAY_MS]
  ]) {
    const headers = new Headers(header === null ? {} : { "retry-after": header });
    assert.equal(retryDelayMs({ headers }, attempt), expected, `Retry-After: ${header}`);
  }

  const baseline = await loadJson(manifestPath);
  const manifest = validateManifest(baseline);
  const desired = resolveDesired(manifest);

  const missingRoot = structuredClone(baseline);
  delete missingRoot.prune;
  expectFailure(() => validateManifest(missingRoot), /root must contain exactly/);

  const extraRoot = structuredClone(baseline);
  extraRoot.labels = [];
  expectFailure(() => validateManifest(extraRoot), /root must contain exactly/);

  // The previous shape, schema 2 with one labels list, names its replacement.
  const previousSchema = { schema_version: 2, prune: baseline.prune, labels: structuredClone(baseline.core) };
  expectFailure(() => validateManifest(previousSchema),
    /schema_version 2 with a labels list is no longer supported: set schema_version 3, rename labels to core/);

  const wrongSchema = structuredClone(baseline);
  wrongSchema.schema_version = 1;
  expectFailure(() => validateManifest(wrongSchema), /schema_version/);

  const invalidPrune = structuredClone(baseline);
  invalidPrune.prune = "true";
  expectFailure(() => validateManifest(invalidPrune), /prune must be a boolean/);

  const emptyLabels = structuredClone(baseline);
  emptyLabels.core = [];
  expectFailure(() => validateManifest(emptyLabels), /core must not be empty/);

  const missingField = structuredClone(baseline);
  delete missingField.core[0].description;
  expectFailure(() => validateManifest(missingField), /name, color, and description/);

  const invalidColor = structuredClone(baseline);
  invalidColor.core[0].color = "#c24e00";
  expectFailure(() => validateManifest(invalidColor), /uppercase hexadecimal/);

  const duplicateName = structuredClone(baseline);
  duplicateName.core[1].name = baseline.core[0].name.toUpperCase();
  expectFailure(() => validateManifest(duplicateName), /duplicates/);

  const longDescription = structuredClone(baseline);
  longDescription.core[0].description = "x".repeat(101);
  expectFailure(() => validateManifest(longDescription), /100-character/);

  const unsorted = structuredClone(baseline);
  [unsorted.core[0], unsorted.core[1]] = [unsorted.core[1], unsorted.core[0]];
  expectFailure(() => validateManifest(unsorted), /core must be sorted/);

  // Overlays: an overlay is a sorted label array; the selection resolves core plus the selected ones.
  // The fixtures start from the core alone, so that the overlays a repository configures
  // do not change what they assert; the configured catalogue is checked by its own selection.
  const coreOnly = { ...structuredClone(baseline), overlays: {}, selected_overlays: [] };
  const core = resolveDesired(coreOnly);
  assert.deepEqual(core, [...manifest.core].sort((left, right) => compareNames(left.name, right.name)));
  const overlayLabel = { name: "zz-regime", color: "ABCDEF", description: "overlay label" };
  const secondLabel = { name: "zz-regime-review", color: "ABCDEF", description: "second overlay label" };
  const overlaid = structuredClone(coreOnly);
  overlaid.overlays = { regime: [overlayLabel, secondLabel], unused: [{ name: "zz-unused", color: "ABCDEF", description: "d" }] };
  overlaid.selected_overlays = ["regime"];
  const resolved = resolveDesired(overlaid);
  assert.deepEqual(resolved, [...core, overlayLabel, secondLabel], "resolved set is core plus the selected overlay");
  assert.deepEqual(catalogueCounts(overlaid), { core: core.length, overlays: 2, selected: 1, resolved: core.length + 2 });
  const unselected = structuredClone(overlaid);
  unselected.selected_overlays = [];
  assert.deepEqual(resolveDesired(unselected), core, "an overlay that is not selected adds nothing");
  const configured = {
    core: manifest.core.length,
    overlays: Object.keys(manifest.overlays).length,
    selected: manifest.selected_overlays.length,
    resolved: desired.length
  };
  assert.deepEqual(catalogueCounts(manifest), configured, "counts follow the configured selection");
  assert.equal(configured.resolved, configured.core
    + manifest.selected_overlays.reduce((sum, name) => sum + manifest.overlays[name].length, 0));

  const unknownSelection = structuredClone(overlaid);
  unknownSelection.selected_overlays = ["missing"];
  expectFailure(() => validateManifest(unknownSelection), /selected_overlays\[0\] selects unknown overlay "missing"/);
  const duplicatedSelection = structuredClone(overlaid);
  duplicatedSelection.selected_overlays = ["regime", "regime"];
  expectFailure(() => validateManifest(duplicatedSelection), /selected_overlays\[1\] selects overlay "regime" again/);
  const unsortedSelection = structuredClone(overlaid);
  unsortedSelection.selected_overlays = ["unused", "regime"];
  expectFailure(() => validateManifest(unsortedSelection), /selected_overlays must be sorted/);
  const notAList = structuredClone(overlaid);
  notAList.selected_overlays = "regime";
  expectFailure(() => validateManifest(notAList), /selected_overlays must be an array/);
  const crossDuplicate = structuredClone(overlaid);
  crossDuplicate.overlays.regime[0].name = baseline.core[0].name.toUpperCase();
  expectFailure(() => validateManifest(crossDuplicate), /overlays\["regime"\]\[0\]\.name duplicates core\[0\]\.name/);
  const overlayDuplicate = structuredClone(overlaid);
  overlayDuplicate.overlays.unused[0].name = overlayLabel.name;
  expectFailure(() => validateManifest(overlayDuplicate), /overlays\["unused"\]\[0\]\.name duplicates overlays\["regime"\]\[0\]\.name/);
  const unsortedOverlay = structuredClone(overlaid);
  unsortedOverlay.overlays.regime.reverse();
  expectFailure(() => validateManifest(unsortedOverlay), /overlays\["regime"\] must be sorted/);
  const unsortedOverlays = structuredClone(overlaid);
  unsortedOverlays.overlays = { unused: overlaid.overlays.unused, regime: overlaid.overlays.regime };
  expectFailure(() => validateManifest(unsortedOverlays), /overlays must be sorted case-insensitively by name: unused precedes regime/);
  const sameOverlayName = structuredClone(overlaid);
  sameOverlayName.overlays = { regime: overlaid.overlays.regime, Regime: overlaid.overlays.unused };
  expectFailure(() => validateManifest(sameOverlayName), /overlays\["Regime"\] duplicates overlay "regime"/);
  const emptyOverlay = structuredClone(overlaid);
  emptyOverlay.overlays.unused = [];
  expectFailure(() => validateManifest(emptyOverlay), /overlays\["unused"\] must not be empty/);
  const blankOverlayName = structuredClone(overlaid);
  blankOverlayName.overlays = { " ": overlaid.overlays.regime };
  blankOverlayName.selected_overlays = [];
  expectFailure(() => validateManifest(blankOverlayName), /overlays\[" "\] name must be non-empty/);
  const overlaysAsList = structuredClone(overlaid);
  overlaysAsList.overlays = [];
  expectFailure(() => validateManifest(overlaysAsList), /overlays must be an object/);
  const overlayNotArray = structuredClone(overlaid);
  overlayNotArray.overlays.regime = { name: "x" };
  expectFailure(() => validateManifest(overlayNotArray), /overlays\["regime"\] must be an array/);

  assert.equal(planChanges(desired, desired, { prune: manifest.prune }).length, 0, "idempotent plan must be empty");

  assert.deepEqual(planChanges(desired, desired.slice(1)).map(change => change.action), ["create"]);

  const changed = structuredClone(desired);
  changed[0].color = "000000";
  changed[0].description = "changed";
  const changedPlan = planChanges(desired, changed);
  assert.deepEqual(changedPlan.map(change => change.action), ["update"]);
  assert.deepEqual(changedPlan[0].fields, ["color", "description"]);

  const extra = [...structuredClone(desired), { name: "extra", color: "ABCDEF", description: "extra" }];
  assert.deepEqual(planChanges(desired, extra, { prune: true }).map(change => change.action), ["delete"]);
  assert.equal(planChanges(desired, extra, { prune: false }).length, 0);

  const caseDrift = structuredClone(desired);
  caseDrift[0].name = desired[0].name.toUpperCase();
  assert.deepEqual(planChanges(desired, caseDrift)[0].fields, ["name"]);

  const summary = renderSummary({
    mode: "plan",
    manifest,
    manifestPath: "fixture-labels.json",
    desiredLabels: desired,
    dryRun: true
  });
  assert.match(summary, /Label manifest: `fixture-labels\.json`/);
  assert.match(summary, new RegExp(`Labels: \\*\\*${desired.length}\\*\\*`));
  const validateSummary = renderSummary({
    mode: "validate",
    manifest,
    manifestPath: "fixture-labels.json",
    desiredLabels: desired,
    changes: null
  });
  assert.match(validateSummary, /Planned changes: \*\*not computed \(live labels not read\)\*\*/);
  assert.match(validateSummary, /Live labels were not read; no changes were planned\./);
  assert.doesNotMatch(validateSummary, /No label changes are required/);
  const failureSummary = renderSummary({
    mode: "failed",
    manifest: null,
    manifestPath: null,
    desiredLabels: [],
    changes: [{ action: "create", name: "new" }],
    partial: true,
    error: "write failed"
  });
  assert.match(failureSummary, /Prune: \*\*unavailable\*\*/);
  assert.match(failureSummary, /Planned changes: \*\*1\*\*/);
  assert.match(failureSummary, /some changes may already be applied/);
  assert.match(summary, new RegExp(`Labels \\(${desired.length}\\):`));
  for (const label of desired) assert.ok(summary.includes(inlineCode(label.name)), label.name);
  assert.ok(summary.includes(`Core / overlays defined / selected / resolved: **${configured.core} / `
    + `${configured.overlays} / ${configured.selected} / ${configured.resolved}**`), summary);
  assert.ok(summary.includes(`Selected overlays (${configured.selected}): ${configured.selected === 0
    ? "none" : manifest.selected_overlays.map(name => inlineCode(name)).join(", ")}`), summary);
  assert.match(failureSummary, /Core \/ overlays defined \/ selected \/ resolved: \*\*unavailable\*\*/);
  const overlaidSummary = renderSummary({
    mode: "plan",
    manifest: validateManifest(overlaid),
    manifestPath: "fixture-labels.json",
    desiredLabels: resolved,
    changes: planChanges(resolved, core)
  });
  assert.match(overlaidSummary, new RegExp(`Core / overlays defined / selected / resolved: \\*\\*${core.length} / 2 / 1 / ${core.length + 2}\\*\\*`));
  assert.match(overlaidSummary, /Selected overlays \(1\): `regime`/);
  assert.ok(overlaidSummary.includes("| create | `zz-regime` | missing live label |"), overlaidSummary);

  // The plan written for another tool is the planner's own change list.
  const planDirectory = await mkdtemp(join(tmpdir(), "labels-plan-"));
  try {
    const planPath = join(planDirectory, "nested", "plan.json");
    await writePlan(planPath, changedPlan);
    assert.deepEqual(await loadJson(planPath), { changes: changedPlan });
  } finally {
    await rm(planDirectory, { recursive: true, force: true });
  }
  expectFailure(() => parseArguments(["--plan-output", "plan.json"]), /--plan-output requires --mode plan/);
  expectFailure(() => parseArguments(["--mode", "validate", "--plan-output", "plan.json"]), /--plan-output requires --mode plan/);
  assert.equal(parseArguments(["--mode", "plan", "--live", "live.json", "--plan-output", "plan.json"]).planOutput, "plan.json");

  // Values are data: no backtick, pipe or line break may close the code span or split the row.
  assert.equal(inlineCode("plain"), "`plain`");
  assert.equal(inlineCode("wip `x`"), "`` wip `x` ``");
  assert.equal(inlineCode("a ``b`` c"), "```a ``b`` c```");
  assert.equal(inlineCode("`edge"), "`` `edge ``");
  assert.equal(inlineCode("a|b"), "`a|b`");
  assert.equal(inlineCode("a|b", { table: true }), "`a\\|b`");
  assert.equal(inlineCode("one\r\ntwo\rthree\nfour"), "`one two three four`");
  assert.equal(markdownEscape("x|y\r\nz"), "x\\|y z");
  const special = [{ name: "wip `x`", color: "ABCDEF", description: "d" },
    { name: "needs|triage", color: "ABCDEF", description: "d" }];
  const specialSummary = renderSummary({
    mode: "plan",
    manifest,
    manifestPath: "fixture-labels.json",
    desiredLabels: desired,
    changes: planChanges(desired, [...structuredClone(desired), ...special])
  });
  assert.ok(specialSummary.includes("| delete | `` wip `x` `` | outside the catalogue |"), specialSummary);
  assert.ok(specialSummary.includes("| delete | `needs\\|triage` | outside the catalogue |"), specialSummary);
  const errorSummary = renderSummary({
    mode: "failed",
    manifest,
    manifestPath: "fixture-labels.json",
    desiredLabels: desired,
    changes: null,
    error: "HTTP 422: {\"message\":\"`bad`\"}\r\nnext line"
  });
  assert.ok(errorSummary.includes("``HTTP 422: {\"message\":\"`bad`\"} next line``"), errorSummary);
  assert.doesNotMatch(errorSummary, /\r/);

  // Pruning refuses labels in use before any change; comma names are checked too.
  // The refusal names each label with the items that use it.
  const asked = [];
  const usage = used => async name => {
    asked.push(name);
    return used[name] ?? null;
  };
  const pruned = planChanges(desired, [...structuredClone(desired),
    { name: "unused", color: "ABCDEF", description: "d" },
    { name: "in-use", color: "ABCDEF", description: "d" },
    { name: "a,b", color: "ABCDEF", description: "d" }]);
  const twoUses = { items: [3, 7], more: false };
  const tenUses = { items: [1, 2, 3, 4, 5, 6, 7, 8, 9, 10], more: false };
  const manyUses = { items: [1, 2, 3, 4, 5, 6, 7, 8, 9, 10], more: true };
  await assert.rejects(assertSafeDeletions(pruned, usage({ "in-use": manyUses, "a,b": twoUses })),
    error => error.message.includes('still used by an issue or pull request: "a,b" (used by #3, #7), '
      + '"in-use" (used by #1, #2, #3, #4, #5, #6, #7, #8, #9, #10 and others). No change was applied'));
  assert.deepEqual(asked, ["a,b", "in-use", "unused"]);
  asked.length = 0;
  await assertSafeDeletions(pruned, usage({}));
  assert.deepEqual(asked, ["a,b", "in-use", "unused"], "an unused comma name can be pruned");
  asked.length = 0;

  // The sample: up to ten numbers, ascending, and "and others" exactly when
  // more items use the label.
  const odd = [21, 19, 17, 15, 13, 11, 9, 7, 5, 3];
  const oddSample = "#3, #5, #7, #9, #11, #13, #15, #17, #19, #21";
  assert.equal(usageSample([]), null);
  assert.deepEqual(usageSample([7, 3, 7]), { items: [3, 7], more: false });
  assert.equal(renderUsage(usageSample([7, 3])), "#3, #7", "under ten");
  assert.equal(renderUsage(usageSample(odd)), oddSample, "exactly ten: no others");
  assert.equal(renderUsage(usageSample([23, ...odd])), `${oddSample} and others`, "eleven");
  assert.equal(renderUsage(usageSample([99, 23, ...odd])), `${oddSample} and others`, "twelve");

  // A plain name costs one request: one page of eleven, oldest first, whose
  // eleventh item only says that there are others.
  const paths = [];
  const pages = responses => {
    let next = 0;
    return async (_repository, _token, method, suffix) => {
      assert.equal(method, "GET");
      paths.push(suffix);
      return responses[next++];
    };
  };
  const issue = (number, ...names) => ({ number, labels: names.map(name => ({ name })) });
  const plainQuery = "/issues?state=all&direction=asc&per_page=11&labels=plain";
  assert.deepEqual(await labelUsage("o/r", "t", "plain", pages([[issue(7, "plain"), issue(3, "plain")]])), twoUses);
  assert.deepEqual(paths, [plainQuery]);
  paths.length = 0;
  assert.equal(await labelUsage("o/r", "t", "plain", pages([[]])), null);
  assert.deepEqual(paths, [plainQuery]);
  paths.length = 0;
  const oddItems = odd.map(number => issue(number, "plain"));
  assert.deepEqual(await labelUsage("o/r", "t", "plain", pages([oddItems])),
    { items: [3, 5, 7, 9, 11, 13, 15, 17, 19, 21], more: false });
  assert.deepEqual(paths, [plainQuery]);
  paths.length = 0;
  assert.deepEqual(await labelUsage("o/r", "t", "plain", pages([[issue(23, "plain"), ...oddItems]])),
    { items: [3, 5, 7, 9, 11, 13, 15, 17, 19, 21], more: true });
  assert.deepEqual(paths, [plainQuery]);
  paths.length = 0;

  // A comma name is not passed to the labels filter: every issue and pull
  // request, open and closed, is read page by page and its real labels compared;
  // the numbers of every page are collected.
  const pageQueries = ["/issues?state=all&per_page=100&page=1", "/issues?state=all&per_page=100&page=2"];
  const full = Array.from({ length: 100 }, (_, index) => issue(200 - index, "a", "b"));
  full[10] = issue(190, "A,B");
  full[90] = issue(110, "a,b", "other");
  assert.deepEqual(await labelUsage("o/r", "t", "a,b", pages([full, [issue(5, "a,b"), issue(4, "x")]])),
    { items: [5, 110, 190], more: false });
  assert.deepEqual(paths, pageQueries);
  paths.length = 0;
  const allUse = Array.from({ length: 100 }, (_, index) => issue(200 - index, "a,b"));
  assert.deepEqual(await labelUsage("o/r", "t", "a,b", pages([allUse, [issue(3, "a,b")]])),
    { items: [3, 101, 102, 103, 104, 105, 106, 107, 108, 109], more: true });
  assert.deepEqual(paths, pageQueries);
  paths.length = 0;
  const noneUse = Array.from({ length: 100 }, (_, index) => issue(200 - index, "a", "b"));
  assert.equal(await labelUsage("o/r", "t", "a,b", pages([noneUse, [issue(1)]])), null);
  assert.equal(paths.length, 2);
  paths.length = 0;
  // Anything it cannot read stops synchronization before a change.
  await assert.rejects(labelUsage("o/r", "t", "a,b", pages([[{ number: 1, labels: null }]])), /malformed labels/);
  await assert.rejects(labelUsage("o/r", "t", "a,b", pages([[{ number: 1, labels: [{ name: 1 }] }]])), /malformed labels/);
  await assert.rejects(labelUsage("o/r", "t", "a,b", pages([{}])), /not an array/);
  await assert.rejects(labelUsage("o/r", "t", "plain", pages([{}])), /not an array/);
  await assert.rejects(labelUsage("o/r", "t", "a,b", async () => full), /exceeded 100 pages while checking "a,b"/);
  for (const number of [undefined, null, 0, -3, 1.5, "7", 2 ** 53]) {
    for (const name of ["plain", "a,b"]) {
      await assert.rejects(labelUsage("o/r", "t", name, pages([[issue(3, name), { number, labels: [{ name }] }]])),
        /an item without a valid number/, `${name}: ${number}`);
    }
  }
  const failing = pages([[], [], {}]);
  await assert.rejects(assertSafeDeletions(pruned, name => labelUsage("o/r", "t", name, failing)),
    /not an array/);
  await assertSafeDeletions(planChanges(desired, [...structuredClone(desired),
    { name: "unused", color: "ABCDEF", description: "d" }]), usage({}));
  assert.deepEqual(asked, ["unused"]);
  asked.length = 0;
  await assertSafeDeletions(planChanges(desired, desired.slice(1)), usage({ anything: twoUses }));
  assert.deepEqual(asked, [], "a plan without deletions makes no query");

  // A dry run previews the deletions a real run would refuse, each with the
  // items that use it; a plan checks none.
  assert.deepEqual(await deletionsInUse(pruned, usage({ "in-use": manyUses })), [{ name: "in-use", ...manyUses }]);
  const blocked = [{ name: "a,b", ...twoUses }, { name: "in-use", ...manyUses }];
  const previewed = renderSummary({
    mode: "reconcile",
    manifest,
    manifestPath: "fixture-labels.json",
    desiredLabels: desired,
    changes: pruned,
    inUse: blocked,
    dryRun: true
  });
  const refusedRow = "outside the catalogue; still in use, a real run would refuse it; used by";
  assert.ok(previewed.includes("- Deletions still in use: **2**"), previewed);
  assert.ok(previewed.includes(`| delete | \`a,b\` | ${refusedRow} #3, #7 |`), previewed);
  assert.ok(previewed.includes(`| delete | \`in-use\` | ${refusedRow} #1, #2, #3, #4, #5, #6, #7, #8, #9, #10 and others |`), previewed);
  assert.ok(previewed.includes("| delete | `unused` | outside the catalogue |"), previewed);
  assert.ok(previewed.includes('A real run would refuse every change. Refusing to prune labels still used by an issue or pull request: '
    + '"a,b" (used by #3, #7), "in-use" (used by #1, #2, #3, #4, #5, #6, #7, #8, #9, #10 and others). No change was applied.'), previewed);
  const exactlyTen = renderSummary({
    mode: "reconcile",
    manifest,
    manifestPath: "fixture-labels.json",
    desiredLabels: desired,
    changes: pruned,
    inUse: [{ name: "IN-USE", ...tenUses }],
    dryRun: true
  });
  assert.ok(exactlyTen.includes(`| delete | \`in-use\` | ${refusedRow} #1, #2, #3, #4, #5, #6, #7, #8, #9, #10 |`), exactlyTen);
  assert.doesNotMatch(exactlyTen, /and others/);
  // A real run's failure summary shows no table: its refusal names the items.
  const refusedRun = renderSummary({
    mode: "failed",
    manifest,
    manifestPath: "fixture-labels.json",
    desiredLabels: desired,
    changes: pruned,
    inUse: blocked,
    error: deletionRefusal(blocked)
  });
  assert.ok(refusedRun.includes('"a,b" (used by #3, #7), "in-use" (used by #1, #2, #3, #4, #5, #6, #7, #8, #9, #10 and others)'), refusedRun);
  assert.doesNotMatch(refusedRun, /\| delete \|/);
  const unchecked = renderSummary({ mode: "plan", manifest, manifestPath: "fixture-labels.json", desiredLabels: desired, changes: pruned });
  assert.ok(unchecked.includes("- Deletions still in use: **not checked (no repository read)**"), unchecked);
  assert.doesNotMatch(unchecked, /would refuse|used by/);
  const clean = renderSummary({ mode: "reconcile", manifest, manifestPath: "fixture-labels.json", desiredLabels: desired, changes: pruned, inUse: [], dryRun: true });
  assert.ok(clean.includes("- Deletions still in use: **0**"), clean);
  assert.doesNotMatch(clean, /would refuse|used by/);

  process.stdout.write("SELF-TEST PASS: retry delays, validation, overlays, reconciliation, safe pruning, label usage, plan output and summary fixtures\n");
}

// What the run knew before a failure, so that the failure summary reports it.
const runState = { summaryBase: null, changes: null, applying: false, verified: null, inUse: null };

async function main() {
  const options = parseArguments(process.argv.slice(2));
  if (options.selfTest) {
    await runSelfTest(options.manifest);
    return;
  }

  const manifest = validateManifest(await loadJson(options.manifest));
  const desired = resolveDesired(manifest);
  const summaryBase = {
    manifest,
    manifestPath: options.manifest,
    desiredLabels: desired
  };
  runState.summaryBase = summaryBase;

  if (options.mode === "validate") {
    await publishSummary(renderSummary({ mode: "validate", ...summaryBase, changes: null }));
    return;
  }

  let live;
  if (options.mode === "plan") {
    if (!options.live) throw new Error("--mode plan requires --live");
    const liveDocument = await loadJson(options.live);
    live = Array.isArray(liveDocument) ? liveDocument : liveDocument.labels;
    if (!Array.isArray(live)) throw new Error("live fixture must be an array or contain a labels array");
  } else if (options.mode === "reconcile") {
    if (!options.repository) throw new Error("--mode reconcile requires --repository or GITHUB_REPOSITORY");
    if (!process.env.GITHUB_TOKEN) throw new Error("--mode reconcile requires GITHUB_TOKEN");
    live = await listLiveLabels(options.repository, process.env.GITHUB_TOKEN);
  } else {
    throw new Error("--mode must be validate, plan, or reconcile");
  }

  const changes = planChanges(desired, live, { prune: manifest.prune });
  runState.changes = changes;
  // A plan reads no repository, so its deletions are not checked. Reconciliation,
  // a dry run included, reads the usage of every planned deletion before anything
  // else: the dry run previews the deletions a real run would refuse.
  const inUse = options.mode === "reconcile"
    ? await deletionsInUse(changes, name => labelUsage(options.repository, process.env.GITHUB_TOKEN, name))
    : null;
  runState.inUse = inUse;
  if (options.mode === "plan" || options.dryRun) {
    if (options.planOutput) await writePlan(options.planOutput, changes);
    await publishSummary(renderSummary({ mode: options.mode, ...summaryBase, changes, inUse, dryRun: options.dryRun }));
    return;
  }

  // Deleting a label cannot be undone: refuse before any change is applied.
  if (inUse.length > 0) throw new Error(deletionRefusal(inUse));
  runState.applying = true;
  await applyChanges(options.repository, process.env.GITHUB_TOKEN, changes);
  const after = await listLiveLabels(options.repository, process.env.GITHUB_TOKEN);
  const remaining = planChanges(desired, after, { prune: manifest.prune });
  const verified = remaining.length === 0;
  if (!verified) {
    runState.verified = false;
    throw new Error(`Post-run verification found ${remaining.length} remaining difference(s)`);
  }
  await publishSummary(renderSummary({ mode: "reconcile", ...summaryBase, changes, verified, inUse }));
}

// Run only when executed directly, so that labels-drift.mjs can import the planner.
if (process.argv[1] && import.meta.url === pathToFileURL(process.argv[1]).href) {
  main().catch(async error => {
    const message = error instanceof Error ? error.message : String(error);
    try {
      await publishSummary(renderSummary({
        mode: "failed",
        ...(runState.summaryBase ?? {
          manifest: null,
          manifestPath: null,
          desiredLabels: []
        }),
        changes: runState.changes,
        verified: runState.verified,
        inUse: runState.inUse,
        partial: runState.applying,
        error: message
      }));
    } finally {
      process.stderr.write(`${message}\n`);
      process.exitCode = 1;
    }
  });
}
