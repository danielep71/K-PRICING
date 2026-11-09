#!/usr/bin/env node

import assert from "node:assert/strict";
import { appendFile, readFile } from "node:fs/promises";
import process from "node:process";

import { codeSpan, selfTestMarkdown, tableCode, tableText } from "./markdown-summary.mjs";

const EXPECTED_SCHEMA_VERSION = 1;
const DEFAULT_MANIFEST = ".github/labels.json";
const DEFAULT_POLICY = ".github/repository-profile.json";
const API_VERSION = "2022-11-28";
const PROFILE_NAMES = ["application", "library", "ui-component"];
const LABELS_PER_PAGE = 100;
const MAX_LABEL_PAGES = 100;
const ISSUES_PER_PAGE = 100;
const MAX_ISSUE_PAGES = 100;
const MAX_RATE_LIMIT_RETRIES = 2;
const MAX_RETRY_DELAY_MS = 5000;

function compareNames(left, right) {
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

function normalizedLiveLabel(label) {
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

export function validateManifest(document) {
  const errors = [];
  if (!document || typeof document !== "object" || Array.isArray(document)) {
    throw new Error("Manifest must be a JSON object.");
  }
  if (!sameKeys(document, ["schema_version", "prune", "core", "overlays"])) {
    errors.push("root must contain exactly schema_version, prune, core, and overlays");
  }
  if (document.schema_version !== EXPECTED_SCHEMA_VERSION) {
    errors.push(`schema_version must be ${EXPECTED_SCHEMA_VERSION}`);
  }
  if (typeof document.prune !== "boolean") errors.push("prune must be a boolean");

  const seen = new Map();
  errors.push(...validateLabelArray(document.core, "core", seen));
  if (Array.isArray(document.core) && document.core.length === 0) errors.push("core must not be empty");

  const overlays = document.overlays;
  if (!overlays || typeof overlays !== "object" || Array.isArray(overlays)) {
    errors.push("overlays must be an object");
  } else if (!sameKeys(overlays, ["profile", "domain"])) {
    errors.push("overlays must contain exactly profile and domain");
  }

  const profiles = overlays?.profile;
  if (!profiles || typeof profiles !== "object" || Array.isArray(profiles)) {
    errors.push("overlays.profile must be an object");
  } else {
    if (!sameKeys(profiles, PROFILE_NAMES)) {
      errors.push(`overlays.profile must contain exactly ${PROFILE_NAMES.join(", ")}`);
    }
    for (const profile of PROFILE_NAMES) {
      errors.push(...validateLabelArray(profiles[profile], `overlays.profile.${profile}`, seen));
    }
  }

  const domains = overlays?.domain;
  if (!domains || typeof domains !== "object" || Array.isArray(domains)) {
    errors.push("overlays.domain must be an object");
  } else {
    for (const name of Object.keys(domains).sort()) {
      if (!/^[a-z0-9]+(?:-[a-z0-9]+)*$/.test(name)) {
        errors.push(`overlays.domain key is not kebab-case: ${name}`);
      }
      errors.push(...validateLabelArray(domains[name], `overlays.domain.${name}`, seen));
    }
  }

  if (errors.length > 0) throw new Error(`Invalid label manifest:\n- ${errors.join("\n- ")}`);
  return structuredClone(document);
}

export function resolveDesired(document, { profile = null, domains = [] } = {}) {
  const manifest = validateManifest(document);
  const labels = [...manifest.core];
  if (profile !== null) {
    if (!PROFILE_NAMES.includes(profile)) throw new Error(`Unknown profile overlay: ${profile}`);
    labels.push(...manifest.overlays.profile[profile]);
  }
  for (const domain of domains) {
    const overlay = manifest.overlays.domain[domain];
    if (!overlay) throw new Error(`Unknown domain overlay: ${domain}`);
    labels.push(...overlay);
  }
  labels.sort((left, right) => compareNames(left.name, right.name));
  return labels;
}

export function resolvePolicySelection(policy, manifest) {
  if (!policy || typeof policy !== "object" || Array.isArray(policy)) {
    throw new Error("Repository policy must be a JSON object");
  }
  if (!["template", "generated"].includes(policy.mode)) {
    throw new Error("Repository policy mode must be template or generated");
  }
  if (
    !Array.isArray(policy.label_domains)
    || policy.label_domains.some(domain => typeof domain !== "string")
  ) {
    throw new Error("Repository policy label_domains must be an array of strings");
  }
  const domains = [...policy.label_domains];
  if (new Set(domains).size !== domains.length) {
    throw new Error("Repository policy label_domains must not contain duplicates");
  }
  if (JSON.stringify(domains) !== JSON.stringify([...domains].sort(compareNames))) {
    throw new Error("Repository policy label_domains must be sorted case-insensitively");
  }
  for (const domain of domains) {
    if (!/^[a-z0-9]+(?:-[a-z0-9]+)*$/.test(domain)) {
      throw new Error(`Repository policy label domain is not kebab-case: ${domain}`);
    }
    if (!manifest.overlays.domain[domain]) {
      throw new Error(`Repository policy selects unknown domain overlay: ${domain}`);
    }
  }
  if (policy.mode === "template") {
    if (policy.profile !== null || domains.length > 0) {
      throw new Error("Template policy must select no profile or domain overlays");
    }
    return { profile: null, domains: [] };
  }
  if (!PROFILE_NAMES.includes(policy.profile)) {
    throw new Error("Generated policy must select one supported profile overlay");
  }
  return { profile: policy.profile, domains };
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
    policy: DEFAULT_POLICY,
    mode: null,
    live: null,
    repository: process.env.GITHUB_REPOSITORY || null,
    dryRun: false,
    selfTest: false
  };
  for (let index = 0; index < argv.length; index += 1) {
    const argument = argv[index];
    if (argument === "--dry-run") {
      parsed.dryRun = true;
    } else if (argument === "--self-test") {
      parsed.selfTest = true;
    } else if (["--manifest", "--policy", "--mode", "--live", "--repository"].includes(argument)) {
      const value = argv[index + 1];
      if (!value || value.startsWith("--")) throw new Error(`${argument} requires a value`);
      parsed[argument.slice(2)] = value;
      index += 1;
    } else {
      throw new Error(`Unknown argument: ${argument}`);
    }
  }
  return parsed;
}

async function loadJson(path) {
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
  return `https://api.github.com/repos/${owner}/${repo}${suffix}`;
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

function retryDelayMs(response, attempt) {
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
        "User-Agent": "canonical-vba-labels-sync"
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

async function listLiveLabels(repository, token) {
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

// Prune protection: a label outside the manifest is deleted only when no issue
// or pull request, open or closed, carries it. The issues endpoint lists pull
// requests too. A name without a comma is checked with the labels filter; the
// filter splits on commas, so a name with one is found by scanning every issue
// and pull request. A malformed response or an exhausted page limit fails, so
// a deletion is never applied on an incomplete check. request(suffix) returns
// the parsed JSON of a GET.
export async function assertSafeDeletions(changes, request) {
  const deletions = changes.filter(change => change.action === "delete").map(change => change.currentName);
  const used = [];
  for (const name of deletions.filter(item => !item.includes(","))) {
    const items = await request(`/issues?state=all&labels=${encodeURIComponent(name)}&per_page=1`);
    if (!Array.isArray(items)) throw new Error(`Malformed issue response for label ${JSON.stringify(name)}`);
    if (items.length > 0) used.push(name);
  }
  const pending = new Map(deletions.filter(item => item.includes(",")).map(name => [name.toLowerCase(), name]));
  for (let page = 1; pending.size > 0; page += 1) {
    if (page > MAX_ISSUE_PAGES) throw new Error(`Label usage scan exceeded ${MAX_ISSUE_PAGES} pages`);
    const items = await request(
      `/issues?state=all&sort=created&direction=asc&per_page=${ISSUES_PER_PAGE}&page=${page}`
    );
    if (!Array.isArray(items)) throw new Error("Malformed issue collection in the label usage scan");
    for (const item of items) {
      if (!item || typeof item !== "object" || !Array.isArray(item.labels)) {
        throw new Error("Malformed issue labels in the label usage scan");
      }
      for (const label of item.labels) {
        const name = typeof label === "string" ? label : label?.name;
        if (typeof name !== "string") throw new Error("Malformed issue label in the label usage scan");
        const key = name.toLowerCase();
        if (pending.has(key)) {
          used.push(pending.get(key));
          pending.delete(key);
        }
      }
    }
    if (items.length < ISSUES_PER_PAGE) break;
  }
  if (used.length > 0) {
    throw new Error(
      "Refusing to prune labels still used by issues or pull requests: "
      + used.sort(compareNames).map(name => JSON.stringify(name)).join(", ")
    );
  }
}

function renderLabelNames(labels) {
  if (labels.length === 0) return "none";
  return labels.map(label => JSON.stringify(label.name)).join(", ");
}

function renderSummary({
  mode,
  manifest,
  manifestPath,
  policyPath,
  desiredLabels,
  profile,
  domains,
  changes = [],
  verified = null,
  dryRun = false,
  error = null
}) {
  const coreLabels = manifest?.core ?? [];
  const profileLabels = profile === null ? [] : manifest?.overlays?.profile?.[profile] ?? [];
  const domainLabels = domains.flatMap(domain => manifest?.overlays?.domain?.[domain] ?? []);
  const resolvedLabels = desiredLabels ?? [];
  const counts = Object.fromEntries(["create", "update", "delete"].map(action => [
    action,
    changes.filter(change => change.action === action).length
  ]));
  const lines = [
    "# Repository label synchronization",
    "",
    `- Mode: \`${mode}\`${dryRun ? " (dry run)" : ""}`,
    `- Policy source: ${codeSpan(policyPath ?? "unavailable")}`,
    `- Label manifest: ${codeSpan(manifestPath ?? "unavailable")}`,
    `- Schema: **${manifest?.schema_version ?? "unavailable"}**`,
    `- Core / selected labels: **${coreLabels.length} / ${resolvedLabels.length}**`,
    `- Profile overlay: **${profile ?? "none"}**`,
    `- Domain overlays: **${domains.length > 0 ? domains.join(", ") : "none"}**`,
    `- Prune: **${manifest?.prune ? "enabled" : "disabled"}**`,
    `- Planned changes: **${changes.length}**`,
    `- Create / update / delete: **${counts.create} / ${counts.update} / ${counts.delete}**`
  ];
  if (verified !== null) lines.push(`- Post-run exact match: **${verified ? "yes" : "no"}**`);
  lines.push(
    "",
    "## Resolved label catalogue",
    "",
    `- Core labels (${coreLabels.length}): ${renderLabelNames(coreLabels)}`,
    `- Profile labels (${profile ?? "none"}; ${profileLabels.length}): ${renderLabelNames(profileLabels)}`,
    `- Domain labels (${domains.length > 0 ? domains.join(", ") : "none"}; ${domainLabels.length}): ${renderLabelNames(domainLabels)}`,
    `- Complete resolved set (${resolvedLabels.length}): ${renderLabelNames(resolvedLabels)}`
  );
  if (error) {
    lines.push("", "## Failure", "", codeSpan(error));
  } else if (changes.length === 0) {
    lines.push("", "No label changes are required.");
  } else {
    lines.push("", "| Action | Label | Detail |", "|---|---|---|");
    for (const change of changes) {
      const detail = change.action === "update" ? change.fields.join(", ") : change.action === "delete" ? "outside selected manifest" : "missing live label";
      lines.push(`| ${change.action} | ${tableCode(change.name)} | ${tableText(detail)} |`);
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

async function runSelfTest(manifestPath, policyPath) {
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
  const policy = await loadJson(policyPath);
  const desired = resolveDesired(manifest);

  const currentSelection = resolvePolicySelection(policy, manifest);
  assert.deepEqual(
    currentSelection,
    policy.mode === "template"
      ? { profile: null, domains: [] }
      : { profile: policy.profile, domains: policy.label_domains }
  );
  const templatePolicy = structuredClone(policy);
  templatePolicy.mode = "template";
  templatePolicy.profile = null;
  templatePolicy.label_domains = [];
  assert.deepEqual(resolvePolicySelection(templatePolicy, manifest), { profile: null, domains: [] });

  const generatedPolicy = structuredClone(templatePolicy);
  generatedPolicy.mode = "generated";
  generatedPolicy.profile = "library";
  assert.deepEqual(resolvePolicySelection(generatedPolicy, manifest), { profile: "library", domains: [] });

  const wrongPolicyMode = structuredClone(policy);
  wrongPolicyMode.mode = "unknown";
  expectFailure(() => resolvePolicySelection(wrongPolicyMode, manifest), /mode/);

  const selectedTemplateProfile = structuredClone(templatePolicy);
  selectedTemplateProfile.profile = "library";
  expectFailure(() => resolvePolicySelection(selectedTemplateProfile, manifest), /Template policy/);

  const missingGeneratedProfile = structuredClone(generatedPolicy);
  missingGeneratedProfile.profile = null;
  expectFailure(() => resolvePolicySelection(missingGeneratedProfile, manifest), /supported profile/);

  const invalidPolicyDomains = structuredClone(generatedPolicy);
  invalidPolicyDomains.label_domains = ["Unknown Domain"];
  expectFailure(() => resolvePolicySelection(invalidPolicyDomains, manifest), /kebab-case/);

  const missingRoot = structuredClone(baseline);
  delete missingRoot.prune;
  expectFailure(() => validateManifest(missingRoot), /root must contain exactly/);

  const wrongSchema = structuredClone(baseline);
  wrongSchema.schema_version = 2;
  expectFailure(() => validateManifest(wrongSchema), /schema_version/);

  const invalidPrune = structuredClone(baseline);
  invalidPrune.prune = "true";
  expectFailure(() => validateManifest(invalidPrune), /prune must be a boolean/);

  const missingField = structuredClone(baseline);
  delete missingField.core[0].description;
  expectFailure(() => validateManifest(missingField), /name, color, and description/);

  const invalidColor = structuredClone(baseline);
  invalidColor.core[0].color = "#c24e00";
  expectFailure(() => validateManifest(invalidColor), /uppercase hexadecimal/);

  const duplicateName = structuredClone(baseline);
  duplicateName.core[1].name = baseline.core[0].name.toUpperCase();
  expectFailure(() => validateManifest(duplicateName), /duplicates/);

  const overlayDuplicate = structuredClone(baseline);
  overlayDuplicate.overlays.profile.library.push(structuredClone(baseline.core[0]));
  expectFailure(() => validateManifest(overlayDuplicate), /duplicates/);

  const longDescription = structuredClone(baseline);
  longDescription.core[0].description = "x".repeat(101);
  expectFailure(() => validateManifest(longDescription), /100-character/);

  const unsorted = structuredClone(baseline);
  [unsorted.core[0], unsorted.core[1]] = [unsorted.core[1], unsorted.core[0]];
  expectFailure(() => validateManifest(unsorted), /sorted/);

  expectFailure(() => resolveDesired(manifest, { profile: "unknown" }), /Unknown profile/);
  expectFailure(() => resolveDesired(manifest, { domains: ["unknown"] }), /Unknown domain/);
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

  const overlay = structuredClone(baseline);
  overlay.overlays.profile.library = [{ name: "library", color: "654321", description: "Library profile label" }];
  overlay.overlays.domain.alpha = [{ name: "alpha", color: "123ABC", description: "Alpha domain label" }];
  overlay.overlays.domain.example = [{ name: "example", color: "ABCDEF", description: "Example domain label" }];
  const selectedOverlay = validateManifest(overlay);
  const selectedLabels = resolveDesired(selectedOverlay, { profile: "library", domains: ["example"] });
  assert.equal(selectedLabels.length, desired.length + 2);

  const domainPolicy = structuredClone(generatedPolicy);
  domainPolicy.label_domains = ["example"];
  assert.deepEqual(resolvePolicySelection(domainPolicy, overlay), { profile: "library", domains: ["example"] });

  const unknownDomainPolicy = structuredClone(generatedPolicy);
  unknownDomainPolicy.label_domains = ["unknown"];
  expectFailure(() => resolvePolicySelection(unknownDomainPolicy, manifest), /unknown domain/);

  const duplicateDomainPolicy = structuredClone(generatedPolicy);
  duplicateDomainPolicy.label_domains = ["example", "example"];
  expectFailure(() => resolvePolicySelection(duplicateDomainPolicy, overlay), /duplicates/);

  const unsortedDomainPolicy = structuredClone(generatedPolicy);
  unsortedDomainPolicy.label_domains = ["example", "alpha"];
  expectFailure(() => resolvePolicySelection(unsortedDomainPolicy, overlay), /sorted/);

  const summary = renderSummary({
    mode: "plan",
    manifest: selectedOverlay,
    manifestPath: "fixture-labels.json",
    policyPath: "fixture-policy.json",
    desiredLabels: selectedLabels,
    profile: "library",
    domains: ["example"],
    dryRun: true
  });
  assert.match(summary, /Policy source: `fixture-policy\.json`/);
  assert.match(summary, /Label manifest: `fixture-labels\.json`/);
  assert.match(summary, /Profile labels \(library; 1\): "library"/);
  assert.match(summary, /Domain labels \(example; 1\): "example"/);
  assert.match(summary, new RegExp(`Complete resolved set \\(${selectedLabels.length}\\):`));
  for (const label of selectedLabels) assert.match(summary, new RegExp(`"${label.name}"`));

  selfTestMarkdown(assert);
  const oddPlan = planChanges(desired, [...structuredClone(desired), { name: "a`|b", color: "ABCDEF", description: "odd" }]);
  const oddSummary = renderSummary({
    mode: "plan",
    manifest,
    manifestPath: "x`y.json",
    policyPath: "p.json",
    desiredLabels: desired,
    profile: null,
    domains: [],
    changes: oddPlan,
    error: null
  });
  assert.match(oddSummary, /^\| delete \| ``a`\\\|b`` \| outside selected manifest \|$/m);
  assert.match(oddSummary, /Label manifest: ``x`y\.json``/);
  assert.match(
    renderSummary({ mode: "failed", manifest: null, manifestPath: null, policyPath: null, desiredLabels: [],
      profile: null, domains: [], error: "bad `x`\nnext" }),
    /## Failure\n\n``bad `x` next``\n/
  );

  await selfTestSafeDeletions();

  process.stdout.write(
    "SELF-TEST PASS: retry delays, validation, policy, overlay, reconciliation, prune protection, and summary fixtures\n"
  );
}

function fakeIssues(pages) {
  const calls = [];
  const request = async suffix => {
    calls.push(suffix);
    if (typeof pages === "function") return pages(suffix);
    const page = Number(new URL(`https://x${suffix}`).searchParams.get("page") ?? "1");
    return pages[page - 1] ?? [];
  };
  return { calls, request };
}

function issuesWith(count, labels = []) {
  return Array.from({ length: count }, () => ({ labels }));
}

async function selfTestSafeDeletions() {
  const del = name => ({ action: "delete", name, currentName: name });
  const none = fakeIssues([]);
  await assertSafeDeletions([{ action: "create", name: "x", target: {} }], none.request);
  assert.equal(none.calls.length, 0, "no deletion needs no usage check");

  const unused = fakeIssues(() => []);
  await assertSafeDeletions([del("stale label")], unused.request);
  assert.deepEqual(unused.calls, ["/issues?state=all&labels=stale%20label&per_page=1"]);

  const used = fakeIssues(suffix => (suffix.includes("labels=in%20use") ? [{ number: 7 }] : []));
  await assert.rejects(
    assertSafeDeletions([del("free"), del("in use")], used.request),
    /still used by issues or pull requests: "in use"$/
  );

  const commaUsed = fakeIssues([issuesWith(ISSUES_PER_PAGE, ["x"]), [{ labels: [{ name: "A, B" }] }]]);
  await assert.rejects(assertSafeDeletions([del("a, b")], commaUsed.request), /"a, b"/);
  assert.equal(commaUsed.calls.length, 2, "a comma name is found by a paginated scan");
  assert.ok(commaUsed.calls.every(call => !call.includes("labels=")));

  const commaFree = fakeIssues([issuesWith(ISSUES_PER_PAGE, ["x"]), issuesWith(3, ["y"])]);
  await assertSafeDeletions([del("a, b")], commaFree.request);
  assert.equal(commaFree.calls.length, 2);

  await assert.rejects(assertSafeDeletions([del("x")], fakeIssues(() => ({})).request), /Malformed issue response/);
  await assert.rejects(assertSafeDeletions([del("a,b")], fakeIssues(() => ({})).request), /Malformed issue collection/);
  await assert.rejects(
    assertSafeDeletions([del("a,b")], fakeIssues(() => [{ labels: null }]).request),
    /Malformed issue labels/
  );
  await assert.rejects(
    assertSafeDeletions([del("a,b")], fakeIssues(() => [{ labels: [{}] }]).request),
    /Malformed issue label in/
  );
  await assert.rejects(
    assertSafeDeletions([del("a,b")], fakeIssues(() => issuesWith(ISSUES_PER_PAGE, ["x"])).request),
    /exceeded 100 pages/
  );
}

async function main() {
  const options = parseArguments(process.argv.slice(2));
  if (options.selfTest) {
    await runSelfTest(options.manifest, options.policy);
    return;
  }

  const manifest = validateManifest(await loadJson(options.manifest));
  const selection = resolvePolicySelection(await loadJson(options.policy), manifest);
  const desired = resolveDesired(manifest, selection);
  const summaryBase = {
    manifest,
    manifestPath: options.manifest,
    policyPath: options.policy,
    desiredLabels: desired,
    profile: selection.profile,
    domains: selection.domains
  };

  if (options.mode === "validate") {
    await publishSummary(renderSummary({ mode: "validate", ...summaryBase }));
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
  if (options.mode === "plan" || options.dryRun) {
    await publishSummary(renderSummary({ mode: options.mode, ...summaryBase, changes, dryRun: options.dryRun }));
    return;
  }

  await assertSafeDeletions(
    changes,
    suffix => githubRequest(options.repository, process.env.GITHUB_TOKEN, "GET", suffix)
  );
  await applyChanges(options.repository, process.env.GITHUB_TOKEN, changes);
  const after = await listLiveLabels(options.repository, process.env.GITHUB_TOKEN);
  const remaining = planChanges(desired, after, { prune: manifest.prune });
  const verified = remaining.length === 0;
  await publishSummary(renderSummary({ mode: "reconcile", ...summaryBase, changes, verified }));
  if (!verified) throw new Error(`Post-run verification found ${remaining.length} remaining difference(s)`);
}

main().catch(async error => {
  const message = error instanceof Error ? error.message : String(error);
  try {
    await publishSummary(renderSummary({
      mode: "failed",
      manifest: null,
      manifestPath: null,
      policyPath: null,
      desiredLabels: [],
      profile: null,
      domains: [],
      error: message
    }));
  } finally {
    process.stderr.write(`${message}\n`);
    process.exitCode = 1;
  }
});
