#!/usr/bin/env node

import assert from "node:assert/strict";
import { execFile } from "node:child_process";
import { appendFile, mkdir, mkdtemp, rm, writeFile } from "node:fs/promises";
import { tmpdir } from "node:os";
import { join } from "node:path";
import process from "node:process";
import { fileURLToPath } from "node:url";
import { promisify } from "node:util";

// The synchronization script owns validation, planning and the GitHub client;
// importing them keeps the drift report and the reconciler on one plan. The
// cross-check then runs that script as a separate process, in plan mode on the
// same live snapshot, and fails if the two plans differ.
import {
  MAX_RETRY_DELAY_MS,
  catalogueCounts,
  compareNames,
  inlineCode,
  listLiveLabels,
  loadJson,
  markdownEscape,
  normalizedLiveLabel,
  planChanges,
  resolveDesired,
  retryDelayMs,
  validateManifest
} from "./labels-sync.mjs";

const TOOL_NAME = "Repository label drift";
const DEFAULT_MANIFEST = ".github/labels.json";
const DEFAULT_RECONCILER = fileURLToPath(new URL("./labels-sync.mjs", import.meta.url));
const run = promisify(execFile);

function parseArguments(argv) {
  const parsed = {
    manifest: DEFAULT_MANIFEST,
    live: null,
    repository: process.env.GITHUB_REPOSITORY || null,
    output: null,
    summary: null,
    reconciler: DEFAULT_RECONCILER,
    selfTest: false
  };
  for (let index = 0; index < argv.length; index += 1) {
    const argument = argv[index];
    if (argument === "--self-test") {
      parsed.selfTest = true;
      continue;
    }
    if (["--manifest", "--live", "--repository", "--output", "--summary", "--reconciler"].includes(argument)) {
      const value = argv[index + 1];
      if (!value || value.startsWith("--")) throw new Error(`${argument} requires a value`);
      parsed[argument.slice(2)] = value;
      index += 1;
      continue;
    }
    throw new Error(`Unknown argument: ${argument}`);
  }
  return parsed;
}

function changeEvidence(change, observedByName) {
  const currentKey = String(change.currentName ?? change.name).toLowerCase();
  return {
    action: change.action,
    name: change.name,
    fields: change.fields ?? [],
    observed: observedByName.get(currentKey) ?? null,
    desired: change.target ?? null
  };
}

function buildReport({ repository, manifestPath, manifest, desiredLabels, liveLabels }) {
  const observed = liveLabels
    .map(normalizedLiveLabel)
    .sort((left, right) => compareNames(left.name, right.name));
  const desired = structuredClone(desiredLabels)
    .sort((left, right) => compareNames(left.name, right.name));
  const changes = planChanges(desired, observed, { prune: manifest.prune });
  const observedByName = new Map(observed.map(label => [label.name.toLowerCase(), label]));
  const differences = changes.map(change => changeEvidence(change, observedByName));
  const counts = Object.fromEntries(["create", "update", "delete"].map(action => [
    action,
    differences.filter(item => item.action === action).length
  ]));
  const catalogue = catalogueCounts(manifest);
  return {
    schema_version: 3,
    tool: TOOL_NAME,
    status: differences.length === 0 ? "pass" : "drift",
    repository,
    manifest: manifestPath,
    prune: manifest.prune,
    counts: {
      core: catalogue.core,
      overlays: catalogue.overlays,
      selected_overlays: catalogue.selected,
      desired: desired.length,
      observed: observed.length,
      differences: differences.length,
      create: counts.create,
      update: counts.update,
      delete: counts.delete
    },
    selected_overlays: [...manifest.selected_overlays],
    desired,
    observed,
    differences,
    // The plan comes from the synchronization script's own planner; the
    // cross-check below confirms it against that script run as a process.
    canonical_plan_verified: true,
    cross_check: null
  };
}

// What the two plans are compared on: the action, the label and the fields
// to update, in the planner's order.
function planKey(change) {
  return { action: change.action, name: change.name, fields: [...(change.fields ?? [])] };
}

function orderedKeys(changes) {
  const order = { update: 0, create: 1, delete: 2 };
  return changes.map(planKey)
    .sort((left, right) => order[left.action] - order[right.action] || compareNames(left.name, right.name));
}

// Run the reconciler in plan mode on the live labels this report read and
// compare its plan with the report's differences. The child inherits the
// environment without GITHUB_STEP_SUMMARY, so its own summary stays out of
// the job summary.
export async function crossCheck(report, { reconciler, manifestPath, liveLabels }) {
  const directory = await mkdtemp(join(tmpdir(), "labels-drift-"));
  try {
    const livePath = join(directory, "live.json");
    const planPath = join(directory, "plan.json");
    await writeFile(livePath, `${JSON.stringify(liveLabels)}\n`, "utf8");
    const { GITHUB_STEP_SUMMARY: _ignored, ...env } = process.env;
    try {
      await run(process.execPath, [reconciler, "--manifest", manifestPath, "--mode", "plan",
        "--live", livePath, "--plan-output", planPath], { env, maxBuffer: 16 * 1024 * 1024 });
    } catch (error) {
      const detail = String(error.stderr || error.message).trim().split("\n")[0];
      throw new Error(`Reconciler plan failed (${reconciler}): ${detail}`);
    }
    const plan = await loadJson(planPath);
    if (!Array.isArray(plan?.changes)) throw new Error(`Reconciler plan in ${planPath} has no changes array`);
    const reported = orderedKeys(report.differences);
    const planned = orderedKeys(plan.changes);
    const agree = JSON.stringify(reported) === JSON.stringify(planned);
    return agree
      ? { status: "agree" }
      : { status: "disagree", reconciler, report_plan: reported, reconciler_plan: planned };
  } finally {
    await rm(directory, { recursive: true, force: true });
  }
}

// The report with its cross-check: a disagreement fails the run, whatever the drift.
export async function crossCheckedReport(report, options) {
  const result = await crossCheck(report, options);
  return {
    ...report,
    status: result.status === "agree" ? report.status : "fail",
    canonical_plan_verified: result.status === "agree",
    cross_check: result
  };
}

function compactLabel(label) {
  if (label === null) return "—";
  return `${inlineCode(label.name, { table: true })} / ${inlineCode(label.color, { table: true })} / `
    + markdownEscape(label.description);
}

function markdownReport(report) {
  const lines = [
    "# Repository label drift",
    "",
    `- Status: **${String(report.status).toUpperCase()}**`,
    `- Repository: ${inlineCode(report.repository ?? "fixture")}`,
    `- Manifest: ${inlineCode(report.manifest)}`,
    `- Core / overlays defined / selected / resolved: **${report.counts.core} / ${report.counts.overlays} / `
    + `${report.counts.selected_overlays} / ${report.counts.desired}**`,
    `- Desired / observed: **${report.counts.desired} / ${report.counts.observed}**`,
    `- Differences: **${report.counts.differences}**`,
    `- Create / update / delete: **${report.counts.create} / ${report.counts.update} / ${report.counts.delete}**`,
    "- Plan: computed by the synchronization script's planner",
    `- Cross-check against the reconciler's plan on the same snapshot: **${
      report.cross_check === null ? "not run" : report.cross_check.status === "agree" ? "agree" : "DISAGREE"}**`,
    ""
  ];
  if (report.cross_check !== null && report.cross_check.status !== "agree") {
    lines.push(
      "## Cross-check",
      "",
      "The drift report and the reconciler planned different changes on the same live labels. "
      + "Neither plan can be trusted until the scripts agree.",
      ""
    );
    for (const [title, plan] of [["This report", report.cross_check.report_plan],
      ["Reconciler", report.cross_check.reconciler_plan]]) {
      lines.push(`### ${title} (${plan.length})`, "");
      if (plan.length === 0) {
        lines.push("No changes planned.", "");
        continue;
      }
      lines.push("| Action | Label | Fields |", "| --- | --- | --- |");
      for (const change of plan) {
        lines.push(`| ${change.action} | ${inlineCode(change.name, { table: true })} | `
          + `${change.fields.length > 0 ? change.fields.join(", ") : "—"} |`);
      }
      lines.push("");
    }
    lines.push("## Differences", "");
  }
  if (report.differences.length === 0) {
    lines.push("No live label drift detected.");
  } else {
    lines.push(
      "| Action | Label | Fields | Observed | Desired |",
      "| --- | --- | --- | --- | --- |"
    );
    for (const difference of report.differences) {
      lines.push(
        `| ${difference.action} | ${inlineCode(difference.name, { table: true })} | `
        + `${difference.fields.length > 0 ? difference.fields.join(", ") : "—"} | `
        + `${compactLabel(difference.observed)} | ${compactLabel(difference.desired)} |`
      );
    }
  }
  return `${lines.join("\n")}\n`;
}

async function writeEvidence(path, content) {
  if (!path) return;
  const slash = path.lastIndexOf("/");
  if (slash > 0) await mkdir(path.slice(0, slash), { recursive: true });
  await writeFile(path, content, "utf8");
}

async function publishSummary(summary) {
  process.stdout.write(summary);
  if (process.env.GITHUB_STEP_SUMMARY) {
    await appendFile(process.env.GITHUB_STEP_SUMMARY, summary, "utf8");
  }
}

async function prepareContract(manifestPath) {
  const manifest = validateManifest(await loadJson(manifestPath));
  return { manifest, desiredLabels: resolveDesired(manifest) };
}

async function runSelfTest(manifestPath, reconciler) {
  for (const [header, attempt, expected] of [
    [null, 0, 1000], [null, 2, 4000], ["", 1, 2000],
    ["invalid", 1, 2000], ["-1", 1, 2000], ["0", 2, 0],
    ["2", 0, 2000], ["999999", 0, MAX_RETRY_DELAY_MS],
    [null, 20, MAX_RETRY_DELAY_MS]
  ]) {
    const headers = new Headers(header === null ? {} : { "retry-after": header });
    assert.equal(retryDelayMs({ headers }, attempt), expected, `Retry-After: ${header}`);
  }

  const { manifest, desiredLabels } = await prepareContract(manifestPath);

  const baseline = structuredClone(desiredLabels);
  const baselineBefore = JSON.stringify(baseline);
  const cleanFirst = buildReport({
    repository: "fixture/clean",
    manifestPath,
    manifest,
    desiredLabels,
    liveLabels: baseline
  });
  const cleanSecond = buildReport({
    repository: "fixture/clean",
    manifestPath,
    manifest,
    desiredLabels,
    liveLabels: baseline
  });
  assert.equal(cleanFirst.status, "pass");
  assert.equal(cleanFirst.counts.differences, 0);
  assert.equal(JSON.stringify(cleanFirst), JSON.stringify(cleanSecond));
  assert.equal(markdownReport(cleanFirst), markdownReport(cleanSecond));
  assert.equal(JSON.stringify(baseline), baselineBefore, "clean check must not mutate live input");

  const drifted = structuredClone(desiredLabels);
  const missing = drifted.shift();
  assert.ok(missing, "fixture requires at least one desired label");
  drifted[0].color = drifted[0].color === "000000" ? "FFFFFF" : "000000";
  drifted[0].description = "simulated out-of-band change";
  drifted.push({ name: "out-of-band", color: "ABCDEF", description: "simulated extra label" });
  const driftedBefore = JSON.stringify(drifted);
  const driftFirst = buildReport({
    repository: "fixture/drift",
    manifestPath,
    manifest,
    desiredLabels,
    liveLabels: drifted
  });
  const driftSecond = buildReport({
    repository: "fixture/drift",
    manifestPath,
    manifest,
    desiredLabels,
    liveLabels: drifted
  });
  assert.equal(driftFirst.status, "drift");
  assert.deepEqual(driftFirst.differences.map(item => item.action), ["update", "create", "delete"]);
  assert.equal(JSON.stringify(driftFirst), JSON.stringify(driftSecond));
  assert.equal(markdownReport(driftFirst), markdownReport(driftSecond));
  assert.equal(JSON.stringify(drifted), driftedBefore, "drift check must not mutate live input");
  assert.match(markdownReport(driftFirst), /simulated out-of-band change/);
  assert.match(markdownReport(driftFirst), new RegExp(missing.name));

  // Names with Markdown syntax are data, not table syntax.
  const special = [...structuredClone(desiredLabels),
    { name: "needs|triage", color: "ABCDEF", description: "pipe" },
    { name: "wip `x`", color: "ABCDEF", description: "backtick" }];
  const specialReport = buildReport({
    repository: "fixture/special", manifestPath, manifest, desiredLabels, liveLabels: special
  });
  assert.equal(specialReport.status, "drift");
  assert.deepEqual(specialReport.differences.map(item => [item.action, item.name]),
    [["delete", "needs|triage"], ["delete", "wip `x`"]]);
  const specialMarkdown = markdownReport(specialReport);
  assert.ok(specialMarkdown.includes("| delete | `needs\\|triage` | — | `needs\\|triage` / `ABCDEF` / pipe | — |"),
    specialMarkdown);
  assert.ok(specialMarkdown.includes("| delete | `` wip `x` `` | — | `` wip `x` `` / `ABCDEF` / backtick | — |"),
    specialMarkdown);

  // Live values come from the API: a line break or pipe in them must not split the row.
  const hostile = buildReport({
    repository: "fixture/hostile", manifestPath, manifest, desiredLabels,
    liveLabels: [...structuredClone(desiredLabels),
      { name: "x", color: "AB|`\r\n", description: "line\r\nbreak | pipe" }]
  });
  const hostileRow = markdownReport(hostile).split("\n").find(line => line.startsWith("| delete |"));
  assert.equal(hostileRow, "| delete | `x` | — | `x` / `` AB\\|`  `` / line break \\| pipe | — |");

  // Counts and the selection are reported before the cross-check runs.
  assert.deepEqual(
    [cleanFirst.counts.core, cleanFirst.counts.overlays, cleanFirst.counts.selected_overlays, cleanFirst.counts.desired],
    [manifest.core.length, Object.keys(manifest.overlays).length, manifest.selected_overlays.length, desiredLabels.length]);
  assert.deepEqual(cleanFirst.selected_overlays, manifest.selected_overlays);
  assert.equal(cleanFirst.cross_check, null);
  assert.match(markdownReport(cleanFirst), /Cross-check against the reconciler's plan on the same snapshot: \*\*not run\*\*/);

  // The reconciler, run in plan mode on the same snapshot, plans the same changes.
  assert.equal(parseArguments(["--reconciler", "other.mjs"]).reconciler, "other.mjs");
  assert.equal(parseArguments([]).reconciler, DEFAULT_RECONCILER);
  for (const [report, liveLabels, status] of [[cleanFirst, baseline, "pass"], [driftFirst, drifted, "drift"],
    [specialReport, special, "drift"]]) {
    const checked = await crossCheckedReport(report, { reconciler, manifestPath, liveLabels });
    assert.equal(checked.status, status);
    assert.deepEqual(checked.cross_check, { status: "agree" });
    assert.equal(checked.canonical_plan_verified, true);
    assert.equal(JSON.stringify(report.differences), JSON.stringify(checked.differences));
    assert.match(markdownReport(checked), /Cross-check against the reconciler's plan on the same snapshot: \*\*agree\*\*/);
    assert.doesNotMatch(markdownReport(checked), /## Cross-check/);
  }
  assert.equal(cleanFirst.cross_check, null, "the cross-check must not mutate the report it was given");

  // A reconciler that plans something else fails the run, and both plans are reported.
  const fakeDirectory = await mkdtemp(join(tmpdir(), "labels-fake-reconciler-"));
  try {
    const fake = join(fakeDirectory, "fake-reconciler.mjs");
    await writeFile(fake, [
      "import { writeFileSync } from \"node:fs\";",
      "const path = process.argv[process.argv.indexOf(\"--plan-output\") + 1];",
      "writeFileSync(path, JSON.stringify({ changes: [{ action: \"delete\", name: \"phantom\", currentName: \"phantom\" }] }));",
      ""
    ].join("\n"), "utf8");
    const disagreed = await crossCheckedReport(cleanFirst, { reconciler: fake, manifestPath, liveLabels: baseline });
    assert.equal(disagreed.status, "fail");
    assert.equal(disagreed.canonical_plan_verified, false);
    assert.deepEqual(disagreed.cross_check, {
      status: "disagree",
      reconciler: fake,
      report_plan: [],
      reconciler_plan: [{ action: "delete", name: "phantom", fields: [] }]
    });
    const disagreedMarkdown = markdownReport(disagreed);
    assert.match(disagreedMarkdown, /- Status: \*\*FAIL\*\*/);
    assert.match(disagreedMarkdown, /Cross-check against the reconciler's plan on the same snapshot: \*\*DISAGREE\*\*/);
    assert.match(disagreedMarkdown, /## Cross-check\n\n[^\n]+\n\n### This report \(0\)\n\nNo changes planned\.\n\n### Reconciler \(1\)\n/);
    assert.ok(disagreedMarkdown.includes("| delete | `phantom` | — |"), disagreedMarkdown);
    assert.match(disagreedMarkdown, /## Differences\n\nNo live label drift detected\./);
    // Drift that the reconciler also plans, with a field more, is a disagreement too.
    const driftChecked = await crossCheckedReport(driftFirst, { reconciler: fake, manifestPath, liveLabels: drifted });
    assert.equal(driftChecked.status, "fail");
    assert.equal(driftChecked.cross_check.report_plan.length, 3);

    // A reconciler that fails or writes no plan stops the run as an error, not as agreement.
    await writeFile(fake, "process.exit(3);\n", "utf8");
    await assert.rejects(crossCheck(cleanFirst, { reconciler: fake, manifestPath, liveLabels: baseline }),
      /Reconciler plan failed/);
    await writeFile(fake, [
      "import { writeFileSync } from \"node:fs\";",
      "writeFileSync(process.argv[process.argv.indexOf(\"--plan-output\") + 1], \"{}\");",
      ""
    ].join("\n"), "utf8");
    await assert.rejects(crossCheck(cleanFirst, { reconciler: fake, manifestPath, liveLabels: baseline }),
      /has no changes array/);
  } finally {
    await rm(fakeDirectory, { recursive: true, force: true });
  }

  process.stdout.write(
    "SELF-TEST PASS: retry delays and deterministic no-drift, create/update/delete and special-name drift fixtures are read-only; "
    + "the reconciler's plan agrees on them and an injected disagreement fails.\n"
  );
}

async function main() {
  const options = parseArguments(process.argv.slice(2));
  if (options.selfTest) {
    await runSelfTest(options.manifest, options.reconciler);
    return;
  }

  const { manifest, desiredLabels } = await prepareContract(options.manifest);

  let liveLabels;
  if (options.live) {
    const liveDocument = await loadJson(options.live);
    liveLabels = Array.isArray(liveDocument) ? liveDocument : liveDocument.labels;
    if (!Array.isArray(liveLabels)) {
      throw new Error("live fixture must be an array or contain a labels array");
    }
  } else {
    if (!options.repository) throw new Error("live check requires --repository or GITHUB_REPOSITORY");
    if (!process.env.GITHUB_TOKEN) throw new Error("live check requires GITHUB_TOKEN");
    liveLabels = await listLiveLabels(options.repository, process.env.GITHUB_TOKEN);
  }

  const report = await crossCheckedReport(buildReport({
    repository: options.repository,
    manifestPath: options.manifest,
    manifest,
    desiredLabels,
    liveLabels
  }), { reconciler: options.reconciler, manifestPath: options.manifest, liveLabels });
  const json = `${JSON.stringify(report, null, 2)}\n`;
  const markdown = markdownReport(report);
  await writeEvidence(options.output, json);
  await writeEvidence(options.summary, markdown);
  await publishSummary(markdown);
  process.exitCode = report.status === "pass" ? 0 : 1;
}

main().catch(error => {
  const message = error instanceof Error ? error.message : String(error);
  process.stderr.write(`ERROR: ${message}\n`);
  process.exitCode = 2;
});
