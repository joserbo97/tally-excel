#!/usr/bin/env node
// Excel-agent benchmark: runs realistic tasks against the real model + the agent's real tools on a
// mock workbook, then scores the resulting cells (formula results via HyperFormula, plus "live
// formula" checks that change an input and re-check the output).
//
//   node scripts/bench/run.mjs                       # all tasks, qwen3.8:latest
//   node scripts/bench/run.mjs --tasks formulas,fix  # by category or task id
//   node scripts/bench/run.mjs --model other:tag --think medium --repeat 3
//   node scripts/bench/run.mjs --suite hard          # core (default) | hard | all
//
// Needs Ollama, and the add-in server (https://localhost:3443) for tasks with attachments.
import fs from "node:fs";
import os from "node:os";
import path from "node:path";
import { execFileSync, spawnSync } from "node:child_process";
import { fileURLToPath } from "node:url";

const HERE = path.dirname(fileURLToPath(import.meta.url));
const ROOT = path.resolve(HERE, "../..");
const CA = path.join(os.homedir(), ".office-addin-dev-certs/ca.crt");

// Node only reads NODE_EXTRA_CA_CERTS at startup: re-exec once with it so https://localhost works.
if (!process.env.NODE_EXTRA_CA_CERTS && fs.existsSync(CA)) {
  const r = spawnSync(process.execPath, process.argv.slice(1), { stdio: "inherit", env: { ...process.env, NODE_EXTRA_CA_CERTS: CA } });
  process.exit(r.status ?? 1);
}

// ---------- args ----------
const args = Object.fromEntries(
  process.argv.slice(2).reduce((acc, a, i, arr) => (a.startsWith("--") ? [...acc, [a.slice(2), arr[i + 1]?.startsWith("--") || arr[i + 1] === undefined ? true : arr[i + 1]]] : acc), [])
);
const settings = {
  model: args.model || "qwen3.8:latest",
  think: args.think || "low",
  numCtx: Number(args.ctx || 32768),
  temperature: Number(args.temperature ?? 0.6),
  maxSteps: Number(args.steps || 30),
  approval: "none",
  web: false, // keep runs deterministic and offline
};
const repeat = Number(args.repeat || 1);
const TIMEOUT_MS = Number(args.timeout || 8) * 60 * 1000;
globalThis.OLLAMA_API_BASE = process.env.OLLAMA_API_BASE || "http://127.0.0.1:11434/api";
globalThis.ADDIN_BASE = process.env.ADDIN_BASE || "https://localhost:3443";

const { TASKS: CORE, CSV_ROWS, CSV_TOTAL } = await import("./tasks.mjs");
const { HARD_TASKS } = await import("./tasks-hard.mjs");
const suite = String(args.suite || "core");
const TASKS = suite === "hard" ? HARD_TASKS : suite === "all" ? [...CORE, ...HARD_TASKS] : CORE;
const { snapshot } = await import("./checks.mjs");
const { createWorkbook } = await import("../lib/mock-excel.mjs");
const { runTurn } = await import(path.join(ROOT, "public/js/agent.js"));
const { workbookContext } = await import(path.join(ROOT, "public/js/tools.js"));
const messages = await import(path.join(ROOT, "public/js/messages.js"));

const pick = args.tasks ? String(args.tasks).split(",") : null;
const tasks = TASKS.filter((t) => !pick || pick.some((p) => t.id === p || t.category === p || t.id.startsWith(p)));
if (!tasks.length) {
  console.error("No tasks match", args.tasks);
  process.exit(1);
}

// ---------- fixtures ----------
function buildFixtures() {
  const dir = fs.mkdtempSync(path.join(os.tmpdir(), "oxa-bench-"));
  fs.writeFileSync(path.join(dir, "transactions.csv"), ["Date,Vendor,Category,Amount", ...CSV_ROWS.map((r) => r.join(","))].join("\n") + "\n");
  const inv = [
    "INVOICE #A-1042        Supplier: Andes Coffee Co.",
    "",
    "Item                     Qty     Unit price      Amount",
    "Arabica beans 1kg         12          18.50      222.00",
    "Paper cups (500)           4          23.75       95.00",
    "Oat milk 1L               30           2.40       72.00",
    "                                     Total       389.00",
  ].join("\n");
  fs.writeFileSync(path.join(dir, "invoice.txt"), inv);
  fs.writeFileSync(path.join(dir, "invoice.pdf"), execFileSync("/usr/sbin/cupsfilter", ["-o", "cpi=12", path.join(dir, "invoice.txt")], { stdio: ["ignore", "pipe", "ignore"] }));
  const memo = `<html><head><meta charset="utf-8"></head><body><h1>Q3 Budget Memo</h1><p>All amounts in Bs.</p>
<table border="1"><tr><th>Department</th><th>Budget</th><th>Actual</th></tr>
<tr><td>Marketing</td><td>120,000</td><td>131,500</td></tr><tr><td>Operations</td><td>340,000</td><td>322,750</td></tr>
<tr><td>IT</td><td>95,000</td><td>97,200</td></tr></table><p>Marketing overspent on the trade fair.</p></body></html>`;
  fs.writeFileSync(path.join(dir, "memo.html"), memo);
  execFileSync("/usr/bin/textutil", ["-inputencoding", "UTF-8", "-convert", "doc", path.join(dir, "memo.html"), "-output", path.join(dir, "budget-memo.doc")]);
  // Not in the public repo (a third-party deck); see fixtures/README.md. Tasks that need it are skipped.
  const deck = path.join(HERE, "fixtures/budget-deck.pptx");
  if (fs.existsSync(deck)) fs.copyFileSync(deck, path.join(dir, "budget-deck.pptx"));
  return dir;
}

async function upload(file) {
  const r = await fetch(`${globalThis.ADDIN_BASE}/docs/upload?name=${encodeURIComponent(path.basename(file))}`, { method: "POST", body: fs.readFileSync(file) });
  const j = await r.json();
  if (j.error) throw new Error(`upload ${path.basename(file)}: ${j.error}`);
  return j;
}

// ---------- run ----------
async function runTask(task, fixtures) {
  const { Excel, api } = createWorkbook(task.workbook);
  globalThis.Excel = Excel;
  const before = snapshot(api);
  const baselineErrors = api.errors();
  const docs = [];
  for (const a of task.attachments || []) docs.push(await upload(path.join(fixtures.dir, a)));
  const ctxLine = messages.contextLine(await workbookContext(), true);
  const history = [{ role: "user", content: messages.userContent({ ctxLine, text: task.prompt, docs }) }];
  const calls = [];
  const emptyTurns = [];
  let toolErrors = 0;
  const ui = {
    step: () => ({
      thinking() {}, content() {}, done() {},
      empty: (info) => emptyTurns.push(info),
      toolStart: (name, a) => {
        const rec = { name, args: JSON.stringify(a).slice(0, 300) };
        calls.push(rec);
        return { approve: async () => true, end: (res) => { if (res?.error) { toolErrors++; rec.error = String(res.error).slice(0, 200); } } };
      },
    }),
    stats: (s) => Object.assign(stats, s),
  };
  const stats = {};
  const ctl = new AbortController();
  const timer = setTimeout(() => ctl.abort(), TIMEOUT_MS);
  const t0 = Date.now();
  let answer = "", crashed = null;
  try {
    answer = await runTurn({ history, settings, signal: ctl.signal, ui, undo: [] });
  } catch (e) {
    crashed = e.name === "AbortError" ? `timed out after ${TIMEOUT_MS / 60000} min` : e.message;
  } finally {
    clearTimeout(timer);
  }
  const seconds = (Date.now() - t0) / 1000;
  let checks;
  try {
    checks = task.check(api, { answer: answer || "", before, baselineErrors, fixtures });
  } catch (e) {
    checks = [{ name: "checker crashed", pass: false, detail: e.stack?.split("\n").slice(0, 2).join(" ") }];
  }
  // Cells the agent changed or added ([ref, formula, value] per sheet), so a run can be inspected or re-scored later.
  const beforeCells = new Map(JSON.parse(before).flatMap(([s, cs]) => cs.map(([ref, v]) => [`${s}!${ref}`, v])));
  const changes = {};
  let changed = 0;
  for (const s of api.sheetNames()) {
    for (const c of api.cells(s)) {
      if (c.spilled || beforeCells.get(`${s}!${c.ref}`) === (c.formula ?? c.value) || ++changed > 5000) continue;
      (changes[s] ||= []).push([c.ref, c.formula, c.value]);
    }
  }
  const passed = checks.filter((c) => c.pass).length;
  const unverifiable = api.unverifiable();
  let status = crashed ? "error" : passed === checks.length ? "pass" : unverifiable.length ? "unverifiable" : "fail";
  return {
    id: task.id, category: task.category, level: task.level, status, score: passed / checks.length,
    checks, seconds: Math.round(seconds), steps: stats.steps || 0, tool_calls: calls.length, tool_errors: toolErrors,
    tokens_out: stats.evalCount || 0, error: crashed, unverifiable, answer: (answer || "").slice(0, 1500), calls, empty_turns: emptyTurns, changes,
  };
}

const outDir = path.join(ROOT, "bench-results");
fs.mkdirSync(outDir, { recursive: true });
const stamp = new Date().toISOString().replace(/[:T]/g, "-").slice(0, 16);
const safeModel = settings.model.replace(/[^\w.-]+/g, "_");
const outFile = path.join(outDir, `${stamp}_${safeModel}_think-${settings.think}_${suite}.json`);
const fixtures = { dir: buildFixtures(), csvTotal: CSV_TOTAL };
const run = { settings, started: new Date().toISOString(), repeat, results: [] };
for (let i = tasks.length - 1; i >= 0; i--) {
  const missing = (tasks[i].attachments || []).filter((a) => !fs.existsSync(path.join(fixtures.dir, a)));
  if (missing.length) {
    console.log(`Skipping ${tasks[i].id}: fixture ${missing.join(", ")} isn't present (see scripts/bench/fixtures/README.md).`);
    tasks.splice(i, 1);
  }
}

console.log(`Benchmark (${suite}): ${tasks.length} tasks × ${repeat} on ${settings.model} (think=${settings.think}, ctx=${settings.numCtx})\n`);
const icon = { pass: "✅", fail: "❌", unverifiable: "❔", error: "💥" };
for (let rep = 0; rep < repeat; rep++) {
  for (const task of tasks) {
    process.stdout.write(`${task.id.padEnd(30)} `);
    const r = await runTask(task, fixtures);
    r.rep = rep;
    run.results.push(r);
    fs.writeFileSync(outFile, JSON.stringify(run, null, 1)); // save as we go
    const failed = r.checks.filter((c) => !c.pass).map((c) => c.name + (c.detail ? ` [${c.detail}]` : ""));
    console.log(`${icon[r.status]} ${(r.score * 100).toFixed(0).padStart(3)}%  ${String(r.seconds).padStart(4)}s  ${String(r.steps).padStart(2)} steps${r.tool_errors ? `  ${r.tool_errors} tool err` : ""}${r.error ? `  ${r.error}` : ""}`);
    for (const f of failed) console.log(`${" ".repeat(32)}✗ ${f.slice(0, 160)}`);
  }
}

// ---------- summary ----------
const R = run.results;
const scored = R.filter((r) => r.status !== "unverifiable");
const pct = (n, d) => (d ? `${Math.round((n / d) * 100)}%` : "–");
const byCat = {};
for (const r of R) (byCat[r.category] ||= []).push(r);
run.finished = new Date().toISOString();
run.summary = {
  tasks: R.length,
  pass_rate: scored.filter((r) => r.status === "pass").length / (scored.length || 1),
  avg_check_score: R.reduce((a, r) => a + r.score, 0) / R.length,
  unverifiable: R.filter((r) => r.status === "unverifiable").length,
  errors: R.filter((r) => r.status === "error").length,
  avg_seconds: Math.round(R.reduce((a, r) => a + r.seconds, 0) / R.length),
  by_category: Object.fromEntries(Object.entries(byCat).map(([c, rs]) => [c, { pass: rs.filter((r) => r.status === "pass").length, total: rs.length }])),
  by_level: Object.fromEntries([1, 2, 3].map((l) => [l, { pass: R.filter((r) => r.level === l && r.status === "pass").length, total: R.filter((r) => r.level === l).length }])),
};
fs.writeFileSync(outFile, JSON.stringify(run, null, 1));

const md = [
  `# Benchmark (${suite}): ${settings.model} (think=${settings.think})`,
  "",
  `${run.started.slice(0, 16).replace("T", " ")} · ${R.length} task runs · avg ${run.summary.avg_seconds}s per task`,
  "",
  `**Pass rate: ${pct(scored.filter((r) => r.status === "pass").length, scored.length)}** (all checks passed) · average check score ${Math.round(run.summary.avg_check_score * 100)}%` +
    (run.summary.unverifiable ? ` · ${run.summary.unverifiable} unverifiable` : "") + (run.summary.errors ? ` · ${run.summary.errors} errors/timeouts` : ""),
  "",
  "| Category | Passed |", "|---|---|",
  ...Object.entries(run.summary.by_category).map(([c, v]) => `| ${c} | ${v.pass}/${v.total} |`),
  "",
  "| Difficulty | Passed |", "|---|---|",
  ...Object.entries(run.summary.by_level).filter(([, v]) => v.total).map(([l, v]) => `| ${"★".repeat(Number(l))} | ${v.pass}/${v.total} |`),
  "",
  "| Task | Result | Score | Time | Steps | Failed checks |", "|---|---|---|---|---|---|",
  ...R.map((r) => `| ${r.id} | ${icon[r.status]} ${r.status} | ${Math.round(r.score * 100)}% | ${r.seconds}s | ${r.steps} | ${r.checks.filter((c) => !c.pass).map((c) => c.name).join("; ").replace(/\|/g, "/") || "—"} |`),
  "",
].join("\n");
fs.writeFileSync(outFile.replace(/\.json$/, ".md"), md);
fs.writeFileSync(path.join(outDir, "latest.md"), md);
console.log(`\n${md.split("\n").slice(4, 6).join("\n")}\nSaved: ${path.relative(ROOT, outFile)} (+ .md)`);
