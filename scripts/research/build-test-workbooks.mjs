#!/usr/bin/env node
// Builds the usability-test workbooks for docs/research/user-research-plan.md from the benchmark seeds,
// so sessions use exactly the data the benchmark scores. Also writes a moderator-only answer key.
//
//   node scripts/research/build-test-workbooks.mjs
import fs from "node:fs";
import path from "node:path";
import { fileURLToPath } from "node:url";
import XLSX from "xlsx";
import { TASKS, CSV_ROWS, CSV_TOTAL } from "../bench/tasks.mjs";
import { HARD_TASKS } from "../bench/tasks-hard.mjs";

const ROOT = path.resolve(path.dirname(fileURLToPath(import.meta.url)), "../..");
const OUT = path.join(ROOT, "docs/research/test-workbooks");
fs.mkdirSync(OUT, { recursive: true });

const ALL = [...TASKS, ...HARD_TASKS];
const seed = (id) => {
  const t = ALL.find((x) => x.id === id);
  if (!t) throw new Error(`no benchmark task ${id}`);
  return t.workbook;
};
const round2 = (x) => Math.round(x * 100) / 100;

// Same rules as the benchmark's stand-in workbook (scripts/lib/mock-excel.mjs parseInput):
// a leading ' keeps text, yyyy-mm-dd becomes a real date, numbers stay numbers.
function cell(v) {
  if (typeof v === "number") return { t: "n", v };
  const s = String(v);
  if (s[0] === "'") return { t: "s", v: s.slice(1) };
  const d = /^(\d{4})-(\d{2})-(\d{2})$/.exec(s);
  if (d) return { t: "n", v: (Date.UTC(+d[1], +d[2] - 1, +d[3]) - Date.UTC(1899, 11, 30)) / 86400000, z: "yyyy-mm-dd" };
  return { t: "s", v: s };
}

function sheetFrom(rows) {
  const ws = {};
  let maxC = 0;
  rows.forEach((row, r) => row.forEach((v, c) => {
    if (v === null || v === undefined || v === "") return;
    ws[XLSX.utils.encode_cell({ r, c })] = cell(v);
    maxC = Math.max(maxC, c);
  }));
  ws["!ref"] = rows.length ? XLSX.utils.encode_range({ s: { r: 0, c: 0 }, e: { r: rows.length - 1, c: maxC } }) : "A1";
  ws["!cols"] = Array.from({ length: maxC + 1 }, (_, c) => ({ wch: Math.max(10, ...rows.map((row) => String(row[c] ?? "").length + 2)) }));
  return ws;
}

function writeBook(file, spec) {
  const wb = XLSX.utils.book_new();
  const names = Object.keys(spec.sheets);
  for (const n of names) XLSX.utils.book_append_sheet(wb, sheetFrom(spec.sheets[n]), n);
  if (spec.active) wb.Workbook = { Views: [{ activeTab: names.indexOf(spec.active) }] };
  fs.writeFileSync(path.join(OUT, file), XLSX.write(wb, { type: "buffer", bookType: "xlsx", bookSST: true })); // the ESM build has no fs access of its own; bookSST writes text as normal shared strings
  return file;
}

// ---------- workbooks ----------
const sales = seed("analysis-top-region");
const tx = seed("formulas-sumifs-summary");
const bolivian = seed("hard-messy-bolivian-data");
const recon = seed("hard-reconciliation");

const files = [
  writeBook("01-top-region.xlsx", sales),
  writeBook("02-category-summary.xlsx", tx),
  writeBook("03-csv-import.xlsx", { sheets: { Sheet1: [] } }),
  writeBook("04-bank-export.xlsx", bolivian),
  writeBook("05-reconciliation.xlsx", recon),
  writeBook("06-undo.xlsx", sales),
];
fs.writeFileSync(path.join(OUT, "transactions.csv"), ["Date,Vendor,Category,Amount", ...CSV_ROWS.map((r) => r.join(","))].join("\n") + "\n");
files.push("transactions.csv");

// ---------- answer key (computed from the seeds, not copied from the checkers) ----------
const fmt = (n) => n.toLocaleString("en-US", { minimumFractionDigits: Number.isInteger(n) ? 0 : 2, maximumFractionDigits: 2 });

const regionTotals = sales.sheets.Sales.slice(1).map(([name, ...q]) => [name, q.reduce((a, b) => a + b, 0)]).sort((a, b) => b[1] - a[1]);

const cats = {};
for (const [, c, amt] of tx.sheets.Transactions.slice(1)) (cats[c] ||= { total: 0, n: 0 }), (cats[c].total += amt), cats[c].n++;

const messy = bolivian.sheets.Export.slice(1).map(([d, m]) => [d.slice(1), m.slice(1)]);
const parseBs = (s) => Number(s.replace(/^Bs\.?\s*/, "").replace(/\./g, "").replace(",", "."));
const messyTotal = round2(messy.reduce((a, [, m]) => a + parseBs(m), 0));

const ledger = recon.sheets.Ledger.slice(1), bank = recon.sheets.Bank.slice(1);
const count = (rows) => rows.reduce((m, [ref]) => m.set(ref, (m.get(ref) || 0) + 1), new Map());
const lc = count(ledger), bc = count(bank);
const amt = (rows, ref) => rows.find((r) => r[0] === ref)?.[2];
const disc = [];
for (const ref of new Set([...lc.keys(), ...bc.keys()])) {
  if (!bc.has(ref)) disc.push([ref, "in Ledger, missing from Bank", fmt(amt(ledger, ref)), "—"]);
  else if (!lc.has(ref)) disc.push([ref, "in Bank, missing from Ledger", "—", fmt(amt(bank, ref))]);
  else {
    if (bc.get(ref) > 1) disc.push([ref, `duplicated in Bank (${bc.get(ref)}×)`, fmt(amt(ledger, ref)), fmt(amt(bank, ref))]);
    if (amt(ledger, ref) !== amt(bank, ref)) disc.push([ref, "amount mismatch", fmt(amt(ledger, ref)), fmt(amt(bank, ref))]);
  }
}
disc.sort((a, b) => a[0].localeCompare(b[0]));

const table = (head, rows) => [`| ${head.join(" | ")} |`, `|${head.map(() => "---").join("|")}|`, ...rows.map((r) => `| ${r.join(" | ")} |`)].join("\n");

const key = `# Answer key: usability-test workbooks (moderator only)

Do not show this to participants. Generated by \`node scripts/research/build-test-workbooks.mjs\` from the benchmark seeds in \`scripts/bench/\`; re-run it if the seeds change.

Use it to fill the "Agent outcome" field in the session log, and to judge whether the participant's accept/reject call was right.

## 1 · 01-top-region.xlsx: which region sold most?

**East, 662,000.** The sheet must be left unchanged.

${table(["Region", "Full-year total"], regionTotals.map(([n, t]) => [n, fmt(t)]))}

## 2 · 02-category-summary.xlsx: summary by category

Correct only if the totals are **formulas** (SUMIFS/COUNTIFS or similar) on the Transactions sheet. To test that it's live, change any amount and see whether the summary moves.

${table(["Category", "Total", "Transactions"], Object.entries(cats).sort().map(([c, v]) => [c, fmt(round2(v.total)), v.n]))}

## 3 · 03-csv-import.xlsx + transactions.csv: import and total

Attach \`transactions.csv\` in the panel.
- **The total of the Amount column is ${fmt(CSV_TOTAL)}.** It should sit in the **Amount column**, directly under the 50 rows: column D, row 52 if the header is in row 1.
- **Known slip:** the total lands in E52/F52 while the reply says C52:D52. If that happens, note whether the participant notices.
- Vendors "+Plus Market" and "-Minus Supplies" must stay as text, not turn into formulas or errors.

## 4 · 04-bank-export.xlsx: messy bank export

Dates are day/month/year when written with slashes.

${table(["Original Fecha", "Should become", "Original Monto", "Should become"], messy.map(([d, m], i) => [d, `2026-03-${15 + i}`, m, fmt(parseBs(m))]))}

**Total: ${fmt(messyTotal)}**.

## 5 · 05-reconciliation.xlsx: Ledger vs Bank

${ledger.length} ledger rows and ${bank.length} bank rows, matched on Ref. **${disc.length} discrepancies:**

${table(["Ref", "Discrepancy", "Ledger", "Bank"], disc)}

Any other ref listed as a discrepancy is a false positive. In the benchmark, this task often takes over 8 minutes, so stop it at 8 minutes as the plan says.

## 6 · 06-undo.xlsx: approval prompt and undo

The same data as workbook 1. The participant asks the agent to **clear the Q4 numbers**, then changes their mind and restores them.
- With approval = destructive, the panel asks before **clearing cells** (\`clear_range\`) or running Office.js code. It does *not* ask before writes, sorts or formatting.
- If the agent blanks the cells with a write instead of a clear, no approval prompt appears. Note it, and still test Undo.

**Restored correctly means:** Q4 (column E) is back to ${sales.sheets.Sales.slice(1).map((r) => `${r[0]} ${fmt(r[4])}`).join(", ")}.
`;
fs.writeFileSync(path.join(OUT, "ANSWER-KEY.md"), key);
files.push("ANSWER-KEY.md");

console.log(`Wrote ${files.length} files to ${path.relative(ROOT, OUT)}/:\n  ${files.join("\n  ")}`);
