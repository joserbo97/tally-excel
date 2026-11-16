#!/usr/bin/env node
// Checker self-test: applies a known-correct solution to every task through the agent's own tools,
// then runs the task's checks. Every task must pass here, or the checker (not the model) is wrong.
//   node scripts/bench/oracle.mjs        (attachment tasks need the add-in server running)
import fs from "node:fs";
import os from "node:os";
import path from "node:path";
import { spawnSync, execFileSync } from "node:child_process";
import { fileURLToPath } from "node:url";

const HERE = path.dirname(fileURLToPath(import.meta.url));
const ROOT = path.resolve(HERE, "../..");
const CA = path.join(os.homedir(), ".office-addin-dev-certs/ca.crt");
if (!process.env.NODE_EXTRA_CA_CERTS && fs.existsSync(CA)) {
  process.exit(spawnSync(process.execPath, process.argv.slice(1), { stdio: "inherit", env: { ...process.env, NODE_EXTRA_CA_CERTS: CA } }).status ?? 1);
}
globalThis.ADDIN_BASE = process.env.ADDIN_BASE || "https://localhost:3443";

const { TASKS, CSV_ROWS, CSV_TOTAL } = await import("./tasks.mjs");
const { HARD_TASKS } = await import("./tasks-hard.mjs");
const ALL = [...TASKS, ...HARD_TASKS];
const { snapshot } = await import("./checks.mjs");
const { createWorkbook } = await import("../lib/mock-excel.mjs");
const { runTool } = await import(path.join(ROOT, "public/js/tools.js"));

const T = (name, args) => runTool(name, args, { undo: [] }).then((r) => {
  if (r?.error) throw new Error(`${name}: ${r.error}`);
  return r;
});
const W = (address, values, sheet) => T("write_range", { address, values, sheet });

const dir = fs.mkdtempSync(path.join(os.tmpdir(), "oxa-oracle-"));
fs.writeFileSync(path.join(dir, "transactions.csv"), ["Date,Vendor,Category,Amount", ...CSV_ROWS.map((r) => r.join(","))].join("\n"));
const upload = async (file, name) => {
  const r = await fetch(`${globalThis.ADDIN_BASE}/docs/upload?name=${name}`, { method: "POST", body: fs.readFileSync(file) });
  return r.json();
};

// Correct solutions, written the way a careful analyst would.
const SOLVE = {
  "analysis-top-region": async () => "East had the highest full-year sales: $662,000.",
  "analysis-find-errors": async () => "D4 shows #DIV/0! because Cocoa's revenue is zero, so the margin divides by zero. E6 shows #N/A because Chai is not in the Prices sheet, so the VLOOKUP finds no match.",
  "analysis-fastest-growth": async () => "Mate grew fastest: +40% (9,000 → 12,600).",
  "formulas-totals": async () => {
    await W("F1", [["Total"], ...[2, 3, 4, 5, 6, 7].map((r) => [`=SUM(B${r}:E${r})`])], "Sales");
    await W("A8", [["Total", "=SUM(B2:B7)", "=SUM(C2:C7)", "=SUM(D2:D7)", "=SUM(E2:E7)", "=SUM(F2:F7)"]], "Sales");
  },
  "formulas-growth-percent": async () => {
    await W("D1", [["Growth %"], ...[2, 3, 4, 5, 6].map((r) => [`=C${r}/B${r}-1`])], "Products");
    await T("format_range", { address: "D2:D6", sheet: "Products", number_format: "0.0%" });
  },
  "formulas-lookup": async () => {
    await W("D1", [["Unit Price", "Line Total"], ...[2, 3, 4, 5, 6, 7, 8, 9].map((r) => [`=XLOOKUP(B${r},Products!$A$2:$A$5,Products!$C$2:$C$5)`, `=C${r}*D${r}`])], "Orders");
  },
  "formulas-sumifs-summary": async () => {
    await T("manage_sheet", { action: "add", name: "Summary" });
    const cats = ["Travel", "Meals", "Software", "Office"];
    await W("A1", [["Category", "Total", "Count"], ...cats.map((c, i) => [c, `=SUMIFS(Transactions!C:C,Transactions!B:B,A${i + 2})`, `=COUNTIFS(Transactions!B:B,A${i + 2})`])], "Summary");
  },
  "formulas-commission-tiers": async () => {
    await W("C1", [["Commission"], ...[2, 3, 4, 5, 6, 7, 8, 9].map((r) => [`=B${r}*IF(B${r}<50000,0.02,IF(B${r}<=100000,0.03,0.05))`])], "Reps");
  },
  "fix-div0": async () => {
    await W("D2", [2, 3, 4, 5].map((r) => [`=IF(B${r}=0,0,(B${r}-C${r})/B${r})`]), "Margins");
  },
  "fix-sum-range": async () => { await W("B12", [["=SUM(B2:B11)"]], "Expenses"); },
  "fix-hardcoded-totals": async () => {
    await W("E2", [2, 3, 4, 5].map((r) => [`=SUM(B${r}:D${r})`]), "Budget");
    return "Done. The IT total was wrong: it said 15,800 but Jan–Mar add up to 15,700. The others were correct.";
  },
  "clean-split-names": async (api) => {
    const rows = api.cells("Contacts").filter((c) => c.c === 0 && c.r > 0);
    await W("C1", [["First Name", "Last Name"], ...rows.map((c) => c.value.split(" "))], "Contacts");
  },
  "clean-text-numbers": async () => {
    await W("B2", [[1234.5], [99], [2000], [15.75], [300], [48.25]], "Payments");
    await W("A8", [["Total", "=SUM(B2:B7)"]], "Payments");
  },
  "clean-dedupe": async (api) => {
    const seen = new Set();
    const rows = [];
    for (let r = 1; ; r++) {
      const id = api.value(`Customers!A${r + 1}`);
      if (id === "" || id === undefined) break;
      if (!seen.has(id)) { seen.add(id); rows.push([id, api.value(`Customers!B${r + 1}`), api.value(`Customers!C${r + 1}`)]); }
    }
    await T("clear_range", { address: "A2:C40", sheet: "Customers" });
    await W("A2", rows, "Customers");
  },
  "clean-sort": async () => { await T("sort_range", { address: "A1:B9", sheet: "Reps", keys: [{ column: 1, ascending: false }] }); },
  "clean-standardize": async (api) => {
    const cats = [2, 3, 4, 5, 6, 7, 8, 9].map((r) => [String(api.value(`Expenses!B${r}`)).trim().toLowerCase().replace(/^./, (m) => m.toUpperCase())]);
    await W("B2", cats, "Expenses");
  },
  "model-loan": async () => {
    await W("A1", [["Loan amount", 250000], ["Annual rate", 0.065], ["Years", 30], ["Monthly payment", "=PMT(B2/12,B3*12,-B1)"], ["Total paid", "=B4*B3*12"], ["Total interest", "=B5-B1"]], "Loan");
  },
  "model-projection": async () => {
    await W("A1", [["Growth", 0.08], ["Cost ratio", 0.65], [], ["", "Year 1", "Year 2", "Year 3"], ["Revenue", 1200000, "=B5*(1+$B$1)", "=C5*(1+$B$1)"], ["Costs", "=B5*$B$2", "=C5*$B$2", "=D5*$B$2"], ["Profit", "=B5-B6", "=C5-C6", "=D5-D6"]].map((r) => [...r, ...Array(4 - r.length).fill("")]), "Plan");
  },
  "model-breakeven": async () => {
    await W("A1", [["Fixed costs", 50000], ["Price per unit", 25], ["Variable cost per unit", 15], ["Break-even units", "=B1/(B2-B3)"], ["Break-even revenue", "=B4*B2"]], "BreakEven");
  },
  "visual-table-chart": async () => {
    await T("create_table", { address: "A1:E7", sheet: "Sales" });
    await T("create_chart", { source: "A1:E7", sheet: "Sales", type: "column", title: "Quarterly sales by region" });
  },
  "visual-format": async () => {
    await T("format_range", { address: "A1:E1", sheet: "Sales", bold: true, fill_color: "#1F4E79", font_color: "#FFFFFF" });
    await T("format_range", { address: "B2:E7", sheet: "Sales", number_format: "#,##0" });
    await T("execute_office_js", { code: "context.workbook.worksheets.getItem('Sales').freezePanes.freezeRows(1); await context.sync();" });
  },
  "file-csv-import": async () => {
    const d = await upload(path.join(dir, "transactions.csv"), "transactions.csv");
    await T("import_from_file", { doc_id: d.doc_id, mode: "values", target_sheet: "Data", target_address: "A1" });
    await W("C52", [["Total", "=SUM(D2:D51)"]], "Data");
  },
  "file-pdf-invoice": async () => {
    await T("manage_sheet", { action: "add", name: "Invoice" });
    await W("A1", [["Item", "Qty", "Unit Price", "Amount"], ["Arabica beans 1kg", 12, 18.5, "=B2*C2"], ["Paper cups (500)", 4, 23.75, "=B3*C3"], ["Oat milk 1L", 30, 2.4, "=B4*C4"], ["Total", "", "", "=SUM(D2:D4)"]], "Invoice");
  },
  "file-pptx-chart": async () => {
    await T("manage_sheet", { action: "add", name: "Revenue" });
    await W("A1", [["Source", "Amount"], ["Taxes", 35742525], ["Licenses, Fees, & Fines", 6127000], ["Charges for Services", 2960699.46], ["Sale of Assets", 18000], ["Intergovernmental Revenues", 891000], ["Other Revenues", 1140400], ["Transfers In", 1223023], ["Total", "=SUM(B2:B8)"]], "Revenue");
  },
  "file-word-table": async () => {
    await T("manage_sheet", { action: "add", name: "Budget" });
    await W("A1", [["Department", "Budget", "Actual", "Variance"], ["Marketing", 120000, 131500, "=B2-C2"], ["Operations", 340000, 322750, "=B3-C3"], ["IT", 95000, 97200, "=B4-C4"]], "Budget");
  },
  "multisheet-quarter": async () => {
    await T("manage_sheet", { action: "add", name: "Q1" });
    await W("A1", [["Product", "Units", "Revenue"], ...[2, 3, 4, 5].map((r) => [`=Jan!A${r}`, `=Jan!B${r}+Feb!B${r}+Mar!B${r}`, `=Jan!C${r}+Feb!C${r}+Mar!C${r}`])], "Q1");
  },

  // ----- hard tier -----
  "hard-amortization-schedule": async () => {
    await W("A1", [["Loan amount", 180000], ["Annual rate", 0.072], ["Years", 15], ["Monthly payment", "=PMT(B2/12,B3*12,-B1)"], ["Total interest", "=B4*B3*12-B1"]], "Loan");
    const rows = [["Month", "Payment", "Interest", "Principal", "Balance"]];
    for (let m = 1; m <= 180; m++) {
      const r = 7 + m; // first data row is 8
      const open = m === 1 ? "$B$1" : `E${r - 1}`;
      rows.push([m, "=$B$4", `=${open}*$B$2/12`, `=B${r}-C${r}`, `=${open}-D${r}`]);
    }
    await W("A7", rows, "Loan");
  },
  "hard-month-category-grid": async () => {
    await T("manage_sheet", { action: "add", name: "Monthly" });
    const cats = ["Travel", "Meals", "Software", "Office"];
    const rows = [["Month", ...cats, "Total"]];
    ["Jan", "Feb", "Mar", "Apr", "May", "Jun"].forEach((m, i) => {
      const r = i + 2;
      rows.push([m, ...cats.map((_, j) => `=SUMIFS(Transactions!$C:$C,Transactions!$B:$B,${String.fromCharCode(66 + j)}$1,Transactions!$A:$A,">="&DATE(2026,${i + 1},1),Transactions!$A:$A,"<"&DATE(2026,${i + 2},1))`), `=SUM(B${r}:E${r})`]);
    });
    rows.push(["Total", ...["B", "C", "D", "E", "F"].map((c) => `=SUM(${c}2:${c}7)`)]);
    await W("A1", rows, "Monthly");
  },
  "hard-reconciliation": async (api) => {
    const read = (sh) => api.cells(sh).filter((c) => c.c === 0 && c.r > 0).map((c) => [c.value, api.value(`${sh}!C${c.r + 1}`)]);
    const L = read("Ledger"), B = read("Bank");
    const out = [["Ref", "Issue", "Ledger", "Bank"]];
    const lm = new Map(L), seen = new Map();
    for (const [ref] of B) seen.set(ref, (seen.get(ref) || 0) + 1);
    for (const [ref, amt] of L) if (!seen.has(ref)) out.push([ref, "Missing from Bank", amt, ""]);
    for (const [ref, amt] of B) if (!lm.has(ref)) out.push([ref, "Missing from Ledger", "", amt]);
    for (const [ref, amt] of L) { const b = B.find((x) => x[0] === ref); if (b && Math.abs(b[1] - amt) > 0.001) out.push([ref, "Amount mismatch", amt, b[1]]); }
    for (const [ref, n] of seen) if (n > 1) out.push([ref, `Duplicate in Bank (${n}x)`, "", ""]);
    await T("manage_sheet", { action: "add", name: "Reconciliation" });
    await W("A1", out, "Reconciliation");
    return `I found ${out.length - 1} discrepancies.`;
  },
  "hard-npv-irr": async () => {
    await W("E1", [["Discount rate", 0.1], ["NPV A", "=NPV(F1,B3:B7)+B2"], ["NPV B", "=NPV(F1,C3:C7)+C2"], ["IRR A", "=IRR(B2:B7)"], ["IRR B", "=IRR(C2:C7)"]], "Projects");
    return "Choose Project A: at 10% its NPV is higher than Project B's, even though B has the higher IRR.";
  },
  "hard-scenario-switch": async () => {
    await W("A1", [
      ["Scenario", "Worst", "Base", "Best"],
      ["Units", 8000, 10000, 13000],
      ["Price", 22, 25, 27],
      ["Variable cost", 15, 14, 13],
      ["Fixed costs", 60000, 60000, 60000],
      ["", "", "", ""],
      ["Selected (1-3)", 2, "", ""],
      ["Revenue", "=INDEX(B2:D2,B7)*INDEX(B3:D3,B7)", "", ""],
      ["Variable costs", "=INDEX(B2:D2,B7)*INDEX(B4:D4,B7)", "", ""],
      ["Fixed costs", "=INDEX(B5:D5,B7)", "", ""],
      ["Profit", "=B8-B9-B10", "", ""],
    ], "Model");
  },
  "hard-large-data-question": async (api) => {
    const t = {};
    for (const c of api.cells("Sales").filter((c) => c.c === 0 && c.r > 0)) {
      const r = c.r + 1;
      const dt = new Date(Date.UTC(1899, 11, 30) + api.value(`Sales!A${r}`) * 86400000);
      if (dt.getUTCMonth() === 2) { const p = api.value(`Sales!B${r}`); t[p] = (t[p] || 0) + api.value(`Sales!C${r}`) * api.value(`Sales!D${r}`); }
    }
    const top = Object.entries(t).sort((a, b) => b[1] - a[1]).slice(0, 3);
    return "Top 3 in March: " + top.map(([p, v], i) => `${i + 1}. ${p} — ${v.toFixed(2)}`).join("; ");
  },
  "hard-messy-bolivian-data": async () => {
    await W("A2", [["2026-03-15", 1234.5], ["2026-03-16", 980], ["2026-03-17", 12500.75], ["2026-03-18", 45.9], ["2026-03-19", 3000], ["2026-03-20", 0.99]], "Export");
    await W("A8", [["Total", "=SUM(B2:B7)"]], "Export");
  },
};

// Other valid ways to solve a task, which the checker must also accept.
const VARIANTS = [
  // Real dates as month labels (=DATE(...) shown as yyyy-mm), as qwen3.8 did on 2026-11-15.
  ["hard-month-category-grid", "date labels", async () => {
    await T("manage_sheet", { action: "add", name: "Monthly" });
    const cats = ["Software", "Meals", "Office", "Travel"];
    const rows = [["Month", ...cats, "Total"]];
    for (let i = 0; i < 6; i++) {
      const r = i + 2;
      rows.push([`=DATE(2026,${i + 1},1)`, ...cats.map((_, j) => `=SUMIFS(Transactions!$C$2:$C$201,Transactions!$A$2:$A$201,">="&$A${r},Transactions!$A$2:$A$201,"<"&EOMONTH($A${r},0)+1,Transactions!$B$2:$B$201,${String.fromCharCode(66 + j)}$1)`), `=SUM(B${r}:E${r})`]);
    }
    rows.push(["Total", ...["B", "C", "D", "E", "F"].map((c) => `=SUM(${c}2:${c}7)`)]);
    await W("A1", rows, "Monthly");
    await T("format_range", { address: "Monthly!A2:A7", number_format: "yyyy-mm" });
  }],
  // Formulas only (spilling FILTER/UNIQUE), the approach the system prompt asks for on large comparisons.
  ["hard-reconciliation", "spilling formulas", async () => {
    await T("manage_sheet", { action: "add", name: "Reconciliation" });
    const L = "Ledger!$A$2:$A$61", LA = "Ledger!$C$2:$C$61", B = "Bank!$A$2:$A$62", BA = "Bank!$C$2:$C$62";
    const mism = `(COUNTIF(${B},${L})>0)*(SUMIF(${B},${L},${BA})<>${LA}*COUNTIF(${B},${L}))`;
    await W("A1", [
      ["Missing from Bank", "Missing from Ledger", "Duplicate in Bank", "Duplicate in Ledger", "Amount mismatch", "Ledger amount", "Bank amount"],
      [
        `=IFERROR(FILTER(${L},COUNTIF(${B},${L})=0),"none")`,
        `=IFERROR(FILTER(${B},COUNTIF(${L},${B})=0),"none")`,
        `=IFERROR(UNIQUE(FILTER(${B},COUNTIF(${B},${B})>1)),"none")`,
        `=IFERROR(UNIQUE(FILTER(${L},COUNTIF(${L},${L})>1)),"none")`,
        `=IFERROR(FILTER(${L},${mism}),"none")`,
        `=IFERROR(FILTER(${LA},${mism}),"")`,
        `=IFERROR(FILTER(SUMIF(${B},${L},${BA})/COUNTIF(${B},${L}),${mism}),"")`,
      ],
    ], "Reconciliation");
    return "I found 4 discrepancies.";
  }],
  // Helper block (all refs + COUNTIF/XLOOKUP flags) on the Reconciliation sheet itself, next to the discrepancy list,
  // as qwen3.8 built it on 2026-11-16.
  ["hard-reconciliation", "helper columns on the same sheet", async () => {
    await T("manage_sheet", { action: "add", name: "Reconciliation" });
    await W("H1", [["Union of Refs", "Ledger Count", "Bank Count"], ["=UNIQUE(VSTACK(Ledger!A2:A61,Bank!A2:A62))"]], "Reconciliation");
    await T("write_range", { sheet: "Reconciliation", address: "I2", values: [["=COUNTIF(Ledger!$A$2:$A$61,H2)", "=COUNTIF(Bank!$A$2:$A$62,H2)"]], fill_down_to_row: 62 });
    await W("A1", [["Ref", "Issue"], ["T1012", "Amount mismatch: 3,716.03 vs 3,806.03"], ["T1028", "Missing from Bank"], ["T1061", "Missing from Ledger"], ["T1041", "Duplicate in Bank"]], "Reconciliation");
    return "I found 4 discrepancies.";
  }],
];

// Wrong answers the checker must reject (a checker that passes these is too lenient).
const MUST_FAIL = [
  ["hard-reconciliation", "lists every ref as a discrepancy", async (api) => {
    await T("manage_sheet", { action: "add", name: "Reconciliation" });
    const refs = api.cells("Ledger").filter((c) => c.c === 0 && c.r > 0).map((c) => [c.value, "check"]);
    await W("A1", [["Ref", "Issue"], ...refs.slice(0, 25), ["T1061", "Missing from Ledger"], ["T1041", "Duplicate"]], "Reconciliation");
    return "I found 4 discrepancies.";
  }],
  ["hard-reconciliation", "helper block only, no discrepancy list", async () => {
    await T("manage_sheet", { action: "add", name: "Reconciliation" });
    await W("H1", [["Union of Refs"], ["=UNIQUE(VSTACK(Ledger!A2:A61,Bank!A2:A62))"]], "Reconciliation");
    return "I found 4 discrepancies.";
  }],
];

const RUNS = [
  ...ALL.map((task) => [task, task.id, SOLVE[task.id]]),
  ...VARIANTS.map(([id, label, fn]) => [ALL.find((t) => t.id === id), `${id} (${label})`, fn]),
];

let ok = 0;
for (const [task, label, solve] of RUNS) {
  const { Excel, api } = createWorkbook(task.workbook);
  globalThis.Excel = Excel;
  const before = snapshot(api);
  let answer = "";
  try {
    if (task.id === "analysis-category-average") {
      const t = api.cells("Transactions").filter((c) => c.c === 1 && c.value === "Travel").map((c) => api.value(`Transactions!C${c.r + 1}`));
      answer = `The average Travel transaction is ${(t.reduce((a, b) => a + b, 0) / t.length).toFixed(2)}.`;
    } else {
      answer = (await solve(api)) || "";
    }
    const checks = task.check(api, { answer, before, baselineErrors: [], fixtures: { csvTotal: CSV_TOTAL } });
    const bad = checks.filter((c) => !c.pass);
    if (!bad.length) ok++;
    console.log(`${bad.length ? "✗" : "✓"} ${label}${bad.length ? "  ← " + bad.map((c) => `${c.name} [${c.detail}]`).join("; ") : ""}${api.unverifiable().length ? "  (unverifiable: " + api.unverifiable().join(", ") + ")" : ""}`);
  } catch (e) {
    console.log(`💥 ${label}: ${e.message}`);
  }
}
console.log(`\n${ok}/${RUNS.length} reference solutions pass their checks`);

let rejected = 0;
for (const [id, label, solve] of MUST_FAIL) {
  const task = ALL.find((t) => t.id === id);
  const { Excel, api } = createWorkbook(task.workbook);
  globalThis.Excel = Excel;
  const answer = (await solve(api)) || "";
  const failed = task.check(api, { answer, before: snapshot(api), baselineErrors: [], fixtures: { csvTotal: CSV_TOTAL } }).filter((c) => !c.pass);
  if (failed.length) rejected++;
  console.log(`${failed.length ? "✓ rejected" : "✗ ACCEPTED"} wrong answer: ${id} (${label})${failed.length ? "  ← " + failed.map((c) => c.name).join("; ") : ""}`);
}
console.log(`${rejected}/${MUST_FAIL.length} wrong answers rejected`);
process.exit(ok === RUNS.length && rejected === MUST_FAIL.length ? 0 : 1);
