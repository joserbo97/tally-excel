// Tests for the compare_lists tool against the stand-in workbook.
//   node scripts/test-compare.mjs
import { createWorkbook } from "./lib/mock-excel.mjs";
import { compareRows, findColumn } from "../public/js/compare.js";

let bad = 0;
const ok = (cond, name, detail = "") => { console.log(`${cond ? "✓" : "✗"} ${name}${cond ? "" : "  " + detail}`); if (!cond) bad++; };

// Pure matching
ok(findColumn(["Ref", "Date", "Amount"], "amount") === 2, "finds a column by header, ignoring case");
ok(findColumn(["Ref", "Date", "Amount"], "C") === 2 && findColumn(["X", "Y"], "C", 1) === 1, "finds a column by letter, relative to the range");
ok(findColumn(["Ref"], "Nope") === -1, "unknown column → -1");
const pure = compareRows(
  [["A", 1], ["B", 2], ["C", 3], ["C", 3], [123, 5]],
  [["A", 1.001], ["B", 2.5], ["D", 4], ["D", 4], ["123", 5]],
  { lk: 0, rk: 0, fields: [{ name: "Amt", l: 1, r: 1 }] },
);
const types = pure.issues.map((x) => `${x.key}:${x.type}`).join(",");
ok(types === "C:missing_right,D:missing_left,B:mismatch,C:duplicate_left,D:duplicate_right", "missing, mismatch and duplicates found in a stable order", types);
ok(!pure.issues.some((x) => x.key === "A"), "differences within the tolerance are not reported");
ok(!pure.issues.some((x) => x.key === "123"), "number key 123 matches text key \"123\"");

// Tool on a workbook
const { Excel, api } = createWorkbook({
  active: "Ledger",
  sheets: {
    Ledger: [["Ref", "Date", "Amount"], ["T1", "2026-04-01", 100], ["T2", "2026-04-02", 200], ["T3", "2026-04-03", 300], ["T4", "2026-04-04", 400]],
    Bank: [["Ref", "Date", "Amount"], ["T1", "2026-04-01", 100], ["T2", "2026-04-02", 250], ["T4", "2026-04-04", 400], ["T4", "2026-04-04", 400], ["T9", "2026-04-09", 90]],
    Reconciliation: [["already here"]],
  },
});
globalThis.Excel = Excel;
const { runTool, undoSnapshots } = await import("../public/js/tools.js");

const undo = [];
const r = await runTool("compare_lists", { left: "Ledger!A:C", right: "Bank!A:C", key: "Ref", compare: ["Amount"] }, { undo });
ok(!r.error, "runs without error", JSON.stringify(r).slice(0, 300));
ok(r.report_sheet === "Reconciliation 2" && api.value("Reconciliation!A1") === "already here", "doesn't overwrite an existing sheet; picks a free name", r.report_sheet);
ok(r.discrepancies === 4, "4 discrepancies: T3 missing from Bank, T9 missing from Ledger, T2 amount, T4 duplicate", JSON.stringify(r.items));
const S = "Reconciliation 2";
const col = (c) => [2, 3, 4, 5].map((n) => api.value(`'${S}'!${c}${n}`));
ok(JSON.stringify(col("A")) === '["T3","T9","T2","T4"]', "report lists the refs", JSON.stringify(col("A")));
ok(JSON.stringify(col("B")) === '["Missing from Bank","Missing from Ledger","Amount differs","Duplicate in Bank (2×)"]', "plain-language issues", JSON.stringify(col("B")));
ok(api.value(`'${S}'!D4`) === 200 && api.value(`'${S}'!E4`) === 250 && api.value(`'${S}'!F4`) === 50, "mismatch row shows both amounts and the difference", JSON.stringify(api.values(`'${S}'!D4:F4`)));
ok(api.value(`'${S}'!D2`) === 300 && api.value(`'${S}'!E2`) === "", "missing row shows the amount from the side that has it");
ok(String(api.raw(`'${S}'!E4`)?.f || "").startsWith("=XLOOKUP("), "amounts are live formulas", JSON.stringify(api.raw(`'${S}'!E4`)));

// Live: fix the bank amount, the difference follows
await runTool("write_range", { sheet: "Bank", address: "C3", values: [[200]] }, { undo: [] });
ok(api.value(`'${S}'!F4`) === 0, "difference updates when the source changes", String(api.value(`'${S}'!F4`)));

// Default compare columns (all shared headers) and labels
const r2 = await runTool("compare_lists", { left: "Ledger!A1:C5", right: "Bank!A1:C6", output_sheet: "Check", left_label: "Books" }, { undo: [] });
ok(r2.report_sheet === "Check" && r2.compared.fields.join() === "Date,Amount", "compares every shared column by default", JSON.stringify(r2.compared));
ok(api.value("Check!B2") === "Missing from Bank" && api.value("Check!D1") === "Books value", "labels name the lists", JSON.stringify(api.values("Check!A1:F2")));

// A range that stops before the end of the list is extended (Bank has 6 rows; the guess covers 5)
const r3 = await runTool("compare_lists", { left: "Ledger!A1:C5", right: "Bank!A1:C5", compare: ["Amount"], output_sheet: "Short" }, { undo: [] });
ok(r3.items?.some((x) => x[0] === "T9") && /A1:C6/.test(String(r3.ranges_extended)), "a short range is extended to the end of the list, and the result says so", JSON.stringify(r3).slice(0, 400));

// Errors
const e1 = await runTool("compare_lists", { left: "Ledger!A:C", right: "Bank!A:C", key: "Invoice" }, { undo: [] });
ok(/Key column "Invoice"/.test(e1.error || ""), "unknown key column → clear error listing the headers", JSON.stringify(e1));
const e2 = await runTool("compare_lists", { left: "Ledger!A:C", right: "Bank!A:C", compare: ["Memo"] }, { undo: [] });
ok(/"Memo" isn't in both/.test(e2.error || ""), "unknown compare column → clear error", JSON.stringify(e2));

// Undo clears the report
await undoSnapshots(undo);
ok(api.value(`'${S}'!A2`) === "" && api.value(`'${S}'!B1`) === "", "undo clears the report", JSON.stringify(api.values(`'${S}'!A1:B3`)));

console.log(bad ? `\n${bad} failed` : "\nall passed");
process.exit(bad ? 1 : 0);
