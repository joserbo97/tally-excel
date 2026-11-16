// Tests for write_range's fill_down_to_row against the stand-in workbook.
//   node scripts/test-fill-down.mjs
import { createWorkbook } from "./lib/mock-excel.mjs";
import { shiftFormulaRows } from "../public/js/formulas.js";

const { Excel, api } = createWorkbook({
  sheets: {
    Ledger: [["Ref", "Amt"], ["A", 1], ["B", 2], ["C", 3], ["D", 4]],
    Bank: [["Ref", "Amt"], ["A", 1], ["B", 5], ["D", 4], ["D", 4]],
  },
});
globalThis.Excel = Excel;
const { runTool, undoSnapshots } = await import("../public/js/tools.js");

let bad = 0;
const ok = (cond, name, detail = "") => { console.log(`${cond ? "✓" : "✗"} ${name}${cond ? "" : "  " + detail}`); if (!cond) bad++; };

// shiftFormulaRows itself
ok(shiftFormulaRows("=A2+$B$2+B$2+$C3+'My Sheet'!D2+Bank!A2:A9", 3) === "=A5+$B$2+B$2+$C6+'My Sheet'!D5+Bank!A5:A12", "relative rows shift; $-anchored rows stay");
ok(shiftFormulaRows('=IF(A2="A2","x",LOG10(A2))', 1) === '=IF(A3="A2","x",LOG10(A3))', "text in quotes and function names are untouched");
ok(shiftFormulaRows("=SUM(A:A)", 4) === "=SUM(A:A)", "whole-column refs are untouched");

// header + first formula row, filled to row 5
const undo = [];
const r = await runTool("write_range", {
  address: "C1", sheet: "Ledger",
  values: [["In bank", "Bank amt", "Note"], ["=COUNTIF(Bank!$A$2:$A$5,A2)", "=SUMIF(Bank!$A$2:$A$5,A2,Bank!$B$2:$B$5)", "check"]],
  fill_down_to_row: 5,
}, { undo });
ok(r.written?.endsWith("C1:E5") && r.rows === 5, "fills through the requested row", JSON.stringify(r).slice(0, 200));
ok(api.raw("Ledger!C5")?.f === "=COUNTIF(Bank!$A$2:$A$5,A5)", "row 5 formula points at A5 with the anchored range kept", JSON.stringify(api.raw("Ledger!C5")));
ok(JSON.stringify(api.values("Ledger!C2:C5")) === "[[1],[1],[0],[2]]", "filled formulas compute (counts 1,1,0,2)", JSON.stringify(api.values("Ledger!C2:C5")));
ok(JSON.stringify(api.values("Ledger!D2:D5")) === "[[1],[5],[0],[8]]", "second column fills too", JSON.stringify(api.values("Ledger!D2:D5")));
ok(api.value("Ledger!E5") === "check", "constants are copied");

// undo restores the whole filled area
await undoSnapshots(undo);
ok(api.value("Ledger!C5") === "" && api.value("Ledger!C1") === "", "undo clears every filled row", JSON.stringify(api.values("Ledger!C1:E5")));

// fill target above the template row → clear error, nothing written
const e = await runTool("write_range", { address: "F1", sheet: "Ledger", values: [["x"], ["=A2"]], fill_down_to_row: 1 }, { undo: [] });
ok(/must be at or below/.test(e.error || "") && api.value("Ledger!F1") === "", "a fill row above the last written row is rejected", JSON.stringify(e));

// fill_down_to_row equal to the template row: no extra rows
const same = await runTool("write_range", { address: "F1", sheet: "Ledger", values: [["x"], ["=A2"]], fill_down_to_row: 2 }, { undo: [] });
ok(same.rows === 2 && api.value("Ledger!F3") === "", "fill row equal to the last row adds nothing", JSON.stringify(same));

// numeric string from the model is accepted
const s = await runTool("write_range", { address: "G2", sheet: "Ledger", values: [["=B2*2"]], fill_down_to_row: "4" }, { undo: [] });
ok(s.rows === 3 && api.value("Ledger!G4") === 6, "fill_down_to_row given as a string works", JSON.stringify(s));

console.log(bad ? `\n${bad} failed` : "\nall passed");
process.exit(bad ? 1 : 0);
