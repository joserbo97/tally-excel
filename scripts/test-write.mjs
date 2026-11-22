// Tests for write_range's warning when it replaces existing cells.
//   node scripts/test-write.mjs
import { createWorkbook } from "./lib/mock-excel.mjs";

let bad = 0;
const ok = (cond, name, detail = "") => { console.log(`${cond ? "✓" : "✗"} ${name}${cond ? "" : "  " + detail}`); if (!cond) bad++; };

const { Excel, api } = createWorkbook({
  sheets: { Sales: [["Region", "Q1"], ["North", 120], ["South", 98], ["Andes", 60]] },
});
globalThis.Excel = Excel;
const { runTool, replacedCells } = await import("../public/js/tools.js");

// The slip from the benchmark: a Total row written over the last data row.
const over = await runTool("write_range", { sheet: "Sales", address: "A4", values: [["Total", "=SUM(B2:B3)"]] }, { undo: [] });
ok(over.replaced_cells?.count === 2 && /already had content/.test(over.warning || ""), "writing a total over a data row warns", JSON.stringify(over));
ok(JSON.stringify(over.replaced_cells?.cells) === '[{"cell":"A4","was":"Andes"},{"cell":"B4","was":60}]', "lists each replaced cell with what was there", JSON.stringify(over.replaced_cells));

// Writing into empty space doesn't warn.
const fresh = await runTool("write_range", { sheet: "Sales", address: "A6", values: [["Total", "=SUM(B2:B4)"]] }, { undo: [] });
ok(!fresh.warning && !fresh.replaced_cells, "writing into empty cells doesn't warn", JSON.stringify(fresh));

// Rewriting the same content doesn't warn (number 120 vs "120", same formula).
const same = await runTool("write_range", { sheet: "Sales", address: "A2", values: [["North", "120"]] }, { undo: [] });
ok(!same.warning, "rewriting identical content doesn't warn", JSON.stringify(same));
const sameF = await runTool("write_range", { sheet: "Sales", address: "B6", values: [["=SUM(B2:B4)"]] }, { undo: [] });
ok(!sameF.warning, "rewriting the same formula doesn't warn", JSON.stringify(sameF));

// Deliberate edits still go through; the warning is information, not a block.
const edit = await runTool("write_range", { sheet: "Sales", address: "B3", values: [[99]] }, { undo: [] });
ok(api.value("Sales!B3") === 99 && edit.replaced_cells?.count === 1, "a deliberate change is written and reported");

// The label-in-the-value-column slip from the benchmark: "Total" and its SUM stacked in one column.
{
  const { Excel: E2 } = createWorkbook({ sheets: { Payments: [["Payee", "Amount"], ["A", "10"], ["B", "20"]] } });
  globalThis.Excel = E2;
  const stackedRes = await runTool("write_range", { sheet: "Payments", address: "B2", values: [["10"], ["20"], ["Total"], ["=SUM(B2:B3)"]] }, { undo: [] });
  ok(/B4 has the label "Total".*B5.*put "Total" in A4 and the formula in B4, then clear B5/.test(stackedRes.layout_warning || ""), "a Total label stacked above its formula is flagged with the fix", stackedRes.layout_warning);
  const rowRes = await runTool("write_range", { sheet: "Payments", address: "A4", values: [["Total", "=SUM(B2:B3)"]] }, { undo: [] });
  ok(!rowRes.layout_warning, "a normal total row (label left of the formula) isn't flagged");
  const headerRes = await runTool("write_range", { sheet: "Payments", address: "D1", values: [["Total"], ["=B2*2"], ["=B3*2"]] }, { undo: [] });
  ok(!headerRes.layout_warning, "a 'Total' column header above per-row formulas isn't flagged", headerRes.layout_warning);
  const oneRow = await runTool("write_range", { sheet: "Payments", address: "E1", values: [["Total"], ["=SUM(B2:C2)"]] }, { undo: [] });
  ok(!oneRow.layout_warning, "…even when there's only one row under the header", oneRow.layout_warning);
  globalThis.Excel = Excel;
}

// Pure helper: column letters and the list limit.
const many = replacedCells([Array(30).fill("x")], [Array(30).fill("y")], 0, 25);
ok(many.count === 30 && many.cells.length === 12 && many.cells[0].cell === "Z1" && many.cells[1].cell === "AA1", "counts all, lists the first 12, letters past Z", JSON.stringify(many.cells.slice(0, 2)));

console.log(bad ? `\n${bad} failed` : "\nall passed");
process.exit(bad ? 1 : 0);
