// Checks the panel's plain-language approval text and undo counts.
//   node scripts/test-describe.mjs
import { cellCount, describeApproval, undoCellCount, undoLabel } from "../public/js/describe.js";

let bad = 0;
const ok = (cond, name, got) => { console.log(`${cond ? "✓" : "✗"} ${name}${cond ? "" : `  (got ${JSON.stringify(got)})`}`); if (!cond) bad++; };

ok(cellCount("E1:F20") === 40, "E1:F20 is 40 cells", cellCount("E1:F20"));
ok(cellCount("Summary!$A$1:$C$5") === 15, "sheet name and $ anchors are ignored", cellCount("Summary!$A$1:$C$5"));
ok(cellCount("B2") === 1, "a single cell is 1");
ok(cellCount("AA10:AB11") === 4, "two-letter columns count correctly", cellCount("AA10:AB11"));
ok(cellCount("C:C") === null && cellCount("5:7") === null, "whole columns/rows have no count");

let d = describeApproval("clear_range", { sheet: "Summary", address: "E1:F20" });
ok(d.text === "Clear 40 cells in Summary!E1:F20." && d.undo === "You can undo this." && d.yes === "Clear cells", "clear names the cells, says it can be undone, and labels the button", d);
d = describeApproval("clear_range", { address: "A1:A3", what: "all" });
ok(/including formatting/.test(d.text) && /not the formatting/.test(d.undo), "clearing everything says formatting won't come back", d);
d = describeApproval("clear_range", { address: "A1", what: "formats" });
ok(d.yes === "Clear formatting" && /can't restore formatting/.test(d.undo), "clearing formats only", d);
d = describeApproval("execute_office_js", { explanation: "Adds conditional formatting to B2:B40." });
ok(d.text === "Run code on the workbook: Adds conditional formatting to B2:B40." && d.yes === "Run code" && /can't reverse/.test(d.undo), "running code uses the model's explanation and warns Undo won't help", d);
d = describeApproval("manage_sheet", { action: "delete", name: "Old" });
ok(d.text === 'Delete the sheet "Old" and everything on it.' && d.yes === "Delete sheet", "deleting a sheet", d);
d = describeApproval("write_range", { sheet: "Data", address: "D52" });
ok(d.text === "Write to 1 cell in Data!D52.", "writes (approval = any change)", d);
d = describeApproval("some_new_tool", {});
ok(d.yes === "Allow" && /some_new_tool/.test(d.text), "unknown tools fall back to a generic prompt", d);

const snaps = [{ formulas: [["a", "b"], ["c", "d"]] }, { formulas: [[1]] }];
ok(undoCellCount(snaps) === 5, "undo counts every cell in every snapshot", undoCellCount(snaps));
ok(undoLabel(snaps) === "Undo 5 cell changes" && undoLabel([{ formulas: [[1]] }]) === "Undo 1 cell change", "undo label with count and singular", undoLabel(snaps));
ok(undoLabel([{ formulas: Array.from({ length: 1200 }, () => [1]) }]) === "Undo 1,200 cell changes", "large counts get a thousands separator");

console.log(bad ? `\n${bad} failed` : "\nall passed");
process.exit(bad ? 1 : 0);
