// Offline checks for spreadsheet handling. Generates sample files with SheetJS, then reads them back.
//   node scripts/test-sheets.mjs [outDir]
import fs from "node:fs";
import os from "node:os";
import path from "node:path";
import assert from "node:assert";
import XLSX from "xlsx";
import { detectDocKind, _internals } from "../web.js";

const { extractSheet, sheetResult, sheetValues } = _internals;
const out = process.argv[2] || fs.mkdtempSync(path.join(os.tmpdir(), "oxa-sheets-"));

// --- build a workbook: Sales (formulas, %, currency, dates) + Notes + Big (3000 rows)
const wb = XLSX.utils.book_new();
const sales = XLSX.utils.aoa_to_sheet([
  ["Region", "Q1", "Q4", "Growth", "Updated"],
  ["North", 125000, 167000, null, null],
  ["South", 98000, 138000, null, null],
  ["Total", null, null, null, null],
]);
for (const r of [2, 3]) {
  sales[`E${r}`] = { t: "n", v: 46037, z: "yyyy-mm-dd" }; // how Excel stores 2026-01-15: a serial number + date format
  sales[`D${r}`] = { t: "n", f: `C${r}/B${r}-1`, v: r === 2 ? 0.336 : 0.408, z: "0.0%" };
  sales[`B${r}`].z = "$#,##0";
  sales[`C${r}`].z = "$#,##0";
}
sales.B4 = { t: "n", f: "SUM(B2:B3)", v: 223000, z: "$#,##0" };
sales.C4 = { t: "n", f: "SUM(C2:C3)", v: 305000, z: "$#,##0" };
XLSX.utils.book_append_sheet(wb, sales, "Sales");
XLSX.utils.book_append_sheet(wb, XLSX.utils.aoa_to_sheet([["Note"], ["Figures in USD | preliminary"], ["Acción: revisar"]]), "Notes");
const big = [["ID", "Amount", "Category"]];
for (let i = 1; i <= 3000; i++) big.push([i, Math.round(i * 13.7 * 100) / 100, i % 3 ? "Ops" : "Sales"]);
XLSX.utils.book_append_sheet(wb, XLSX.utils.aoa_to_sheet(big), "Big");

const files = {
  "sample.xlsx": XLSX.write(wb, { type: "buffer", bookType: "xlsx" }),
  "sample.xls": XLSX.write(wb, { type: "buffer", bookType: "biff8" }),
  "sample.ods": XLSX.write(wb, { type: "buffer", bookType: "ods" }),
  "sample.csv": Buffer.from("Region,Q1,Q4\nNorth,125000,167000\nSouth,98000,138000\n\"Lima, Peru\",1,2\nPhone,+591 70012345,-\n"),
};
for (const [n, b] of Object.entries(files)) fs.writeFileSync(path.join(out, n), b);
console.log("sample files in", out);

// --- detection
assert.equal(detectDocKind(files["sample.xlsx"], "x.xlsx"), "xlsx");
assert.equal(detectDocKind(files["sample.xlsx"], "noext"), "xlsx");
assert.equal(detectDocKind(files["sample.xls"], "x.xls"), "xls");
assert.equal(detectDocKind(files["sample.ods"], "x.ods"), "ods");
assert.equal(detectDocKind(files["sample.csv"], "x.csv"), "csv");
console.log("✓ format detection");

for (const [n, b] of Object.entries(files)) {
  const kind = detectDocKind(b, n);
  const e = extractSheet(b, kind, n);
  const first = sheetResult(e, { maxChars: 1600, start: 0 });
  console.log(`\n=== ${n}: sheets ${JSON.stringify(first.sheets)}\n${first.content.slice(0, 700)}`);
  if (kind === "csv") {
    assert.match(e.text, /\| 4 \| Lima, Peru \| 1 \| 2 \|/, "quoted CSV field with comma");
    const cv = sheetValues(e, {});
    assert.deepEqual(cv.values[4], ["Phone", "'+591 70012345", "'-"], "text that Excel would parse as a formula is apostrophe-escaped");
    assert.equal(cv.values[1][1], 125000, "real numbers stay numbers");
    console.log(`✓ ${n}: quoted fields, +591/- text escaped for import, numbers kept`);
    continue;
  }
  assert.equal(e.sheets.length, 3);
  assert.match(e.text, /\| # \| A \| B \| C \| D \| E \|/, "column letters header");
  assert.match(e.text, /\| 2 \| North \|/, "row numbers");
  // SheetJS CE can't *write* formulas to .xls, so only check where the sample really has them.
  const hasFormulas = e.sheets[0].formulas > 0;
  if (hasFormulas) assert.match(e.text, /⟨=SUM\(B2:B3\)⟩/, "formulas shown");
  assert.match(e.text, /Figures in USD \/ preliminary/, "pipes in text escaped");
  assert.match(e.text, /Acción/);
  const jump = sheetResult(e, { maxChars: 500, start: 0, sheet: "big" });
  assert.match(jump.content, /^=== Sheet "Big" — A1:C3001 \(3001 rows × 3 cols\)/, "sheet jump");

  const v = sheetValues(e, { sheet: "Sales" });
  assert.deepEqual(v.values[0], ["Region", "Q1", "Q4", "Growth", "Updated"]);
  assert.equal(v.values[1][1], 125000);
  assert.equal(v.values[1][4], "2026-01-15", "dates as ISO");
  const vf = sheetValues(e, { sheet: "Sales", range: "B4:C4", includeFormulas: true });
  if (hasFormulas) assert.deepEqual(vf.values, [["=SUM(B2:B3)", "=SUM(C2:C3)"]]);
  else assert.deepEqual(vf.values, [[223000, 305000]], "cached values when no formulas");
  if (kind === "xlsx") {
    assert.equal(v.formats?.[1]?.[3], "0.0%", "number formats carried for import");
    assert.equal(v.formats?.[1]?.[1], "$#,##0");
  }
  const bigV = sheetValues(e, { sheet: "Big" });
  assert.equal(bigV.rows, 3001);
  assert.equal(bigV.values[3000][1], 41100);
  assert.throws(() => sheetValues(e, { sheet: "Nope" }), /No sheet named/);
  console.log(`✓ ${n}: text view, sheet jump, values, ${hasFormulas ? "formulas" : "no formulas in sample"}${kind === "xlsx" ? ", number formats" : ""}, 3001-row import`);
}
console.log("\nall spreadsheet checks passed");
