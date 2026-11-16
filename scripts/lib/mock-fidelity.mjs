// Fidelity test for the mock workbook: common Excel formulas must evaluate to what Excel gives.
// HyperFormula (the engine) differs from Excel in a few places; mock-excel.mjs rewrites those.
// Any failure here means the benchmark could score correct Excel work as wrong.
//   node scripts/lib/mock-fidelity.mjs
import { createWorkbook } from "./mock-excel.mjs";

const { api } = createWorkbook({
  sheets: {
    Data: [
      ["Product", "Price", "Qty", "Date", "Cat"],
      ["Coffee", 12.5, 4, "2026-01-31", "Food"],
      ["Tea", 8, 10, "2026-02-01", "Food"],
      ["Mate", 6.75, 3, "2026-02-28", "Drink"],
    ],
    "My Sheet": [["x", 5]],
    Out: [],
  },
});

// [formula, expected Excel result]
const CASES = [
  ['=VLOOKUP("Tea",Data!A:C,2,FALSE)', 8],
  ['=VLOOKUP("Tea",Data!A1:C4,3,FALSE)', 10],
  ['=IFERROR(VLOOKUP("Chai",Data!A:C,2,FALSE),"missing")', "missing"],
  ["=IF(TRUE,1,2)", 1],
  ["=IF(Data!B2>10,TRUE,FALSE)", true],
  ["=AND(Data!C3>5,TRUE)", true],
  ['=COUNTIF(Data!E:E,"Food")', 2],
  ["=INDEX(Data!B1:D1,2)", "Qty"],
  ["=INDEX(Data!A2:A4,3)", "Mate"],
  ["=INDEX(Data!A1:E4,3,2)", 8],
  ['=INDEX(Data!A2:A4,MATCH(MAX(Data!B2:B4),Data!B2:B4,0))', "Coffee"],
  ['=XLOOKUP("Mate",Data!A2:A4,Data!B2:B4)', 6.75],
  ['=SUMIFS(Data!C:C,Data!D:D,">="&DATE(2026,2,1),Data!D:D,"<"&DATE(2026,3,1))', 13],
  ['=SUMPRODUCT(Data!B2:B4,Data!C2:C4)', 150.25],
  ["='My Sheet'!B1*2", 10],
  ["=ROUND(PMT(0.065/12,360,-250000),2)", 1580.17],
  ['=TEXT(1234.5,"#,##0.00")', "1,234.50"],
  ['=LEFT("Ana Quispe",FIND(" ","Ana Quispe")-1)', "Ana"],
  ['=TRIM("  hi  ")', "hi"],
  ['="TRUE story"&" "&IF(FALSE,"a","b")', "TRUE story b"],
  ["=NPV(0.1,100,100)+(-100)", 73.55371900826447],
  ["=MAX(Data!B2:B4)-MIN(Data!B2:B4)", 5.75],
];

let bad = 0;
CASES.forEach(([f, want], i) => {
  api.setValue(`Out!A${i + 1}`, 0);
  api.sheet("Out").cells.set(`${i},0`, { f });
  const got = api.value(`Out!A${i + 1}`);
  const ok = typeof want === "number" ? typeof got === "number" && Math.abs(got - want) < 1e-6 : got === want;
  if (!ok) bad++;
  console.log(`${ok ? "✓" : "✗"} ${f.padEnd(70)} → ${JSON.stringify(got)}${ok ? "" : `   (Excel: ${JSON.stringify(want)})`}`);
});
console.log(`\n${CASES.length - bad}/${CASES.length} formulas match Excel`);
process.exit(bad ? 1 : 0);
