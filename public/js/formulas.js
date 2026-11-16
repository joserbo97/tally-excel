// Formula text helpers shared by the Excel tools and the benchmark's stand-in workbook.

// Shift the relative row numbers in an A1-style formula by dRow, as Excel does when a formula
// is filled down or a row moves. $-anchored rows, whole-column refs and "string literals" are left alone.
export function shiftFormulaRows(f, dRow) {
  if (!dRow || typeof f !== "string") return f;
  return f
    .split(/("(?:[^"]|"")*")/)
    .map((part, i) =>
      i % 2
        ? part
        : part.replace(/(^|[^A-Za-z0-9_.$])(\$?)([A-Za-z]{1,3})(\$?)(\d+)(?![\d(A-Za-z_])/g, (m, pre, cAbs, col, rAbs, row) =>
            rAbs ? m : `${pre}${cAbs}${col}${rAbs}${Number(row) + dRow}`
          )
    )
    .join("");
}
