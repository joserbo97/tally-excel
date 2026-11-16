// Scoring helpers for benchmark tasks. Checks look for results by meaning (a "Total" label, a
// "Growth" header) rather than fixed addresses, so reasonable layouts all score the same.

export const near = (a, b, tol = 0.005) =>
  typeof a === "number" && typeof b === "number" && Math.abs(a - b) <= Math.max(Math.abs(b) * tol, 1e-6);

export const check = (name, pass, detail = "") => ({ name, pass: Boolean(pass), detail: String(detail ?? "") });

// All numbers mentioned in a text ("$1,234.50", "40.8%", "(2,200)").
export function numbersIn(text) {
  return [...String(text).matchAll(/\(?-?\$?\s?(\d[\d,]*(?:\.\d+)?)\s?(%?)\)?/g)].map((m) => {
    const n = Number(m[1].replace(/,/g, ""));
    const neg = m[0].startsWith("(") && m[0].endsWith(")") || m[0].includes("-");
    return { n: neg ? -n : n, pct: m[2] === "%" };
  });
}
export const mentionsNumber = (text, x, tol = 0.005) =>
  numbersIn(text).some(({ n, pct }) => near(n, x, tol) || (pct && near(n / 100, x, tol)) || near(Math.abs(n), Math.abs(x), tol));
export const mentions = (text, re) => (re instanceof RegExp ? re : new RegExp(re, "i")).test(String(text));

// Cells (any sheet, or one sheet) whose text matches.
export function findLabel(api, re, sheet) {
  const names = sheet ? [sheet] : api.sheetNames();
  const out = [];
  for (const s of names) for (const c of api.cells(s)) if (typeof c.value === "string" && re.test(c.value.trim())) out.push({ ...c, sheet: s });
  return out;
}

// Numbers in the same row to the right of a label cell.
export function rowNumbersRight(api, cell, max = 20) {
  const row = api.cells(cell.sheet).filter((c) => c.r === cell.r && c.c > cell.c && c.c <= cell.c + max);
  return row.filter((c) => typeof c.value === "number");
}

// Column under a header matching re (searches the first `scan` rows). Returns { sheet, c, headerRow }.
export function findColumn(api, re, sheet, scan = 6) {
  const names = sheet ? [sheet] : api.sheetNames();
  for (const s of names) {
    for (const c of api.cells(s)) {
      if (c.r < scan && typeof c.value === "string" && re.test(c.value.trim())) return { sheet: s, c: c.c, headerRow: c.r };
    }
  }
  return null;
}
export const columnCells = (api, col, fromRow, toRow) =>
  api.cells(col.sheet).filter((x) => x.c === col.c && x.r >= fromRow && x.r <= toRow);

// Value in a column (col) on the row whose column-A-ish label matches labelRe.
export function valueAt(api, col, labelRe) {
  const lab = findLabel(api, labelRe, col.sheet).find((l) => l.c < col.c);
  if (!lab) return undefined;
  return api.cells(col.sheet).find((x) => x.r === lab.r && x.c === col.c);
}

// Formula liveness: change an input, verify the output follows, then put the input back.
export function dynamicCheck(name, api, inputRef, newValue, outputFn, expectFn) {
  const before = api.value(inputRef);
  api.setValue(inputRef, newValue);
  let got, want;
  try {
    got = outputFn();
    want = expectFn();
  } finally {
    api.setValue(inputRef, before);
  }
  return check(name, typeof got === "number" && near(got, want, 0.002), `after setting ${inputRef}=${newValue}: got ${got}, expected ${want}`);
}

export const noNewErrors = (api, baseline = []) => {
  const errs = api.errors().filter((e) => !baseline.includes(e));
  return check("no #ERROR values in the workbook", !errs.length, errs.slice(0, 5).join("; "));
};

// Snapshot of every sheet's cell contents; scratch work that's cleaned up afterwards still counts as unchanged.
export const snapshot = (api) => JSON.stringify(api.sheetNames().map((s) => [s, api.cells(s).map((c) => [c.ref, c.formula ?? c.value])]));
export const unchanged = (api, snap) => check("workbook left unchanged (question only)", snapshot(api) === snap, "cell contents differ from the original");
