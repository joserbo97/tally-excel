// Matching logic for the compare_lists tool: reconcile two lists by a key column.
// Pure functions (no Excel), shared by the tool and its tests.

const norm = (v) => String(v ?? "").trim();

/** Finds a column by header name (case-insensitive) or by column letter. Returns a 0-based index into `headers`, or -1. */
export function findColumn(headers, want, firstColumnIndex = 0) {
  if (want == null || want === "") return -1;
  const w = norm(want).toLowerCase();
  const byName = headers.findIndex((h) => norm(h).toLowerCase() === w);
  if (byName > -1) return byName;
  if (/^[a-z]{1,3}$/.test(w)) {
    const n = [...w.toUpperCase()].reduce((a, ch) => a * 26 + ch.charCodeAt(0) - 64, 0) - 1;
    const i = n - firstColumnIndex;
    if (i >= 0 && i < headers.length) return i;
  }
  return -1;
}

function same(a, b, tolerance) {
  if (typeof a === "number" && typeof b === "number") return Math.abs(a - b) <= tolerance;
  const na = Number(a), nb = Number(b);
  if (norm(a) !== "" && norm(b) !== "" && Number.isFinite(na) && Number.isFinite(nb)) return Math.abs(na - nb) <= tolerance;
  return norm(a).toLowerCase() === norm(b).toLowerCase();
}

function index(rows, k) {
  const map = new Map();
  rows.forEach((row, i) => {
    const key = norm(row[k]);
    if (!key) return;
    if (!map.has(key)) map.set(key, []);
    map.get(key).push(i);
  });
  return map;
}

/**
 * Compares two lists (rows without their header row).
 *   lk / rk        key column index on each side
 *   fields         [{ name, l, r }] columns to compare (indexes on each side)
 * Returns issues in a stable order: missing from right, missing from left, mismatches, duplicates.
 * Each issue: { key, raw (the key as stored in the sheet), type, field?, left?, right?, count? }; left/right hold the first matching row's value.
 */
export function compareRows(leftRows, rightRows, { lk, rk, fields = [], tolerance = 0.005 }) {
  const L = index(leftRows, lk), R = index(rightRows, rk);
  const first = fields[0];
  const val = (rows, i, c) => (c == null || c < 0 ? "" : rows[i][c]);
  const rawL = (li) => leftRows[li[0]][lk], rawR = (ri) => rightRows[ri[0]][rk];
  const missingRight = [], missingLeft = [], mismatches = [], duplicates = [];
  for (const [key, li] of L) {
    if (!R.has(key)) { missingRight.push({ key, raw: rawL(li), type: "missing_right", field: first?.name, left: val(leftRows, li[0], first?.l), right: "" }); continue; }
    const ri = R.get(key);
    for (const f of fields) {
      const a = val(leftRows, li[0], f.l), b = val(rightRows, ri[0], f.r);
      if (!same(a, b, tolerance)) mismatches.push({ key, raw: rawL(li), type: "mismatch", field: f.name, left: a, right: b });
    }
  }
  for (const [key, ri] of R) if (!L.has(key)) missingLeft.push({ key, raw: rawR(ri), type: "missing_left", field: first?.name, left: "", right: val(rightRows, ri[0], first?.r) });
  for (const [key, li] of L) if (li.length > 1) duplicates.push({ key, raw: rawL(li), type: "duplicate_left", count: li.length, field: first?.name, left: val(leftRows, li[0], first?.l), right: "" });
  for (const [key, ri] of R) if (ri.length > 1) duplicates.push({ key, raw: rawR(ri), type: "duplicate_right", count: ri.length, field: first?.name, left: "", right: val(rightRows, ri[0], first?.r) });
  return {
    issues: [...missingRight, ...missingLeft, ...mismatches, ...duplicates],
    left_keys: L.size,
    right_keys: R.size,
    matched_keys: [...L.keys()].filter((k) => R.has(k)).length,
  };
}

export function issueText(issue, leftName, rightName) {
  switch (issue.type) {
    case "missing_right": return `Missing from ${rightName}`;
    case "missing_left": return `Missing from ${leftName}`;
    case "mismatch": return `${issue.field} differs`;
    case "duplicate_left": return `Duplicate in ${leftName} (${issue.count}×)`;
    case "duplicate_right": return `Duplicate in ${rightName} (${issue.count}×)`;
    default: return issue.type;
  }
}
