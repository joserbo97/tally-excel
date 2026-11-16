// Headless end-to-end test of the agent loop against the real local Ollama
// model, with a small in-memory mock of the Office.js Excel API.
//   node scripts/test-agent.mjs            (needs Ollama running)
globalThis.OLLAMA_API_BASE = process.env.OLLAMA_API_BASE || "http://127.0.0.1:11434/api";
// Web tools go through the add-in server; run with NODE_EXTRA_CA_CERTS=~/.office-addin-dev-certs/ca.crt
globalThis.ADDIN_BASE = process.env.ADDIN_BASE || "https://localhost:3443";

// ---------------- mock Excel ----------------
const colToNum = (s) => [...s.toUpperCase()].reduce((n, ch) => n * 26 + ch.charCodeAt(0) - 64, 0) - 1;
const numToCol = (n) => { let s = ""; n++; while (n > 0) { const m = (n - 1) % 26; s = String.fromCharCode(65 + m) + s; n = Math.floor((n - 1) / 26); } return s; };
function parseA1(a) {
  const [p1, p2 = p1] = a.split(":");
  const one = (p) => { const m = /^([A-Z]*)(\d*)$/i.exec(p); return { c: m[1] ? colToNum(m[1]) : null, r: m[2] ? Number(m[2]) - 1 : null }; };
  const A = one(p1), B = one(p2);
  return { r0: A.r ?? 0, c0: A.c ?? 0, r1: B.r ?? 1048575, c1: B.c ?? 16383 };
}

const log = [];
class Sheet {
  constructor(name) { this.name = name; this.cells = new Map(); this.tables = { add: (addr, h) => { log.push(`table ${this.name}!${addr}`); return { name: "Table1", style: "", load() {}, getRange: () => this.getRange(addr) }; } };
    this.charts = { add: (type, range) => { log.push(`chart ${type} ${range.address}`); return { name: "Chart 1", title: {}, setPosition() {}, load() {} }; } }; this.visibility = "Visible"; }
  getRange(a) { const p = parseA1(a); return new Range(this, p.r0, p.c0, p.r1 - p.r0 + 1, p.c1 - p.c0 + 1); }
  getUsedRange() { return this.getUsedRangeOrNullObject(); }
  getUsedRangeOrNullObject() {
    if (!this.cells.size) return Object.assign(new Range(this, 0, 0, 1, 1), { isNullObject: true });
    let r0 = 1e9, c0 = 1e9, r1 = 0, c1 = 0;
    for (const k of this.cells.keys()) { const [r, c] = k.split(",").map(Number); r0 = Math.min(r0, r); c0 = Math.min(c0, c); r1 = Math.max(r1, r); c1 = Math.max(c1, c); }
    return new Range(this, r0, c0, r1 - r0 + 1, c1 - c0 + 1);
  }
  activate() {} load() {}
  findAllOrNullObject(q) {
    const hits = [...this.cells.entries()].filter(([, v]) => String(v).toLowerCase().includes(q.toLowerCase()));
    if (!hits.length) return { isNullObject: true, load() {}, areas: { load() {} } };
    return { isNullObject: false, load() {}, areas: { load() {}, items: hits.map(([k]) => { const [r, c] = k.split(",").map(Number); return new Range(this, r, c, 1, 1); }) } };
  }
}
// Formulas are evaluated by HyperFormula (a real spreadsheet engine) over all mock sheets.
const { HyperFormula } = await import("hyperformula");
function evalCell(sheet, v, r, c) {
  if (typeof v !== "string" || !v.startsWith("=")) return v;
  const sheetsData = {};
  for (const s of wb.worksheets.items) {
    const grid = [];
    for (const [k, val] of s.cells) { const [rr, cc] = k.split(",").map(Number); (grid[rr] ||= [])[cc] = val; }
    sheetsData[s.name] = Array.from(grid, (row) => Array.from(row || [], (x) => (x === undefined ? null : x)));
  }
  const hf = HyperFormula.buildFromSheets(sheetsData, { licenseKey: "gpl-v3" });
  const out = hf.getCellValue({ sheet: hf.getSheetId(sheet.name), row: r, col: c });
  hf.destroy();
  return out && typeof out === "object" ? out.value || `#${out.type}` : out ?? "";
}
class Range {
  constructor(ws, r, c, rows, cols) { Object.assign(this, { ws, rowIndex: r, columnIndex: c, rowCount: rows, columnCount: cols, isNullObject: false }); this.format = mkFormat(this); this.sort = { apply: (f) => log.push(`sort ${this.address} ${JSON.stringify(f)}`) }; }
  get address() { const a = `${numToCol(this.columnIndex)}${this.rowIndex + 1}`, b = `${numToCol(this.columnIndex + this.columnCount - 1)}${this.rowIndex + this.rowCount}`; return `${this.ws.name}!${a === b ? a : a + ":" + b}`; }
  grid(fn) { return Array.from({ length: Math.min(this.rowCount, 2000) }, (_, i) => Array.from({ length: Math.min(this.columnCount, 200) }, (_, j) => fn(this.rowIndex + i, this.columnIndex + j))); }
  get formulas() { return this.grid((r, c) => this.ws.cells.get(`${r},${c}`) ?? ""); }
  set formulas(v) { v.forEach((row, i) => row.forEach((x, j) => { const k = `${this.rowIndex + i},${this.columnIndex + j}`; x === "" ? this.ws.cells.delete(k) : this.ws.cells.set(k, x); })); log.push(`write ${this.address}`); }
  get values() { return this.grid((r, c) => evalCell(this.ws, this.ws.cells.get(`${r},${c}`) ?? "", r, c)); }
  get text() { return this.values.map((r) => r.map(String)); }
  set numberFormat(v) { log.push(`numfmt ${this.address} ${v[0][0]}`); }
  load() {} select() { log.push(`select ${this.address}`); } merge() {} unmerge() {}
  clear(w) { log.push(`clear ${this.address} ${w}`); for (const row of this.grid((r, c) => `${r},${c}`)) row.forEach((k) => this.ws.cells.delete(k)); }
  insert(d) { log.push(`insert ${this.address} ${d}`); } delete(d) { log.push(`delete ${this.address} ${d}`); }
  getCell(r, c) { return new Range(this.ws, this.rowIndex + r, this.columnIndex + c, 1, 1); }
  getResizedRange(dr, dc) { return new Range(this.ws, this.rowIndex, this.columnIndex, this.rowCount + dr, this.columnCount + dc); }
  getIntersectionOrNullObject(o) {
    const r0 = Math.max(this.rowIndex, o.rowIndex), c0 = Math.max(this.columnIndex, o.columnIndex);
    const r1 = Math.min(this.rowIndex + this.rowCount, o.rowIndex + o.rowCount), c1 = Math.min(this.columnIndex + this.columnCount, o.columnIndex + o.columnCount);
    if (r1 <= r0 || c1 <= c0 || o.isNullObject) return Object.assign(new Range(this.ws, 0, 0, 1, 1), { isNullObject: true });
    return new Range(this.ws, r0, c0, r1 - r0, c1 - c0);
  }
}
function mkFormat(range) {
  const rec = (path) => new Proxy({}, {
    get: (_, k) => k === "getItem" ? (e) => rec(`${path}.${e}`) : typeof k === "string" && /^(autofitColumns|clear)$/.test(k) ? () => log.push(`${path}.${k}() ${range.address}`) : rec(`${path}.${String(k)}`),
    set: (_, k, v) => { log.push(`${path}.${String(k)}=${v} ${range.address}`); return true; },
  });
  return rec("format");
}

const sales = new Sheet("Sales");
const data = [["Region", "Q1", "Q2", "Q3", "Q4"], ["North", 120, 135, 150, 160], ["South", 90, 95, 110, 105], ["East", 200, 190, 210, 230], ["West", 75, 80, 85, 95]];
sales.formulas; new Range(sales, 0, 0, 5, 5).formulas = data; log.length = 0;
const wb = {
  worksheets: Object.assign([sales], {
    items: [sales], load() {}, getItem: (n) => wb.worksheets.items.find((s) => s.name === n) || (() => { throw new Error(`Sheet ${n} not found`); })(),
    getActiveWorksheet: () => sales, add: (n) => { const s = new Sheet(n); wb.worksheets.items.push(s); log.push(`add sheet ${n}`); return s; },
  }),
  getSelectedRange: () => sales.getRange("A1:E5"),
  tables: { items: [], load() {} }, names: { items: [], load() {} },
};
globalThis.Excel = { run: async (fn) => fn({ workbook: wb, sync: async () => {} }) };

// ---------------- run ----------------
const { runTurn } = await import("../public/js/agent.js");
const { workbookContext } = await import("../public/js/tools.js");

const settings = { model: process.env.MODEL || "qwen3.8:latest", think: "low", numCtx: 32768, temperature: 0.6, maxSteps: 15, approval: "none" };
const prompt = process.argv[2] || "Add a Total column (F) with a formula summing Q1-Q4 for each region, and a Total row under the data summing each quarter. Bold the headers and format numbers with thousands separators.";
const c = await workbookContext();
const history = [{ role: "user", content: `[Excel context: sheet "${c.active_sheet}" · selection ${c.selection} · used ${c.used_range}]\n\n${prompt}` }];
const t0 = Date.now();
const ui = {
  step: () => ({ thinking() {}, content() {}, done() {},
    toolStart: (n, a) => { process.stdout.write(`→ ${n} ${JSON.stringify(a).slice(0, 160)}\n`); return { approve: async () => true, end: (r) => process.stdout.write(`  ← ${JSON.stringify(r).slice(0, 200)}\n`) }; } }),
  stats: (s) => console.log(`\n[${s.steps} steps, ${((Date.now() - t0) / 1000).toFixed(1)}s, ${(s.evalCount / (s.evalNs / 1e9)).toFixed(1)} tok/s]`),
};
const answer = await runTurn({ history, settings, signal: new AbortController().signal, ui, undo: [] });
console.log("\nFINAL:\n" + answer);
console.log("\nMOCK LOG:\n" + log.join("\n"));
for (const ws of wb.worksheets.items) {
  if (!ws.cells.size) continue;
  console.log(`\nSHEET ${ws.name} (values):`);
  for (const row of ws.getUsedRange().values) console.log(row.map((x) => String(x).slice(0, 44).padEnd(14)).join(""));
}
