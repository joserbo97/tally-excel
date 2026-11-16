// In-memory stand-in for the Office.js Excel API, good enough to run the agent's real tools
// headlessly and to score the result. Formulas are evaluated by HyperFormula across all sheets.
//
//   const { Excel, api } = createWorkbook({ sheets: { Sales: [["Region","Q1"],["North",120]] }, active: "Sales" });
//   globalThis.Excel = Excel;   // then run tools / the agent
//   api.value("Sales!B2")       // computed value, api.formula(...), api.style(...), api.setValue(...) ...
import XLSX from "xlsx";
import { shiftFormulaRows } from "../../public/js/formulas.js";
import { HyperFormula, FunctionPlugin, FunctionArgumentType } from "hyperformula";

// HyperFormula's TEXT() only knows a few formats ("#,##0.00" comes out as "1235,##0.00").
// Replace it with SheetJS's implementation of Excel number formats.
class ExcelTextPlugin extends FunctionPlugin {
  text(ast, state) {
    return this.runFunction(ast.args, state, this.metadata("TEXT"), (value, fmt) => {
      try {
        return XLSX.SSF.format(fmt, value);
      } catch {
        return String(value);
      }
    });
  }
}
ExcelTextPlugin.implementedFunctions = {
  TEXT: { method: "text", parameters: [{ argumentType: FunctionArgumentType.SCALAR }, { argumentType: FunctionArgumentType.STRING }] },
};
if (!globalThis.__oxaTextPlugin) {
  HyperFormula.unregisterFunction("TEXT");
  HyperFormula.registerFunctionPlugin(ExcelTextPlugin, { enGB: { TEXT: "TEXT" } });
  globalThis.__oxaTextPlugin = true;
}

const HF_FUNCS = new Set(HyperFormula.getRegisteredFunctionNames("enGB"));
const ERR = { DIV_BY_ZERO: "#DIV/0!", NAME: "#NAME?", VALUE: "#VALUE!", REF: "#REF!", NA: "#N/A", NUM: "#NUM!", NULL: "#NULL!", CYCLE: "#CYCLE!", SPILL: "#SPILL!", ERROR: "#ERROR!" };
const EXCEL_ERR = /^#(REF!|NAME\?|VALUE!|DIV\/0!|N\/A|NUM!|NULL!|SPILL!|CALC!|CYCLE!|ERROR!)/;

// ---------- addresses ----------
export const colToNum = (s) => [...s.toUpperCase()].reduce((n, ch) => n * 26 + ch.charCodeAt(0) - 64, 0) - 1;
export const numToCol = (n) => {
  let s = "";
  n++;
  while (n > 0) {
    const m = (n - 1) % 26;
    s = String.fromCharCode(65 + m) + s;
    n = Math.floor((n - 1) / 26);
  }
  return s;
};
const MAXR = 1048575, MAXC = 16383;

export function parseA1(a) {
  const clean = String(a).replace(/\$/g, "").trim();
  const [p1, p2 = p1] = clean.split(":");
  const one = (p) => {
    const m = /^([A-Za-z]*)(\d*)$/.exec(p);
    if (!m || (!m[1] && !m[2])) throw new Error(`Invalid address "${a}"`);
    return { c: m[1] ? colToNum(m[1]) : null, r: m[2] ? Number(m[2]) - 1 : null };
  };
  const A = one(p1), B = one(p2);
  return {
    r0: A.r ?? 0, c0: A.c ?? 0,
    r1: B.r ?? (A.r === null ? MAXR : A.r), c1: B.c ?? (A.c === null ? MAXC : A.c),
    wholeRows: A.c === null, wholeCols: A.r === null,
  };
}

const quoteSheet = (n) => (/^[A-Za-z_][\w.]*$/.test(n) ? n : `'${n.replace(/'/g, "''")}'`);
export function splitRef(ref) {
  const i = String(ref).lastIndexOf("!");
  if (i < 0) return { sheet: null, addr: String(ref) };
  return { sheet: ref.slice(0, i).replace(/^'(.*)'$/, "$1").replace(/''/g, "'"), addr: ref.slice(i + 1) };
}

// Shift relative row references in a formula by dRow, as Excel does when a sort moves a formula
// to another row ($-anchored rows stay put). Skips string literals and function names like LOG10(.
export { shiftFormulaRows }; // shared with the real tools (fill_down_to_row)

// HyperFormula differs from Excel on INDEX(<single-row range>, n): Excel returns the n-th column,
// HyperFormula treats n as a row number (#NUM!). Rewrite to INDEX(range, 1, n) before evaluating
// so valid Excel isn't scored as an error. Handles nested calls in the index argument.
export function excelIndexCompat(f) {
  if (typeof f !== "string" || !/INDEX\s*\(/i.test(f)) return f;
  let out = "", i = 0, quoted = false;
  while (i < f.length) {
    if (f[i] === '"') quoted = !quoted; // never rewrite inside "string literals"
    const m = /INDEX\s*\(/iy;
    m.lastIndex = i;
    if (quoted || !m.test(f) || (i > 0 && /[A-Za-z0-9_.]/.test(f[i - 1]))) { out += f[i++]; continue; }
    // parse balanced arguments
    let j = m.lastIndex, depth = 0, inStr = false;
    const args = [""];
    for (; j < f.length; j++) {
      const ch = f[j];
      if (ch === '"') inStr = !inStr;
      if (!inStr && ch === "(") depth++;
      if (!inStr && ch === ")") { if (depth === 0) break; depth--; }
      if (!inStr && ch === "," && depth === 0) args.push("");
      else args[args.length - 1] += ch;
    }
    const inner = args.map((a) => excelIndexCompat(a));
    const oneRow = /^\s*(?:(?:'[^']+'|[\w.]+)!)?\$?[A-Za-z]{1,3}\$?(\d+):\$?[A-Za-z]{1,3}\$?(\d+)\s*$/.exec(inner[0] || "");
    if (inner.length === 2 && oneRow && oneRow[1] === oneRow[2]) inner.splice(1, 0, "1");
    out += `INDEX(${inner.join(",")})`;
    i = j + 1;
  }
  return out;
}

// HyperFormula rejects Excel's bare TRUE/FALSE literals (#NAME?) — e.g. every exact-match
// VLOOKUP(..., FALSE). It accepts TRUE()/FALSE(), so rewrite bare ones (outside "strings").
export function excelBoolCompat(f) {
  if (typeof f !== "string" || !/\b(TRUE|FALSE)\b/i.test(f)) return f;
  return f
    .split(/("(?:[^"]|"")*")/)
    .map((part, i) => (i % 2 ? part : part.replace(/(^|[^A-Za-z0-9_.!$'])(TRUE|FALSE)(?!\s*\(|[A-Za-z0-9_!])/gi, (m, pre, b) => `${pre}${b.toUpperCase()}()`)))
    .join("");
}

// Everything that makes a valid Excel formula evaluate the way Excel would in HyperFormula.
export const toEngineFormula = (f) => excelBoolCompat(excelIndexCompat(f));

// Excel interprets written strings like typed input.
const isoToSerial = (y, m, d) => (Date.UTC(y, m - 1, d) - Date.UTC(1899, 11, 30)) / 86400000;
function parseInput(v) {
  if (v === null || v === undefined) return { empty: true };
  if (typeof v === "number" || typeof v === "boolean") return { v };
  const s = String(v);
  if (s === "") return { empty: true };
  if (s[0] === "'") return { v: s.slice(1), text: true };
  if (s[0] === "=") return { f: s };
  if (/^[+-]?\d+(\.\d+)?(e[+-]?\d+)?$/i.test(s.trim())) return { v: Number(s) };
  if (/^[+-]?\d{1,3}(,\d{3})+(\.\d+)?$/.test(s.trim())) return { v: Number(s.replace(/,/g, "")) };
  const pct = /^([+-]?\d+(?:\.\d+)?)%$/.exec(s.trim());
  if (pct) return { v: Number(pct[1]) / 100, fmt: "0%" };
  if (/^(true|false)$/i.test(s)) return { v: /^true$/i.test(s) };
  const d = /^(\d{4})-(\d{2})-(\d{2})$/.exec(s.trim());
  if (d) return { v: isoToSerial(+d[1], +d[2], +d[3]), fmt: "yyyy-mm-dd" };
  if (s[0] === "+" || s[0] === "-" || s[0] === "@") return { f: "=" + s.replace(/^\+/, "") }; // like typing it
  return { v: s };
}

export function createWorkbook(spec = {}) {
  const sheets = []; // Sheet[]
  let version = 0;
  let hf = null, hfVersion = -1;
  let active = null;
  let selection = { sheet: null, addr: "A1" };
  const log = [];
  const touched = () => version++;

  // ---------- formula engine ----------
  function engine() {
    if (hf && hfVersion === version) return hf;
    hf?.destroy();
    const data = {};
    for (const s of sheets) {
      const grid = [];
      for (const [k, cell] of s.cells) {
        const [r, c] = k.split(",").map(Number);
        (grid[r] ||= [])[c] = cell.f !== undefined ? toEngineFormula(cell.f) : cell.v;
      }
      data[s.name] = Array.from(grid, (row) => Array.from(row || [], (x) => (x === undefined ? null : x)));
    }
    hf = HyperFormula.buildFromSheets(data, { licenseKey: "gpl-v3", useArrayArithmetic: true, nullDate: { year: 1899, month: 12, day: 30 } });
    hfVersion = version;
    return hf;
  }
  function computed(sheet, r, c) {
    const cell = sheet.cells.get(`${r},${c}`);
    if (!cell) {
      // Spilled array results land in cells with no content of their own.
      const v = engine().getCellValue({ sheet: engine().getSheetId(sheet.name), row: r, col: c });
      return v === null || v === undefined ? "" : typeof v === "object" ? ERR[v.type] || "#ERROR!" : v;
    }
    if (cell.f === undefined) return cell.v;
    const v = engine().getCellValue({ sheet: engine().getSheetId(sheet.name), row: r, col: c });
    if (v && typeof v === "object") return ERR[v.type] || "#ERROR!";
    return v ?? 0;
  }
  // Functions Excel has but HyperFormula doesn't: results there can't be scored.
  const unknownFunctions = (f) => [...String(f).matchAll(/([A-Z][A-Z0-9.]*)\s*\(/gi)].map((m) => m[1].toUpperCase()).filter((n) => !HF_FUNCS.has(n));

  // ---------- Range ----------
  class Range {
    constructor(ws, r, c, rows, cols) {
      Object.assign(this, { ws, rowIndex: r, columnIndex: c, rowCount: rows, columnCount: cols, isNullObject: false });
      this.format = makeFormat(this);
      this.sort = { apply: (fields, matchCase, hasHeaders) => sortRange(this, fields, hasHeaders) };
    }
    get address() {
      const a = `${numToCol(this.columnIndex)}${this.rowIndex + 1}`;
      const b = `${numToCol(this.columnIndex + this.columnCount - 1)}${this.rowIndex + this.rowCount}`;
      return `${quoteSheet(this.ws.name)}!${a === b ? a : `${a}:${b}`}`;
    }
    *each() {
      const maxR = Math.min(this.rowIndex + this.rowCount, this.rowIndex + 5000);
      const maxC = Math.min(this.columnIndex + this.columnCount, this.columnIndex + 300);
      for (let r = this.rowIndex; r < maxR; r++) for (let c = this.columnIndex; c < maxC; c++) yield [r, c];
    }
    grid(fn) {
      const rows = Math.min(this.rowCount, 5000), cols = Math.min(this.columnCount, 300);
      return Array.from({ length: rows }, (_, i) => Array.from({ length: cols }, (_, j) => fn(this.rowIndex + i, this.columnIndex + j)));
    }
    get values() { return this.grid((r, c) => computed(this.ws, r, c)); }
    set values(v) { this.write(v); }
    get formulas() { return this.grid((r, c) => { const x = this.ws.cells.get(`${r},${c}`); return x ? (x.f ?? x.v) : ""; }); }
    set formulas(v) { this.write(v); }
    get text() {
      return this.grid((r, c) => {
        const v = computed(this.ws, r, c);
        const fmt = this.ws.styles.get(`${r},${c}`)?.numberFormat;
        return typeof v === "number" && fmt ? XLSX.SSF.format(fmt, v) : String(v);
      });
    }
    get numberFormat() { return this.grid((r, c) => this.ws.styles.get(`${r},${c}`)?.numberFormat || "General"); }
    set numberFormat(v) {
      v.forEach((row, i) => row.forEach((fmt, j) => style(this.ws, this.rowIndex + i, this.columnIndex + j).numberFormat = fmt));
      touched();
    }
    write(v) {
      if (!Array.isArray(v) || !v.every(Array.isArray)) throw new Error("InvalidArgument: expected a 2D array");
      if (v.length !== this.rowCount || v.some((row) => row.length !== this.columnCount)) {
        throw new Error(`InvalidArgument: array is ${v.length}x${v[0]?.length}, range is ${this.rowCount}x${this.columnCount}`);
      }
      v.forEach((row, i) => row.forEach((x, j) => {
        const k = `${this.rowIndex + i},${this.columnIndex + j}`;
        const p = parseInput(x);
        if (p.empty) this.ws.cells.delete(k);
        else this.ws.cells.set(k, p.f !== undefined ? { f: p.f } : { v: p.v });
        if (p.fmt) style(this.ws, this.rowIndex + i, this.columnIndex + j).numberFormat ??= p.fmt;
      }));
      touched();
      log.push(`write ${this.address}`);
    }
    load() { return this; }
    select() { selection = { sheet: this.ws.name, addr: this.address.split("!").pop() }; active = this.ws; }
    getCell(r, c) { return new Range(this.ws, this.rowIndex + r, this.columnIndex + c, 1, 1); }
    getResizedRange(dr, dc) { return new Range(this.ws, this.rowIndex, this.columnIndex, this.rowCount + dr, this.columnCount + dc); }
    getOffsetRange(dr, dc) { return new Range(this.ws, this.rowIndex + dr, this.columnIndex + dc, this.rowCount, this.columnCount); }
    getIntersectionOrNullObject(o) {
      const r0 = Math.max(this.rowIndex, o.rowIndex), c0 = Math.max(this.columnIndex, o.columnIndex);
      const r1 = Math.min(this.rowIndex + this.rowCount, o.rowIndex + o.rowCount), c1 = Math.min(this.columnIndex + this.columnCount, o.columnIndex + o.columnCount);
      if (o.isNullObject || r1 <= r0 || c1 <= c0) return nullRange(this.ws);
      return new Range(this.ws, r0, c0, r1 - r0, c1 - c0);
    }
    clear(what = "All") {
      for (const [r, c] of this.each()) {
        if (what !== "Formats") this.ws.cells.delete(`${r},${c}`);
        if (what !== "Contents") this.ws.styles.delete(`${r},${c}`);
      }
      touched();
      log.push(`clear ${this.address} ${what}`);
    }
    insert(shift) { shiftCells(this, shift, +1); }
    delete(shift) { shiftCells(this, shift, -1); }
    merge() { this.ws.merges.push(this.address); }
    unmerge() { this.ws.merges = this.ws.merges.filter((m) => m !== this.address); }
    getUsedRange() { return this; }
  }
  const nullRange = (ws) => Object.assign(new Range(ws, 0, 0, 1, 1), { isNullObject: true });

  function style(ws, r, c) {
    const k = `${r},${c}`;
    if (!ws.styles.has(k)) ws.styles.set(k, {});
    return ws.styles.get(k);
  }
  function makeFormat(range) {
    const setAll = (key, val) => { for (const [r, c] of range.each()) style(range.ws, r, c)[key] = val; };
    const font = {};
    for (const k of ["bold", "italic", "underline", "color", "size", "name"]) {
      Object.defineProperty(font, k, { set: (v) => setAll(k === "color" ? "fontColor" : k, v), get: () => style(range.ws, range.rowIndex, range.columnIndex)[k] });
    }
    const fill = { clear: () => setAll("fill", null) };
    Object.defineProperty(fill, "color", { set: (v) => setAll("fill", v), get: () => style(range.ws, range.rowIndex, range.columnIndex).fill });
    const borders = { getItem: (edge) => ({ set style(v) { setAll(`border_${edge}`, v); }, set color(v) {} }) };
    const fmt = { font, fill, borders, autofitColumns() {}, autofitRows() {} };
    for (const k of ["horizontalAlignment", "verticalAlignment", "wrapText", "columnWidth", "rowHeight"]) {
      Object.defineProperty(fmt, k, { set: (v) => setAll(k, v), get: () => style(range.ws, range.rowIndex, range.columnIndex)[k] });
    }
    return fmt;
  }

  function sortRange(range, fields, hasHeaders) {
    const start = range.rowIndex + (hasHeaders ? 1 : 0);
    const end = range.rowIndex + range.rowCount;
    const rows = [];
    for (let r = start; r < end; r++) {
      const cells = [], styles = [], keys = [];
      for (let c = range.columnIndex; c < range.columnIndex + range.columnCount; c++) {
        cells.push(range.ws.cells.get(`${r},${c}`));
        styles.push(range.ws.styles.get(`${r},${c}`));
      }
      for (const f of fields) keys.push(computed(range.ws, r, range.columnIndex + (Number(f.key) || 0)));
      rows.push({ cells, styles, keys, from: r });
    }
    rows.sort((a, b) => {
      for (let i = 0; i < fields.length; i++) {
        const x = a.keys[i], y = b.keys[i];
        const cmp = typeof x === "number" && typeof y === "number" ? x - y : String(x).localeCompare(String(y));
        if (cmp) return fields[i].ascending === false ? -cmp : cmp;
      }
      return 0;
    });
    rows.forEach((row, i) => row.cells.forEach((cell, j) => {
      const k = `${start + i},${range.columnIndex + j}`;
      const moved = cell?.f !== undefined ? { f: shiftFormulaRows(cell.f, start + i - row.from) } : cell;
      moved ? range.ws.cells.set(k, moved) : range.ws.cells.delete(k);
      row.styles[j] ? range.ws.styles.set(k, row.styles[j]) : range.ws.styles.delete(k);
    }));
    touched();
    log.push(`sort ${range.address}`);
  }

  // Insert/delete whole rows/columns or a block (cells shift; formula references are NOT adjusted).
  function shiftCells(range, shift, dir) {
    const ws = range.ws;
    const vertical = /down|up/i.test(shift) || range.columnCount > 16000;
    const n = vertical ? range.rowCount : range.columnCount;
    const moved = new Map(), movedS = new Map();
    for (const [store, out] of [[ws.cells, moved], [ws.styles, movedS]]) {
      for (const [k, v] of store) {
        let [r, c] = k.split(",").map(Number);
        const inBand = vertical ? c >= range.columnIndex && c < range.columnIndex + range.columnCount : r >= range.rowIndex && r < range.rowIndex + range.rowCount;
        if (inBand) {
          if (vertical && r >= range.rowIndex) {
            if (dir < 0 && r < range.rowIndex + n) continue;
            r += dir * n;
          } else if (!vertical && c >= range.columnIndex) {
            if (dir < 0 && c < range.columnIndex + n) continue;
            c += dir * n;
          }
        }
        out.set(`${r},${c}`, v);
      }
    }
    ws.cells = moved;
    ws.styles = movedS;
    touched();
    log.push(`${dir > 0 ? "insert" : "delete"} ${range.address} ${shift}`);
  }

  // ---------- Worksheet ----------
  class Sheet {
    constructor(name) {
      Object.assign(this, { name, id: name, cells: new Map(), styles: new Map(), merges: [], tablesList: [], chartsList: [], frozen: { rows: 0, cols: 0 }, visibility: "Visible" });
      this.tables = { add: (addr, hasHeaders) => addTable(this, addr, hasHeaders), items: this.tablesList, load() {} };
      this.charts = { add: (type, range, seriesBy) => addChart(this, type, range, seriesBy), items: this.chartsList, load() {} };
      this.freezePanes = {
        freezeRows: (n = 1) => { this.frozen.rows = n; },
        freezeColumns: (n = 1) => { this.frozen.cols = n; },
        freezeAt: (r) => { const p = parseA1(typeof r === "string" ? r : r.address.split("!").pop()); this.frozen = { rows: p.r1 + 1, cols: p.c1 + 1 }; },
        unfreeze: () => { this.frozen = { rows: 0, cols: 0 }; },
      };
    }
    get position() { return sheets.indexOf(this); }
    set position(p) { sheets.splice(sheets.indexOf(this), 1); sheets.splice(p, 0, this); }
    load() { return this; }
    getRange(a = "A1") {
      if (a.includes("!")) throw new Error("InvalidArgument: worksheet.getRange takes an address without a sheet name");
      const p = parseA1(a);
      return new Range(this, p.r0, p.c0, p.r1 - p.r0 + 1, p.c1 - p.c0 + 1);
    }
    getRangeByIndexes(r, c, rows, cols) { return new Range(this, r, c, rows, cols); }
    getCell(r, c) { return new Range(this, r, c, 1, 1); }
    usedBounds(valuesOnly) {
      let r0 = Infinity, c0 = Infinity, r1 = -1, c1 = -1;
      const keys = valuesOnly ? [...this.cells.keys()] : [...this.cells.keys(), ...this.styles.keys()];
      for (const k of keys) {
        const [r, c] = k.split(",").map(Number);
        r0 = Math.min(r0, r); c0 = Math.min(c0, c); r1 = Math.max(r1, r); c1 = Math.max(c1, c);
      }
      // Spilled results count as used, as in Excel.
      if ([...this.cells.values()].some((x) => x.f !== undefined)) {
        engine().getSheetValues(engine().getSheetId(this.name)).forEach((row, r) => row.forEach((v, c) => {
          if (v === null || v === undefined || v === "") return;
          r0 = Math.min(r0, r); c0 = Math.min(c0, c); r1 = Math.max(r1, r); c1 = Math.max(c1, c);
        }));
      }
      return r1 < 0 ? null : { r0, c0, r1, c1 };
    }
    getUsedRangeOrNullObject(valuesOnly) {
      const b = this.usedBounds(valuesOnly);
      return b ? new Range(this, b.r0, b.c0, b.r1 - b.r0 + 1, b.c1 - b.c0 + 1) : nullRange(this);
    }
    getUsedRange(valuesOnly) {
      const r = this.getUsedRangeOrNullObject(valuesOnly);
      return r.isNullObject ? new Range(this, 0, 0, 1, 1) : r;
    }
    activate() { active = this; }
    delete() {
      if (sheets.length === 1) throw new Error("InvalidOperation: a workbook must contain at least one sheet");
      sheets.splice(sheets.indexOf(this), 1);
      if (active === this) active = sheets[0];
      touched();
    }
    copy() {
      let n = 2;
      while (sheets.some((s) => s.name === `${this.name} (${n})`)) n++;
      const s = addSheet(`${this.name} (${n})`);
      for (const [k, v] of this.cells) s.cells.set(k, { ...v });
      for (const [k, v] of this.styles) s.styles.set(k, { ...v });
      touched();
      return s;
    }
    findAllOrNullObject(text, { completeMatch = false, matchCase = false } = {}) {
      const want = matchCase ? String(text) : String(text).toLowerCase();
      const hits = [];
      for (const k of this.cells.keys()) {
        const [r, c] = k.split(",").map(Number);
        let v = String(computed(this, r, c));
        if (!matchCase) v = v.toLowerCase();
        if (completeMatch ? v === want : v.includes(want)) hits.push(new Range(this, r, c, 1, 1));
      }
      if (!hits.length) return { isNullObject: true, load() {}, areas: { load() {}, items: [] } };
      return { isNullObject: false, address: hits.map((h) => h.address).join(","), load() {}, areas: { load() {}, items: hits } };
    }
  }
  function addSheet(name) {
    if (!name || /[\\/?*[\]:]/.test(name) || name.length > 31) throw new Error(`InvalidArgument: invalid sheet name "${name}"`);
    if (sheets.some((s) => s.name.toLowerCase() === name.toLowerCase())) throw new Error(`ItemAlreadyExists: a sheet named "${name}" already exists`);
    const s = new Sheet(name);
    sheets.push(s);
    touched();
    return s;
  }
  function getSheet(name) {
    const s = sheets.find((x) => x.name.toLowerCase() === String(name).toLowerCase() || x.id === name);
    if (!s) throw new Error(`ItemNotFound: sheet "${name}" doesn't exist`);
    return s;
  }
  function addTable(ws, addr, hasHeaders) {
    const range = typeof addr === "string" ? ws.getRange(splitRef(addr).addr) : addr;
    const t = { name: `Table${sheets.reduce((n, s) => n + s.tablesList.length, 0) + 1}`, style: "TableStyleMedium2", hasHeaders, sheet: ws.name, address: range.address, load() {}, getRange: () => range };
    ws.tablesList.push(t);
    log.push(`table ${range.address}`);
    return t;
  }
  function addChart(ws, type, range, seriesBy) {
    const c = { name: `Chart ${ws.chartsList.length + 1}`, type, source: range.address, seriesBy, title: { text: "", visible: false }, position: null, load() {}, setPosition(start) { this.position = typeof start === "string" ? start : start.address; } };
    ws.chartsList.push(c);
    log.push(`chart ${type} ${range.address}`);
    return c;
  }

  // ---------- build from spec ----------
  for (const [name, content] of Object.entries(spec.sheets || { Sheet1: [] })) {
    const s = addSheet(name);
    if (Array.isArray(content)) {
      content.forEach((row, r) => row.forEach((v, c) => {
        const p = parseInput(v);
        if (!p.empty) s.cells.set(`${r},${c}`, p.f !== undefined ? { f: p.f } : { v: p.v });
        if (p.fmt) style(s, r, c).numberFormat = p.fmt;
      }));
    }
    for (const [addr, fmt] of Object.entries(spec.formats?.[name] || {})) {
      const p = parseA1(addr);
      for (let r = p.r0; r <= p.r1; r++) for (let c = p.c0; c <= p.c1; c++) style(s, r, c).numberFormat = fmt;
    }
  }
  active = spec.active ? getSheet(spec.active) : sheets[0];
  selection = { sheet: active.name, addr: spec.selection || "A1" };
  log.length = 0;

  const worksheets = {
    get items() { return sheets; },
    load() {},
    getItem: getSheet,
    getItemOrNullObject: (n) => { try { return getSheet(n); } catch { return { isNullObject: true, load() {} }; } },
    getActiveWorksheet: () => active,
    add: (name) => addSheet(name || `Sheet${sheets.length + 1}`),
    getCount: () => ({ value: sheets.length }),
  };
  const settingsStore = new Map();
  const workbook = {
    name: spec.name || "Book1.xlsx",
    worksheets,
    load() {},
    getSelectedRange: () => getSheet(selection.sheet).getRange(selection.addr),
    get tables() { return { items: sheets.flatMap((s) => s.tablesList), load() {} }; },
    names: { items: [], load() {} },
    settings: {
      getItemOrNullObject: (k) => (settingsStore.has(k) ? { isNullObject: false, value: settingsStore.get(k), load() {} } : { isNullObject: true, load() {} }),
      add: (k, v) => settingsStore.set(k, v),
    },
    calculate() {},
    application: { calculate() {}, load() {} },
    insertWorksheetsFromBase64(b64, opts = {}) {
      const src = XLSX.read(Buffer.from(b64, "base64"), { type: "buffer", cellFormula: true, cellNF: true });
      const names = opts.sheetNamesToInsert?.length ? opts.sheetNamesToInsert : src.SheetNames;
      const ids = [];
      for (const n of names) {
        const ws = src.Sheets[n];
        if (!ws) throw new Error(`ItemNotFound: sheet "${n}" not in file`);
        let name = n, i = 2;
        while (sheets.some((s) => s.name.toLowerCase() === name.toLowerCase())) name = `${n} (${i++})`;
        const s = addSheet(name);
        for (const k of Object.keys(ws)) {
          if (k[0] === "!") continue;
          const a = XLSX.utils.decode_cell(k), cell = ws[k];
          s.cells.set(`${a.r},${a.c}`, cell.f ? { f: "=" + cell.f } : { v: cell.v instanceof Date ? isoToSerial(cell.v.getUTCFullYear(), cell.v.getUTCMonth() + 1, cell.v.getUTCDate()) : cell.v });
          if (cell.z && cell.z !== "General") style(s, a.r, a.c).numberFormat = cell.z;
        }
        ids.push(s.id);
      }
      touched();
      return { value: ids };
    },
  };
  const context = { workbook, sync: async () => {}, trackedObjects: { add() {}, remove() {} } };

  const Excel = {
    run: async (fn) => fn(context),
    WorksheetPositionType: { none: "None", before: "Before", after: "After", beginning: "Beginning", end: "End" },
    InsertShiftDirection: { down: "Down", right: "Right" },
    DeleteShiftDirection: { up: "Up", left: "Left" },
    ChartType: new Proxy({}, { get: (_, k) => String(k)[0].toUpperCase() + String(k).slice(1) }),
    ChartSeriesBy: { auto: "Auto", columns: "Columns", rows: "Rows" },
    ClearApplyTo: { all: "All", contents: "Contents", formats: "Formats" },
  };

  // ---------- scoring API ----------
  const at = (ref, defaultSheet) => {
    const { sheet, addr } = splitRef(ref);
    const s = getSheet(sheet || defaultSheet || active.name);
    const p = parseA1(addr);
    return { s, r: p.r0, c: p.c0, p };
  };
  const api = {
    log,
    get version() { return version; },
    sheetNames: () => sheets.map((s) => s.name),
    hasSheet: (n) => sheets.some((s) => s.name.toLowerCase() === String(n).toLowerCase()),
    sheet: getSheet,
    value: (ref) => { const { s, r, c } = at(ref); return computed(s, r, c); },
    formula: (ref) => { const { s, r, c } = at(ref); const x = s.cells.get(`${r},${c}`); return x?.f ?? null; },
    raw: (ref) => { const { s, r, c } = at(ref); return s.cells.get(`${r},${c}`) ?? null; },
    style: (ref) => { const { s, r, c } = at(ref); return s.styles.get(`${r},${c}`) || {}; },
    values: (ref) => { const { s, p } = at(ref); return new Range(s, p.r0, p.c0, p.r1 - p.r0 + 1, p.c1 - p.c0 + 1).values; },
    setValue: (ref, v) => { const { s, r, c } = at(ref); s.cells.set(`${r},${c}`, { v }); touched(); },
    usedRange: (name) => { const r = getSheet(name).getUsedRangeOrNullObject(true); return r.isNullObject ? null : r.address; },
    // Every non-empty cell as { ref, value, formula }.
    // Includes cells filled by a spilling formula (FILTER, UNIQUE, SORT...), as in Excel.
    cells: (name) => {
      const s = getSheet(name);
      const out = [...s.cells.keys()].map((k) => {
        const [r, c] = k.split(",").map(Number);
        return { ref: `${numToCol(c)}${r + 1}`, r, c, value: computed(s, r, c), formula: s.cells.get(k).f ?? null };
      });
      if ([...s.cells.values()].some((x) => x.f !== undefined)) {
        const grid = engine().getSheetValues(engine().getSheetId(s.name));
        grid.forEach((row, r) => row.forEach((v, c) => {
          if (v === null || v === undefined || v === "" || s.cells.has(`${r},${c}`)) return;
          out.push({ ref: `${numToCol(c)}${r + 1}`, r, c, value: typeof v === "object" ? ERR[v.type] || "#ERROR!" : v, formula: null, spilled: true });
        }));
      }
      return out.sort((a, b) => a.r - b.r || a.c - b.c);
    },
    find: (name, pred) => api.cells(name).filter((x) => (typeof pred === "function" ? pred(x.value, x) : pred instanceof RegExp ? pred.test(String(x.value)) : x.value === pred)),
    charts: () => sheets.flatMap((s) => s.chartsList.map((c) => ({ ...c, sheet: s.name }))),
    tables: () => sheets.flatMap((s) => s.tablesList.map((t) => ({ ...t, sheet: s.name }))),
    frozen: (name) => getSheet(name).frozen,
    errors: () => sheets.flatMap((s) => api.cells(s.name).filter((x) => typeof x.value === "string" && EXCEL_ERR.test(x.value)).map((x) => `${s.name}!${x.ref}=${x.value}`)),
    // Formulas using Excel functions the scoring engine can't evaluate (results can't be judged).
    unverifiable: () => sheets.flatMap((s) => api.cells(s.name).filter((x) => x.formula && unknownFunctions(x.formula).length).map((x) => `${s.name}!${x.ref}: ${unknownFunctions(x.formula).join(",")}`)),
    selection: () => selection,
  };
  return { Excel, api, context };
}
