// Plain-language descriptions for the panel: approval prompts and undo counts.
// Pure functions (no DOM, no Office.js) so they can be tested in Node.

const colNum = (letters) => [...letters.toUpperCase()].reduce((n, ch) => n * 26 + ch.charCodeAt(0) - 64, 0);

/** Number of cells in an A1 range like "B2", "E1:F20" or "Sheet1!A1:C5"; null for whole rows/columns or unknown forms. */
export function cellCount(address) {
  const a = String(address || "").split("!").pop().replace(/\$/g, "").trim();
  const m = /^([A-Z]{1,3})(\d+)(?::([A-Z]{1,3})(\d+))?$/i.exec(a);
  if (!m) return null;
  if (!m[3]) return 1;
  const cols = Math.abs(colNum(m[3]) - colNum(m[1])) + 1;
  const rows = Math.abs(Number(m[4]) - Number(m[2])) + 1;
  return cols * rows;
}

const plural = (n, one, many = one + "s") => `${n.toLocaleString("en-US")} ${n === 1 ? one : many}`;

function where(args) {
  const addr = args.address || "";
  return addr ? `${args.sheet ? args.sheet + "!" : ""}${addr}` : args.sheet || "";
}

function cellsIn(args) {
  const n = cellCount(args.address);
  const w = where(args);
  if (!w) return "cells";
  return n ? `${plural(n, "cell")} in ${w}` : w;
}

/**
 * What an approval prompt should say for a tool call: the action in plain words, whether Undo covers it,
 * and button labels that name the action.
 * @returns {{ text: string, undo: string, yes: string, no: string }}
 */
export function describeApproval(name, args = {}) {
  switch (name) {
    case "clear_range": {
      const what = args.what || "contents";
      if (what === "formats") return { text: `Remove the formatting from ${cellsIn(args)}.`, undo: "Undo can't restore formatting.", yes: "Clear formatting", no: "Don't clear" };
      return {
        text: `Clear ${cellsIn(args)}${what === "all" ? ", including formatting" : ""}.`,
        undo: what === "all" ? "Undo restores the values and formulas, not the formatting." : "You can undo this.",
        yes: "Clear cells", no: "Don't clear",
      };
    }
    case "execute_office_js":
      return { text: `Run code on the workbook${args.explanation ? `: ${String(args.explanation).replace(/\.$/, "")}` : ""}.`, undo: "Undo can't reverse what this code does.", yes: "Run code", no: "Don't run" };
    case "write_range":
      return { text: `Write to ${cellsIn(args)}.`, undo: "You can undo this.", yes: "Write cells", no: "Don't write" };
    case "insert_delete_cells":
      return args.action === "delete"
        ? { text: `Delete ${where(args) || "cells"}; the cells after it move to fill the gap.`, undo: "Undo can't reverse this.", yes: "Delete cells", no: "Don't delete" }
        : { text: `Insert cells at ${where(args) || "the selection"}.`, undo: "Undo can't reverse this.", yes: "Insert cells", no: "Don't insert" };
    case "manage_sheet": {
      const act = String(args.action || "change").toLowerCase();
      const sheet = args.name ? `"${args.name}"` : "a sheet";
      if (act === "delete") return { text: `Delete the sheet ${sheet} and everything on it.`, undo: "Undo can't bring a sheet back.", yes: "Delete sheet", no: "Don't delete" };
      if (act === "rename") return { text: `Rename the sheet ${sheet}${args.new_name ? ` to "${args.new_name}"` : ""}.`, undo: "Undo can't reverse this.", yes: "Rename sheet", no: "Don't rename" };
      return { text: `${act[0].toUpperCase() + act.slice(1)} the sheet ${sheet}.`, undo: "Undo can't reverse this.", yes: "Allow", no: "Don't allow" };
    }
    case "sort_range":
      return { text: `Sort ${cellsIn(args)}.`, undo: "Undo can't reverse a sort.", yes: "Sort", no: "Don't sort" };
    case "format_range":
      return { text: `Change the formatting of ${cellsIn(args)}.`, undo: "Undo can't restore formatting.", yes: "Format cells", no: "Don't format" };
    case "import_from_file":
      return { text: args.mode === "copy_sheet" ? "Copy a sheet from the attached file into this workbook." : `Import data into ${args.target_sheet ? args.target_sheet + "!" : ""}${args.target_address || "A1"}.`, undo: args.mode === "copy_sheet" ? "Undo can't remove a copied sheet." : "You can undo this.", yes: "Import", no: "Don't import" };
    case "create_table":
      return { text: `Turn ${where(args) || "the range"} into a table.`, undo: "Undo can't reverse this.", yes: "Create table", no: "Don't create" };
    case "create_chart":
      return { text: `Add a chart from ${where(args) || "the data"}.`, undo: "Undo can't remove the chart.", yes: "Add chart", no: "Don't add" };
    default:
      return { text: `Make a change to the workbook (${name}).`, undo: "Undo may not cover this.", yes: "Allow", no: "Don't allow" };
  }
}

/** Cells covered by the undo snapshots of a turn (each snapshot holds a 2D formulas array). */
export function undoCellCount(snapshots = []) {
  return snapshots.reduce((n, s) => n + (Array.isArray(s.formulas) ? s.formulas.reduce((m, row) => m + (Array.isArray(row) ? row.length : 1), 0) : 0), 0);
}

export function undoLabel(snapshots) {
  const n = undoCellCount(snapshots);
  return n ? `Undo ${plural(n, "cell change")}` : "Undo cell changes";
}
