/* global Excel */
// Excel tools the agent can call. Each tool has:
//   schema  - the JSON-schema function definition sent to Ollama
//   risk    - "read" | "write" | "destructive" (drives the approval prompt)
//   run     - async (args, ctx) => result object (serialized back to the model)

import { shiftFormulaRows } from "./formulas.js";
import { findColumn, compareRows, issueText } from "./compare.js";

const MAX_READ_CELLS = 3000;
const MAX_FILL_ROWS = 100000;
const MAX_RESULT_CHARS = 14000;

// ---------- helpers ----------
export function colLetter(n) {
  // 0-based column index -> letters
  let s = "";
  n += 1;
  while (n > 0) {
    const m = (n - 1) % 26;
    s = String.fromCharCode(65 + m) + s;
    n = Math.floor((n - 1) / 26);
  }
  return s;
}

function splitAddress(sheet, address) {
  // Accepts "Sheet1!A1:B2", "'My Sheet'!A1", or plain "A1:B2" + sheet arg.
  let a = String(address || "").trim();
  let s = sheet ? String(sheet).trim() : "";
  const bang = a.lastIndexOf("!");
  if (bang > -1) {
    s = a.slice(0, bang).replace(/^'(.*)'$/, "$1").replace(/''/g, "'");
    a = a.slice(bang + 1);
  }
  a = a.replace(/\$/g, "");
  return { sheet: s, address: a };
}

function getSheet(context, sheetName) {
  return sheetName
    ? context.workbook.worksheets.getItem(sheetName)
    : context.workbook.worksheets.getActiveWorksheet();
}

function getRange(context, sheet, address) {
  const p = splitAddress(sheet, address);
  if (!p.address) throw new Error("An address like 'A1' or 'B2:D10' is required.");
  const ws = getSheet(context, p.sheet);
  return { ws, range: ws.getRange(p.address) };
}

function to2D(values) {
  if (typeof values === "string") {
    try {
      values = JSON.parse(values);
    } catch {
      values = [[values]];
    }
  }
  if (!Array.isArray(values)) values = [[values]];
  if (!values.every(Array.isArray)) values = [values]; // single row given as flat list
  const width = Math.max(1, ...values.map((r) => r.length));
  return values.map((r) =>
    Array.from({ length: width }, (_, i) => {
      const v = r[i];
      if (v === null || v === undefined) return "";
      if (typeof v === "object") return JSON.stringify(v);
      return v;
    })
  );
}

const cellAddr = (row0, col0) => `${colLetter(col0)}${row0 + 1}`;

const ERROR_RX = /^#(REF!|NAME\?|VALUE!|DIV\/0!|N\/A|NUM!|NULL!|SPILL!|CALC!)/;

function collectErrors(values, rowIndex, columnIndex, limit = 20) {
  const errs = [];
  values.forEach((row, r) =>
    row.forEach((v, c) => {
      if (typeof v === "string" && ERROR_RX.test(v) && errs.length < limit) {
        errs.push(`${cellAddr(rowIndex + r, columnIndex + c)}: ${v}`);
      }
    })
  );
  return errs;
}

// ---------- tool implementations ----------
const tools = {
  get_workbook_info: {
    risk: "read",
    schema: {
      description:
        "Get an overview of the workbook: every sheet with its used range, tables, named ranges, the active sheet and the user's current selection. Call this first when you don't know the workbook layout.",
      parameters: { type: "object", properties: {} },
    },
    async run() {
      return Excel.run(async (context) => {
        const sheets = context.workbook.worksheets;
        sheets.load("items/name,items/visibility,items/position");
        const active = sheets.getActiveWorksheet();
        active.load("name");
        const sel = context.workbook.getSelectedRange();
        sel.load("address");
        const tables = context.workbook.tables;
        tables.load("items/name");
        const names = context.workbook.names;
        names.load("items/name,items/formula");
        await context.sync();

        const used = sheets.items.map((ws) => {
          const u = ws.getUsedRangeOrNullObject(true);
          u.load("address,rowCount,columnCount");
          return u;
        });
        const tableInfo = tables.items.map((t) => {
          const r = t.getRange();
          r.load("address");
          return { t, r };
        });
        await context.sync();

        return {
          active_sheet: active.name,
          selection: sel.address,
          sheets: sheets.items.map((ws, i) => ({
            name: ws.name,
            visible: ws.visibility === "Visible",
            used_range: used[i].isNullObject ? null : used[i].address.split("!").pop(),
            rows: used[i].isNullObject ? 0 : used[i].rowCount,
            columns: used[i].isNullObject ? 0 : used[i].columnCount,
          })),
          tables: tableInfo.map(({ t, r }) => ({ name: t.name, range: r.address })),
          named_ranges: names.items.map((n) => ({ name: n.name, refers_to: n.formula })),
        };
      });
    },
  },

  read_range: {
    risk: "read",
    schema: {
      description:
        "Read cell values and formulas from a range. Use A1 notation, optionally with a sheet prefix (e.g. 'Sales!A1:F50'). Whole columns/rows like 'A:D' are clipped to the used range. Returns raw values row by row plus a map of cells that contain formulas.",
      parameters: {
        type: "object",
        properties: {
          address: { type: "string", description: "Range in A1 notation, e.g. 'A1:D20' or 'Sheet2!B:B'." },
          sheet: { type: "string", description: "Sheet name. Defaults to the active sheet." },
          as_displayed: {
            type: "boolean",
            description: "Return the text as shown in Excel (formatted dates, %, currency) instead of raw values.",
          },
        },
        required: ["address"],
      },
    },
    async run({ address, sheet, as_displayed }) {
      return Excel.run(async (context) => {
        const { ws, range: requested } = getRange(context, sheet, address);
        const used = ws.getUsedRangeOrNullObject(true);
        used.load("address");
        await context.sync();
        let range = used.isNullObject ? null : requested.getIntersectionOrNullObject(used);
        if (range) {
          range.load("address,rowCount,columnCount");
          await context.sync();
        }
        if (!range || range.isNullObject) return { address: splitAddress(sheet, address).address, empty: true, values: [] };

        let truncated = false;
        const maxRows = Math.max(1, Math.floor(MAX_READ_CELLS / range.columnCount));
        if (range.rowCount * range.columnCount > MAX_READ_CELLS) {
          truncated = true;
          range = range.getCell(0, 0).getResizedRange(Math.min(range.rowCount, maxRows) - 1, Math.min(range.columnCount, 60) - 1);
        }
        range.load("address,rowIndex,columnIndex,values,formulas,text,rowCount,columnCount");
        await context.sync();

        const formulas = {};
        range.formulas.forEach((row, r) =>
          row.forEach((f, c) => {
            if (typeof f === "string" && f.startsWith("=")) formulas[cellAddr(range.rowIndex + r, range.columnIndex + c)] = f;
          })
        );
        const out = {
          address: range.address,
          top_left: cellAddr(range.rowIndex, range.columnIndex),
          rows: range.rowCount,
          columns: range.columnCount,
          values: as_displayed ? range.text : range.values,
        };
        if (Object.keys(formulas).length) out.formulas = formulas;
        const errs = collectErrors(range.values, range.rowIndex, range.columnIndex);
        if (errs.length) out.errors = errs;
        if (truncated) out.note = `Range was larger than ${MAX_READ_CELLS} cells; only ${range.address} returned. Read the rest in chunks.`;
        return out;
      });
    },
  },

  write_range: {
    risk: "write",
    schema: {
      description:
        "Write values and/or formulas into cells. Give the top-left cell (e.g. 'B2') or the full range, plus a 2D array (list of rows). Strings starting with '=' are entered as formulas. The range is automatically sized to the array. Returns the resulting computed values so you can check for errors. To put the same formulas on many rows, don't type every row: write the header and the FIRST formula row only, and set fill_down_to_row — the last row you give is copied down to that row with relative references adjusted, like Excel's fill handle.",
      parameters: {
        type: "object",
        properties: {
          address: { type: "string", description: "Top-left cell or range, e.g. 'A1' or 'Sheet1!A1:C3'." },
          sheet: { type: "string", description: "Sheet name. Defaults to the active sheet." },
          values: {
            type: "array",
            description: "2D array of rows, e.g. [[\"Name\",\"Qty\"],[\"Apples\",3],[\"Total\",\"=SUM(B2:B2)\"]].",
            items: { type: "array", items: {} },
          },
          fill_down_to_row: {
            type: "integer",
            description: "Optional sheet row number (1-based). Copies the last row of `values` down through this row, shifting relative row references (A2 → A3 → …; $A$2 stays). E.g. address 'D1', values [[\"Match\"],[\"=COUNTIF(Bank!A:A,A2)\"]], fill_down_to_row 61 fills D2:D61.",
          },
        },
        required: ["address", "values"],
      },
    },
    async run({ address, sheet, values, fill_down_to_row }, ctx) {
      let data = to2D(values);
      return Excel.run(async (context) => {
        const { ws, range: anchor } = getRange(context, sheet, address);
        if (fill_down_to_row != null) {
          anchor.load("rowIndex");
          await context.sync();
          const lastRow = anchor.rowIndex + data.length; // 1-based row of the template (last given row)
          const extra = Math.floor(Number(fill_down_to_row)) - lastRow;
          if (!(extra >= 0)) return { error: `fill_down_to_row (${fill_down_to_row}) must be at or below the last written row (${lastRow}).` };
          if (extra > MAX_FILL_ROWS) return { error: `fill_down_to_row would fill ${extra} rows; the limit is ${MAX_FILL_ROWS}.` };
          const template = data[data.length - 1];
          const filled = Array.from({ length: extra }, (_, k) => template.map((v) => (typeof v === "string" && v.startsWith("=") ? shiftFormulaRows(v, k + 1) : v)));
          data = [...data, ...filled];
        }
        const target = anchor.getCell(0, 0).getResizedRange(data.length - 1, data[0].length - 1);
        target.load("address,formulas,numberFormat");
        ws.load("name");
        await context.sync();
        ctx?.undo?.push({ sheet: ws.name, address: target.address.split("!").pop(), formulas: target.formulas });

        target.formulas = data;
        target.load("address,values,rowIndex,columnIndex");
        await context.sync();
        const out = { written: target.address, rows: data.length, columns: data[0].length };
        const preview = target.values.slice(0, 40).map((r) => r.slice(0, 20));
        out.result_values = preview;
        const errs = collectErrors(target.values, target.rowIndex, target.columnIndex);
        if (errs.length) out.errors = errs;
        return out;
      });
    },
  },

  clear_range: {
    risk: "destructive",
    schema: {
      description: "Clear a range's contents, formats, or both.",
      parameters: {
        type: "object",
        properties: {
          address: { type: "string" },
          sheet: { type: "string" },
          what: { type: "string", enum: ["all", "contents", "formats"], description: "Default 'contents'." },
        },
        required: ["address"],
      },
    },
    async run({ address, sheet, what = "contents" }, ctx) {
      return Excel.run(async (context) => {
        const { ws, range } = getRange(context, sheet, address);
        range.load("address,formulas");
        ws.load("name");
        await context.sync();
        if (what !== "formats") ctx?.undo?.push({ sheet: ws.name, address: range.address.split("!").pop(), formulas: range.formulas });
        const map = { all: "All", contents: "Contents", formats: "Formats" };
        range.clear(map[what] || "Contents");
        await context.sync();
        return { cleared: range.address, what };
      });
    },
  },

  format_range: {
    risk: "write",
    schema: {
      description:
        "Format a range: font, fill, number format, alignment, borders, column width / row height, wrap, merge. Only the properties you pass are changed.",
      parameters: {
        type: "object",
        properties: {
          address: { type: "string" },
          sheet: { type: "string" },
          bold: { type: "boolean" },
          italic: { type: "boolean" },
          underline: { type: "boolean" },
          font_size: { type: "number" },
          font_color: { type: "string", description: "Hex like '#1F4E79'." },
          fill_color: { type: "string", description: "Hex background color, or 'none' to remove." },
          number_format: {
            type: "string",
            description: "Excel number format, e.g. '#,##0.00', '0.0%', '$#,##0', 'yyyy-mm-dd'.",
          },
          horizontal_alignment: { type: "string", enum: ["Left", "Center", "Right", "General"] },
          vertical_alignment: { type: "string", enum: ["Top", "Center", "Bottom"] },
          wrap_text: { type: "boolean" },
          borders: { type: "string", enum: ["all", "outline", "bottom", "none"] },
          border_color: { type: "string", description: "Hex border color. Default '#BFBFBF'." },
          column_width: { type: "number", description: "Width in points." },
          row_height: { type: "number", description: "Height in points." },
          autofit: { type: "boolean", description: "Autofit column widths to content." },
          merge: { type: "boolean", description: "true to merge the cells, false to unmerge." },
        },
        required: ["address"],
      },
    },
    async run(a) {
      return Excel.run(async (context) => {
        const { range } = getRange(context, a.sheet, a.address);
        const f = range.format;
        if (a.bold !== undefined) f.font.bold = !!a.bold;
        if (a.italic !== undefined) f.font.italic = !!a.italic;
        if (a.underline !== undefined) f.font.underline = a.underline ? "Single" : "None";
        if (a.font_size) f.font.size = a.font_size;
        if (a.font_color) f.font.color = a.font_color;
        if (a.fill_color) a.fill_color === "none" ? f.fill.clear() : (f.fill.color = a.fill_color);
        if (a.horizontal_alignment) f.horizontalAlignment = a.horizontal_alignment;
        if (a.vertical_alignment) f.verticalAlignment = a.vertical_alignment;
        if (a.wrap_text !== undefined) f.wrapText = !!a.wrap_text;
        if (a.column_width) f.columnWidth = a.column_width;
        if (a.row_height) f.rowHeight = a.row_height;
        if (a.merge === true) range.merge(false);
        if (a.merge === false) range.unmerge();
        if (a.number_format) {
          range.load("rowCount,columnCount");
          await context.sync();
          range.numberFormat = Array.from({ length: range.rowCount }, () => Array(range.columnCount).fill(a.number_format));
        }
        if (a.borders) {
          const edges = {
            all: ["EdgeTop", "EdgeBottom", "EdgeLeft", "EdgeRight", "InsideHorizontal", "InsideVertical"],
            outline: ["EdgeTop", "EdgeBottom", "EdgeLeft", "EdgeRight"],
            bottom: ["EdgeBottom"],
            none: ["EdgeTop", "EdgeBottom", "EdgeLeft", "EdgeRight", "InsideHorizontal", "InsideVertical"],
          }[a.borders];
          for (const e of edges || []) {
            const b = f.borders.getItem(e);
            if (a.borders === "none") b.style = "None";
            else {
              b.style = "Continuous";
              b.color = a.border_color || "#BFBFBF";
            }
          }
        }
        if (a.autofit) f.autofitColumns();
        range.load("address");
        await context.sync();
        return { formatted: range.address };
      });
    },
  },

  manage_sheet: {
    risk: "write",
    schema: {
      description: "Add, rename, delete, copy or activate a worksheet.",
      parameters: {
        type: "object",
        properties: {
          action: { type: "string", enum: ["add", "rename", "delete", "copy", "activate"] },
          name: { type: "string", description: "Sheet to act on (or the new sheet's name for 'add')." },
          new_name: { type: "string", description: "New name for 'rename' or 'copy'." },
        },
        required: ["action", "name"],
      },
    },
    async run({ action, name, new_name }) {
      return Excel.run(async (context) => {
        const sheets = context.workbook.worksheets;
        let ws;
        switch (action) {
          case "add":
            ws = sheets.add(name);
            ws.activate();
            break;
          case "rename":
            ws = sheets.getItem(name);
            ws.name = new_name;
            break;
          case "delete":
            sheets.getItem(name).delete();
            await context.sync();
            return { deleted: name };
          case "copy":
            ws = sheets.getItem(name).copy("End");
            if (new_name) ws.name = new_name;
            ws.activate();
            break;
          case "activate":
            ws = sheets.getItem(name);
            ws.activate();
            break;
          default:
            throw new Error(`Unknown action ${action}`);
        }
        ws.load("name");
        await context.sync();
        return { action, sheet: ws.name };
      });
    },
  },

  insert_delete_cells: {
    risk: "write",
    schema: {
      description:
        "Insert or delete whole rows ('5:7'), whole columns ('C:D') or a block of cells ('B2:C4', shifting cells down/right on insert or up/left on delete).",
      parameters: {
        type: "object",
        properties: {
          action: { type: "string", enum: ["insert", "delete"] },
          address: { type: "string", description: "e.g. '5:7' for rows, 'C:D' for columns." },
          sheet: { type: "string" },
          shift: { type: "string", enum: ["down", "right", "up", "left"], description: "Only for cell blocks." },
        },
        required: ["action", "address"],
      },
    },
    async run({ action, address, sheet, shift }) {
      return Excel.run(async (context) => {
        const { range } = getRange(context, sheet, address);
        const addr = splitAddress(sheet, address).address;
        const isRows = /^\d+(:\d+)?$/.test(addr);
        const isCols = /^[A-Za-z]+(:[A-Za-z]+)?$/.test(addr);
        if (action === "insert") {
          range.insert(isRows ? "Down" : isCols ? "Right" : shift === "right" ? "Right" : "Down");
        } else {
          range.delete(isRows ? "Up" : isCols ? "Left" : shift === "left" ? "Left" : "Up");
        }
        await context.sync();
        return { action, address: addr };
      });
    },
  },

  create_table: {
    risk: "write",
    schema: {
      description: "Convert a range into an Excel table (with filter buttons and banded rows).",
      parameters: {
        type: "object",
        properties: {
          address: { type: "string" },
          sheet: { type: "string" },
          has_headers: { type: "boolean", description: "Default true." },
          name: { type: "string", description: "Table name, no spaces." },
          style: { type: "string", description: "e.g. 'TableStyleMedium2' (default), 'TableStyleLight9'." },
        },
        required: ["address"],
      },
    },
    async run({ address, sheet, has_headers = true, name, style }) {
      return Excel.run(async (context) => {
        const { ws } = getRange(context, sheet, address);
        const t = ws.tables.add(splitAddress(sheet, address).address, has_headers);
        if (name) t.name = name.replace(/\s+/g, "_");
        t.style = style || "TableStyleMedium2";
        t.load("name");
        const r = t.getRange();
        r.load("address");
        await context.sync();
        return { table: t.name, range: r.address };
      });
    },
  },

  create_chart: {
    risk: "write",
    schema: {
      description: "Create a chart from a data range (first row/column used as labels).",
      parameters: {
        type: "object",
        properties: {
          source: { type: "string", description: "Data range including headers, e.g. 'A1:C13'." },
          sheet: { type: "string" },
          type: {
            type: "string",
            enum: ["column", "bar", "line", "pie", "doughnut", "area", "scatter", "stacked_column", "combo_line_markers"],
          },
          title: { type: "string" },
          position: { type: "string", description: "Top-left cell to place the chart, e.g. 'H2'." },
          series_by: { type: "string", enum: ["auto", "columns", "rows"] },
        },
        required: ["source", "type"],
      },
    },
    async run({ source, sheet, type, title, position, series_by = "auto" }) {
      const types = {
        column: "ColumnClustered",
        bar: "BarClustered",
        line: "Line",
        pie: "Pie",
        doughnut: "Doughnut",
        area: "Area",
        scatter: "XYScatter",
        stacked_column: "ColumnStacked",
        combo_line_markers: "LineMarkers",
      };
      return Excel.run(async (context) => {
        const { ws, range } = getRange(context, sheet, source);
        const by = { auto: "Auto", columns: "Columns", rows: "Rows" }[series_by] || "Auto";
        const chart = ws.charts.add(types[type] || "ColumnClustered", range, by);
        if (title) {
          chart.title.text = title;
          chart.title.visible = true;
        }
        if (position) {
          const start = ws.getRange(splitAddress(null, position).address);
          chart.setPosition(start);
        }
        chart.load("name");
        await context.sync();
        return { chart: chart.name, type: types[type] };
      });
    },
  },

  sort_range: {
    risk: "write",
    schema: {
      description: "Sort a range by one or more columns.",
      parameters: {
        type: "object",
        properties: {
          address: { type: "string", description: "Whole data block including the header row if any." },
          sheet: { type: "string" },
          keys: {
            type: "array",
            description: "Sort keys in priority order.",
            items: {
              type: "object",
              properties: {
                column: { type: "integer", description: "0-based column offset inside the range." },
                ascending: { type: "boolean" },
              },
              required: ["column"],
            },
          },
          has_headers: { type: "boolean", description: "Default true." },
        },
        required: ["address", "keys"],
      },
    },
    async run({ address, sheet, keys, has_headers = true }) {
      return Excel.run(async (context) => {
        const { range } = getRange(context, sheet, address);
        const fields = (keys || []).map((k) => ({ key: Number(k.column) || 0, ascending: k.ascending !== false }));
        range.sort.apply(fields, false, has_headers, "Rows");
        range.load("address");
        await context.sync();
        return { sorted: range.address, keys: fields };
      });
    },
  },

  compare_lists: {
    risk: "write",
    schema: {
      description:
        "Reconcile or compare two lists by a key column, e.g. ledger vs bank statement, invoices vs payments, this month's list vs last month's. Does all the matching in one step and writes a discrepancy report to a new sheet: keys missing from either side, values that differ (both values plus the difference), and duplicate keys. The report's values are live XLOOKUP formulas into the source sheets. Use this for any list matching or reconciliation instead of comparing rows by reading them or building COUNTIF/XLOOKUP helper columns yourself. Both ranges need a header row. Afterwards, report the counts it returns.",
      parameters: {
        type: "object",
        properties: {
          left: { type: "string", description: "First list including its header row. Prefer whole columns, e.g. 'Ledger!A:C'; a range that stops early is extended to the end of the list." },
          right: { type: "string", description: "Second list including its header row, e.g. 'Bank!A:C'. Don't assume it has as many rows as the first list." },
          key: { type: "string", description: "Header name (or column letter) of the column to match on, e.g. 'Ref'. Defaults to the first column. Use right_key if the second list names it differently." },
          right_key: { type: "string", description: "Key column header in the second list, if different from `key`." },
          compare: {
            type: "array",
            items: { type: "string" },
            description: "Headers of the columns whose values must agree, e.g. ['Amount']. Defaults to every column with the same header on both sides (except the key).",
          },
          output_sheet: { type: "string", description: "Name for the new report sheet. Default 'Reconciliation'; a number is added if it already exists." },
          left_label: { type: "string", description: "Name for the first list in the report. Defaults to its sheet name." },
          right_label: { type: "string", description: "Name for the second list in the report. Defaults to its sheet name." },
          tolerance: { type: "number", description: "Largest numeric difference treated as equal. Default 0.005." },
        },
        required: ["left", "right"],
      },
    },
    async run(a, ctx) {
      return Excel.run(async (context) => {
        const readList = async (addr) => {
          const { ws, range: requested } = getRange(context, "", addr);
          const used = ws.getUsedRangeOrNullObject(true);
          ws.load("name");
          used.load("address");
          await context.sync();
          let range = used.isNullObject ? null : requested.getIntersectionOrNullObject(used);
          if (range) range.load("values,rowIndex,columnIndex,rowCount,columnCount");
          if (!used.isNullObject) used.load("rowIndex,rowCount");
          await context.sync();
          if (!range || range.isNullObject || range.rowCount < 2) throw new Error(`${addr} has no data rows under a header row.`);
          // A list is the whole block: if rows right below the given range still hold data in its columns
          // (e.g. a guessed 'Bank!A1:C61' when the bank has 62 rows), extend the range through them.
          let extended = null;
          const below = used.rowIndex + used.rowCount - (range.rowIndex + range.rowCount);
          if (below > 0) {
            const tail = range.getCell(range.rowCount - 1, 0).getOffsetRange(1, 0).getResizedRange(below - 1, range.columnCount - 1);
            tail.load("values");
            await context.sync();
            let more = 0;
            while (more < tail.values.length && tail.values[more].some((v) => v !== "" && v != null)) more++;
            if (more) {
              const from = range.address.split("!").pop();
              range = range.getResizedRange(more, 0);
              range.load("address,values,rowIndex,columnIndex,rowCount,columnCount");
              await context.sync();
              extended = `${ws.name}!${from} → ${range.address.split("!").pop()} (the list continues below the given range)`;
            }
          }
          if (range.rowCount * range.columnCount > 200000) throw new Error(`${addr} is too large to compare (over 200,000 cells).`);
          return { sheet: ws.name, values: range.values, rowIndex: range.rowIndex, columnIndex: range.columnIndex, rowCount: range.rowCount, extended };
        };
        const left = await readList(a.left);
        const right = await readList(a.right);
        const lh = left.values[0], rh = right.values[0];
        const lk = a.key ? findColumn(lh, a.key, left.columnIndex) : 0;
        const rk = a.right_key || a.key ? findColumn(rh, a.right_key || a.key, right.columnIndex) : 0;
        if (lk < 0) return { error: `Key column "${a.key}" isn't in the header row of ${a.left}. Headers: ${lh.join(", ")}` };
        if (rk < 0) return { error: `Key column "${a.right_key || a.key}" isn't in the header row of ${a.right}. Headers: ${rh.join(", ")}` };
        let names = Array.isArray(a.compare) && a.compare.length ? a.compare.map(String)
          : lh.filter((h, i) => i !== lk && String(h).trim() && findColumn(rh, h) > -1 && findColumn(rh, h) !== rk).map(String);
        const fields = [];
        for (const name of names) {
          const l = findColumn(lh, name, left.columnIndex), r = findColumn(rh, name, right.columnIndex);
          if (l < 0 || r < 0) return { error: `Column "${name}" isn't in both header rows. ${a.left}: ${lh.join(", ")}. ${a.right}: ${rh.join(", ")}` };
          fields.push({ name: String(lh[l]).trim() || name, l, r });
        }
        const res = compareRows(left.values.slice(1), right.values.slice(1), { lk, rk, fields, tolerance: Number.isFinite(Number(a.tolerance)) ? Number(a.tolerance) : 0.005 });
        let lName = a.left_label || left.sheet, rName = a.right_label || right.sheet;
        if (lName === rName) { lName = a.left_label || "List 1"; rName = a.right_label || "List 2"; }

        // New report sheet with a free name.
        const base = String(a.output_sheet || "Reconciliation").slice(0, 28);
        let outName = base;
        for (let n = 2; ; n++) {
          const probe = context.workbook.worksheets.getItemOrNullObject(outName);
          probe.load("name");
          await context.sync();
          if (probe.isNullObject) break;
          outName = `${base} ${n}`;
        }
        const out = context.workbook.worksheets.add(outName);

        // Live lookups into the source lists: ='Bank'!$C$2:$C$62 style absolute ranges.
        const ref = (list, col) => {
          const q = `'${list.sheet.replace(/'/g, "''")}'!`;
          const c = colLetter(list.columnIndex + col);
          return `${q}$${c}$${list.rowIndex + 2}:$${c}$${list.rowIndex + list.rowCount}`;
        };
        const fieldOf = (name) => fields.find((f) => f.name === name);
        const issues = res.issues.slice(0, 5000);
        const rows = [[String(lh[lk]).trim() || "Key", "Issue", "Field", `${lName} value`, `${rName} value`, `Difference (${rName} − ${lName})`]];
        issues.forEach((it, i) => {
          const r = i + 2, f = fieldOf(it.field);
          const look = (list, kc, vc) => `=XLOOKUP($A${r},${ref(list, kc)},${ref(list, vc)},"")`;
          rows.push([
            it.raw,
            issueText(it, lName, rName),
            it.field || "",
            f ? look(left, lk, f.l) : "",
            f ? look(right, rk, f.r) : "",
            f ? `=IF(AND(ISNUMBER(D${r}),ISNUMBER(E${r})),E${r}-D${r},"")` : "",
          ]);
        });
        if (!issues.length) rows.push(["", "No discrepancies", "", "", "", ""]);
        const counts = {
          [`missing from ${rName}`]: res.issues.filter((x) => x.type === "missing_right").length,
          [`missing from ${lName}`]: res.issues.filter((x) => x.type === "missing_left").length,
          value_mismatches: res.issues.filter((x) => x.type === "mismatch").length,
          [`duplicates in ${lName}`]: res.issues.filter((x) => x.type === "duplicate_left").length,
          [`duplicates in ${rName}`]: res.issues.filter((x) => x.type === "duplicate_right").length,
        };
        rows.push(["", "", "", "", "", ""]);
        rows.push(["Discrepancies", res.issues.length, "", "", "", ""]);
        rows.push(["Compared", `${res.left_keys} keys in ${lName}, ${res.right_keys} in ${rName}, ${res.matched_keys} on both`, "", "", "", ""]);
        const now = new Date(); // local date: toISOString() is UTC and can be a day ahead in the evening
        rows.push(["Checked", `${now.getFullYear()}-${String(now.getMonth() + 1).padStart(2, "0")}-${String(now.getDate()).padStart(2, "0")}`, "", "", "", ""]);

        const target = out.getRange("A1").getResizedRange(rows.length - 1, 5);
        target.load("address,formulas");
        await context.sync();
        ctx?.undo?.push({ sheet: outName, address: target.address.split("!").pop(), formulas: target.formulas });
        target.formulas = rows;
        const header = out.getRange("A1:F1");
        header.format.font.bold = true;
        header.format.fill.color = "#E7F0EC";
        if (issues.length) out.getRange(`D2:F${issues.length + 1}`).numberFormat = Array.from({ length: issues.length }, () => ["#,##0.00", "#,##0.00", "#,##0.00"]);
        out.getRange(`A${rows.length - 2}`).format.font.bold = true;
        target.format.autofitColumns();
        out.activate();
        target.load("address");
        await context.sync();
        return {
          written: target.address,
          report_sheet: outName,
          discrepancies: res.issues.length,
          counts,
          compared: { [lName]: res.left_keys, [rName]: res.right_keys, on_both: res.matched_keys, fields: fields.map((f) => f.name) },
          items: res.issues.slice(0, 40).map((x) => [x.raw, issueText(x, lName, rName), x.left, x.right]),
          ...([left.extended, right.extended].some(Boolean) ? { ranges_extended: [left.extended, right.extended].filter(Boolean) } : {}),
          ...(res.issues.length > 40 ? { note: `Showing the first 40 of ${res.issues.length}; all are on ${outName}.` } : {}),
        };
      });
    },
  },

  find_cells: {
    risk: "read",
    schema: {
      description: "Search the workbook (or one sheet) for cells containing text or a number. Returns up to 50 matching cell addresses with values.",
      parameters: {
        type: "object",
        properties: {
          query: { type: "string" },
          sheet: { type: "string", description: "Limit to one sheet." },
          match_case: { type: "boolean" },
          whole_cell: { type: "boolean", description: "Match the entire cell content." },
        },
        required: ["query"],
      },
    },
    async run({ query, sheet, match_case = false, whole_cell = false }) {
      return Excel.run(async (context) => {
        const sheets = context.workbook.worksheets;
        sheets.load("items/name");
        await context.sync();
        const targets = sheet ? sheets.items.filter((s) => s.name === sheet) : sheets.items;
        const found = targets.map((ws) => {
          const r = ws.findAllOrNullObject(String(query), { completeMatch: whole_cell, matchCase: match_case });
          r.load("address");
          r.areas.load("items/address,items/values");
          return { ws, r };
        });
        await context.sync();
        const matches = [];
        for (const { r } of found) {
          if (r.isNullObject) continue;
          for (const area of r.areas.items) {
            if (matches.length >= 50) break;
            matches.push({ address: area.address, value: area.values[0][0] });
          }
        }
        return { query, count: matches.length, matches };
      });
    },
  },

  select_range: {
    risk: "read",
    schema: {
      description: "Select (and scroll to) a range so the user can see it.",
      parameters: {
        type: "object",
        properties: { address: { type: "string" }, sheet: { type: "string" } },
        required: ["address"],
      },
    },
    async run({ address, sheet }) {
      return Excel.run(async (context) => {
        const { ws, range } = getRange(context, sheet, address);
        ws.activate();
        range.select();
        range.load("address");
        await context.sync();
        return { selected: range.address };
      });
    },
  },

  read_document: {
    risk: "read",
    schema: {
      description:
        "Read a file the user attached in the panel — PDF, Word (.docx/.doc/.rtf/.odt) or spreadsheet (.xlsx/.xls/.xlsm/.csv/.ods) — identified by doc_id from the '[Attached …]' note in their message. Tables come back as | rows, headings as # lines. PDFs: text is marked '=== Page N of M ===', use page=N to jump to a page and as_image=true to SEE a page (charts, scans). Word: the first read returns an outline of headings with start offsets; pass start=<offset> to jump. Spreadsheets: shown as '| row | A | B | … |' with formulas in ⟨⟩; use sheet=<name> to jump to a sheet; to copy data into this workbook use import_from_file, not write_range. Presentations (.pptx/.odp): '=== Slide N of M: title ===' with bullets, tables, '[Chart …]' data tables (the chart's underlying numbers) and speaker notes; use page=N to jump to slide N. Use start=next_start to continue reading.",
      parameters: {
        type: "object",
        properties: {
          doc_id: { type: "string", description: "The doc_id from the attachment note." },
          page: { type: "integer", description: "PDFs: 1-based page to jump to. Presentations: slide number." },
          sheet: { type: "string", description: "Spreadsheets: sheet name to jump to." },
          start: { type: "integer", description: "Character offset (next_start) to continue reading." },
          as_image: { type: "boolean", description: "Also attach an image of the page." },
          max_chars: { type: "integer", description: "1000-20000, default 8000." },
        },
        required: ["doc_id"],
      },
    },
    async run({ doc_id, page, sheet, start, as_image, max_chars }) {
      return postJson("/docs/read", { doc_id, page, sheet, start, as_image, max_chars });
    },
  },

  import_from_file: {
    risk: "write",
    schema: {
      description:
        "Bring data from an attached (or web-fetched) spreadsheet file into this workbook exactly, without retyping. mode='copy_sheet' inserts a whole sheet as a new worksheet (for .xlsx/.xlsm keeps formatting, column widths and formulas). mode='values' writes a sheet or range (source_range) into a target location, keeping number formats (%, currency, dates); include_formulas=true copies formulas instead of values. ALWAYS use this instead of write_range when moving more than a handful of cells from a file.",
      parameters: {
        type: "object",
        properties: {
          doc_id: { type: "string", description: "doc_id of the attached/fetched spreadsheet." },
          sheet: { type: "string", description: "Source sheet name (default: first sheet)." },
          mode: { type: "string", enum: ["copy_sheet", "values"], description: "Default 'values'." },
          source_range: { type: "string", description: "values mode: e.g. 'A1:F200' (default: whole used range)." },
          target_address: { type: "string", description: "values mode: top-left destination cell, e.g. 'A1' or 'Data!B3'." },
          target_sheet: { type: "string", description: "values mode: destination sheet (default: active). Created if it doesn't exist." },
          include_formulas: { type: "boolean", description: "values mode: copy formulas instead of their results." },
          new_sheet_name: { type: "string", description: "copy_sheet mode: name for the inserted sheet." },
        },
        required: ["doc_id"],
      },
    },
    async run(a, ctx) {
      return a.mode === "copy_sheet" ? copySheetFromFile(a) : importValues(a, ctx);
    },
  },

  web_search: {
    risk: "read",
    web: true,
    schema: {
      description:
        "Search the web (DuckDuckGo). Returns titles, URLs and snippets. Use for current facts the workbook doesn't contain: exchange rates, prices, company figures, definitions, statistics. Then use web_fetch on the best result to read the actual numbers.",
      parameters: {
        type: "object",
        properties: {
          query: { type: "string", description: "Short search query. Never include data copied from the workbook." },
          max_results: { type: "integer", description: "1-10, default 6." },
        },
        required: ["query"],
      },
    },
    async run({ query, max_results }) {
      return postJson("/web/search", { query, max_results });
    },
  },

  web_fetch: {
    risk: "read",
    web: true,
    schema: {
      description:
        "Read a public web page or PDF as plain text (table columns kept as | separated rows). Long documents are paged: pass the returned next_start as start to continue. For PDFs, text is marked '=== Page N of M ===', use page=N to jump straight to a page (e.g. after reading the table of contents), and as_image=true to SEE a page (charts, complex tables, scanned PDFs; scanned PDFs return an image automatically). Web pages also return up to 25 links.",
      parameters: {
        type: "object",
        properties: {
          url: { type: "string", description: "Full http(s) URL, e.g. from web_search results." },
          start: { type: "integer", description: "Character offset to continue a long page." },
          page: { type: "integer", description: "PDFs only: 1-based page number to jump to." },
          as_image: { type: "boolean", description: "PDFs only: also attach an image of the page (requires page or defaults to page 1)." },
          max_chars: { type: "integer", description: "1000-20000, default 8000." },
        },
        required: ["url"],
      },
    },
    async run({ url, start, page, as_image, max_chars }) {
      return postJson("/web/fetch", { url, start, page, as_image, max_chars });
    },
  },

  execute_office_js: {
    risk: "destructive",
    schema: {
      description:
        "Escape hatch for anything the other tools can't do (conditional formatting, data validation, freeze panes, pivot-like summaries, hiding columns, etc.). The code is the BODY of an async function with `context` (Excel.RequestContext) and `Excel` in scope; it must `await context.sync()` before reading loaded properties and may `return` a JSON-serializable result. Example: `const ws = context.workbook.worksheets.getActiveWorksheet(); ws.freezePanes.freezeRows(1); await context.sync(); return 'ok';`",
      parameters: {
        type: "object",
        properties: {
          code: { type: "string", description: "JavaScript function body using the Office.js Excel API." },
          explanation: { type: "string", description: "One sentence for the user on what this code does." },
        },
        required: ["code"],
      },
    },
    async run({ code }) {
      const AsyncFunction = Object.getPrototypeOf(async function () {}).constructor;
      const fn = new AsyncFunction("context", "Excel", code);
      return Excel.run(async (context) => {
        const result = await fn(context, Excel);
        await context.sync();
        return { result: result === undefined ? "done" : result };
      });
    },
  },
};

// ---------- importing from attached spreadsheets ----------
async function copySheetFromFile({ doc_id, sheet, new_sheet_name }) {
  const file = await postJson("/docs/file", { doc_id });
  if (file.error) return file;
  const src = sheet ? file.sheet_names.find((n) => n.toLowerCase() === String(sheet).toLowerCase()) : file.sheet_names[0];
  if (!src) return { error: `No sheet named "${sheet}". Sheets: ${file.sheet_names.join(", ")}` };
  return Excel.run(async (context) => {
    const before = context.workbook.worksheets;
    before.load("items/name");
    await context.sync();
    const existing = new Set(before.items.map((w) => w.name.toLowerCase()));
    const ids = context.workbook.insertWorksheetsFromBase64(file.base64, {
      sheetNamesToInsert: [src],
      positionType: Excel.WorksheetPositionType.end,
    });
    await context.sync();
    const ws = context.workbook.worksheets.getItem(ids.value[0]);
    ws.load("name");
    await context.sync();
    let wanted = new_sheet_name || (existing.has(src.toLowerCase()) ? null : src);
    if (wanted && wanted !== ws.name) {
      // Excel renames on clashes ("Sales (2)"); apply the requested name when it's free.
      if (!existing.has(wanted.toLowerCase())) ws.name = wanted.slice(0, 31);
    }
    ws.activate();
    const used = ws.getUsedRangeOrNullObject();
    used.load("address");
    ws.load("name");
    await context.sync();
    return {
      inserted_sheet: ws.name,
      from: `${file.name} › ${src}`,
      used_range: used.isNullObject ? null : used.address,
      formatting: file.formatting,
    };
  });
}

async function importValues({ doc_id, sheet, source_range, target_address, target_sheet, include_formulas }, ctx) {
  const data = await postJson("/docs/values", { doc_id, sheet, range: source_range, include_formulas: !!include_formulas });
  if (data.error) return data;
  const { values, formats, rows, cols } = data;
  const p = splitAddress(target_sheet, target_address || "A1");
  return Excel.run(async (context) => {
    let ws;
    if (p.sheet) {
      ws = context.workbook.worksheets.getItemOrNullObject(p.sheet);
      await context.sync();
      if (ws.isNullObject) ws = context.workbook.worksheets.add(p.sheet);
    } else ws = context.workbook.worksheets.getActiveWorksheet();
    ws.load("name");
    const anchor = ws.getRange(p.address || "A1").getCell(0, 0);
    const whole = anchor.getResizedRange(rows - 1, cols - 1);
    whole.load("address");
    // Snapshot for "Undo cell changes" on moderate imports.
    let snapshot = null;
    if (rows * cols <= 50000) {
      whole.load("formulas");
      snapshot = true;
    }
    await context.sync();
    if (snapshot) ctx?.undo?.push({ sheet: ws.name, address: whole.address.split("!").pop(), formulas: whole.formulas });

    const chunkRows = Math.max(1, Math.floor(20000 / cols));
    for (let r0 = 0; r0 < rows; r0 += chunkRows) {
      const n = Math.min(chunkRows, rows - r0);
      const rg = anchor.getOffsetRange(r0, 0).getResizedRange(n - 1, cols - 1);
      const part = values.slice(r0, r0 + n);
      if (include_formulas) rg.formulas = part;
      else rg.values = part;
      if (formats) rg.numberFormat = formats.slice(r0, r0 + n);
      await context.sync();
    }
    whole.format.autofitColumns();
    whole.load("values,rowIndex,columnIndex");
    await context.sync();
    const out = {
      imported: whole.address,
      from: `${data.name} › ${data.sheet}!${data.address}`,
      rows,
      columns: cols,
      preview: whole.values.slice(0, 5).map((r) => r.slice(0, 12)),
    };
    if (!snapshot) out.note = "Large import: 'Undo cell changes' isn't available for this one.";
    const errs = collectErrors(whole.values.slice(0, 2000), whole.rowIndex, whole.columnIndex);
    if (errs.length) out.errors = errs;
    return out;
  });
}

async function postJson(path, body) {
  const base = globalThis.ADDIN_BASE || "";
  const r = await fetch(base + path, {
    method: "POST",
    headers: { "Content-Type": "application/json" },
    body: JSON.stringify(body),
  });
  return r.json();
}

// ---------- public API ----------
export function getToolDefinitions({ web = true } = {}) {
  return Object.entries(tools)
    .filter(([, t]) => web || !t.web)
    .map(([name, t]) => ({ type: "function", function: { name, ...t.schema } }));
}
export const toolDefinitions = getToolDefinitions();

export function isWebTool(name) {
  return Boolean(tools[name]?.web);
}

export function toolRisk(name) {
  return tools[name]?.risk || "destructive";
}

export async function runTool(name, args, ctx) {
  const tool = tools[name];
  if (!tool) return { error: `Unknown tool '${name}'. Available: ${Object.keys(tools).join(", ")}` };
  try {
    const result = await tool.run(args || {}, ctx);
    return result;
  } catch (e) {
    const info = e?.debugInfo ? { code: e.code, location: e.debugInfo.errorLocation, statement: e.debugInfo.statement } : undefined;
    return { error: e?.message || String(e), ...(info ? { debug: info } : {}) };
  }
}

export function serializeResult(result) {
  let s = JSON.stringify(result);
  if (s === undefined) s = "null";
  if (s.length > MAX_RESULT_CHARS) s = s.slice(0, MAX_RESULT_CHARS) + `… [truncated ${s.length - MAX_RESULT_CHARS} chars — request a smaller range]`;
  return s;
}

// Restores snapshots captured by write_range / clear_range (values & formulas only).
export async function undoSnapshots(snapshots) {
  return Excel.run(async (context) => {
    for (const snap of [...snapshots].reverse()) {
      const ws = context.workbook.worksheets.getItem(snap.sheet);
      ws.getRange(snap.address).formulas = snap.formulas;
    }
    await context.sync();
  });
}

// Workbook context line attached to each user message (cheap, keeps prompt cache warm).
export async function workbookContext() {
  try {
    return await Excel.run(async (context) => {
      const ws = context.workbook.worksheets.getActiveWorksheet();
      ws.load("name");
      const used = ws.getUsedRangeOrNullObject(true);
      used.load("address");
      const sheets = context.workbook.worksheets;
      sheets.load("items/name");
      await context.sync();
      const out = {
        active_sheet: ws.name,
        sheets: sheets.items.map((s) => s.name),
        selection: "(a chart or shape is selected, not cells)",
        selection_size: "0x0",
        used_range: used.isNullObject ? "(empty sheet)" : used.address,
      };
      try {
        // Throws when a chart/shape is selected instead of cells.
        const sel = context.workbook.getSelectedRange();
        sel.load("address,rowCount,columnCount");
        await context.sync();
        out.selection = sel.address;
        out.selection_size = `${sel.rowCount}x${sel.columnCount}`;
      } catch {
        /* keep the placeholder */
      }
      return out;
    });
  } catch {
    return null;
  }
}
