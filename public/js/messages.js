// Builds the user message exactly as the model sees it. Shared by the panel (app.js) and the
// benchmark (scripts/bench), so benchmark scores reflect what the real app sends.
import { ATTACH_MARK } from "./chats.js";

// One-line workbook context from tools.workbookContext().
export function contextLine(c, inExcel = true) {
  if (!inExcel) return "not running inside Excel (preview mode) — workbook tools are unavailable";
  if (!c) return "";
  return `sheet "${c.active_sheet}" · selection ${c.selection} (${c.selection_size}) · used ${c.used_range} · sheets: ${c.sheets.join(", ")}`;
}

// Note for one attached document, from the /docs/upload response.
export function attachmentNote(d) {
  const head =
    d.kind === "slides"
      ? `[Attached presentation: ${d.name} — doc_id ${d.doc_id}, ${d.format?.toUpperCase()}, ${d.slides} slides${d.charts ? `, ${d.charts} charts with data` : ""}${d.notes ? `, speaker notes on ${d.notes} slides` : ""}. Slide titles: ${(d.titles || []).map((t, i) => `${i + 1}. ${t}`).join("; ")}]`
      : d.kind === "sheet"
      ? `[Attached spreadsheet: ${d.name} — doc_id ${d.doc_id}, ${d.format?.toUpperCase()}, sheets: ${(d.sheets || [])
          .map((s) => `"${s.name}" ${s.range || "(empty)"}${s.rows ? ` (${s.rows} rows${s.formulas ? `, ${s.formulas} formulas` : ""})` : ""}`)
          .join("; ")}]`
      : d.kind === "word"
      ? `[Attached Word document: ${d.name} — doc_id ${d.doc_id}, ${d.format?.toUpperCase()}, ${d.words} words]`
      : `[Attached PDF: ${d.name} — doc_id ${d.doc_id}, ${d.pages} page${d.pages === 1 ? "" : "s"}${
          d.scanned ? ", scanned (no text layer: use read_document to see page images)" : ""
        }${d.title && d.title !== d.name ? `, title "${d.title}"` : ""}]`;
  const how = d.kind === "sheet" ? "(Use read_document to read it; use import_from_file to bring data into this workbook.)" : "(Use read_document to read it.)";
  return d.full_text ? `${head}\nFull text:\n${d.full_text}${d.kind === "sheet" ? `\n${how}` : ""}` : `${head}\n${how}`;
}

// Full stored content of a user message.
export function userContent({ ctxLine = "", note = "", text, docs = [] }) {
  const ctx = ctxLine ? `[Excel context: ${ctxLine}]\n\n` : "";
  const n = note ? `[Note: ${note}]\n\n` : "";
  const att = docs.length ? ATTACH_MARK + docs.map(attachmentNote).join("\n\n") : "";
  return `${ctx}${n}${text}${att}`;
}
