// Offline checks for Word handling.
//   node scripts/test-word.mjs [file.docx file.doc file.rtf ...]
import fs from "node:fs";
import path from "node:path";
import assert from "node:assert";
import { docxXmlToText, detectDocKind, _internals } from "../web.js";

const P = (text, style, list) =>
  `<w:p><w:pPr>${style ? `<w:pStyle w:val="${style}"/>` : ""}${list ? '<w:numPr><w:ilvl w:val="0"/></w:numPr>' : ""}</w:pPr><w:r><w:t xml:space="preserve">${text}</w:t></w:r></w:p>`;
const cell = (inner) => `<w:tc><w:tcPr/>${inner}</w:tc>`;
const row = (...cells) => `<w:tr>${cells.map((c) => cell(P(c))).join("")}</w:tr>`;

const xml = `<?xml version="1.0"?><w:document><w:body>
${P("Q3 Budget Memo", "Title")}
${P("Resumen", "Ttulo1")}
<w:p><w:r><w:t>Amounts in Bs. </w:t></w:r><w:r><w:delText>DELETED</w:delText></w:r><w:r><w:instrText> PAGE </w:instrText></w:r><w:r><w:t>R&amp;D &lt;draft&gt;</w:t></w:r></w:p>
${P("Department budgets", "Heading2")}
<w:tbl>${row("Department", "Budget", "Actual")}${row("Marketing", "120,000", "(11,500)")}
<w:tr>${cell(P("IT"))}${cell(`<w:tbl>${row("a", "b")}</w:tbl>`)}${cell(P("97,200"))}</w:tr></w:tbl>
${P("Marketing overspent.", null, true)}
${P("Operations | saved", null, true)}
<w:p/>
${P("Acción requerida")}
</w:body></w:document>`;

const text = docxXmlToText(xml);
console.log(text, "\n");
assert.match(text, /^# Q3 Budget Memo/m, "Title style -> heading");
assert.match(text, /^# Resumen/m, "Spanish heading style id (Ttulo1) -> heading");
assert.match(text, /^## Department budgets/m);
assert.match(text, /Amounts in Bs\. R&D <draft>/, "entities decoded, deleted text and field codes skipped");
assert.ok(!/DELETED|PAGE/.test(text));
assert.match(text, /^\| Department \| Budget \| Actual \|$/m, "table header row");
assert.match(text, /^\| Marketing \| 120,000 \| \(11,500\) \|$/m);
assert.match(text, /^\| IT \| a, b \| 97,200 \|$/m, "nested table flattened into its cell");
assert.match(text, /^- Marketing overspent\.$/m, "numbered paragraph -> list item");
assert.match(text, /^- Operations \| saved$/m);
assert.match(text, /Acción requerida/);
console.log("✓ docx XML conversion");

assert.equal(detectDocKind(Buffer.from("%PDF-1.7"), "x.pdf"), "pdf");
assert.equal(detectDocKind(Buffer.from("{\\rtf1\\ansi"), "x.rtf"), "rtf");
assert.equal(detectDocKind(Buffer.from([0xd0, 0xcf, 0x11, 0xe0, 0, 0, 0, 0]), "x.doc"), "doc");
assert.equal(detectDocKind(Buffer.from("hello world"), "x.docx"), null);
console.log("✓ format detection");

for (const f of process.argv.slice(2)) {
  const buf = fs.readFileSync(f);
  const kind = detectDocKind(buf, f);
  const w = await _internals.extractWord(buf, kind, path.basename(f));
  console.log(`\n=== ${path.basename(f)} (${kind}, ${w.words} words)\n${w.text}`);
  assert.ok(w.text.includes("Marketing"), `${f}: content missing`);
  assert.match(w.text, /\| Marketing \| 120,000 \| 131,500 \| \(11,500\) \|/, `${f}: table row not preserved`);
  assert.ok(w.text.includes("Acción"), `${f}: accents broken`);
}
console.log("\nall Word checks passed");
