// Offline checks for PDF handling (no network): column cleanup, scanned-PDF detection, page rendering.
//   node scripts/test-pdf.mjs <scanned.pdf> <text.pdf>
import fs from "node:fs";
import assert from "node:assert";
import { _internals } from "../web.js";

const { extractPdf, pdfResult, cleanPdfPage } = _internals;

assert.equal(
  cleanPdfPage("   Insurance premiums earned        $     88,257   $   83,403\n  Total revenues      371,433"),
  "Insurance premiums earned | $88,257 | $83,403\nTotal revenues | 371,433"
);
assert.equal(cleanPdfPage("Risk Factors . . . . . . . . . . K-24"), "Risk Factors … K-24");
console.log("✓ layout cleanup");

const [scannedPath, textPath] = process.argv.slice(2);
const opts = { maxChars: 2000, start: 0 };

const scanned = await extractPdf(fs.readFileSync(scannedPath), "https://example.org/scan.pdf");
assert.ok(scanned.scanned, "image-only PDF should be detected as scanned");
const s = await pdfResult(scanned, "https://example.org/scan.pdf", opts);
assert.ok(s._images?.[0]?.length > 1000, "scanned PDF should return a page image");
assert.equal(s.content, undefined);
console.log(`✓ scanned PDF -> ${s.pages} page(s), image ${Math.round(s._images[0].length / 1024)} KB base64, note: ${s.note}`);

const text = await extractPdf(fs.readFileSync(textPath), "https://example.org/doc.pdf");
assert.ok(!text.scanned, "text PDF should not be flagged as scanned");
const t = await pdfResult(text, "https://example.org/doc.pdf", { ...opts, page: 1, asImage: true });
assert.ok(t.content.startsWith("=== Page 1 of"), "page jump should start at the page marker");
assert.ok(t._images?.length === 1, "as_image should attach a rendering");
console.log(`✓ text PDF -> ${t.pages} page(s), ${t.total_chars} chars, page 1 + image attached`);

for (const p of [scanned, text]) fs.unlinkSync(p.file);
console.log("all PDF checks passed");
