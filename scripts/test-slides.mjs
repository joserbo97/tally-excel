// Offline checks for presentation handling.
//   node scripts/test-slides.mjs [deck.pptx ...]
import fs from "node:fs";
import path from "node:path";
import assert from "node:assert";
import { chartToTable, slideXmlToText, detectDocKind, _internals } from "../web.js";

// --- chart XML -> table (bar chart, 2 series, % format on one)
const ser = (name, cats, vals, fmt) => `<c:ser><c:tx><c:strRef><c:strCache><c:pt idx="0"><c:v>${name}</c:v></c:pt></c:strCache></c:strRef></c:tx>
<c:dLbls><c:dLbl><c:idx val="9"/></c:dLbl></c:dLbls>
<c:cat><c:strRef><c:strCache>${cats.map((c, i) => `<c:pt idx="${i}"><c:v>${c}</c:v></c:pt>`).join("")}</c:strCache></c:strRef></c:cat>
<c:val><c:numRef><c:numCache>${fmt ? `<c:formatCode>${fmt}</c:formatCode>` : ""}${vals.map((v, i) => `<c:pt idx="${i}"><c:v>${v}</c:v></c:pt>`).join("")}</c:numCache></c:numRef></c:val></c:ser>`;
const chart = `<c:chartSpace><c:chart><c:title><c:tx><c:rich><a:p><a:r><a:t>Revenue &amp; Growth</a:t></a:r></a:p></c:rich></c:tx></c:title>
<c:plotArea><c:barChart>${ser("Revenue", ["North", "South"], [125000, 98000])}${ser("Growth", ["North", "South"], [0.336, 0.408], "0.0%")}</c:barChart></c:plotArea></c:chart></c:chartSpace>`;
const t = chartToTable(chart);
console.log(t, "\n");
assert.equal(t, "[Chart: Revenue & Growth (bar)]\n| Category | Revenue | Growth [percent] |\n| North | 125000 | 0.336 |\n| South | 98000 | 0.408 |");
assert.equal(chartToTable("<c:chartSpace><c:plotArea><c:pieChart></c:pieChart></c:plotArea></c:chartSpace>"), "", "empty chart skipped");
// Unnamed series whose data label has its own rich-text <c:tx>: name must not be taken from the categories.
const unnamed = `<c:plotArea><c:pieChart><c:ser><c:idx val="0"/><c:order val="0"/>
<c:dLbls><c:dLbl><c:idx val="0"/><c:tx><c:rich><a:p><a:r><a:t>label</a:t></a:r></a:p></c:rich></c:tx></c:dLbl></c:dLbls>
<c:cat><c:strRef><c:strCache><c:pt idx="0"><c:v>Taxes</c:v></c:pt></c:strCache></c:strRef></c:cat>
<c:val><c:numRef><c:numCache><c:pt idx="0"><c:v>5</c:v></c:pt></c:numCache></c:numRef></c:val></c:ser></c:pieChart></c:plotArea>`;
assert.match(chartToTable(unnamed), /\| Category \| Series 1 \|\n\| Taxes \| 5 \|/, "unnamed series stays 'Series 1'");
// Series named with rich text instead of a cell reference.
assert.match(chartToTable(unnamed.replace("<c:order val=\"0\"/>", '<c:order val="0"/><c:tx><c:rich><a:p><a:r><a:t>Budget</a:t></a:r></a:p></c:rich></c:tx>')), /\| Category \| Budget \|/);
console.log("✓ chart data extraction");

// --- slide XML -> text
const sp = (ph, paras) => `<p:sp><p:nvSpPr><p:cNvPr id="2" name="x"/><p:nvPr>${ph ? `<p:ph ${ph}/>` : ""}</p:nvPr></p:nvSpPr><p:txBody>${paras}</p:txBody></p:sp>`;
const p = (txt, lvl, bu) => `<a:p>${lvl || bu ? `<a:pPr${lvl ? ` lvl="${lvl}"` : ""}>${bu || ""}</a:pPr>` : ""}<a:r><a:t>${txt}</a:t></a:r></a:p>`;
const slide = `<p:sld><p:cSld><p:spTree>
${sp('type="title"', p("Q3 Results"))}
${sp('idx="1"', p("Revenue up 12%") + p("North led", 1) + p("Plain line", 0, "<a:buNone/>"))}
${sp(null, p("Text box") + p("Starred", 0, '<a:buChar char="•"/>'))}
${sp('type="sldNum"', p("7"))}
<p:pic><p:nvPicPr><p:cNvPr id="9" name="Picture 8" descr="Map of regions"/></p:nvPicPr></p:pic>
<p:graphicFrame><a:graphic><a:graphicData><a:tbl><a:tr><a:tc><a:txBody>${p("Region")}</a:txBody></a:tc><a:tc><a:txBody>${p("Q3")}</a:txBody></a:tc></a:tr>
<a:tr><a:tc><a:txBody>${p("North | East")}</a:txBody></a:tc><a:tc><a:txBody>${p("1,200")}</a:txBody></a:tc></a:tr></a:tbl></a:graphicData></a:graphic></p:graphicFrame>
<p:graphicFrame><a:graphic><a:graphicData><c:chart r:id="rId7"/></a:graphicData></a:graphic></p:graphicFrame>
</p:spTree></p:cSld></p:sld>`;
const s = slideXmlToText(slide, (rid) => (rid === "rId7" ? "[Chart (pie)]\n| Category | V |" : ""));
console.log(s.title, "\n" + s.text, "\n");
assert.equal(s.title, "Q3 Results");
assert.match(s.text, /^- Revenue up 12%$/m, "body placeholder -> bullet");
assert.match(s.text, /^ {2}- North led$/m, "indent level");
assert.match(s.text, /^Plain line$/m, "buNone -> no bullet");
assert.match(s.text, /^Text box$/m, "text box -> plain");
assert.match(s.text, /^- Starred$/m, "explicit bullet char");
assert.ok(!/^7$/m.test(s.text), "slide number skipped");
assert.match(s.text, /^\[Picture: Map of regions\]$/m);
assert.match(s.text, /^\| Region \| Q3 \|\n\| North \/ East \| 1,200 \|$/m, "table with escaped pipe");
assert.match(s.text, /^\[Chart \(pie\)\]/m, "chart inserted where it sits");
console.log("✓ slide text, bullets, pictures, tables, charts");

assert.equal(detectDocKind(Buffer.from([0xd0, 0xcf, 0x11, 0xe0, 0, 0, 0, 0]), "old.ppt"), "ppt");
console.log("✓ legacy .ppt detected (rejected with a Save As hint)");

// --- real decks
for (const f of process.argv.slice(2)) {
  const buf = fs.readFileSync(f);
  const kind = detectDocKind(buf, f);
  if (!kind) {
    console.log(`- ${path.basename(f)}: not a presentation (skipped)`);
    continue;
  }
  const d = await _internals.extractSlides(buf, kind, path.basename(f));
  const jump = _internals.slidesResult(d, { maxChars: 400, start: 0, page: Math.min(3, d.pages) });
  assert.match(jump.content, new RegExp(`^=== Slide ${Math.min(3, d.pages)} of ${d.pages}`));
  console.log(`✓ ${path.basename(f)}: ${d.pages} slides, ${d.charts} charts, ${d.withNotes} with notes, ${d.text.length} chars; slide jump OK`);
}
console.log("\nall presentation checks passed");
