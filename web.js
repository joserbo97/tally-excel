// Web search (DuckDuckGo, no API key) and page reading for the agent.
// Only search queries and fetched URLs leave the Mac. Page fetches are restricted to
// public internet addresses so a malicious page can't steer the agent at localhost,
// Ollama, or devices on the LAN (checked at connect time, on every redirect hop).
import http from "node:http";
import https from "node:https";
import dns from "node:dns";
import net from "node:net";
import zlib from "node:zlib";
import fs from "node:fs";
import os from "node:os";
import path from "node:path";
import crypto from "node:crypto";
import { execFile } from "node:child_process";
import { promisify } from "node:util";
import { parse } from "node-html-parser";
import XLSX from "xlsx";

const run = promisify(execFile);
const UA = "Mozilla/5.0 (Macintosh; Intel Mac OS X 10_15_7) AppleWebKit/605.1.15 (KHTML, like Gecko) Version/17.0 Safari/605.1.15";
const MAX_BYTES = 40 * 1024 * 1024; // annual reports and statistical PDFs can be large
const TIMEOUT_MS = 15000;

// ---------- address safety ----------
// Separate lists: a single BlockList also matches IPv4 addresses against IPv4-mapped
// IPv6 rules, which would block every IPv4 address.
const blocked4 = new net.BlockList();
const blocked6 = new net.BlockList();
for (const [net4, bits] of [
  ["0.0.0.0", 8], ["10.0.0.0", 8], ["100.64.0.0", 10], ["127.0.0.0", 8], ["169.254.0.0", 16], ["172.16.0.0", 12],
  ["192.0.0.0", 24], ["192.168.0.0", 16], ["198.18.0.0", 15], ["224.0.0.0", 3],
]) blocked4.addSubnet(net4, bits, "ipv4");
for (const [net6, bits] of [
  ["::", 96], // unspecified, loopback, IPv4-compatible
  ["::ffff:0:0", 96], // IPv4-mapped (also re-checked as IPv4 below)
  ["64:ff9b::", 96], ["64:ff9b:1::", 48], // NAT64
  ["fc00::", 7], ["fe80::", 10], ["fec0::", 10], ["ff00::", 8],
]) blocked6.addSubnet(net6, bits, "ipv6");

// Expand an IPv6 literal to its 8 hextets (handles "::" and a dotted IPv4 tail).
function hextets(ip) {
  let s = ip.toLowerCase().split("%")[0];
  const v4 = /(\d+\.\d+\.\d+\.\d+)$/.exec(s);
  if (v4) {
    const [a, b, c, d] = v4[1].split(".").map(Number);
    s = s.slice(0, -v4[1].length) + `${((a << 8) | b).toString(16)}:${((c << 8) | d).toString(16)}`;
  }
  const [head, tail] = s.split("::");
  const h = head ? head.split(":") : [];
  const t = tail !== undefined ? (tail ? tail.split(":") : []) : [];
  const fill = tail !== undefined ? Array(8 - h.length - t.length).fill("0") : [];
  return [...h, ...fill, ...t].map((x) => parseInt(x || "0", 16));
}

function isPublicIp(ip) {
  if (net.isIPv4(ip)) return !blocked4.check(ip, "ipv4");
  if (!net.isIPv6(ip)) return false;
  return !blocked6.check(ip, "ipv6");
}

// Embedded IPv4 in mapped/compatible/NAT64 forms is covered by the /96 blocks above;
// this extra check guards against unusual spellings slipping past BlockList parsing.
function embeddedV4IsPrivate(ip) {
  if (!net.isIPv6(ip)) return false;
  const h = hextets(ip);
  const v4 = `${h[6] >> 8}.${h[6] & 255}.${h[7] >> 8}.${h[7] & 255}`;
  const prefixZero = h.slice(0, 5).every((x) => x === 0);
  return prefixZero && (h[5] === 0 || h[5] === 0xffff) && blocked4.check(v4, "ipv4");
}

function safeLookup(hostname, options, cb) {
  dns.lookup(hostname, { ...options, all: true }, (err, addrs) => {
    if (err) return cb(err);
    const bad = addrs.find((a) => !isPublicIp(a.address) || embeddedV4IsPrivate(a.address));
    if (bad) return cb(new Error(`Blocked: ${hostname} resolves to a private/local address (${bad.address}).`));
    if (options.all) return cb(null, addrs);
    cb(null, addrs[0].address, addrs[0].family);
  });
}

// ---------- HTTP ----------
function request(url, { method = "GET", body, headers = {}, redirects = 5 } = {}) {
  return new Promise((resolve, reject) => {
    let u;
    try {
      u = new URL(url);
    } catch {
      return reject(new Error(`Invalid URL: ${url}`));
    }
    if (!/^https?:$/.test(u.protocol)) return reject(new Error("Only http and https URLs are allowed."));
    if (u.username || u.password) return reject(new Error("URLs with credentials are not allowed."));
    // Node skips `lookup` for literal IPs, so check those here.
    const host = u.hostname.replace(/^\[|\]$/g, "");
    if (net.isIP(host) && (!isPublicIp(host) || embeddedV4IsPrivate(host))) return reject(new Error(`Blocked: ${host} is a private/local address.`));
    const lib = u.protocol === "https:" ? https : http;
    const req = lib.request(
      u,
      {
        method,
        lookup: safeLookup,
        timeout: TIMEOUT_MS,
        headers: {
          "User-Agent": UA,
          Accept: "text/html,application/xhtml+xml,application/json,text/plain;q=0.9,*/*;q=0.5",
          "Accept-Language": "en-US,en;q=0.9,es;q=0.8",
          "Accept-Encoding": "gzip, deflate, br",
          ...headers,
        },
      },
      (res) => {
        if ([301, 302, 303, 307, 308].includes(res.statusCode) && res.headers.location) {
          res.resume();
          if (redirects <= 0) return reject(new Error("Too many redirects."));
          const next = new URL(res.headers.location, u).toString();
          const keep = res.statusCode === 307 || res.statusCode === 308;
          return resolve(request(next, { method: keep ? method : "GET", body: keep ? body : undefined, headers, redirects: redirects - 1 }));
        }
        let stream = res;
        const enc = res.headers["content-encoding"];
        if (enc === "gzip") stream = res.pipe(zlib.createGunzip());
        else if (enc === "deflate") stream = res.pipe(zlib.createInflate());
        else if (enc === "br") stream = res.pipe(zlib.createBrotliDecompress());
        const chunks = [];
        let size = 0;
        stream.on("data", (d) => {
          size += d.length;
          if (size > MAX_BYTES) {
            req.destroy();
            resolve({ status: res.statusCode, headers: res.headers, url: u.toString(), body: Buffer.concat(chunks), truncated: true });
          } else chunks.push(d);
        });
        stream.on("end", () => resolve({ status: res.statusCode, headers: res.headers, url: u.toString(), body: Buffer.concat(chunks) }));
        stream.on("error", reject);
      }
    );
    req.on("timeout", () => req.destroy(new Error(`Timed out after ${TIMEOUT_MS / 1000}s.`)));
    req.on("error", reject);
    if (body) req.write(body);
    req.end();
  });
}

function decode(buf, contentType = "") {
  const m = /charset=([\w-]+)/i.exec(contentType);
  try {
    return new TextDecoder(m ? m[1] : "utf-8").decode(buf);
  } catch {
    return new TextDecoder("utf-8").decode(buf);
  }
}

// ---------- search ----------
export async function webSearch(query, max = 6) {
  query = String(query || "").trim().slice(0, 300);
  if (!query) throw new Error("Empty query.");
  const res = await request("https://html.duckduckgo.com/html/", {
    method: "POST",
    body: new URLSearchParams({ q: query, kl: "wt-wt" }).toString(),
    headers: { "Content-Type": "application/x-www-form-urlencoded" },
  });
  const html = decode(res.body, res.headers["content-type"]);
  if (res.status !== 200 || /anomaly|unusual traffic/i.test(html.slice(0, 5000))) {
    throw new Error(`Search provider refused the request (HTTP ${res.status}). Try again in a minute.`);
  }
  const root = parse(html);
  const results = [];
  for (const r of root.querySelectorAll(".result")) {
    if (r.classList.contains("result--ad")) continue;
    const a = r.querySelector("a.result__a");
    if (!a) continue;
    let href = a.getAttribute("href") || "";
    // Older DDG markup wraps links in a redirect: //duckduckgo.com/l/?uddg=<encoded>
    const uddg = /[?&]uddg=([^&]+)/.exec(href);
    if (uddg) href = decodeURIComponent(uddg[1]);
    if (!/^https?:\/\//.test(href)) continue;
    results.push({
      title: a.text.trim(),
      url: href,
      snippet: (r.querySelector(".result__snippet")?.text || "").replace(/\s+/g, " ").trim(),
    });
    if (results.length >= Math.min(10, Math.max(1, max))) break;
  }
  return { query, results };
}

// ---------- page reading ----------
const DROP = "script,style,noscript,svg,iframe,form,nav,footer,header,aside,button,select,template,canvas,dialog";
const BLOCK = new Set(["p", "div", "section", "article", "main", "br", "ul", "ol", "dl", "dt", "dd", "blockquote", "pre", "figure", "figcaption", "hr", "address"]);

function htmlToText(root, baseUrl) {
  root.querySelectorAll(DROP).forEach((n) => n.remove());
  const out = [];
  const push = (s) => out.push(s);
  const walk = (node) => {
    if (node.nodeType === 3) return push(node.text.replace(/\s+/g, " "));
    if (node.nodeType !== 1) return;
    const tag = node.rawTagName?.toLowerCase();
    if (tag === "table") {
      push("\n");
      for (const tr of node.querySelectorAll("tr").slice(0, 400)) {
        const cells = tr.querySelectorAll("th,td").map((c) => c.text.replace(/\s+/g, " ").trim());
        if (cells.some(Boolean)) push("| " + cells.join(" | ") + " |\n");
      }
      return push("\n");
    }
    if (/^h[1-6]$/.test(tag)) return push(`\n\n${"#".repeat(Math.min(3, Number(tag[1])))} ${node.text.replace(/\s+/g, " ").trim()}\n`);
    if (tag === "li") push("\n- ");
    else if (BLOCK.has(tag) || tag === "tr") push("\n");
    node.childNodes.forEach(walk);
    if (BLOCK.has(tag)) push("\n");
  };
  walk(root);
  return out
    .join("")
    .replace(/[ \t]+\n/g, "\n")
    .replace(/\n[ \t]+/g, "\n")
    .replace(/\n{3,}/g, "\n\n")
    .replace(/[ \t]{2,}/g, " ")
    .trim();
}

// ---------- PDF (Poppler: pdftotext / pdfinfo / pdftoppm) ----------
const POPPLER = process.env.POPPLER_BIN || ["/opt/homebrew/bin", "/usr/local/bin"].find((d) => fs.existsSync(`${d}/pdftotext`));
const pdfCache = new Map(); // url -> { file, title, pages, text, pageStarts, scanned, at }
const PDF_TTL = 20 * 60 * 1000;

function cachePdf(url, entry) {
  // Map.set on an existing key keeps its old (oldest) position, so the fresh entry could be evicted
  // right away. Delete first so it becomes the newest, and clean up the stale entry's file.
  const prev = pdfCache.get(url);
  if (prev) {
    pdfCache.delete(url);
    if (prev.file !== entry.file) fs.unlink(prev.file, () => {});
  }
  pdfCache.set(url, entry);
  for (const [k, v] of pdfCache) {
    if (k === url) break; // never evict the entry just stored (it's last in order)
    if (pdfCache.size <= 6 && Date.now() - v.at < PDF_TTL) break;
    pdfCache.delete(k);
    fs.unlink(v.file, () => {});
  }
}

// Layout text -> compact rows: indentation dropped, column gaps (3+ spaces) become " | ".
function cleanPdfPage(raw) {
  return raw
    .split("\n")
    .map((l) =>
      l
        .replace(/\s+$/, "")
        .replace(/^\s+/, "")
        .replace(/(?:\s*\.){4,}\s*/g, " … ") // table-of-contents dot leaders
        .replace(/ {3,}/g, " | ")
        .replace(/([$€£¥]|Bs\.?|US\$) \| (?=[\d(])/g, "$1") // "$ | 88,257" -> "$88,257" keeps columns aligned
    )
    .join("\n")
    .replace(/\n{3,}/g, "\n\n")
    .trim();
}

async function extractPdf(buf, url) {
  if (!POPPLER) throw new Error("PDF support needs Poppler. Install it with: brew install poppler");
  // Unique per extraction: two fetches of the same URL (or a re-fetch after expiry) must not share a file.
  const file = path.join(os.tmpdir(), `oxa-pdf-${crypto.createHash("sha1").update(url).digest("hex").slice(0, 10)}-${crypto.randomBytes(4).toString("hex")}.pdf`);
  fs.writeFileSync(file, buf);
  let title = "";
  let pages = 0;
  try {
    const { stdout } = await run(`${POPPLER}/pdfinfo`, [file], { timeout: 20000 });
    title = /^Title:\s*(.+)$/m.exec(stdout)?.[1]?.trim() || "";
    pages = Number(/^Pages:\s*(\d+)/m.exec(stdout)?.[1] || 0);
  } catch (e) {
    fs.unlink(file, () => {});
    throw new Error(`Not a readable PDF (${String(e.stderr || e.message).trim().split("\n")[0]}).`);
  }
  const { stdout } = await run(`${POPPLER}/pdftotext`, ["-layout", "-enc", "UTF-8", "-q", file, "-"], {
    timeout: 90000,
    maxBuffer: 128 * 1024 * 1024,
  });
  const rawPages = stdout.split("\f");
  if (rawPages.length > pages) rawPages.length = Math.max(pages, 1);
  let text = "";
  const pageStarts = [];
  let letters = 0;
  rawPages.forEach((p, i) => {
    if (i) text += "\n\n";
    pageStarts.push(text.length);
    const clean = cleanPdfPage(p);
    letters += (clean.match(/[A-Za-zÀ-ÿ0-9]/g) || []).length;
    text += `=== Page ${i + 1} of ${pages} ===\n${clean}`;
  });
  // Scanned PDFs have (almost) no text layer.
  const scanned = letters < 40 * Math.max(1, pages);
  // Many PDFs carry junk metadata titles ("printmgr file", "Microsoft Word - x.docx").
  if (!title || title.length < 4 || /^(printmgr|microsoft (word|excel|powerpoint)|untitled|document\d*$|\d+$)/i.test(title)) {
    title = decodeURIComponent(new URL(url).pathname.split("/").pop() || "document.pdf");
  }
  return { file, title, pages, text, pageStarts, scanned, at: Date.now() };
}

async function renderPdfPage(pdf, page) {
  const prefix = path.join(os.tmpdir(), `oxa-page-${process.pid}-${Date.now()}`);
  await run(
    `${POPPLER}/pdftoppm`,
    ["-f", String(page), "-l", String(page), "-jpeg", "-jpegopt", "quality=85", "-scale-to", "1400", "-singlefile", pdf.file, prefix],
    { timeout: 60000 }
  );
  const jpg = fs.readFileSync(`${prefix}.jpg`);
  fs.unlink(`${prefix}.jpg`, () => {});
  return jpg.toString("base64");
}

async function pdfResult(pdf, url, { maxChars, start, page, asImage }) {
  if (page) {
    page = Math.min(Math.max(1, Math.round(page)), pdf.pages || 1);
    start = pdf.pageStarts[page - 1] ?? 0;
  }
  const out = { url, type: "pdf", title: pdf.title, pages: pdf.pages };
  if (pdf.scanned || asImage) {
    // Show the page itself to the vision model (scans, charts, complex layouts).
    const p = page || 1;
    out.page = p;
    out._images = [await renderPdfPage(pdf, p)];
    out.note = pdf.scanned
      ? `This PDF is scanned (no text layer). An image of page ${p} is attached; read it visually. Use page=N for other pages.`
      : `An image of page ${p} is attached.`;
    if (pdf.scanned) return out;
  }
  out.content = pdf.text.slice(start, start + maxChars);
  out.total_chars = pdf.text.length;
  if (start + maxChars < pdf.text.length) out.next_start = start + maxChars;
  if (!page && start === 0 && pdf.pages > 3) {
    out.tip = `Long PDF (${pdf.pages} pages). Use page=N to jump to a page (e.g. from a table of contents), or next_start to continue.`;
  }
  return out;
}

// ---------- Word documents (.docx parsed directly; .doc/.rtf/.odt via macOS textutil) ----------
const XML_ENT = { amp: "&", lt: "<", gt: ">", quot: '"', apos: "'" };
const unxml = (s) =>
  s.replace(/&(#x[0-9a-f]+|#\d+|\w+);/gi, (m, e) =>
    e[0] === "#" ? String.fromCodePoint(e[1].toLowerCase() === "x" ? parseInt(e.slice(2), 16) : Number(e.slice(1))) : XML_ENT[e] ?? m
  );

// WordprocessingML -> text with "# " headings, "- " list items and "| a | b |" table rows.
// textutil flattens .docx tables to one cell per line, so tables are rebuilt from w:tbl here.
export function docxXmlToText(xml) {
  const body = xml.slice(Math.max(0, xml.indexOf("<w:body")));
  const out = [];
  const tables = []; // stack of { rows: [[cell text]], row: [], cell: [] }
  let para = "";
  let paraPrefix = "";
  let inText = false;
  let skip = 0; // inside deleted text / field instructions
  const rx = /<(\/?)(w:[A-Za-z]+)\b([^>]*?)(\/?)>|([^<]+)/g;
  let m;
  while ((m = rx.exec(body))) {
    const [, close, tag, attrs, selfClose, txt] = m;
    if (txt !== undefined) {
      if (inText && !skip) para += unxml(txt);
      continue;
    }
    const t = tables[tables.length - 1];
    if (!close) {
      if (tag === "w:t") inText = !selfClose;
      else if (tag === "w:delText" || tag === "w:instrText") skip += selfClose ? 0 : 1;
      else if (tag === "w:tab") para += "\t";
      else if (tag === "w:br" || tag === "w:cr") para += t ? " " : "\n";
      else if (tag === "w:p") { para = ""; paraPrefix = ""; }
      else if (tag === "w:pStyle") {
        const v = /w:val="([^"]+)"/.exec(attrs)?.[1] || "";
        const h = /^(?:Heading|Ttulo|Titre|berschrift)(\d)$/i.exec(v) || (/^Title$/i.test(v) ? [0, "1"] : null);
        if (h) paraPrefix = "#".repeat(Math.min(3, Number(h[1]))) + " ";
      } else if (tag === "w:numPr" && !paraPrefix) paraPrefix = "- ";
      else if (tag === "w:tbl") tables.push({ rows: [], row: [], cell: [] });
      else if (tag === "w:tr" && t) t.row = [];
      else if (tag === "w:tc" && t) t.cell = [];
      if (selfClose && tag === "w:p") { if (t) t.cell.push(""); else out.push(""); }
    } else {
      if (tag === "w:t") inText = false;
      else if (tag === "w:delText" || tag === "w:instrText") skip = Math.max(0, skip - 1);
      else if (tag === "w:p") {
        const line = para.replace(/[ \t]+/g, " ").trim();
        if (t) t.cell.push(line);
        else out.push(line ? paraPrefix + line : "");
      } else if (tag === "w:tc" && t) t.row.push(t.cell.filter(Boolean).join(" ").replace(/\|/g, "/"));
      else if (tag === "w:tr" && t) { if (t.row.some(Boolean)) t.rows.push(t.row); }
      else if (tag === "w:tbl" && t) {
        tables.pop();
        const parent = tables[tables.length - 1];
        if (parent) parent.cell.push(t.rows.map((r) => r.filter(Boolean).join(", ")).join("; ")); // nested table: "a, b; c, d"
        else out.push("", t.rows.map((r) => `| ${r.join(" | ")} |`).join("\n"), "");
      }
    }
  }
  return out.join("\n").replace(/\n{3,}/g, "\n\n").trim();
}

const WORD_EXT = { docx: "docx", doc: "doc", rtf: "rtf", odt: "odt" };

export const SHEET_KINDS = new Set(["xlsx", "xlsm", "xlsb", "xls", "ods", "csv"]);

export function detectDocKind(buf, name = "") {
  const ext = (/\.(\w+)$/.exec(name)?.[1] || "").toLowerCase();
  if (buf.length < 8) return ["csv", "tsv"].includes(ext) ? "csv" : null;
  const head = buf.subarray(0, 8);
  if (head.subarray(0, 5).toString() === "%PDF-") return "pdf";
  if (head.subarray(0, 5).toString() === "{\\rtf") return "rtf";
  if (head.readUInt32BE(0) === 0xd0cf11e0) return ext === "xls" ? "xls" : ["ppt", "pps", "pot"].includes(ext) ? "ppt" : "doc"; // legacy OLE
  if (head.readUInt32BE(0) === 0x504b0304) {
    if (["xlsx", "xlsm", "xlsb", "ods", "odt", "docx", "pptx", "odp", "key"].includes(ext)) return ext;
    if (["ppsx", "potx", "pptm"].includes(ext)) return "pptx";
    const s = buf.toString("latin1", 0, Math.min(buf.length, 8192));
    if (s.includes("opendocument.spreadsheet")) return "ods";
    if (s.includes("opendocument.presentation")) return "odp";
    if (s.includes("opendocument.text")) return "odt";
    if (s.includes("xl/")) return "xlsx";
    if (s.includes("ppt/")) return "pptx";
    if (s.includes("word/")) return "docx";
  }
  if (["csv", "tsv"].includes(ext)) return "csv";
  return null;
}

// ---------- presentations (.pptx parsed directly; .odp basic) ----------
export const SLIDE_KINDS = new Set(["pptx", "odp"]);

function relsMap(xml) {
  const m = {};
  for (const r of (xml || "").matchAll(/<Relationship\b[^>]*?Id="([^"]+)"[^>]*?Target="([^"]+)"/g)) m[r[1]] = r[2];
  for (const r of (xml || "").matchAll(/<Relationship\b[^>]*?Target="([^"]+)"[^>]*?Id="([^"]+)"/g)) m[r[2]] ??= r[1];
  return m;
}

const resolvePath = (base, target) => path.posix.normalize(path.posix.join(path.posix.dirname(base), target));
const tdText = (s) => unxml(s.replace(/<[^>]+>/g, "")).replace(/\s+/g, " ").replace(/\|/g, "/").trim();

// Chart XML (DrawingML) -> data table from the cached category/value points.
export function chartToTable(xml) {
  const titleXml = /<c:title>([\s\S]*?)<\/c:title>/.exec(xml)?.[1] || "";
  const title = /<c:autoTitleDeleted val="1"/.test(xml) ? "" : [...titleXml.matchAll(/<a:t>([^<]*)<\/a:t>/g)].map((m) => unxml(m[1])).join("").trim();
  const type = /<c:plotArea>[\s\S]*?<c:(\w+?)Chart>/.exec(xml)?.[1] || "chart";
  const pts = (block) => {
    const out = [];
    for (const m of (block || "").matchAll(/<c:pt idx="(\d+)"[^>]*>\s*<c:v>([^<]*)<\/c:v>/g)) out[Number(m[1])] = unxml(m[2]);
    return out;
  };
  const series = [];
  for (const s of xml.matchAll(/<c:ser>([\s\S]*?)<\/c:ser>/g)) {
    const body = s[1];
    // The series name is the <c:tx> that's a direct child of <c:ser> (right after idx/order).
    // Data labels have their own <c:tx>, so an unanchored match could run into the categories.
    const txBlock = /^\s*(?:<c:idx\b[^>]*\/>\s*)?(?:<c:order\b[^>]*\/>\s*)?<c:tx>([\s\S]*?)<\/c:tx>/.exec(body)?.[1] || "";
    const txText = /<c:v>([^<]*)<\/c:v>/.exec(txBlock)?.[1] ?? [...txBlock.matchAll(/<a:t>([^<]*)<\/a:t>/g)].map((x) => x[1]).join("");
    let name = unxml(txText || "").trim() || `Series ${series.length + 1}`;
    const cats = pts(/<c:(?:cat|xVal)>([\s\S]*?)<\/c:(?:cat|xVal)>/.exec(body)?.[1]);
    const valBlock = /<c:(?:val|yVal)>([\s\S]*?)<\/c:(?:val|yVal)>/.exec(body)?.[1];
    const vals = pts(valBlock);
    const fmt = unxml(/<c:formatCode>([^<]*)<\/c:formatCode>/.exec(valBlock || "")?.[1] || "");
    // Only the hints that change how to read the numbers: 0.34 with a % format means 34%.
    if (/%/.test(fmt)) name += " [percent]";
    else if (/[$€£¥]|Bs/.test(fmt)) name += " [currency]";
    series.push({ name, cats, vals });
  }
  if (!series.length) return "";
  const cats = series.find((s) => s.cats.length)?.cats || [];
  const n = Math.max(cats.length, ...series.map((s) => s.vals.length));
  const rows = [`| Category | ${series.map((s) => s.name.replace(/\|/g, "/")).join(" | ")} |`];
  for (let i = 0; i < n; i++) rows.push(`| ${(cats[i] ?? String(i + 1)).replace(/\|/g, "/")} | ${series.map((s) => s.vals[i] ?? "").join(" | ")} |`);
  return `[Chart${title ? `: ${title}` : ""} (${type})]\n${rows.join("\n")}`;
}

// One slide's XML -> { title, text } with bullets, tables and charts in reading order.
export function slideXmlToText(xml, chartFor = () => "") {
  const lines = [];
  let title = "";
  const shapes = []; // stack of { ph: null | { type } }
  const tables = [];
  let para = null;
  let inText = false;
  let inPic = false;
  const SKIP_PH = new Set(["sldNum", "dt", "ftr"]); // slide number, date, footer: noise on every slide
  const rx = /<(\/?)((?:p|a|c):[A-Za-z]+)\b([^>]*?)(\/?)>|([^<]+)/g;
  let m;
  while ((m = rx.exec(xml))) {
    const [, close, tag, attrs, selfClose, txt] = m;
    if (txt !== undefined) {
      if (inText && para) para.text += unxml(txt);
      continue;
    }
    const shape = shapes[shapes.length - 1];
    const tbl = tables[tables.length - 1];
    if (!close) {
      if (tag === "p:sp") shapes.push({ ph: null });
      else if (tag === "p:pic") inPic = true;
      else if (tag === "p:cNvPr" && inPic) {
        const alt = unxml(/descr="([^"]*)"/.exec(attrs)?.[1] || "").replace(/\s+/g, " ").trim();
        lines.push(`[Picture${alt ? `: ${alt.slice(0, 200)}` : ""}]`);
        inPic = false;
      } else if (tag === "p:ph" && shape) shape.ph = { type: /type="(\w+)"/.exec(attrs)?.[1] || "body" };
      else if (tag === "a:p") para = { text: "", lvl: 0, bullet: null };
      else if (tag === "a:pPr" && para) para.lvl = Number(/lvl="(\d+)"/.exec(attrs)?.[1] || 0);
      else if ((tag === "a:buChar" || tag === "a:buAutoNum") && para) para.bullet = true;
      else if (tag === "a:buNone" && para) para.bullet = false;
      else if (tag === "a:t") inText = !selfClose;
      else if (tag === "a:br" && para) para.text += tbl ? " " : "\n";
      else if (tag === "a:tab" && para) para.text += " ";
      else if (tag === "a:tbl") tables.push({ rows: [], row: [], cell: [] });
      else if (tag === "a:tr" && tbl) tbl.row = [];
      else if (tag === "a:tc" && tbl) tbl.cell = [];
      else if (tag === "c:chart") {
        const rid = /r:id="([^"]+)"/.exec(attrs)?.[1];
        const t = rid && chartFor(rid);
        if (t) lines.push("", t, "");
      }
    } else {
      if (tag === "a:t") inText = false;
      else if (tag === "a:p" && para) {
        const t = para.text.replace(/[ \t]+/g, " ").trim();
        if (tbl) tbl.cell.push(t);
        else if (t && !SKIP_PH.has(shape?.ph?.type)) {
          const ph = shape?.ph?.type;
          if ((ph === "title" || ph === "ctrTitle") && !title) title = t.replace(/\n/g, " ");
          else {
            const isBody = ph === "body" || ph === "obj";
            const bullet = para.bullet ?? isBody;
            lines.push(`${"  ".repeat(para.lvl)}${bullet ? "- " : ""}${t}`);
          }
        }
        para = null;
      } else if (tag === "a:tc" && tbl) tbl.row.push(tbl.cell.filter(Boolean).join(" ").replace(/\|/g, "/"));
      else if (tag === "a:tr" && tbl) { if (tbl.row.some(Boolean)) tbl.rows.push(tbl.row); }
      else if (tag === "a:tbl" && tbl) {
        tables.pop();
        if (tbl.rows.length) lines.push("", tbl.rows.map((r) => `| ${r.join(" | ")} |`).join("\n"), "");
      } else if (tag === "p:sp") shapes.pop();
    }
  }
  return { title, text: lines.join("\n").replace(/\n{3,}/g, "\n\n").trim() };
}

// Speaker notes: only the notes body placeholder (not the slide image or slide number).
function notesText(xml) {
  const parts = [];
  for (const sp of (xml || "").matchAll(/<p:sp>([\s\S]*?)<\/p:sp>/g)) {
    if (!/<p:ph [^>]*type="body"/.test(sp[1])) continue;
    for (const p of sp[1].matchAll(/<a:p>([\s\S]*?)<\/a:p>/g)) {
      const t = [...p[1].matchAll(/<a:t>([^<]*)<\/a:t>/g)].map((x) => unxml(x[1])).join("").trim();
      if (t) parts.push(t);
    }
  }
  return parts.join("\n");
}

async function pptxSlides(buf) {
  const dir = fs.mkdtempSync(path.join(os.tmpdir(), "oxa-pptx-"));
  const file = path.join(dir, "deck.pptx");
  fs.writeFileSync(file, buf);
  try {
    await run("/usr/bin/unzip", ["-qq", "-o", file, "ppt/*", "-d", path.join(dir, "x")], { timeout: 60000, maxBuffer: 16 * 1024 * 1024 });
    const root = path.join(dir, "x");
    const rd = (p) => {
      const full = path.join(root, p);
      return full.startsWith(root) && fs.existsSync(full) ? fs.readFileSync(full, "utf8") : "";
    };
    const presRels = relsMap(rd("ppt/_rels/presentation.xml.rels"));
    const order = [...rd("ppt/presentation.xml").matchAll(/<p:sldId\b[^>]*r:id="([^"]+)"/g)].map((x) => presRels[x[1]]).filter(Boolean);
    const slides = [];
    for (const target of order) {
      const sp = resolvePath("ppt/presentation.xml", target);
      const rels = relsMap(rd(path.posix.join(path.posix.dirname(sp), "_rels", path.posix.basename(sp) + ".rels")));
      const chartFor = (rid) => (rels[rid] ? chartToTable(rd(resolvePath(sp, rels[rid]))) : "");
      const { title, text } = slideXmlToText(rd(sp), chartFor);
      const notesTarget = Object.values(rels).find((t) => /notesSlide\d*\.xml$/.test(t));
      const notes = notesTarget ? notesText(rd(resolvePath(sp, notesTarget))) : "";
      const hidden = /<p:sld\b[^>]*show="0"/.test(rd(sp).slice(0, 2000));
      slides.push({ title, text, notes, hidden });
    }
    return slides;
  } finally {
    fs.rm(dir, { recursive: true, force: true }, () => {});
  }
}

async function odpSlides(buf) {
  const dir = fs.mkdtempSync(path.join(os.tmpdir(), "oxa-odp-"));
  const file = path.join(dir, "deck.odp");
  fs.writeFileSync(file, buf);
  try {
    const { stdout } = await run("/usr/bin/unzip", ["-p", file, "content.xml"], { maxBuffer: 64 * 1024 * 1024, encoding: "utf8", timeout: 30000 });
    return [...stdout.matchAll(/<draw:page\b([^>]*)>([\s\S]*?)<\/draw:page>/g)].map(([, attrs, body]) => {
      const notesXml = /<presentation:notes>([\s\S]*?)<\/presentation:notes>/.exec(body)?.[1] || "";
      const main = body.replace(/<presentation:notes>[\s\S]*?<\/presentation:notes>/, "");
      const rows = [...main.matchAll(/<table:table-row\b[^>]*>([\s\S]*?)<\/table:table-row>/g)].map(
        (r) => `| ${[...r[1].matchAll(/<table:table-cell\b[^>]*>([\s\S]*?)<\/table:table-cell>/g)].map((c) => tdText(c[1])).join(" | ")} |`
      );
      const noTables = main.replace(/<table:table\b[\s\S]*?<\/table:table>/g, "");
      const paras = [...noTables.matchAll(/<text:(?:p|h)\b[^>]*>([\s\S]*?)<\/text:(?:p|h)>/g)].map((p) => tdText(p[1])).filter(Boolean);
      const titled = /presentation:class="title"[\s\S]*?<text:p\b[^>]*>([\s\S]*?)<\/text:p>/.exec(main);
      const title = titled ? tdText(titled[1]) : unxml(/draw:name="([^"]*)"/.exec(attrs)?.[1] || "");
      const text = [...paras.filter((p) => p !== title), ...(rows.length ? ["", ...rows] : [])].join("\n");
      const notes = [...notesXml.matchAll(/<text:p\b[^>]*>([\s\S]*?)<\/text:p>/g)].map((p) => tdText(p[1])).filter(Boolean).join("\n");
      return { title, text, notes, hidden: false };
    });
  } finally {
    fs.rm(dir, { recursive: true, force: true }, () => {});
  }
}

async function extractSlides(buf, kind, name) {
  let slides;
  try {
    slides = kind === "odp" ? await odpSlides(buf) : await pptxSlides(buf);
  } catch (e) {
    throw new Error(`Couldn't read this ${kind.toUpperCase()} file (${String(e.stderr || e.message).trim().split("\n")[0]}).`);
  }
  if (!slides.length) throw new Error("No slides found in this presentation.");
  let text = "";
  const pageStarts = [];
  let charts = 0, tables = 0, withNotes = 0;
  slides.forEach((s, i) => {
    if (i) text += "\n\n";
    pageStarts.push(text.length);
    text += `=== Slide ${i + 1} of ${slides.length}${s.title ? `: ${s.title}` : ""}${s.hidden ? " (hidden)" : ""} ===\n`;
    if (s.text) text += s.text + "\n";
    if (s.notes) {
      text += `Speaker notes: ${s.notes}\n`;
      withNotes++;
    }
    charts += (s.text.match(/^\[Chart/gm) || []).length;
    tables += (s.text.match(/^\| /gm) ? 1 : 0);
  });
  return {
    kind: "slides", format: kind, title: name, text: text.trimEnd(), pageStarts, pages: slides.length,
    outline: slides.map((s, i) => ({ slide: i + 1, title: s.title || "(untitled)" })),
    charts, tables, withNotes, at: Date.now(),
  };
}

function slidesResult(entry, { maxChars, start, page }) {
  if (page) start = entry.pageStarts[Math.min(Math.max(1, Math.round(page)), entry.pages) - 1] ?? 0;
  const out = { type: entry.format, title: entry.title, slides: entry.pages };
  out.content = entry.text.slice(start, start + maxChars);
  out.total_chars = entry.text.length;
  if (start + maxChars < entry.text.length) out.next_start = start + maxChars;
  if (!page && start === 0 && entry.text.length > maxChars) {
    out.outline = entry.outline;
    out.tip = "Use page=N to jump to slide N. Chart data appears as [Chart …] tables you can write into the sheet.";
  }
  return out;
}

// ---------- spreadsheets (SheetJS) ----------
const MAX_TEXT_ROWS = 5000; // per sheet in the text view; import_from_file handles any size
const MAX_TEXT_COLS = 60;

// SheetJS 0.20 returns Excel date serials as Dates at UTC midnight (verified: serial 46037 ->
// 2026-01-15T00:00Z), so read the UTC parts — local parts would shift dates a day west of UTC.
const pad2 = (n) => String(n).padStart(2, "0");
function localDateString(d) {
  const day = `${d.getUTCFullYear()}-${pad2(d.getUTCMonth() + 1)}-${pad2(d.getUTCDate())}`;
  const h = d.getUTCHours(), m = d.getUTCMinutes(), s = d.getUTCSeconds();
  return h || m || s ? `${day} ${pad2(h)}:${pad2(m)}:${pad2(s)}` : day;
}

function cellText(cell) {
  if (!cell) return "";
  let t = cell.w ?? (cell.v instanceof Date ? localDateString(cell.v) : String(cell.v ?? ""));
  t = t.replace(/\s+/g, " ").replace(/\|/g, "/").trim();
  if (t.length > 80) t = t.slice(0, 77) + "…";
  if (cell.f) t += ` ⟨=${cell.f.slice(0, 80)}⟩`;
  return t;
}

function extractSheet(buf, kind, name) {
  let wb;
  try {
    wb = XLSX.read(buf, { type: "buffer", cellFormula: true, cellDates: true, cellNF: true, cellText: true, FS: kind === "csv" && /\.tsv$/i.test(name) ? "\t" : undefined });
  } catch (e) {
    throw new Error(`Couldn't read this ${kind.toUpperCase()} file (${e.message}).`);
  }
  let text = "";
  let cells = 0; // for the docs store's memory estimate
  const sheets = [];
  for (const sn of wb.SheetNames) {
    const ws = wb.Sheets[sn];
    // Files often declare a range far bigger than the data (formatted-but-empty cells). Use the real extent.
    let minR = Infinity, minC = Infinity, maxR = -1, maxC = -1;
    for (const k of Object.keys(ws)) {
      if (k[0] === "!") continue;
      cells++;
      const c = ws[k];
      if (c.f || (c.v !== undefined && c.v !== null && c.v !== "")) {
        const a = XLSX.utils.decode_cell(k);
        if (a.r < minR) minR = a.r;
        if (a.c < minC) minC = a.c;
        if (a.r > maxR) maxR = a.r;
        if (a.c > maxC) maxC = a.c;
      }
    }
    // Keep A1 as the origin when the data starts near it, so row/column numbers match Excel.
    const ref = maxR < 0 ? null : XLSX.utils.encode_range({ s: { r: minR <= 10 ? 0 : minR, c: minC <= 10 ? 0 : minC }, e: { r: maxR, c: maxC } });
    ws["!used"] = ref;
    const info = { name: sn, ref: ref || null, rows: 0, cols: 0, formulas: 0, start: text.length };
    if (!ref) {
      text += `${text ? "\n\n" : ""}=== Sheet "${sn}" — empty ===\n`;
      sheets.push(info);
      continue;
    }
    const r = XLSX.utils.decode_range(ref);
    info.rows = r.e.r - r.s.r + 1;
    info.cols = r.e.c - r.s.c + 1;
    if (text) text += "\n\n";
    info.start = text.length;
    text += `=== Sheet "${sn}" — ${ref} (${info.rows} rows × ${info.cols} cols) ===\n`;
    const lastC = Math.min(r.e.c, r.s.c + MAX_TEXT_COLS - 1);
    const colHdr = [];
    for (let c = r.s.c; c <= lastC; c++) colHdr.push(XLSX.utils.encode_col(c));
    text += `| # | ${colHdr.join(" | ")} |\n`;
    const lastR = Math.min(r.e.r, r.s.r + MAX_TEXT_ROWS - 1);
    for (let R = r.s.r; R <= lastR; R++) {
      const cells = [];
      let any = false;
      for (let c = r.s.c; c <= lastC; c++) {
        const cell = ws[XLSX.utils.encode_cell({ r: R, c })];
        if (cell?.f) info.formulas++;
        const t = cellText(cell);
        if (t) any = true;
        cells.push(t);
      }
      if (any) text += `| ${R + 1} | ${cells.join(" | ")} |\n`;
    }
    if (lastR < r.e.r) text += `… ${r.e.r - lastR} more rows not shown as text (use import_from_file to bring them in).\n`;
    if (lastC < r.e.c) text += `… columns after ${XLSX.utils.encode_col(lastC)} not shown as text.\n`;
    sheets.push(info);
  }
  const keepBuf = kind === "xlsx" || kind === "xlsm"; // other formats are re-encoded from wb for sheet copies
  return { kind: "sheet", format: kind, title: name, wb, buf: keepBuf ? buf : null, cells, sheets, text: text.trimEnd(), at: Date.now() };
}

function sheetSummary(entry) {
  return entry.sheets.map((s) => ({ name: s.name, range: s.ref, rows: s.rows, cols: s.cols, formulas: s.formulas || undefined }));
}

function sheetResult(entry, { maxChars, start, sheet }) {
  if (sheet) {
    const s = entry.sheets.find((x) => x.name.toLowerCase() === String(sheet).toLowerCase());
    if (!s) throw new Error(`No sheet named "${sheet}". Sheets: ${entry.sheets.map((x) => x.name).join(", ")}`);
    start = s.start;
  }
  const out = { type: entry.format, title: entry.title, sheets: sheetSummary(entry) };
  out.content = entry.text.slice(start, start + maxChars);
  out.total_chars = entry.text.length;
  if (start + maxChars < entry.text.length) out.next_start = start + maxChars;
  out.tip = "Cells show as | row | A | B | … with formulas in ⟨⟩. Use sheet=<name> to jump to a sheet. To bring data into this workbook use import_from_file (exact and fast) — don't retype values.";
  return out;
}

const toCellValue = (cell, withFormulas) => {
  if (!cell) return "";
  if (withFormulas && cell.f) return "=" + cell.f;
  if (cell.t === "e") return cell.w || "#N/A";
  if (cell.v instanceof Date) return localDateString(cell.v);
  // Excel parses written strings like typed input: "+591 7001…", "-", "=x", "@user" would become
  // formulas/numbers/errors. A leading apostrophe keeps them as text (and isn't shown in the cell).
  if (typeof cell.v === "string" && /^[=+\-@]/.test(cell.v)) return "'" + cell.v;
  return cell.v ?? "";
};

// Raw cell values (+ number formats) for import into the open workbook.
export function sheetValues(entry, { sheet, range, includeFormulas = false, maxCells = 300000 } = {}) {
  const sn = sheet ? entry.wb.SheetNames.find((n) => n.toLowerCase() === String(sheet).toLowerCase()) : entry.wb.SheetNames[0];
  if (!sn) throw new Error(`No sheet named "${sheet}". Sheets: ${entry.wb.SheetNames.join(", ")}`);
  const ws = entry.wb.Sheets[sn];
  const used = ws["!used"] ?? ws["!ref"];
  if (!used) throw new Error(`Sheet "${sn}" is empty.`);
  const r = XLSX.utils.decode_range(range ? String(range).replace(/\$/g, "").split("!").pop() : used);
  const rows = r.e.r - r.s.r + 1;
  const cols = r.e.c - r.s.c + 1;
  if (rows * cols > maxCells) throw new Error(`That's ${rows * cols} cells; the limit is ${maxCells}. Import a smaller range.`);
  const values = [];
  const formats = [];
  let hasFormats = false;
  for (let R = r.s.r; R <= r.e.r; R++) {
    const vr = [];
    const fr = [];
    for (let C = r.s.c; C <= r.e.c; C++) {
      const cell = ws[XLSX.utils.encode_cell({ r: R, c: C })];
      vr.push(toCellValue(cell, includeFormulas));
      let z = cell?.z && cell.z !== "General" ? String(cell.z) : "General";
      if (cell?.v instanceof Date && z === "General") z = "yyyy-mm-dd";
      if (z !== "General") hasFormats = true;
      fr.push(z);
    }
    values.push(vr);
    formats.push(fr);
  }
  return { sheet: sn, address: XLSX.utils.encode_range(r), rows, cols, values, formats: hasFormats ? formats : undefined };
}

// An .xlsx copy of the file for Excel's insertWorksheetsFromBase64 (keeps formatting for real .xlsx/.xlsm).
export function sheetFileBase64(entry) {
  if (entry.format === "xlsx" || entry.format === "xlsm") return { base64: entry.buf.toString("base64"), formatting: "preserved" };
  const out = XLSX.write(entry.wb, { type: "base64", bookType: "xlsx" });
  return { base64: out, formatting: "values and formulas only (converted from " + entry.format.toUpperCase() + ")" };
}

async function wordToText(buf, kind) {
  const tmp = path.join(os.tmpdir(), `oxa-word-${process.pid}-${Date.now()}.${WORD_EXT[kind]}`);
  fs.writeFileSync(tmp, buf);
  try {
    if (kind === "docx") {
      const { stdout } = await run("/usr/bin/unzip", ["-p", tmp, "word/document.xml"], { maxBuffer: 128 * 1024 * 1024, encoding: "utf8", timeout: 30000 });
      if (stdout.includes("<w:body")) return docxXmlToText(stdout);
    }
    // .doc / .rtf / .odt (and a fallback for odd .docx): textutil -> HTML keeps tables.
    const { stdout } = await run("/usr/bin/textutil", ["-convert", "html", "-stdout", tmp], { maxBuffer: 128 * 1024 * 1024, timeout: 60000 });
    return htmlToText(parse(stdout).querySelector("body") || parse(stdout), "");
  } catch (e) {
    throw new Error(`Couldn't read this ${kind.toUpperCase()} file (${String(e.stderr || e.message).trim().split("\n")[0]}).`);
  } finally {
    fs.unlink(tmp, () => {});
  }
}

// Headings with their character offsets, so the agent can jump around long documents.
function outlineOf(text) {
  const out = [];
  const rx = /^#{1,3} (.+)$/gm;
  let m;
  while ((m = rx.exec(text)) && out.length < 60) out.push({ heading: m[1].slice(0, 100), start: m.index });
  return out;
}

async function extractWord(buf, kind, name) {
  const text = await wordToText(buf, kind);
  const words = (text.match(/[\p{L}\p{N}]+/gu) || []).length;
  return { kind: "word", format: kind, title: name, text, words, outline: outlineOf(text), at: Date.now() };
}

function wordResult(doc, { maxChars, start, asImage }) {
  const out = { type: doc.format, title: doc.title, words: doc.words };
  out.content = doc.text.slice(start, start + maxChars);
  out.total_chars = doc.text.length;
  if (start + maxChars < doc.text.length) out.next_start = start + maxChars;
  if (start === 0 && doc.outline.length > 1 && doc.text.length > maxChars) {
    out.outline = doc.outline;
    out.tip = "Use start=<outline start> to jump to a section.";
  }
  if (asImage) out.note = "Page images are only available for PDFs; Word documents are read as text.";
  return out;
}

// ---------- documents uploaded from the panel (PDF and Word) ----------
const docs = new Map(); // doc_id -> { pdf, name, at }
const DOC_TTL = 6 * 60 * 60 * 1000;
const INLINE_CHARS = 8000; // small docs go straight into the message, no tool call needed

const MAX_DOCS_KEPT = 10;
const MAX_DOC_BYTES = 512 * 1024 * 1024; // the server runs all day as a login service

// Rough in-memory size: text (UTF-16) + raw file kept for sheet copies + parsed SheetJS cells.
function docBytes(d) {
  if (d.bytes) return d.bytes;
  const text = (d.pdf || d.word || d.slides || d.sheet)?.text || "";
  let b = text.length * 2;
  if (d.sheet) b += (d.sheet.buf?.length || 0) + (d.sheet.cells || 0) * 150;
  return (d.bytes = b);
}

// docs is kept in least-recently-used order (touchDoc moves an entry to the end), so evict from the front.
function pruneDocs(keepId) {
  let total = 0;
  for (const d of docs.values()) total += docBytes(d);
  for (const [id, d] of docs) {
    if (id === keepId) continue;
    const expired = Date.now() - d.at >= DOC_TTL;
    if (!expired && docs.size <= MAX_DOCS_KEPT && total <= MAX_DOC_BYTES) break;
    docs.delete(id);
    total -= docBytes(d);
    if (d.pdf) fs.unlink(d.pdf.file, () => {});
  }
}

function touchDoc(id, d) {
  d.at = Date.now();
  docs.delete(id);
  docs.set(id, d);
}

export async function uploadDoc(buf, name) {
  if (!buf?.length) throw new Error("Empty upload.");
  const kind = detectDocKind(buf, name);
  if (!kind) throw new Error(`${name} isn't a supported file (PDF, Word, RTF, ODT, Excel .xlsx/.xls/.xlsm, CSV, ODS, PowerPoint .pptx or ODP).`);
  if (kind === "ppt") throw new Error(`${name} is an old-format PowerPoint (.ppt). Open it in PowerPoint and Save As .pptx, then attach that.`);
  if (kind === "key") throw new Error(`${name} is a Keynote file. In Keynote use File › Export To › PowerPoint, then attach the .pptx.`);
  const id = crypto.randomBytes(5).toString("hex");
  const safeName = String(name || `document.${kind}`).replace(/[^\w.\- ()À-ÿ]/g, "_").slice(0, 120);

  if (SLIDE_KINDS.has(kind)) {
    const slides = await extractSlides(buf, kind, safeName);
    docs.set(id, { slides, name: safeName, at: Date.now() });
    pruneDocs(id);
    const out = {
      doc_id: id, name: safeName, kind: "slides", format: kind, slides: slides.pages,
      charts: slides.charts, notes: slides.withNotes, total_chars: slides.text.length,
      titles: slides.outline.slice(0, 40).map((o) => o.title),
    };
    if (slides.text.length <= INLINE_CHARS) out.full_text = slides.text;
    return out;
  }

  if (SHEET_KINDS.has(kind)) {
    const sheet = extractSheet(buf, kind, safeName);
    docs.set(id, { sheet, name: safeName, at: Date.now() });
    pruneDocs(id);
    const out = { doc_id: id, name: safeName, kind: "sheet", format: kind, sheets: sheetSummary(sheet), total_chars: sheet.text.length };
    if (sheet.text.length <= INLINE_CHARS) out.full_text = sheet.text;
    return out;
  }

  if (kind !== "pdf") {
    const word = await extractWord(buf, kind, safeName);
    docs.set(id, { word, name: safeName, at: Date.now() });
    pruneDocs(id);
    const out = { doc_id: id, name: safeName, kind: "word", format: kind, words: word.words, total_chars: word.text.length };
    if (word.text.length <= INLINE_CHARS) out.full_text = word.text;
    if (!word.text.trim()) out.warning = "No text found in this document.";
    return out;
  }

  const pdf = await extractPdf(buf, `upload://${id}/${encodeURIComponent(safeName)}`);
  if (/^(printmgr|microsoft|untitled)/i.test(pdf.title) || pdf.title.startsWith("upload")) pdf.title = safeName;
  docs.set(id, { pdf, name: safeName, at: Date.now() });
  pruneDocs(id);
  const out = { doc_id: id, name: safeName, kind: "pdf", title: pdf.title, pages: pdf.pages, scanned: pdf.scanned, total_chars: pdf.text.length };
  if (!pdf.scanned && pdf.text.length <= INLINE_CHARS) out.full_text = pdf.text;
  return out;
}

function getDoc(id) {
  id = String(id || "");
  const d = docs.get(id);
  if (!d || (d.pdf && !fs.existsSync(d.pdf.file))) {
    throw new Error("That file is no longer available (attachments are kept for 6 hours or until the server restarts). Ask the user to attach it again.");
  }
  touchDoc(id, d);
  return d;
}

export function docValues(id, opts) {
  const d = getDoc(id);
  if (!d.sheet) throw new Error("import_from_file only works with spreadsheet files (xlsx, xls, xlsm, csv, ods).");
  return { doc_id: id, name: d.name, ...sheetValues(d.sheet, opts) };
}

export function docFileBase64(id) {
  const d = getDoc(id);
  if (!d.sheet) throw new Error("Only spreadsheet files can be inserted as sheets.");
  return { doc_id: id, name: d.name, format: d.sheet.format, sheet_names: d.sheet.wb.SheetNames, ...sheetFileBase64(d.sheet) };
}

export async function readDoc(id, { maxChars = 8000, start = 0, page, asImage = false, sheet } = {}) {
  const d = getDoc(id);
  maxChars = Math.min(20000, Math.max(1000, Number(maxChars) || 8000));
  start = Math.max(0, Number(start) || 0);
  if (d.word) return { doc_id: id, name: d.name, ...wordResult(d.word, { maxChars, start, asImage }) };
  if (d.sheet) return { doc_id: id, name: d.name, ...sheetResult(d.sheet, { maxChars, start, sheet }) };
  if (d.slides) {
    const out = { doc_id: id, name: d.name, ...slidesResult(d.slides, { maxChars, start, page: Number(page) || undefined }) };
    if (asImage) out.note = "Slide images aren't available (no renderer on this Mac); slide text, tables, chart data and notes are included instead.";
    return out;
  }
  const out = await pdfResult(d.pdf, d.name, { maxChars, start, page: Number(page) || undefined, asImage });
  delete out.url;
  return { doc_id: id, name: d.name, ...out };
}

export const _internals = { extractPdf, pdfResult, cleanPdfPage, extractWord, extractSheet, sheetResult, sheetValues, extractSlides, slidesResult }; // for scripts/test-*.mjs

export async function webFetch(url, { maxChars = 8000, start = 0, page, asImage = false } = {}) {
  maxChars = Math.min(20000, Math.max(1000, Number(maxChars) || 8000));
  start = Math.max(0, Number(start) || 0);
  page = Number(page) || undefined;

  const cached = pdfCache.get(url);
  if (cached && Date.now() - cached.at < PDF_TTL && fs.existsSync(cached.file)) {
    return pdfResult(cached, url, { maxChars, start, page, asImage });
  }

  const res = await request(url);
  const type = String(res.headers["content-type"] || "");
  if (res.status >= 400) throw new Error(`HTTP ${res.status} for ${res.url}`);

  // Spreadsheets linked from web pages: read as text and register as a doc so the agent can import them.
  const sheetKind =
    /spreadsheetml|ms-excel|opendocument\.spreadsheet/i.test(type) || /\.(xlsx|xlsm|xlsb|xls|ods)($|[?#])/i.test(res.url)
      ? detectDocKind(res.body, new URL(res.url).pathname)
      : null;
  if (sheetKind && SHEET_KINDS.has(sheetKind) && sheetKind !== "csv") {
    if (res.truncated) throw new Error("Spreadsheet is larger than 40 MB; can't read it.");
    const name = decodeURIComponent(new URL(res.url).pathname.split("/").pop() || `data.${sheetKind}`);
    const id = crypto.randomBytes(5).toString("hex");
    const sheet = extractSheet(res.body, sheetKind, name);
    docs.set(id, { sheet, name, at: Date.now() });
    pruneDocs(id);
    return { url: res.url, doc_id: id, ...sheetResult(sheet, { maxChars, start, sheet: undefined }), note: `Registered as doc_id ${id}: use read_document (sheet=…) to read more or import_from_file to bring data into the workbook.` };
  }

  // Presentations linked from web pages.
  const slideKind =
    /presentationml|ms-powerpoint|opendocument\.presentation/i.test(type) || /\.(pptx|ppsx|odp)($|[?#])/i.test(res.url)
      ? detectDocKind(res.body, new URL(res.url).pathname)
      : null;
  if (slideKind && SLIDE_KINDS.has(slideKind)) {
    if (res.truncated) throw new Error("Presentation is larger than 40 MB; can't read it.");
    const name = decodeURIComponent(new URL(res.url).pathname.split("/").pop() || `slides.${slideKind}`);
    const slides = await extractSlides(res.body, slideKind, name);
    return { url: res.url, ...slidesResult(slides, { maxChars, start, page: Number(page) || undefined }) };
  }

  // Word / RTF / ODT documents linked from web pages.
  const wordKind =
    /wordprocessingml|msword|rtf|opendocument\.text/i.test(type) || /\.(docx?|rtf|odt)($|[?#])/i.test(res.url)
      ? detectDocKind(res.body, new URL(res.url).pathname)
      : null;
  if (wordKind && wordKind !== "pdf") {
    if (res.truncated) throw new Error("Document is larger than 40 MB; can't read it.");
    const name = decodeURIComponent(new URL(res.url).pathname.split("/").pop() || `document.${wordKind}`);
    const word = await extractWord(res.body, wordKind, name);
    return { url: res.url, ...wordResult(word, { maxChars, start, asImage }) };
  }

  if (/application\/pdf/i.test(type) || res.body.subarray(0, 5).toString() === "%PDF-") {
    if (res.truncated) throw new Error("PDF is larger than 40 MB; can't read it.");
    const pdf = await extractPdf(res.body, res.url);
    cachePdf(url, pdf);
    return pdfResult(pdf, res.url, { maxChars, start, page, asImage });
  }
  let title = "";
  let text;
  let links = [];
  if (/html|xml/i.test(type) || /^\s*</.test(res.body.subarray(0, 200).toString())) {
    const root = parse(decode(res.body, type), { comment: false });
    title = root.querySelector("title")?.text.trim() || "";
    links = root
      .querySelectorAll("a[href]")
      .map((a) => {
        try {
          return { text: a.text.replace(/\s+/g, " ").trim().slice(0, 80), url: new URL(a.getAttribute("href"), res.url).toString() };
        } catch {
          return null;
        }
      })
      .filter((l) => l && l.text && /^https?:/.test(l.url))
      .slice(0, 25);
    const body = root.querySelector("body") || root;
    // Prefer the main content area when the page marks one.
    const main = body.querySelector("main") || body.querySelector("article") || body.querySelector("[role=main]");
    text = htmlToText(main && main.text.trim().length > 400 ? main : body, res.url);
  } else if (/json|text\/|csv/i.test(type)) {
    text = decode(res.body, type);
  } else {
    throw new Error(`Can't read this content type (${type || "unknown"}). Supported: web pages, PDF, Word (.docx/.doc), RTF, ODT, text, CSV and JSON.`);
  }
  const slice = text.slice(start, start + maxChars);
  const out = { url: res.url, title, content: slice, total_chars: text.length };
  if (start + maxChars < text.length) out.next_start = start + maxChars;
  if (start === 0 && links.length) out.links = links;
  if (res.truncated) out.note = "Page was very large; only the first 40 MB were read.";
  return out;
}
