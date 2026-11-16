/* global Office, Excel */
import { marked } from "/vendor/marked.esm.js";
import DOMPurify from "/vendor/purify.es.mjs";
import { runTurn, listModels, modelInfo } from "./agent.js";
import { runTool, undoSnapshots, workbookContext, toolRisk } from "./tools.js";
import { describeApproval, undoLabel } from "./describe.js";
import { Recorder, transcribe, sttStatus } from "./dictation.js";
import * as tts from "./speech.js";
import { prepareImage, attachableFilesFrom, isDoc, MAX_IMAGES } from "./images.js";
import * as chats from "./chats.js";
import * as messages from "./messages.js";

const $ = (id) => document.getElementById(id);
const els = {
  messages: $("messages"), empty: $("empty"), input: $("input"), send: $("send"), stop: $("stop"),
  status: $("status"), statusText: $("statusText"), ctxChip: $("ctxChip"), settings: $("settings"),
  model: $("model"), think: $("think"), numCtx: $("numCtx"), approval: $("approval"),
  temperature: $("temperature"), tempVal: $("tempVal"), maxSteps: $("maxSteps"), emptyModel: $("emptyModel"),
  mic: $("mic"), sttLang: $("sttLang"), sttAutoSend: $("sttAutoSend"),
  ttsAuto: $("ttsAuto"), ttsVoice: $("ttsVoice"), ttsRate: $("ttsRate"), ttsRateVal: $("ttsRateVal"),
  attach: $("attach"), fileInput: $("fileInput"), attachments: $("attachments"), web: $("web"), showStats: $("showStats"),
  chatPanel: $("chatPanel"), chatList: $("chatList"), chatPanelTitle: $("chatPanelTitle"),
  legacyBanner: $("legacyBanner"), legacyTitle: $("legacyTitle"),
};

// ---------- storage (best effort) ----------
const store = {
  get(k, d) { try { const v = localStorage.getItem(k); return v ? JSON.parse(v) : d; } catch { return d; } },
  set(k, v) { try { localStorage.setItem(k, JSON.stringify(v)); } catch { /* ignore */ } },
};

const DEFAULTS = {
  model: "qwen3.8:latest", think: "low", numCtx: 32768, approval: "destructive", temperature: 0.6, maxSteps: 30,
  sttLang: "auto", sttAutoSend: false,
  ttsAuto: false, ttsVoice: "auto", ttsRate: 190,
  web: true,
  showStats: false,
};
const settings = { ...DEFAULTS, ...store.get("oxa.settings", {}) };
let history = []; // current chat's messages (Ollama format)
// Per-workbook chats: wb.id comes from a setting stored in the workbook file (see chats.js).
const wb = { id: null, name: "", data: null };
let chat = null; // current chat record in wb.data.chats (null = unsaved new chat)
let legacyHidden = false;
let pendingImages = []; // [{ dataUrl, b64, name }]
let pendingDocs = []; // [{ name, uploading, doc_id, pages, scanned, full_text, error }]
const MAX_DOCS = 3;

// Images are kept in memory for the session but not persisted (localStorage is ~5 MB).
const leanHistory = (h) =>
  h.map((m) => (m.images?.length ? { ...m, images: undefined, _previews: undefined, imageCount: m.images.length } : m));

// Write messages into a chat record. Used for the current chat and for a turn that
// finishes after the user switched to another chat (so it lands where it started).
function persistChat(c, h) {
  if (!wb.id || !wb.data || !c || !h.length) return;
  if (!wb.data.chats.includes(c)) return; // deleted meanwhile
  c.history = leanHistory(h);
  c.updated = Date.now();
  c.title = chats.titleFor(h);
  wb.data.name = wb.name;
  chats.saveWorkbook(wb.id, wb.data);
  if (!els.chatPanel.hidden) renderChatList();
}

function saveHistory() {
  if (!wb.id || !wb.data || !history.length) return;
  if (!chat) {
    chat = chats.newChat();
    wb.data.chats.unshift(chat);
  }
  wb.data.current = chat.id;
  persistChat(chat, history);
}

// The workbook gets its id (stored in the file) the first time the user actually chats in it.
async function ensureWorkbook() {
  if (wb.id) return;
  wb.id = chats.newId();
  if (inExcel) {
    try {
      await chats.assignWorkbookId(wb.id);
    } catch (e) {
      console.warn("Couldn't store workbook id; chats for this workbook will last until it's closed.", e);
    }
  }
  wb.data = chats.loadWorkbook(wb.id);
}
let inExcel = false;
let busy = false;
let controller = null;
let pendingNote = "";

// ---------- settings UI ----------
function syncSettingsUI() {
  els.think.value = settings.think;
  els.numCtx.value = String(settings.numCtx);
  els.approval.value = settings.approval;
  els.temperature.value = settings.temperature;
  els.tempVal.textContent = `(${settings.temperature})`;
  els.maxSteps.value = settings.maxSteps;
  els.sttLang.value = settings.sttLang;
  els.sttAutoSend.checked = !!settings.sttAutoSend;
  els.ttsAuto.checked = !!settings.ttsAuto;
  els.web.checked = settings.web !== false;
  els.showStats.checked = !!settings.showStats;
  els.ttsVoice.value = settings.ttsVoice;
  els.ttsRate.value = settings.ttsRate;
  els.ttsRateVal.textContent = `(${settings.ttsRate} wpm)`;
  els.emptyModel.textContent = settings.model.replace(/:latest$/, "");
}
function saveSettings() {
  settings.think = els.think.value;
  settings.numCtx = Number(els.numCtx.value);
  settings.approval = els.approval.value;
  settings.temperature = Number(els.temperature.value);
  settings.maxSteps = Math.max(1, Math.min(100, Number(els.maxSteps.value) || DEFAULTS.maxSteps));
  settings.sttLang = els.sttLang.value;
  settings.sttAutoSend = els.sttAutoSend.checked;
  settings.ttsAuto = els.ttsAuto.checked;
  settings.web = els.web.checked;
  settings.showStats = els.showStats.checked;
  settings.ttsVoice = els.ttsVoice.value || "auto";
  settings.ttsRate = Number(els.ttsRate.value) || DEFAULTS.ttsRate;
  if (els.model.value) settings.model = els.model.value;
  store.set("oxa.settings", settings);
  syncSettingsUI();
}
for (const el of [els.think, els.numCtx, els.approval, els.temperature, els.maxSteps, els.model, els.sttLang, els.sttAutoSend, els.ttsAuto, els.ttsVoice, els.ttsRate, els.web, els.showStats]) {
  el.addEventListener("change", saveSettings);
}
els.temperature.addEventListener("input", () => (els.tempVal.textContent = `(${els.temperature.value})`));
els.ttsRate.addEventListener("input", () => (els.ttsRateVal.textContent = `(${els.ttsRate.value} wpm)`));
$("openSettings").onclick = () => {
  els.settings.hidden = !els.settings.hidden;
  if (!els.settings.hidden) els.chatPanel.hidden = true;
};

async function refreshModels() {
  try {
    const models = await listModels();
    if (!models.length) throw new Error("no models pulled");
    if (!models.includes(settings.model)) {
      settings.model = models.find((m) => /^qwen3\.8/.test(m)) || models.find((m) => /qwen/i.test(m)) || models[0];
      store.set("oxa.settings", settings);
    }
    els.model.innerHTML = "";
    for (const m of models) els.model.add(new Option(m, m, false, m === settings.model));
    setStatus("ok", settings.model.replace(/:latest$/, ""));
    syncSettingsUI();
  } catch (e) {
    setStatus("err", "Ollama offline");
    console.warn(e);
  }
}
function setStatus(kind, text) {
  els.status.className = `status ${kind}`;
  els.statusText.textContent = text;
  els.status.title = kind === "err" ? "Start Ollama (ollama serve) and the add-in server (npm start)." : "Connected to Ollama";
}

// ---------- rendering ----------
marked.setOptions({ gfm: true, breaks: true });
const md = (text) => DOMPurify.sanitize(marked.parse(text || ""));

function nearBottom() {
  const m = els.messages;
  return m.scrollHeight - m.scrollTop - m.clientHeight < 80;
}
function scroll(force) {
  if (force || nearBottom()) els.messages.scrollTop = els.messages.scrollHeight;
}
function el(tag, cls, html) {
  const e = document.createElement(tag);
  if (cls) e.className = cls;
  if (html !== undefined) e.innerHTML = html;
  return e;
}
const esc = (s) => String(s).replace(/[&<>"']/g, (c) => ({ "&": "&amp;", "<": "&lt;", ">": "&gt;", '"': "&quot;", "'": "&#39;" })[c]);

function addUser(text, ctxLine, images = [], goneCount = 0, docs = []) {
  els.empty.hidden = true;
  const wrap = el("div", "msg user");
  if (images.length || goneCount || docs?.length) {
    const strip = el("div", "imgs");
    for (const d of docs || []) {
      const chip = el("div", "doc-chip");
      chip.append(docIcon(d), el("span", "doc-meta", `<b>${esc(d.name)}</b><small>${esc(docLabel(d))}</small>`));
      strip.append(chip);
    }
    for (const im of images) {
      const i = document.createElement("img");
      i.src = im.dataUrl;
      i.alt = im.name;
      i.onclick = () => openLightbox(im.dataUrl);
      strip.append(i);
    }
    if (goneCount) strip.append(el("span", "gone", `🖼 ${goneCount} image${goneCount > 1 ? "s" : ""} (not kept after reload)`));
    wrap.append(strip);
  }
  const b = el("div", "bubble");
  if (!text) b.hidden = true;
  b.textContent = text;
  if (ctxLine) b.append(el("span", "ctxline", esc(ctxLine)));
  wrap.append(b);
  els.messages.append(wrap);
  scroll(true);
}

const ICON_OK = '<svg viewBox="0 0 20 20"><path d="M5 10.5l3.2 3L15 7" /></svg>';
const ICON_FAIL = '<svg viewBox="0 0 20 20"><path d="M6 6l8 8M14 6l-8 8" /></svg>';
const ICON_WAIT = '<svg viewBox="0 0 20 20"><circle cx="10" cy="10" r="6.5" /><path d="M10 6.5V10l2.2 1.6" /></svg>';

function toolLabel(name, a) {
  const addr = a.address || a.source || "";
  const where = addr ? ` <a class="addr" data-addr="${esc(addr)}" data-sheet="${esc(a.sheet || "")}">${esc((a.sheet ? a.sheet + "!" : "") + addr)}</a>` : "";
  switch (name) {
    case "get_workbook_info": return "<b>Read</b> workbook structure";
    case "read_range": return `<b>Read</b>${where}`;
    case "write_range": return `<b>Write</b>${where}`;
    case "clear_range": return `<b>Clear</b>${where}`;
    case "format_range": return `<b>Format</b>${where}`;
    case "manage_sheet": return `<b>${esc(a.action || "Edit")} sheet</b> ${esc(a.name || "")}${a.new_name ? " → " + esc(a.new_name) : ""}`;
    case "insert_delete_cells": return `<b>${esc(a.action === "delete" ? "Delete" : "Insert")}</b>${where}`;
    case "create_table": return `<b>Create table</b>${where}`;
    case "create_chart": return `<b>Create ${esc(a.type || "")} chart</b>${where}`;
    case "sort_range": return `<b>Sort</b>${where}`;
    case "find_cells": return `<b>Search</b> “${esc(a.query || "")}”`;
    case "select_range": return `<b>Select</b>${where}`;
    case "execute_office_js": return `<b>Run Office.js</b> ${esc(a.explanation || "")}`;
    case "read_document": {
      const d = history.flatMap((m) => m._docs || []).find((x) => x.doc_id === a.doc_id);
      const unit = d?.kind === "slides" ? "slide" : "page";
      const extra = a.page ? ` · ${unit} ${Number(a.page)}${a.as_image && unit === "page" ? " (image)" : ""}` : a.sheet ? ` · ${a.sheet}` : "";
      return `<b>Read</b> 📄 ${esc(d?.name || "attached document")}${esc(extra)}`;
    }
    case "import_from_file": {
      const d = history.flatMap((m) => m._docs || []).find((x) => x.doc_id === a.doc_id);
      const src = `${d?.name || "file"}${a.sheet ? " › " + a.sheet : ""}${a.source_range ? "!" + a.source_range : ""}`;
      const dest = a.mode === "copy_sheet" ? " as new sheet" : a.target_address || a.target_sheet ? ` → ${(a.target_sheet ? a.target_sheet + "!" : "") + (a.target_address || "A1")}` : "";
      return `<b>Import</b> 📊 ${esc(src)}${esc(dest)}`;
    }
    case "web_search": return `<b>Search the web</b> “${esc(a.query || "")}”`;
    case "web_fetch": {
      let host = a.url || "";
      try { host = new URL(a.url).hostname.replace(/^www\./, ""); } catch { /* keep raw */ }
      const docExt = /\.(pdf|docx?|rtf|odt)($|[?#])/i.exec(a.url || "")?.[1];
      const extra = a.page ? ` · page ${Number(a.page)}${a.as_image ? " (image)" : ""}` : docExt ? ` · ${docExt.toUpperCase()}` : "";
      // Only real web URLs become links: a model-supplied "javascript:" URL would otherwise run
      // in this pane (which has Office.js access to the workbook) when clicked.
      const link = /^https?:\/\//i.test(a.url || "")
        ? `<a class="ext" href="${esc(a.url)}" target="_blank" rel="noopener noreferrer">${esc(host)}</a>`
        : esc(host);
      return `<b>Read</b> ${link}${esc(extra)}`;
    }
    default: return `<b>${esc(name)}</b>`;
  }
}

els.messages.addEventListener("click", (e) => {
  const a = e.target.closest(".addr");
  if (!a || !inExcel) return;
  e.preventDefault();
  runTool("select_range", { address: a.dataset.addr, sheet: a.dataset.sheet || undefined });
});

// The range a tool actually touched, from its result (the model may pass just the top-left cell,
// e.g. "A1" for a 5-row write). Returns args with sheet/address replaced, or the original args.
function actualArgs(args, res) {
  const full = res && !res.error && (res.written || res.cleared || res.formatted || res.sorted || (typeof res.range === "string" && res.range));
  if (typeof full !== "string" || !full.includes("!")) return args;
  const i = full.lastIndexOf("!");
  return { ...args, sheet: full.slice(0, i).replace(/^'(.*)'$/, "$1").replace(/''/g, "'"), address: full.slice(i + 1) };
}

function toolCard(container, name, args, result) {
  const card = el("details", "tool");
  const sum = el("summary");
  const ico = el("span", "ico", '<span class="spin"></span>');
  const label = el("span", "label", toolLabel(name, args));
  sum.append(ico, label);
  card.append(sum);
  const argPre = el("pre");
  argPre.textContent = name === "execute_office_js" ? args.code || "" : JSON.stringify(args, null, 1);
  card.append(argPre);
  container.append(card);

  const view = {
    end(res, images) {
      const failed = res && res.error;
      card.classList.add(failed ? "fail" : "ok");
      const actual = actualArgs(args, res);
      if (actual !== args) label.innerHTML = toolLabel(name, actual);
      ico.innerHTML = failed ? ICON_FAIL : ICON_OK;
      if (images?.length) {
        const strip = el("div", "tool-imgs");
        for (const b64 of images) {
          const i = document.createElement("img");
          i.src = `data:image/jpeg;base64,${b64}`;
          i.alt = "Rendered page";
          i.onclick = (e) => { e.preventDefault(); openLightbox(i.src); };
          strip.append(i);
        }
        card.append(strip);
        card.open = true;
      }
      const out = el("pre");
      const s = JSON.stringify(res, null, 1) || "";
      out.textContent = s.length > 4000 ? s.slice(0, 4000) + "\n…" : s;
      card.append(out);
      if (failed) card.title = res.error;
    },
    // Asks in plain words what will happen and whether Undo covers it. The raw arguments stay in the
    // (collapsed) card above, for anyone who wants the details.
    approve() {
      ico.innerHTML = ICON_WAIT;
      const d = describeApproval(name, args);
      const box = el("div", "approval");
      box.append(el("div", "why", esc(d.text)), el("div", "undo-note", esc(d.undo)));
      const btns = el("div", "btns");
      const yes = el("button", "yes", esc(d.yes));
      const no = el("button", "", esc(d.no));
      btns.append(yes, no);
      box.append(btns);
      card.classList.add("asking");
      card.after(box);
      scroll(true);
      return new Promise((resolve) => {
        const done = (v) => { box.remove(); card.classList.remove("asking"); ico.innerHTML = '<span class="spin"></span>'; resolve(v); };
        yes.onclick = () => done(true);
        no.onclick = () => done(false);
        controller?.signal.addEventListener("abort", () => done(false), { once: true });
      });
    },
  };
  if (result !== undefined) view.end(result);
  scroll();
  return view;
}

function assistantBlock() {
  const wrap = el("div", "msg assistant");
  els.messages.append(wrap);
  const calls = []; // { name, args, ok } for every tool call in this turn, for the outcome summary
  return {
    wrap,
    calls,
    step() {
      const stepEl = el("div", "step");
      const waiting = el("div", "meta pulse", "Working…");
      stepEl.append(waiting);
      wrap.append(stepEl);
      scroll();
      let think, thinkBody, content;
      return {
        thinking(t) {
          waiting.remove();
          if (!think) {
            think = el("details", "think");
            const s = el("summary", "pulse", "Thinking…");
            thinkBody = el("div", "think-body");
            think.append(s, thinkBody);
            stepEl.prepend(think);
          }
          thinkBody.textContent = t;
          thinkBody.scrollTop = thinkBody.scrollHeight;
          scroll();
        },
        content(t) {
          waiting.remove();
          if (think) { think.querySelector("summary").className = ""; think.querySelector("summary").textContent = "Thought process"; }
          if (!content) { content = el("div", "md"); stepEl.append(content); }
          content.innerHTML = md(t);
          scroll();
        },
        toolStart(name, args) {
          waiting.remove();
          if (think) { think.querySelector("summary").className = ""; think.querySelector("summary").textContent = "Thought process"; }
          const rec = { name, args, ok: null };
          calls.push(rec);
          const view = toolCard(stepEl, name, args);
          const end = view.end;
          view.end = (res, images) => { rec.ok = !(res && res.error); rec.args = actualArgs(args, res); end(res, images); };
          return view;
        },
        done() {
          waiting.remove();
          if (think && think.querySelector("summary").className) { think.querySelector("summary").className = ""; think.querySelector("summary").textContent = "Thought process"; }
          if (content && !content.textContent.trim()) content.remove();
        },
      };
    },
  };
}

// Rebuild the transcript from saved history after the pane reloads.
function renderHistory() {
  if (!history.length) return;
  els.empty.hidden = true;
  let block = null;
  const pendingCalls = [];
  for (const m of history) {
    if (m._nudge) continue; // the agent's own "continue" prompt after an empty step
    if (m.role === "user") {
      const { ctx, text } = chats.parseUserMessage(m.content);
      addUser(text, ctx, [], m.imageCount || 0, m._docs || []);
      block = null;
    } else if (m.role === "assistant") {
      block = block || assistantBlock();
      const step = el("div", "step");
      block.wrap.append(step);
      if (m.content) step.append(el("div", "md", md(m.content)));
      for (const c of m.tool_calls || []) pendingCalls.push({ step, c });
      if (m.content && !m.tool_calls?.length && m.content !== "(stopped by user)") {
        const meta = el("div", "meta");
        meta.append(speakButton(m.content));
        block.wrap.append(meta);
      }
    } else if (m.role === "tool") {
      const p = pendingCalls.shift();
      let res;
      try { res = JSON.parse(m.content); } catch { res = m.content; }
      if (p) toolCard(p.step, p.c.function?.name, p.c.function?.arguments || {}, res);
    }
  }
  scroll(true);
}

// ---------- selection context ----------
async function updateCtxChip() {
  if (!inExcel) return;
  const c = await workbookContext();
  els.ctxChip.textContent = !c
    ? "No selection"
    : c.selection_size === "0x0"
      ? `${c.active_sheet} · chart/shape selected`
      : `${c.selection}${c.selection_size !== "1x1" ? `  (${c.selection_size})` : ""}`;
}

// ---------- turn outcome ----------
// The reply is the model's own account of what it did; the tool results are what actually happened.
// When changes failed (or were declined), say so above the reply; list what did change, with
// clickable addresses, below it. Returns the "Changed" row (or null) so Undo can sit in it.
const MAX_CHANGES_LISTED = 6;
function renderOutcome(block) {
  const changes = block.calls.filter((c) => c.ok !== null && toolRisk(c.name) !== "read");
  if (!changes.length) return null;
  const failed = changes.filter((c) => !c.ok);
  if (failed.length) {
    const n = changes.length;
    const banner = el(
      "div",
      "outcome warn",
      `<b>${failed.length === n ? (n === 1 ? "The change" : `All ${n} changes`) : `${failed.length} of ${n} changes`} didn't go through.</b> ` +
        "The reply may not match the workbook; check the steps marked ✕."
    );
    banner.setAttribute("role", "status");
    const reply = [...block.wrap.querySelectorAll(".step > .md")].pop();
    if (reply) reply.before(banner);
    else block.wrap.append(banner);
  }
  const done = [...new Set(changes.filter((c) => c.ok).map((c) => toolLabel(c.name, c.args)))];
  if (!done.length) return null;
  const shown = done.slice(0, MAX_CHANGES_LISTED).map((h) => `<span class="chg">${h}</span>`).join("");
  const more = done.length > MAX_CHANGES_LISTED ? `<span class="chg more">+${done.length - MAX_CHANGES_LISTED} more</span>` : "";
  const row = el("div", "changes", `<span class="k">Changed</span>${shown}${more}`);
  block.wrap.append(row);
  return row;
}

// ---------- sending ----------
function setBusy(b) {
  busy = b;
  els.send.hidden = b;
  els.stop.hidden = !b;
  els.input.placeholder = b ? "Working… (Stop to interrupt)" : "Ask about this workbook…";
}

async function send(text) {
  text = text.trim();
  if ((!text && !pendingImages.length && !pendingDocs.length) || busy) return;
  if (pendingDocs.some((d) => d.uploading)) return flashMic("Still reading the document — one moment.");
  const images = pendingImages;
  const docs = pendingDocs;
  if (!text && docs.length) text = `Summarize ${docs.length > 1 ? "these documents" : "this document"} and point out any data I could put into this workbook.`;
  else if (!text && images.length) text = "What's in this image? If it contains data that belongs in this workbook, tell me how you'd use it.";
  pendingImages = [];
  pendingDocs = [];
  renderAttachments();
  els.input.value = "";
  autosize();

  const ctxLine = messages.contextLine(inExcel ? await workbookContext() : null, inExcel);
  const note = pendingNote;
  pendingNote = "";
  const userMsg = { role: "user", content: messages.userContent({ ctxLine, note, text, docs }) };
  if (images.length) {
    userMsg.images = images.map((i) => i.b64);
    userMsg._previews = images.map((i) => ({ dataUrl: i.dataUrl, name: i.name })); // UI only; stripped before sending
  }
  if (docs.length) {
    userMsg._docs = docs.map((d) => ({ doc_id: d.doc_id, name: d.name, kind: d.kind, format: d.format, words: d.words, pages: d.pages, scanned: d.scanned, sheets: d.sheets, slides: d.slides, charts: d.charts })); // UI only
  }
  history.push(userMsg);
  addUser(text, ctxLine, images, 0, userMsg._docs);
  await ensureWorkbook();
  els.legacyBanner.hidden = true;
  saveHistory(); // keep the question even if the turn is interrupted
  // The turn belongs to this chat even if the user switches chats while it runs.
  const turnHistory = history;
  const turnChat = chat;
  const stillHere = () => history === turnHistory;

  const block = assistantBlock();
  const undo = [];
  controller = new AbortController();
  setBusy(true);
  const t0 = performance.now();
  let statsLine = "";
  let finalText = "";
  tts.stop();
  try {
    finalText = await runTurn({
      history: turnHistory, settings, signal: controller.signal, undo,
      ui: {
        step: () => block.step(),
        stats: (s) => {
          const tps = s.evalNs ? (s.evalCount / (s.evalNs / 1e9)).toFixed(1) : "–";
          statsLine = `${s.steps} step${s.steps > 1 ? "s" : ""} · ${((performance.now() - t0) / 1000).toFixed(1)}s · ${tps} tok/s · ${s.promptCount} ctx tokens`;
        },
      },
    });
    if (/^Stopped after/.test(finalText)) block.wrap.append(el("div", "md", md(finalText)));
  } catch (e) {
    if (e.name === "AbortError") {
      block.wrap.append(el("div", "meta", "Stopped."));
      turnHistory.push({ role: "assistant", content: "(stopped by user)" });
    } else {
      block.wrap.append(el("div", "error-box", esc(e.message || String(e))));
      console.error(e);
    }
  } finally {
    setBusy(false);
    controller = null;
    if (stillHere()) saveHistory();
    else persistChat(turnChat, turnHistory); // user switched chats mid-turn: save to the chat it started in
    const meta = el("div", "meta");
    let readBtn = null;
    if (finalText && !/^Stopped after/.test(finalText)) {
      readBtn = speakButton(finalText);
      meta.append(readBtn);
    }
    if (statsLine && settings.showStats) meta.append(el("span", "", esc(statsLine)));
    const changesRow = renderOutcome(block);
    if (undo.length) {
      const u = el("button", "undo-btn", esc(undoLabel(undo)));
      u.title = "Restores the values and formulas this reply changed (not formatting, tables or charts)";
      u.onclick = async () => {
        u.disabled = true;
        try {
          await undoSnapshots(undo);
          u.replaceWith(el("span", "undo-done", "Reverted ✓"));
          pendingNote = "the user reverted the cell value/formula changes you made in your previous turn (formatting, tables and charts were not reverted)";
        } catch (err) {
          u.replaceWith(el("span", "undo-done", "Undo failed: " + esc(err.message)));
        }
      };
      (changesRow || block.wrap).append(u);
    }
    if (meta.childNodes.length) block.wrap.append(meta);
    scroll(true);
    updateCtxChip();
    if (readBtn && settings.ttsAuto && micState === "idle" && stillHere()) readBtn.click();
  }
}

function autosize() {
  els.input.style.height = "auto";
  els.input.style.height = Math.min(els.input.scrollHeight, 160) + "px";
}
els.input.addEventListener("input", autosize);
els.input.addEventListener("keydown", (e) => {
  if (e.key === "Enter" && !e.shiftKey && !e.isComposing) {
    e.preventDefault();
    send(els.input.value);
  }
});
els.send.onclick = () => send(els.input.value);
els.stop.onclick = () => controller?.abort();
document.querySelectorAll(".suggestions button").forEach((b) => (b.onclick = () => send(b.dataset.prompt)));
function clearTranscript() {
  els.messages.querySelectorAll(".msg").forEach((m) => m.remove());
  els.empty.hidden = false;
}

// New chat keeps the previous one in this workbook's history.
function startNewChat() {
  if (busy) controller?.abort();
  tts.stop();
  saveHistory();
  pendingImages = [];
  pendingDocs = [];
  renderAttachments();
  history = [];
  chat = null;
  pendingNote = "";
  if (wb.data) {
    wb.data.current = null;
    chats.saveWorkbook(wb.id, wb.data);
  }
  clearTranscript();
  updateLegacyBanner();
  if (!els.chatPanel.hidden) renderChatList();
  els.input.focus();
}
$("newChat").onclick = startNewChat;
$("newChat2").onclick = () => {
  startNewChat();
  els.chatPanel.hidden = true;
};

function switchToChat(id) {
  if (chat?.id === id) return (els.chatPanel.hidden = true);
  if (busy) controller?.abort();
  tts.stop();
  saveHistory();
  const next = wb.data?.chats.find((c) => c.id === id);
  if (!next) return;
  chat = next;
  history = structuredClone(next.history);
  pendingNote = "";
  wb.data.current = id;
  chats.saveWorkbook(wb.id, wb.data);
  clearTranscript();
  renderHistory();
  els.chatPanel.hidden = true;
}

function deleteChat(id) {
  if (!wb.data) return;
  wb.data.chats = wb.data.chats.filter((c) => c.id !== id);
  if (chat?.id === id) {
    chat = null;
    history = [];
    wb.data.current = null;
    clearTranscript();
  }
  chats.saveWorkbook(wb.id, wb.data);
  renderChatList();
}

function renderChatList() {
  els.chatPanelTitle.textContent = `Chats in ${wb.name || "this workbook"}`;
  els.chatList.innerHTML = "";
  const legacy = chats.legacyHistory();
  if (legacy) {
    const item = el("div", "chat-item legacy-item");
    item.append(
      el("span", "ci-text", `<span class="ci-title">${esc(chats.titleFor(legacy))}</span><span class="ci-meta">Earlier conversation, not in any workbook yet · ${legacy.length} messages</span>`)
    );
    const mv = el("button", "link-btn", "Move here");
    mv.onclick = (e) => { e.stopPropagation(); moveLegacyHere(); };
    item.append(mv);
    els.chatList.append(item);
  }
  const list = wb.data?.chats || [];
  if (!list.length && !legacy) els.chatList.append(el("div", "chat-empty", "No saved chats in this workbook yet."));
  for (const c of list) {
    const item = el("div", `chat-item${chat?.id === c.id ? " current" : ""}`);
    const msgs = c.history.filter((m) => m.role === "user" && !m._nudge).length;
    item.append(el("span", "ci-text", `<span class="ci-title">${esc(c.title || "New chat")}</span><span class="ci-meta">${esc(chats.timeAgo(c.updated))} · ${msgs} message${msgs === 1 ? "" : "s"}</span>`));
    const del = el("button", "ci-del", "Delete");
    del.title = "Delete this chat";
    del.onclick = (e) => {
      e.stopPropagation();
      if (del.classList.contains("confirm")) return deleteChat(c.id);
      del.classList.add("confirm");
      del.textContent = "Sure?";
      setTimeout(() => { del.classList.remove("confirm"); del.textContent = "Delete"; }, 3000);
    };
    item.append(del);
    item.onclick = () => switchToChat(c.id);
    els.chatList.append(item);
  }
}

$("openChats").onclick = () => {
  els.chatPanel.hidden = !els.chatPanel.hidden;
  if (!els.chatPanel.hidden) {
    els.settings.hidden = true;
    renderChatList();
  }
};

// ---------- the old single shared conversation ----------
function updateLegacyBanner() {
  const legacy = chats.legacyHistory();
  const show = Boolean(legacy) && !legacyHidden && !history.length;
  els.legacyBanner.hidden = !show;
  if (show) els.legacyTitle.textContent = `“${chats.titleFor(legacy)}”`;
}

async function moveLegacyHere() {
  const legacy = chats.legacyHistory();
  if (!legacy) return;
  if (busy) controller?.abort();
  saveHistory();
  await ensureWorkbook();
  const c = chats.newChat();
  c.history = legacy;
  c.title = chats.titleFor(legacy);
  c.updated = Date.now();
  wb.data.chats.unshift(c);
  wb.data.current = c.id;
  wb.data.name = wb.name;
  if (chats.saveWorkbook(wb.id, wb.data)) chats.clearLegacy();
  chat = c;
  history = structuredClone(legacy);
  clearTranscript();
  renderHistory();
  els.legacyBanner.hidden = true;
  els.chatPanel.hidden = true;
}
$("legacyMove").onclick = moveLegacyHere;
$("legacyHide").onclick = () => {
  legacyHidden = true;
  els.legacyBanner.hidden = true;
};

// ---------- image attachments (vision) ----------
let visionOk = null; // null = unknown

async function checkVision() {
  const info = await modelInfo(settings.model).catch(() => null);
  visionOk = info ? (info.capabilities || []).includes("vision") : null;
  els.attach.title = visionOk === false
    ? `${settings.model} can't see images (PDF/Word text still works) — pick a vision model in settings`
    : "Attach images, PDFs or Word documents (or paste with ⌘V / drag & drop)";
  return visionOk;
}

async function addFiles(files) {
  const docFiles = files.filter(isDoc);
  const images = files.filter((f) => !isDoc(f));
  if (images.length) addImages(images);
  for (const f of docFiles) addDoc(f);
}

async function addDoc(file) {
  if (pendingDocs.filter((d) => !d.error).length >= MAX_DOCS) return flashMic(`Up to ${MAX_DOCS} documents per message.`);
  if (file.size > 40 * 1024 * 1024) return flashMic(`${file.name} is larger than 40 MB.`);
  const doc = { name: file.name || "document", uploading: true };
  pendingDocs.push(doc);
  renderAttachments();
  try {
    const r = await fetch(`/docs/upload?name=${encodeURIComponent(doc.name)}`, {
      method: "POST",
      headers: { "Content-Type": "application/octet-stream" },
      body: file,
    });
    const j = await r.json();
    if (j.error) throw new Error(j.error);
    Object.assign(doc, j, { uploading: false });
  } catch (e) {
    pendingDocs.splice(pendingDocs.indexOf(doc), 1);
    flashMic(`Couldn't read ${doc.name}: ${e.message}`);
  }
  renderAttachments();
}

function docLabel(d) {
  if (d.uploading) return "reading…";
  if (d.kind === "word") return `${Number(d.words || 0).toLocaleString()} words`;
  if (d.kind === "slides") return `${d.slides} slide${d.slides === 1 ? "" : "s"}${d.charts ? ` · ${d.charts} chart${d.charts === 1 ? "" : "s"}` : ""}`;
  if (d.kind === "sheet") {
    const sheets = d.sheets || [];
    const rows = sheets.reduce((n, s) => n + (s.rows || 0), 0);
    return `${sheets.length} sheet${sheets.length === 1 ? "" : "s"} · ${rows.toLocaleString()} rows`;
  }
  return `${d.pages} page${d.pages === 1 ? "" : "s"}${d.scanned ? " · scanned" : ""}`;
}

function docIcon(d) {
  const fileExt = (/\.(\w+)$/.exec(d.name || "")?.[1] || "").toUpperCase();
  const ext = d.format ? d.format.toUpperCase() : d.kind === "pdf" || fileExt === "PDF" ? "PDF" : fileExt || "FILE";
  const cls =
    d.kind === "sheet" || /^(XLS|XLSX|XLSM|XLSB|CSV|TSV|ODS)$/.test(ext) ? " sheet"
    : d.kind === "slides" || /^(PPTX|PPSX|PPTM|PPT|ODP|KEY)$/.test(ext) ? " slides"
    : ext === "PDF" ? "" : " word";
  return el("span", `doc-ico${cls}`, esc(ext.slice(0, 4)));
}

async function addImages(files) {
  if (!files.length) return;
  if (visionOk === false || (visionOk === null && (await checkVision()) === false)) {
    return flashMic(`${settings.model} doesn't support images. Choose a vision model (like qwen3.8) in settings.`);
  }
  for (const f of files) {
    if (pendingImages.length >= MAX_IMAGES) {
      flashMic(`Up to ${MAX_IMAGES} images per message.`);
      break;
    }
    try {
      pendingImages.push(await prepareImage(f));
    } catch (e) {
      flashMic(e.message);
    }
  }
  renderAttachments();
  els.input.focus();
}

function renderAttachments() {
  els.attachments.innerHTML = "";
  els.attachments.hidden = !pendingImages.length && !pendingDocs.length;
  pendingDocs.forEach((d, idx) => {
    const chip = el("div", `doc-chip${d.uploading ? " busy" : ""}`);
    chip.title = d.title && d.title !== d.name ? `${d.name} — ${d.title}` : d.name;
    chip.append(docIcon(d), el("span", "doc-meta", `<b>${esc(d.name)}</b><small>${esc(docLabel(d))}</small>`));
    if (!d.uploading) {
      const x = el("button", "", "×");
      x.title = "Remove";
      x.onclick = () => {
        pendingDocs.splice(idx, 1);
        renderAttachments();
      };
      chip.append(x);
    }
    els.attachments.append(chip);
  });
  updateSendEnabled();
  pendingImages.forEach((im, idx) => {
    const t = el("div", "thumb");
    const i = document.createElement("img");
    i.src = im.dataUrl;
    i.alt = im.name;
    i.title = `${im.name} (${im.width}×${im.height})`;
    i.onclick = () => openLightbox(im.dataUrl);
    const x = el("button", "", "×");
    x.title = "Remove";
    x.onclick = () => {
      pendingImages.splice(idx, 1);
      renderAttachments();
    };
    t.append(i, x);
    els.attachments.append(t);
  });
}

function openLightbox(src) {
  const lb = el("div", "lightbox");
  const i = document.createElement("img");
  i.src = src;
  lb.append(i);
  lb.onclick = () => lb.remove();
  document.body.append(lb);
}

els.attach.onclick = () => els.fileInput.click();
els.fileInput.onchange = () => {
  addFiles([...els.fileInput.files]);
  els.fileInput.value = "";
};
els.input.addEventListener("paste", (e) => {
  const files = attachableFilesFrom(e.clipboardData);
  if (files.length) {
    e.preventDefault();
    addFiles(files);
  }
});

// Don't send while a document is still being read.
function updateSendEnabled() {
  els.send.disabled = pendingDocs.some((d) => d.uploading);
}
let dragDepth = 0;
document.addEventListener("dragenter", (e) => {
  if ([...(e.dataTransfer?.types || [])].includes("Files")) {
    dragDepth++;
    document.body.classList.add("dragging");
  }
});
document.addEventListener("dragleave", () => {
  if (--dragDepth <= 0) {
    dragDepth = 0;
    document.body.classList.remove("dragging");
  }
});
document.addEventListener("dragover", (e) => e.preventDefault());
document.addEventListener("drop", (e) => {
  e.preventDefault();
  dragDepth = 0;
  document.body.classList.remove("dragging");
  addFiles(attachableFilesFrom(e.dataTransfer));
});
els.model.addEventListener("change", () => { visionOk = null; checkVision(); });

// ---------- read aloud (local macOS voices) ----------
let voices = [];
const SPEAK_ICON =
  '<svg viewBox="0 0 20 20"><path d="M3.5 7.6h2.9l4.1-3.4v11.6l-4.1-3.4H3.5z" /><path d="M13.4 7.2a4 4 0 0 1 0 5.6M15.6 5a7 7 0 0 1 0 10" /></svg>';

function speakButton(text) {
  const b = el("button", "speak-btn", `${SPEAK_ICON}<span>Read aloud</span>`);
  const label = b.querySelector("span");
  b.title = "Read this reply aloud";
  b.onclick = () => {
    if (b.classList.contains("on")) return tts.stop();
    tts.speak(text, {
      voices, voice: settings.ttsVoice, rate: settings.ttsRate,
      onStart: () => { b.classList.add("on"); label.textContent = "Stop"; },
      onEnd: () => { b.classList.remove("on"); label.textContent = "Read aloud"; },
    });
  };
  return b;
}

async function loadVoices() {
  voices = await tts.listVoices();
  const groups = [
    ["English", voices.filter((v) => v.lang.startsWith("en"))],
    ["Español", voices.filter((v) => v.lang.startsWith("es"))],
    ["Other languages", voices.filter((v) => !/^(en|es)/.test(v.lang))],
  ];
  els.ttsVoice.innerHTML = '<option value="auto">Auto (match reply language)</option>';
  for (const [label, list] of groups) {
    if (!list.length) continue;
    const og = document.createElement("optgroup");
    og.label = label;
    for (const v of list) og.append(new Option(`${v.name} — ${v.lang}`, v.name));
    els.ttsVoice.append(og);
  }
  els.ttsVoice.value = voices.some((v) => v.name === settings.ttsVoice) ? settings.ttsVoice : "auto";
}

// ---------- dictation (local Whisper) ----------
let recorder = null;
let micState = "idle"; // idle | rec | busy
const micLevel = els.mic.querySelector(".level");

function setMic(state, title) {
  micState = state;
  els.mic.classList.toggle("rec", state === "rec");
  els.mic.classList.toggle("busy", state === "busy");
  els.mic.disabled = state === "busy";
  micLevel.style.height = "0";
  els.mic.title = title || (state === "rec" ? "Stop and transcribe (Esc cancels)" : "Dictate (local Whisper)");
  if (state === "rec") els.input.placeholder = "Listening… click the mic again when done";
  else if (state === "busy") els.input.placeholder = "Transcribing…";
  else els.input.placeholder = busy ? "Working… (Stop to interrupt)" : "Ask about this workbook…";
}

function insertText(text) {
  const ta = els.input;
  const before = ta.value.slice(0, ta.selectionStart ?? ta.value.length);
  const after = ta.value.slice(ta.selectionEnd ?? ta.value.length);
  const sep = before && !/\s$/.test(before) ? " " : "";
  ta.value = before + sep + text + after;
  const pos = (before + sep + text).length;
  ta.focus();
  ta.setSelectionRange(pos, pos);
  autosize();
}

async function toggleDictation() {
  if (micState === "busy") return;
  if (micState === "rec") {
    const rec = recorder;
    recorder = null;
    setMic("busy");
    try {
      const { wav, seconds } = await rec.stop();
      if (seconds < 0.4) return setMic("idle");
      const { text } = await transcribe(wav, settings.sttLang);
      setMic("idle");
      if (!text) return flashMic("Didn't catch anything — try again");
      insertText(text);
      if (settings.sttAutoSend && !busy) send(els.input.value);
    } catch (e) {
      setMic("idle");
      flashMic(e.message);
    }
    return;
  }
  tts.stop(); // don't let the mic hear the agent talking
  try {
    recorder = new Recorder({ onLevel: (l) => (micLevel.style.height = `${Math.round(l * 100)}%`) });
    await recorder.start();
    setMic("rec");
  } catch (e) {
    recorder = null;
    setMic("idle");
    flashMic(e.name === "NotAllowedError" ? "Microphone access was denied for Excel." : e.message);
  }
}

function flashMic(msg) {
  const box = el("div", "msg");
  box.append(el("div", "error-box", esc(msg)));
  els.empty.hidden = true;
  els.messages.append(box);
  scroll(true);
  setTimeout(() => box.remove(), 8000);
}

els.mic.onclick = toggleDictation;
document.addEventListener("keydown", (e) => {
  if (e.key === "Escape" && tts.isSpeaking()) tts.stop();
  if (e.key === "Escape" && micState === "rec") {
    recorder?.cancel();
    recorder = null;
    setMic("idle");
  }
});
sttStatus().then((s) => {
  if (!s.available) els.mic.title = `Dictation not set up yet — ${s.hint}`;
});

// ---------- boot ----------
syncSettingsUI();
loadVoices();
refreshModels();
setInterval(refreshModels, 30000);

async function boot(info) {
  inExcel = info?.host === Office.HostType.Excel;
  if (inExcel) {
    document.body.classList.add("in-office");
    Excel.run(async (context) => {
      context.workbook.onSelectionChanged.add(updateCtxChip);
      await context.sync();
    }).catch(() => {});
    updateCtxChip();
    try {
      const ident = await chats.readWorkbookIdentity();
      wb.name = ident.name;
      wb.id = ident.id; // null until the first message in this workbook
    } catch (e) {
      console.warn("Couldn't read workbook identity", e);
    }
  } else {
    els.ctxChip.textContent = "Preview mode — open inside Excel for workbook tools";
    wb.id = "preview";
    wb.name = "preview";
  }
  if (wb.id) {
    wb.data = chats.loadWorkbook(wb.id);
    chat = wb.data.chats.find((c) => c.id === wb.data.current) || null;
    history = chat ? structuredClone(chat.history) : [];
  }
  renderHistory();
  updateLegacyBanner();
}
if (window.Office?.onReady) Office.onReady(boot);
else boot(null);
