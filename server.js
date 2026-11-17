// Local server for the Tally Excel add-in.
// - Serves the task pane over HTTPS (Office requires HTTPS for add-in pages).
// - Proxies /ollama/* to the local Ollama server, so the task pane never hits
//   CORS or mixed-content (https page -> http Ollama) problems.
import https from "node:https";
import http from "node:http";
import fs from "node:fs";
import path from "node:path";
import os from "node:os";
import { fileURLToPath } from "node:url";
import { execFile, spawn } from "node:child_process";
import { webSearch, webFetch, uploadDoc, readDoc, docValues, docFileBase64 } from "./web.js";

const __dirname = path.dirname(fileURLToPath(import.meta.url));
const PORT = Number(process.env.PORT || 3443);
const OLLAMA = new URL(process.env.OLLAMA_HOST || "http://127.0.0.1:11434");
const CERT_DIR = path.join(os.homedir(), ".office-addin-dev-certs");
const USE_HTTP = process.env.HTTP === "1";

const PUBLIC = path.join(__dirname, "public");
// Office.js served locally so the add-in works offline (office.js loads its host/locale files
// relative to its own URL, so the whole dist folder is exposed under /vendor/office/).
const OFFICE_DIST = path.join(__dirname, "node_modules/@microsoft/office-js/dist");
let servedOffice = false; // logged once, to confirm the local copy is in use
const VENDOR = {
  "/vendor/marked.esm.js": path.join(__dirname, "node_modules/marked/lib/marked.esm.js"),
  "/vendor/purify.es.mjs": path.join(__dirname, "node_modules/dompurify/dist/purify.es.mjs"),
};

const TYPES = {
  ".html": "text/html; charset=utf-8",
  ".js": "text/javascript; charset=utf-8",
  ".mjs": "text/javascript; charset=utf-8",
  ".css": "text/css; charset=utf-8",
  ".png": "image/png",
  ".svg": "image/svg+xml",
  ".json": "application/json",
  ".xml": "application/xml",
};

function serveStatic(req, res) {
  const urlPath = decodeURIComponent(new URL(req.url, "http://x").pathname);
  let file = VENDOR[urlPath];
  let cache = "no-store";
  if (!file && urlPath.startsWith("/vendor/office/")) {
    file = path.join(OFFICE_DIST, path.normalize(urlPath.slice("/vendor/office".length)));
    if (!file.startsWith(OFFICE_DIST + path.sep)) return send(res, 403, "Forbidden");
    cache = "max-age=86400"; // pinned npm version; fine to cache for a day
    if (!servedOffice && /\/office\.js$/.test(urlPath)) {
      servedOffice = true;
      console.log(`Office.js served locally (${urlPath}) — the add-in no longer needs Microsoft's CDN.`);
    }
  } else if (!file) {
    const rel = urlPath === "/" ? "/taskpane.html" : urlPath;
    file = path.join(PUBLIC, path.normalize(rel));
    if (!file.startsWith(PUBLIC + path.sep)) return send(res, 403, "Forbidden");
  }
  fs.readFile(file, (err, data) => {
    if (err) return send(res, 404, "Not found");
    res.writeHead(200, {
      "Content-Type": TYPES[path.extname(file)] || "application/octet-stream",
      "Cache-Control": cache,
    });
    res.end(data);
  });
}

function send(res, code, body) {
  res.writeHead(code, { "Content-Type": "text/plain" });
  res.end(body);
}

function proxyOllama(req, res) {
  const target = new URL(req.url.replace(/^\/ollama/, "") || "/", OLLAMA);
  const upstream = http.request(
    target,
    { method: req.method, headers: { "content-type": req.headers["content-type"] || "application/json" } },
    (up) => {
      res.writeHead(up.statusCode || 502, {
        "Content-Type": up.headers["content-type"] || "application/json",
        "Cache-Control": "no-store",
      });
      up.pipe(res);
    }
  );
  upstream.on("error", (e) => {
    if (!res.headersSent) send(res, 502, `Cannot reach Ollama at ${OLLAMA.origin}: ${e.message}`);
    else res.end();
  });
  // If the task pane aborts (Stop button), stop generation upstream too.
  res.on("close", () => upstream.destroy());
  req.pipe(upstream);
}

// ---------- local speech-to-text (whisper.cpp) ----------
const WHISPER_BINS = ["/opt/homebrew/bin/whisper-cli", "/usr/local/bin/whisper-cli", path.join(__dirname, "bin/whisper-cli")];
const WHISPER_MODELS = ["ggml-large-v3-turbo-q5_0.bin", "ggml-large-v3-turbo.bin", "ggml-small.bin", "ggml-base.bin", "ggml-base.en.bin"];

function sttStatus() {
  // Re-checked on each call, so installing whisper or a model later works without a restart.
  const bin = process.env.WHISPER_BIN || WHISPER_BINS.find((p) => fs.existsSync(p));
  const model =
    process.env.WHISPER_MODEL || WHISPER_MODELS.map((f) => path.join(__dirname, "models", f)).find((p) => fs.existsSync(p));
  return {
    available: Boolean(bin && model),
    bin: bin || null,
    model: model || null,
    hint: !bin ? "Install whisper.cpp:  brew install whisper-cpp" : !model ? "Download a model:  npm run stt-model" : null,
  };
}

function transcribe(req, res) {
  const st = sttStatus();
  if (!st.available) return sendJson(res, 503, { error: `Speech-to-text not set up. ${st.hint}` });
  const lang = (new URL(req.url, "http://x").searchParams.get("lang") || "auto").replace(/[^a-z]/gi, "").slice(0, 8) || "auto";
  readBody(req, 50 * 1024 * 1024, (err, audio) => { // ~25 min of 16 kHz mono audio
    if (err) return bodyError(res, err);
    const wav = path.join(os.tmpdir(), `oxa-${process.pid}-${Date.now()}.wav`);
    fs.writeFileSync(wav, audio);
    const t0 = Date.now();
    execFile(
      st.bin,
      ["-m", st.model, "-f", wav, "-l", lang, "-nt", "-np"],
      { timeout: 120000, maxBuffer: 4 * 1024 * 1024 },
      (err, stdout, stderr) => {
        fs.unlink(wav, () => {});
        if (err) return sendJson(res, 500, { error: `whisper failed: ${(stderr || err.message).trim().slice(-400)}` });
        const text = stdout.replace(/\[[^\]]*\]/g, " ").replace(/\s+/g, " ").trim();
        sendJson(res, 200, { text, ms: Date.now() - t0, model: path.basename(st.model) });
      }
    );
  });
}

// ---------- local text-to-speech (macOS `say`, offline system voices) ----------
const NOVELTY_VOICES = /^(Albert|Bad News|Bahh|Bells|Boing|Bubbles|Cellos|Fred|Good News|Jester|Junior|Organ|Ralph|Superstar|Trinoids|Whisper|Wobble|Zarvox)$/;
let voiceCache = null;
let speaking = null; // { proc, id }

function listVoices(cb) {
  if (voiceCache) return cb(voiceCache);
  execFile("/usr/bin/say", ["-v", "?"], (err, stdout) => {
    if (err) return cb([]);
    voiceCache = stdout
      .split("\n")
      .map((l) => /^(.+?)\s+([a-z]{2,3}_[A-Z0-9]{2,3})\s+#/.exec(l))
      .filter(Boolean)
      .map(([, name, lang]) => ({ name: name.trim(), lang: lang.replace("_", "-") }))
      .filter((v) => !NOVELTY_VOICES.test(v.name));
    cb(voiceCache);
  });
}

function stopSpeaking() {
  if (speaking) {
    speaking.proc.kill("SIGTERM");
    speaking = null;
  }
}

// Collects a request body up to `limit` bytes. Past the limit it keeps draining (without storing)
// so we can still answer with 413 — destroying the socket would leave the client hanging.
function readBody(req, limit, cb) {
  const chunks = [];
  let size = 0;
  let tooBig = false;
  req.on("data", (d) => {
    size += d.length;
    if (size > limit) tooBig = true;
    else chunks.push(d);
  });
  req.on("end", () => cb(tooBig ? Object.assign(new Error(`Request body is larger than ${Math.round(limit / 1048576)} MB.`), { status: 413 }) : null, Buffer.concat(chunks)));
  req.on("error", (e) => cb(e));
}

function readJsonBody(req, cb, limit = 1024 * 1024) {
  readBody(req, limit, (err, buf) => {
    if (err) return cb(err);
    try {
      cb(null, JSON.parse(buf.toString("utf8") || "{}"));
    } catch {
      cb(Object.assign(new Error("Bad JSON"), { status: 400 }));
    }
  });
}

const bodyError = (res, err) => sendJson(res, err.status || 400, { error: err.message });

function speak(req, res) {
  readJsonBody(req, (err, body) => {
    if (err) return bodyError(res, err);
    const text = String(body.text || "").slice(0, 20000);
    if (!text.trim()) return sendJson(res, 400, { error: "Nothing to say" });
    listVoices((voices) => {
      stopSpeaking();
      const args = [];
      const voice = voices.find((v) => v.name === body.voice);
      if (voice) args.push("-v", voice.name);
      const rate = Math.round(Number(body.rate));
      if (rate >= 90 && rate <= 400) args.push("-r", String(rate));
      args.push("-f", "-"); // text via stdin: no arg-length or quoting issues
      const proc = spawn("/usr/bin/say", args, { stdio: ["pipe", "ignore", "pipe"] });
      const me = { proc, id: Date.now() };
      speaking = me;
      let stderr = "";
      proc.stderr.on("data", (d) => (stderr += d));
      proc.on("close", (code, signal) => {
        if (speaking === me) speaking = null;
        if (!res.headersSent) sendJson(res, 200, { done: true, interrupted: Boolean(signal), error: code && !signal ? stderr.trim() : undefined });
      });
      proc.stdin.end(text);
      // If the panel goes away mid-speech, don't keep talking.
      res.on("close", () => {
        if (!res.writableEnded && speaking === me) stopSpeaking();
      });
    });
  });
}

function sendJson(res, code, obj) {
  res.writeHead(code, { "Content-Type": "application/json", "Cache-Control": "no-store" });
  res.end(JSON.stringify(obj));
}

// ---------- request origin checks ----------
// Only the add-in's own page (or a non-browser client, which sends no Origin) may call the API.
// Without this, any website could POST here blindly (simple requests skip CORS preflight):
// make the Mac speak, upload files, or reach Ollama through the proxy — which drops the
// browser's Origin and so would bypass Ollama's own OLLAMA_ORIGINS protection.
const SCHEME = USE_HTTP ? "http" : "https";
const ALLOWED_HOSTS = new Set([`localhost:${PORT}`, `127.0.0.1:${PORT}`, `[::1]:${PORT}`]);
const ALLOWED_ORIGINS = new Set([...ALLOWED_HOSTS].map((h) => `${SCHEME}://${h}`));

function allowed(req) {
  // Host check defeats DNS rebinding (attacker.com resolving to 127.0.0.1).
  if (!ALLOWED_HOSTS.has(String(req.headers.host || "").toLowerCase())) return false;
  const origin = req.headers.origin;
  if (!origin) return true; // same-origin GETs, curl, the test scripts
  return ALLOWED_ORIGINS.has(origin.toLowerCase());
}

function handler(req, res) {
  if (!allowed(req)) return send(res, 403, "Forbidden: requests must come from the add-in.");
  if (req.url.startsWith("/ollama/")) return proxyOllama(req, res);
  if (req.url === "/stt/status") return sendJson(res, 200, sttStatus());
  if (req.url.startsWith("/stt/transcribe") && req.method === "POST") return transcribe(req, res);
  if (req.url.startsWith("/docs/upload") && req.method === "POST") {
    const name = new URL(req.url, "http://x").searchParams.get("name") || "document";
    return readBody(req, 40 * 1024 * 1024, async (err, buf) => {
      if (err) return bodyError(res, Object.assign(err, { message: err.status === 413 ? "File is larger than 40 MB." : err.message }));
      try {
        sendJson(res, 200, await uploadDoc(buf, name));
      } catch (e) {
        sendJson(res, 200, { error: e.message });
      }
    });
  }
  if (req.url === "/docs/read" && req.method === "POST") {
    return readJsonBody(req, async (err, b) => {
      if (err) return bodyError(res, err);
      try {
        sendJson(res, 200, await readDoc(b.doc_id, { maxChars: b.max_chars, start: b.start, page: b.page, asImage: b.as_image, sheet: b.sheet }));
      } catch (e) {
        sendJson(res, 200, { error: e.message });
      }
    });
  }
  if ((req.url === "/docs/values" || req.url === "/docs/file") && req.method === "POST") {
    return readJsonBody(req, (err, b) => {
      if (err) return bodyError(res, err);
      try {
        sendJson(res, 200, req.url === "/docs/values"
          ? docValues(b.doc_id, { sheet: b.sheet, range: b.range, includeFormulas: !!b.include_formulas })
          : docFileBase64(b.doc_id));
      } catch (e) {
        sendJson(res, 200, { error: e.message });
      }
    });
  }
  if ((req.url === "/web/search" || req.url === "/web/fetch") && req.method === "POST") {
    return readJsonBody(req, async (err, b) => {
      if (err) return bodyError(res, err);
      try {
        const out = req.url === "/web/search"
          ? await webSearch(b.query, b.max_results)
          : await webFetch(b.url, { maxChars: b.max_chars, start: b.start, page: b.page, asImage: b.as_image });
        sendJson(res, 200, out);
      } catch (e) {
        sendJson(res, 200, { error: e.message });
      }
    });
  }
  if (req.url === "/tts/voices") return listVoices((v) => sendJson(res, 200, { voices: v }));
  if (req.url === "/tts/speak" && req.method === "POST") return speak(req, res);
  if (req.url === "/tts/stop" && req.method === "POST") {
    stopSpeaking();
    return sendJson(res, 200, { stopped: true });
  }
  if (req.method !== "GET") return send(res, 405, "Method not allowed");
  serveStatic(req, res);
}

let makeServer;
if (USE_HTTP) {
  makeServer = () => http.createServer(handler);
} else {
  const cert = path.join(CERT_DIR, "localhost.crt");
  const key = path.join(CERT_DIR, "localhost.key");
  if (!fs.existsSync(cert) || !fs.existsSync(key)) {
    console.error(
      "No dev certificate found in ~/.office-addin-dev-certs.\n" +
        "Run once:  npm run certs   (macOS will ask for your password to trust it)"
    );
    process.exit(1);
  }
  const tls = { cert: fs.readFileSync(cert), key: fs.readFileSync(key) };
  makeServer = () => https.createServer(tls, handler);
}

// Loopback only (IPv4 + IPv6) — never exposed to the network.
for (const host of ["127.0.0.1", "::1"]) {
  makeServer()
    .on("error", (e) => {
      if (host === "::1" && e.code === "EADDRNOTAVAIL") return;
      console.error(`Listen on ${host}:${PORT} failed: ${e.message}`);
      process.exit(1);
    })
    .listen(PORT, host);
}
console.log(`Tally running at ${USE_HTTP ? "http" : "https"}://localhost:${PORT}`);
console.log(`Proxying /ollama -> ${OLLAMA.origin}`);
