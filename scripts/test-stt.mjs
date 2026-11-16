// Tests the dictation pipeline without a microphone:
// 48 kHz float WAV (e.g. from macOS `say`) -> the panel's downsample + WAV encoder -> /stt/transcribe.
//   say -o /tmp/t.wav --data-format=LEF32@48000 "hello world"
//   NODE_EXTRA_CA_CERTS=~/.office-addin-dev-certs/ca.crt node scripts/test-stt.mjs /tmp/t.wav [lang]
import fs from "node:fs";
import { downsample, encodeWav } from "../public/js/dictation.js";

const [file, lang = "auto"] = process.argv.slice(2);
const buf = fs.readFileSync(file);

// Minimal RIFF parser for 32-bit float (fmt 3) or 16-bit PCM (fmt 1) mono WAV.
let off = 12, fmt, rate, bits, data;
while (off < buf.length) {
  const id = buf.toString("ascii", off, off + 4), size = buf.readUInt32LE(off + 4);
  if (id === "fmt ") { fmt = buf.readUInt16LE(off + 8); rate = buf.readUInt32LE(off + 12); bits = buf.readUInt16LE(off + 22); }
  if (id === "data") { data = buf.subarray(off + 8, off + 8 + size); break; }
  off += 8 + size + (size % 2);
}
const n = data.length / (bits / 8);
const pcm = new Float32Array(n);
for (let i = 0; i < n; i++) pcm[i] = fmt === 3 ? data.readFloatLE(i * 4) : data.readInt16LE(i * 2) / 32768;

const wav = encodeWav(downsample(pcm, rate, 16000), 16000);
console.log(`input ${rate} Hz, ${(n / rate).toFixed(2)}s -> 16 kHz WAV ${wav.size} bytes`);

const t0 = Date.now();
const r = await fetch(`https://localhost:3443/stt/transcribe?lang=${lang}`, { method: "POST", headers: { "Content-Type": "audio/wav" }, body: wav });
const j = await r.json();
console.log(`HTTP ${r.status} in ${Date.now() - t0} ms:`, j);
