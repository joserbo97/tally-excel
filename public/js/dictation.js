// Microphone capture -> 16 kHz mono WAV -> local whisper.cpp via /stt/transcribe.
// Audio never leaves this Mac.

const TARGET_RATE = 16000;

export async function sttStatus() {
  try {
    const r = await fetch("/stt/status");
    return await r.json();
  } catch {
    return { available: false, hint: "Add-in server not reachable." };
  }
}

export class Recorder {
  constructor({ onLevel } = {}) {
    this.onLevel = onLevel;
    this.chunks = [];
  }

  async start() {
    if (!navigator.mediaDevices?.getUserMedia) {
      throw new Error("This Excel panel can't access the microphone (getUserMedia unavailable).");
    }
    this.stream = await navigator.mediaDevices.getUserMedia({
      audio: { channelCount: 1, echoCancellation: true, noiseSuppression: true, autoGainControl: true },
    });
    const Ctx = window.AudioContext || window.webkitAudioContext;
    this.ctx = new Ctx();
    if (this.ctx.state === "suspended") await this.ctx.resume();
    this.rate = this.ctx.sampleRate;
    this.source = this.ctx.createMediaStreamSource(this.stream);
    // ScriptProcessor is deprecated but universally supported (incl. Office's WKWebView).
    this.proc = this.ctx.createScriptProcessor(4096, 1, 1);
    this.proc.onaudioprocess = (e) => {
      const data = e.inputBuffer.getChannelData(0);
      this.chunks.push(new Float32Array(data));
      if (this.onLevel) {
        let sum = 0;
        for (let i = 0; i < data.length; i += 8) sum += data[i] * data[i];
        this.onLevel(Math.min(1, Math.sqrt(sum / (data.length / 8)) * 4));
      }
    };
    this.source.connect(this.proc);
    this.proc.connect(this.ctx.destination);
    this.startedAt = performance.now();
  }

  async stop() {
    this.proc?.disconnect();
    this.source?.disconnect();
    this.stream?.getTracks().forEach((t) => t.stop());
    await this.ctx?.close().catch(() => {});
    const seconds = (performance.now() - this.startedAt) / 1000;
    const pcm = concat(this.chunks);
    this.chunks = [];
    return { wav: encodeWav(downsample(pcm, this.rate, TARGET_RATE), TARGET_RATE), seconds };
  }

  cancel() {
    this.proc?.disconnect();
    this.source?.disconnect();
    this.stream?.getTracks().forEach((t) => t.stop());
    this.ctx?.close().catch(() => {});
    this.chunks = [];
  }
}

export async function transcribe(wavBlob, lang = "auto") {
  const r = await fetch(`/stt/transcribe?lang=${encodeURIComponent(lang)}`, {
    method: "POST",
    headers: { "Content-Type": "audio/wav" },
    body: wavBlob,
  });
  const j = await r.json().catch(() => ({ error: `HTTP ${r.status}` }));
  if (!r.ok || j.error) throw new Error(j.error || `HTTP ${r.status}`);
  return j;
}

// ---------- audio helpers ----------
function concat(chunks) {
  const out = new Float32Array(chunks.reduce((n, c) => n + c.length, 0));
  let o = 0;
  for (const c of chunks) {
    out.set(c, o);
    o += c.length;
  }
  return out;
}

// Box-filter downsample (averaging) — adequate anti-aliasing for speech.
export function downsample(buf, from, to) {
  if (from === to) return buf;
  const ratio = from / to;
  const out = new Float32Array(Math.floor(buf.length / ratio));
  for (let i = 0; i < out.length; i++) {
    const start = Math.floor(i * ratio);
    const end = Math.min(buf.length, Math.floor((i + 1) * ratio));
    let sum = 0;
    for (let j = start; j < end; j++) sum += buf[j];
    out[i] = sum / Math.max(1, end - start);
  }
  return out;
}

export function encodeWav(samples, rate) {
  const buf = new ArrayBuffer(44 + samples.length * 2);
  const v = new DataView(buf);
  const str = (o, s) => [...s].forEach((ch, i) => v.setUint8(o + i, ch.charCodeAt(0)));
  str(0, "RIFF");
  v.setUint32(4, 36 + samples.length * 2, true);
  str(8, "WAVE");
  str(12, "fmt ");
  v.setUint32(16, 16, true);
  v.setUint16(20, 1, true); // PCM
  v.setUint16(22, 1, true); // mono
  v.setUint32(24, rate, true);
  v.setUint32(28, rate * 2, true);
  v.setUint16(32, 2, true);
  v.setUint16(34, 16, true);
  str(36, "data");
  v.setUint32(40, samples.length * 2, true);
  for (let i = 0; i < samples.length; i++) {
    const s = Math.max(-1, Math.min(1, samples[i]));
    v.setInt16(44 + i * 2, s < 0 ? s * 0x8000 : s * 0x7fff, true);
  }
  return new Blob([buf], { type: "audio/wav" });
}
