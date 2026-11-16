// Read agent replies aloud with macOS system voices via the local server (`say`).
// Fully offline. One utterance at a time; starting a new one stops the previous.

let current = null; // { id, onEnd }
let seq = 0;

export async function listVoices() {
  try {
    const r = await fetch("/tts/voices");
    return (await r.json()).voices || [];
  } catch {
    return [];
  }
}

// Markdown -> natural spoken text. Tables and code are summarized, not read cell by cell.
export function toSpeech(md, lang = "en") {
  const es = lang.startsWith("es");
  let t = String(md || "");
  t = t.replace(/```[\s\S]*?```/g, es ? " (código omitido). " : " (code omitted). ");
  // Markdown tables: keep a short spoken pointer instead of reading every cell.
  t = t.replace(/(^\|.*\|[ \t]*\n?)+/gm, es ? " La tabla está en el panel. \n" : " The table is in the panel. \n");
  t = t
    .replace(/!\[[^\]]*\]\([^)]*\)/g, "")
    .replace(/\[([^\]]+)\]\([^)]*\)/g, "$1")
    .replace(/`([^`]+)`/g, "$1")
    .replace(/(\*\*|__)(.*?)\1/g, "$2")
    .replace(/(^|\s)[*_](\S[^*_]*?)[*_](?=\s|[.,;:!?)]|$)/g, "$1$2")
    .replace(/^#{1,6}\s*/gm, "")
    .replace(/^\s*[-*+]\s+/gm, "")
    .replace(/^\s*(\d+)\.\s+/gm, "$1. ")
    .replace(/^>\s?/gm, "");
  // Cell ranges: "B2:E2" -> "B2 to E2" (or "a" in Spanish); "Sheet1!A1" -> "Sheet1, A1".
  t = t.replace(/\b([A-Z]{1,3}\$?\d+):\$?([A-Z]{1,3}\$?\d+)\b/g, es ? "$1 a $2" : "$1 to $2");
  t = t.replace(/([A-Za-z0-9_]+)!([A-Z]{1,3}\d+)/g, "$1, $2");
  t = t.replace(/\$/g, (m, i, s) => (/\d/.test(s[i + 1] || "") ? m : ""));
  // Emoji and decorative symbols.
  t = t.replace(/[\u{1F000}-\u{1FAFF}\u{2600}-\u{27BF}\u{FE0F}\u{2190}-\u{21FF}\u{2B00}-\u{2BFF}]/gu, "");
  // Line breaks become sentence pauses.
  t = t
    .split("\n")
    .map((l) => l.trim())
    .filter(Boolean)
    .map((l) => (/[.!?:;,]$/.test(l) ? l : l + "."))
    .join(" ");
  return t.replace(/\s{2,}/g, " ").trim();
}

// Tiny English/Spanish detector (the two languages this panel is used in).
export function detectLang(text) {
  const words = String(text).toLowerCase().match(/[a-záéíóúñü]+/g) || [];
  const es = new Set("el la los las de del que y en un una por para con es está son se su sus al lo como más pero hoja celda columna fila tabla total ventas".split(" "));
  const en = new Set("the and of to in is are for with on this that it as by from was were be sheet cell column row table total sales".split(" "));
  let s = 0;
  for (const w of words) {
    if (es.has(w)) s++;
    else if (en.has(w)) s--;
  }
  if (/[ñ¿¡áéíóú]/i.test(text)) s += 2;
  return s > 0 ? "es" : "en";
}

const PREFERRED = { en: ["Samantha", "Daniel", "Karen", "Moira"], es: ["Paulina", "Mónica", "Flo (Spanish (Mexico))"] };

export function pickVoice(voices, lang, chosen) {
  if (chosen && chosen !== "auto") return chosen;
  const forLang = voices.filter((v) => v.lang.startsWith(lang));
  // Downloaded Enhanced/Premium voices sound far better — prefer them when present.
  const hq = forLang.find((v) => /\((Premium|Enhanced)\)/.test(v.name));
  if (hq) return hq.name;
  for (const n of PREFERRED[lang] || []) if (forLang.some((v) => v.name === n)) return n;
  return forLang[0]?.name;
}

/** Speak markdown text. Resolves when finished or interrupted. */
export async function speak(md, { voices, voice = "auto", rate = 190, onStart, onEnd } = {}) {
  await stop();
  const lang = detectLang(md);
  const text = toSpeech(md, lang);
  if (!text) return;
  const id = ++seq;
  current = { id, onEnd };
  onStart?.();
  try {
    await fetch("/tts/speak", {
      method: "POST",
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify({ text, voice: pickVoice(voices || [], lang, voice), rate }),
    });
  } catch {
    /* server gone or request aborted */
  } finally {
    if (current?.id === id) {
      current = null;
      onEnd?.();
    }
  }
}

export async function stop() {
  const c = current;
  current = null;
  if (c) {
    c.onEnd?.();
    await fetch("/tts/stop", { method: "POST" }).catch(() => {});
  }
}

export const isSpeaking = () => Boolean(current);
