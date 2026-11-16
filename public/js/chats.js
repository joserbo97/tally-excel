/* global Excel */
// Per-workbook chat storage.
// - Each workbook gets a hidden id stored in the file itself (add-in setting "oxa.workbookId"),
//   written lazily on the first message so merely opening the panel doesn't dirty the workbook.
// - localStorage key "oxa.wb.<id>" holds { name, chats: [{ id, title, created, updated, history }], current }.
// - The pre-per-workbook global chat lives in "oxa.history" until the user moves it somewhere.

const PREFIX = "oxa.wb.";
const LEGACY_KEY = "oxa.history";
const MAX_CHATS_PER_WORKBOOK = 30;

const rid = () => Math.random().toString(36).slice(2, 10) + Date.now().toString(36).slice(-4);

function read(key, fallback) {
  try {
    const v = localStorage.getItem(key);
    return v ? JSON.parse(v) : fallback;
  } catch {
    return fallback;
  }
}

// Write; if storage is full, drop the oldest chats (from any workbook, never the one being saved) and retry.
function write(key, value, protectChatId) {
  for (let attempt = 0; attempt < 40; attempt++) {
    try {
      localStorage.setItem(key, JSON.stringify(value));
      return true;
    } catch {
      if (!pruneOldest(protectChatId, key, value)) return false;
    }
  }
  return false;
}

function pruneOldest(protectChatId, pendingKey, pendingValue) {
  let oldest = null;
  const consider = (key, data) => {
    for (const c of data?.chats || []) {
      if (c.id === protectChatId) continue;
      if (!oldest || c.updated < oldest.updated) oldest = { key, id: c.id, updated: c.updated };
    }
  };
  try {
    for (let i = 0; i < localStorage.length; i++) {
      const k = localStorage.key(i);
      if (k?.startsWith(PREFIX) && k !== pendingKey) consider(k, read(k));
    }
  } catch {
    return false;
  }
  consider(pendingKey, pendingValue);
  if (!oldest) return false;
  if (oldest.key === pendingKey) {
    pendingValue.chats = pendingValue.chats.filter((c) => c.id !== oldest.id);
  } else {
    const d = read(oldest.key);
    d.chats = d.chats.filter((c) => c.id !== oldest.id);
    try {
      localStorage.setItem(oldest.key, JSON.stringify(d));
    } catch {
      /* will retry */
    }
  }
  return true;
}

// ---------- workbook identity ----------
export async function readWorkbookIdentity() {
  return Excel.run(async (context) => {
    const s = context.workbook.settings.getItemOrNullObject("oxa.workbookId");
    s.load("value");
    context.workbook.load("name");
    await context.sync();
    return { id: s.isNullObject ? null : String(s.value), name: context.workbook.name || "" };
  });
}

export async function assignWorkbookId(id) {
  await Excel.run(async (context) => {
    context.workbook.settings.add("oxa.workbookId", id);
    await context.sync();
  });
}

export const newId = rid;

// ---------- chat store ----------
export function loadWorkbook(wbId) {
  return read(PREFIX + wbId, { name: "", chats: [], current: null });
}

export function saveWorkbook(wbId, data) {
  data.chats.sort((a, b) => b.updated - a.updated);
  if (data.chats.length > MAX_CHATS_PER_WORKBOOK) data.chats.length = MAX_CHATS_PER_WORKBOOK;
  return write(PREFIX + wbId, data, data.current);
}

export function newChat() {
  const now = Date.now();
  return { id: rid(), title: "", created: now, updated: now, history: [] };
}

// ---------- legacy (pre per-workbook) conversation ----------
export function legacyHistory() {
  const h = read(LEGACY_KEY, []);
  return Array.isArray(h) && h.length ? h : null;
}

export function clearLegacy() {
  try {
    localStorage.removeItem(LEGACY_KEY);
  } catch {
    /* ignore */
  }
}

// ---------- stored user-message format ----------
// "[Excel context: …]\n\n" + optional "[Note: …]\n\n" + the user's words + ATTACH_MARK + attachment notes.
export const ATTACH_MARK = "\n\n<<attachments>>\n";

export function parseUserMessage(content = "") {
  const m = /^(?:\[Excel context:([^\]]*)\]\n\n)?([\s\S]*)$/.exec(content);
  const rest = m ? m[2] : content;
  return {
    ctx: m?.[1]?.trim() || "",
    text: rest.split(ATTACH_MARK)[0].replace(/^\[Note:[^\]]*\]\n\n/, "").trim(),
  };
}

// The user's own words from a stored user message (drops the Excel context line, notes and attachment notes).
export const userText = (content) => parseUserMessage(content).text;

// What the model sees: the marker becomes a plain paragraph break.
export const forModel = (content = "") => content.replace(ATTACH_MARK, "\n\n");

export function titleFor(history) {
  const first = history.find((m) => m.role === "user");
  const t = first ? userText(first.content).replace(/\s+/g, " ") : "";
  return t.length > 60 ? t.slice(0, 57) + "…" : t || "New chat";
}

export function timeAgo(ts) {
  const s = Math.round((Date.now() - ts) / 1000);
  if (s < 60) return "just now";
  if (s < 3600) return `${Math.round(s / 60)} min ago`;
  if (s < 86400) return `${Math.round(s / 3600)} h ago`;
  const d = Math.round(s / 86400);
  return d === 1 ? "yesterday" : d < 30 ? `${d} days ago` : new Date(ts).toLocaleDateString();
}
