// Image attachments for the vision model: downscale, encode, hand back base64.

const MAX_SIDE = 1600; // Retina screenshots are huge; this keeps text legible but prompts fast.
export const MAX_IMAGES = 4;

export async function prepareImage(file) {
  if (!file || !/^image\//.test(file.type)) throw new Error(`${file?.name || "File"} isn't an image.`);
  if (file.size > 30 * 1024 * 1024) throw new Error(`${file.name} is too large (max 30 MB).`);
  const url = URL.createObjectURL(file);
  try {
    const img = await new Promise((resolve, reject) => {
      const i = new Image();
      i.onload = () => resolve(i);
      i.onerror = () => reject(new Error(`Couldn't read ${file.name || "the image"}.`));
      i.src = url;
    });
    const scale = Math.min(1, MAX_SIDE / Math.max(img.naturalWidth, img.naturalHeight));
    const w = Math.max(1, Math.round(img.naturalWidth * scale));
    const h = Math.max(1, Math.round(img.naturalHeight * scale));
    const canvas = document.createElement("canvas");
    canvas.width = w;
    canvas.height = h;
    const ctx = canvas.getContext("2d");
    ctx.fillStyle = "#fff"; // flatten transparency (JPEG has no alpha)
    ctx.fillRect(0, 0, w, h);
    ctx.imageSmoothingQuality = "high";
    ctx.drawImage(img, 0, 0, w, h);
    const dataUrl = canvas.toDataURL("image/jpeg", 0.9);
    return { dataUrl, b64: dataUrl.split(",")[1], width: w, height: h, name: file.name || "screenshot" };
  } finally {
    URL.revokeObjectURL(url);
  }
}

const DOC_TYPES =
  /^(application\/(pdf|msword|rtf|vnd\.openxmlformats-officedocument\.(wordprocessingml\.document|spreadsheetml\.sheet|presentationml\.presentation)|vnd\.oasis\.opendocument\.(text|spreadsheet|presentation)|vnd\.ms-excel(\.sheet\.macroenabled\.12)?|vnd\.ms-powerpoint)|text\/(rtf|csv|tab-separated-values))$/i;

// Documents, spreadsheets and presentations (read on the server), as opposed to images (sent to vision).
export const isDoc = (f) =>
  f && (DOC_TYPES.test(f.type) || /\.(pdf|docx?|rtf|odt|xlsx|xlsm|xlsb|xls|ods|csv|tsv|pptx|ppsx|pptm|ppt|odp|key)$/i.test(f.name || ""));

// Images and documents from a paste / drop / file picker.
export function attachableFilesFrom(dataTransfer) {
  if (!dataTransfer) return [];
  const files = [...(dataTransfer.files || [])];
  const fromItems = [...(dataTransfer.items || [])].filter((it) => it.kind === "file").map((it) => it.getAsFile());
  const all = files.length ? files : fromItems;
  return all.filter((f) => f && (f.type.startsWith("image/") || isDoc(f)));
}

export function imageFilesFrom(dataTransfer) {
  if (!dataTransfer) return [];
  const files = [...(dataTransfer.files || [])].filter((f) => f.type.startsWith("image/"));
  if (files.length) return files;
  // Pasted screenshots arrive as items, not files.
  return [...(dataTransfer.items || [])]
    .filter((it) => it.kind === "file" && it.type.startsWith("image/"))
    .map((it) => it.getAsFile())
    .filter(Boolean);
}
