// Renders the add-in's PNG icons from the Tally mark SVGs (needs: brew install librsvg).
//   npm run icons
// 16 and 32 px use the simplified small mark; 64, 80 and 128 px use the full mark.
import { execFileSync } from "node:child_process";
import path from "node:path";
import { fileURLToPath } from "node:url";

const assets = path.join(path.dirname(fileURLToPath(import.meta.url)), "../public/assets");
for (const size of [16, 32, 64, 80, 128]) {
  const src = path.join(assets, size <= 32 ? "tally-mark-small.svg" : "tally-mark.svg");
  execFileSync("rsvg-convert", ["-w", String(size), "-h", String(size), src, "-o", path.join(assets, `icon-${size}.png`)]);
  console.log(`icon-${size}.png`);
}
