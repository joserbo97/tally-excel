#!/bin/bash
# Tally installer for macOS (Apple silicon + Microsoft Excel for Mac).
#
#   curl -fsSL https://raw.githubusercontent.com/joserbo97/tally-excel/main/install.sh | bash
#
# Or from a clone:   ./install.sh                 install or update everything
#                    ./install.sh --check         only check this Mac, change nothing
#                    ./install.sh --with-dictation  also set up local voice dictation (~600 MB)
#
# Safe to run again: every step skips what is already done.
# Settings (environment): TALLY_DIR (default ~/tally-excel), TALLY_MODEL (default qwen3.8).
set -euo pipefail

REPO_URL="${TALLY_REPO:-https://github.com/joserbo97/tally-excel.git}"
MODEL="${TALLY_MODEL:-qwen3.8}"
MIN_OLLAMA="0.32.12"
CHECK_ONLY=0
DICTATION=0

usage() {
  cat <<'USAGE'
Tally installer for macOS (Apple silicon + Microsoft Excel for Mac).

  curl -fsSL https://raw.githubusercontent.com/joserbo97/tally-excel/main/install.sh | bash
  curl -fsSL https://raw.githubusercontent.com/joserbo97/tally-excel/main/install.sh | bash -s -- --check

Options:
  --check            only check this Mac, change nothing
  --with-dictation   also set up local voice dictation (~600 MB)

Safe to run again: every step skips what is already done.
Settings (environment): TALLY_DIR (default ~/tally-excel), TALLY_MODEL (default qwen3.8).
USAGE
}

for arg in "$@"; do
  case "$arg" in
    --check) CHECK_ONLY=1 ;;
    --with-dictation) DICTATION=1 ;;
    -h|--help) usage; exit 0 ;;
    *) echo "Unknown option: $arg (try --help)" >&2; exit 2 ;;
  esac
done

bold=$'\033[1m'; green=$'\033[32m'; yellow=$'\033[33m'; red=$'\033[31m'; reset=$'\033[0m'
step() { printf '\n%s==> %s%s\n' "$bold" "$*" "$reset"; }
ok()   { printf '  %s✓%s %s\n' "$green" "$reset" "$*"; }
warn() { printf '  %s!%s %s\n' "$yellow" "$reset" "$*"; }
fail() { printf '  %s✗ %s%s\n' "$red" "$*" "$reset" >&2; exit 1; }

# Version comparison: returns 0 if $1 >= $2.
version_ge() { [ "$(printf '%s\n%s\n' "$2" "$1" | sort -V | head -1)" = "$2" ]; }

# If this script sits in a Tally checkout, install from there; otherwise clone to TALLY_DIR.
SCRIPT_DIR=""
if [ -n "${BASH_SOURCE[0]:-}" ] && [ -f "${BASH_SOURCE[0]}" ]; then
  SCRIPT_DIR="$(cd "$(dirname "${BASH_SOURCE[0]}")" && pwd)"
fi
if [ -n "$SCRIPT_DIR" ] && [ -f "$SCRIPT_DIR/manifest.xml" ] && [ -f "$SCRIPT_DIR/server.js" ]; then
  DIR="$SCRIPT_DIR"
else
  DIR="${TALLY_DIR:-$HOME/tally-excel}"
fi

# The steps live in main(), called on the last line: with `curl | bash`, nothing runs until the whole
# script has downloaded, and no command can swallow the rest of the script from stdin.
main() {
# ---------------------------------------------------------------- 1. this Mac
step "Checking this Mac"
[ "$(uname -s)" = "Darwin" ] || fail "Tally runs on macOS only."
[ "$(uname -m)" = "arm64" ] || fail "Tally needs a Mac with Apple silicon (M1 or later) to run the model on the GPU."
mem_gb=$(( $(sysctl -n hw.memsize) / 1073741824 ))
if [ "$mem_gb" -ge 32 ]; then ok "${mem_gb} GB memory"
elif [ "$mem_gb" -ge 24 ]; then warn "${mem_gb} GB memory: the 17 GB model will fit but leave little room; 32 GB or more is recommended."
else fail "${mem_gb} GB memory: the model needs about 17 GB on its own. 32 GB or more is recommended (set TALLY_MODEL to a smaller Ollama model to try anyway)."
fi
free_gb=$(df -g "$HOME" | awk 'NR==2 {print $4}')
if [ "$free_gb" -ge 25 ]; then ok "${free_gb} GB free disk"; else warn "${free_gb} GB free disk: the model download needs about 17 GB."; fi
if [ -d "/Applications/Microsoft Excel.app" ]; then ok "Microsoft Excel is installed"
else fail "Microsoft Excel for Mac isn't in /Applications. Install it first (Microsoft 365)."
fi
# A fresh Homebrew install isn't on PATH until the shell profile is updated; find it in its usual places.
if ! command -v brew >/dev/null 2>&1; then
  for b in /opt/homebrew/bin/brew /usr/local/bin/brew; do [ -x "$b" ] && eval "$("$b" shellenv)" && break; done
fi
command -v brew >/dev/null 2>&1 && ok "Homebrew is installed" || fail "Homebrew is needed to install the tools. Install it with the one-line command at https://brew.sh, then run this again."

if [ "$CHECK_ONLY" = 1 ]; then
  step "What the install would do"
  todo() { if eval "$2" >/dev/null 2>&1; then ok "$1: already done"; else warn "$1: will be set up"; fi; }
  todo "Git" "command -v git"
  todo "Node.js 18 or later" "[ \"\$(node -p 'process.versions.node.split(\".\")[0]')\" -ge 18 ]"
  todo "PDF reader (poppler)" "command -v pdftotext"
  todo "Ollama" "command -v ollama || [ -d /Applications/Ollama.app ]"
  todo "Model $MODEL (about 17 GB)" "curl -fsS --max-time 2 http://127.0.0.1:11434/api/tags | grep -q '\"$MODEL[:\"]'"
  todo "Tally files in $DIR" "[ -f '$DIR/manifest.xml' ]"
  todo "Certificate for https://localhost (asks for your password)" "[ -f \"\$HOME/.office-addin-dev-certs/ca.crt\" ] && security verify-cert -c \"\$HOME/.office-addin-dev-certs/localhost.crt\""
  todo "Tally registered with Excel" "[ -f \"\$HOME/Library/Containers/com.microsoft.Excel/Data/Documents/wef/tally.xml\" ]"
  todo "Background server" "curl -fsk --max-time 2 https://localhost:3443/taskpane.html"
  step "Check finished: this Mac can run Tally. Nothing was changed."
  exit 0
fi

# ---------------------------------------------------------------- 2. tools
step "Installing tools"
need_brew() { # $1 = command to look for, $2 = Homebrew formula
  if command -v "$1" >/dev/null 2>&1; then ok "$2 is installed"; else brew install "$2"; ok "$2 installed"; fi
}
need_brew git git
need_brew node node
need_brew pdftotext poppler
node_major=$(node -p 'process.versions.node.split(".")[0]')
[ "$node_major" -ge 18 ] || fail "Node.js $node_major is too old; run: brew upgrade node"

# ---------------------------------------------------------------- 3. Ollama + model
step "Setting up Ollama and the model"
if ! command -v ollama >/dev/null 2>&1 && [ ! -d "/Applications/Ollama.app" ]; then
  brew install ollama
fi
ollama_up() { curl -fsS --max-time 2 http://127.0.0.1:11434/api/version >/dev/null 2>&1; }
if ! ollama_up; then
  if [ -d "/Applications/Ollama.app" ]; then open -g -a Ollama; else brew services start ollama >/dev/null; fi
  for _ in $(seq 1 30); do ollama_up && break; sleep 1; done
fi
ollama_up || fail "Ollama didn't start. Open the Ollama app (or run: brew services start ollama), then run this again."
ollama_version=$(curl -fsS http://127.0.0.1:11434/api/version | sed -E 's/.*"version":"([^"]+)".*/\1/')
version_ge "$ollama_version" "$MIN_OLLAMA" || fail "Ollama $ollama_version is too old for $MODEL; update to $MIN_OLLAMA or later (Ollama app: check for updates; Homebrew: brew upgrade ollama)."
ok "Ollama $ollama_version is running"
OLLAMA_BIN="$(command -v ollama || echo /Applications/Ollama.app/Contents/Resources/ollama)"
if "$OLLAMA_BIN" list 2>/dev/null | awk 'NR>1 {print $1}' | grep -qx "${MODEL}\(:latest\)\{0,1\}"; then
  ok "Model $MODEL is already downloaded"
else
  echo "  Downloading $MODEL (about 17 GB; this can take a while)…"
  "$OLLAMA_BIN" pull "$MODEL"
  ok "Model $MODEL downloaded"
fi

# ---------------------------------------------------------------- 4. Tally
step "Getting Tally"
if [ "$DIR" = "$SCRIPT_DIR" ]; then
  ok "Using this folder: $DIR"
elif [ -d "$DIR/.git" ]; then
  git -C "$DIR" pull --ff-only --quiet && ok "Updated $DIR"
elif [ -e "$DIR" ]; then
  fail "$DIR exists but isn't a Tally checkout. Move it, or set TALLY_DIR to another folder."
else
  git clone --quiet "$REPO_URL" "$DIR" && ok "Downloaded to $DIR"
fi
cd "$DIR"
# npm ci deletes node_modules before reinstalling; on macOS that can fail with ENOTEMPTY when Finder has
# dropped .DS_Store files into it. Remove the folder ourselves first, and retry once if it still fails.
install_deps() { rm -rf node_modules && npm ci --no-audit --no-fund --loglevel=error; }
if ! install_deps; then
  warn "Dependency install failed; retrying once with a clean folder"
  sleep 2
  install_deps || fail "npm couldn't install the dependencies. Run: cd \"$DIR\" && rm -rf node_modules && npm ci"
fi
ok "Dependencies installed"

# ---------------------------------------------------------------- 5. HTTPS certificate
step "Trusting a certificate for https://localhost (Office requires HTTPS)"
if npx --no-install office-addin-dev-certs verify >/dev/null 2>&1; then
  ok "Certificate already trusted"
else
  echo "  macOS will ask for your password to trust a certificate that only covers localhost."
  npx --no-install office-addin-dev-certs install --days 365
  ok "Certificate trusted"
fi

# ---------------------------------------------------------------- 6. Excel
step "Adding Tally to Excel"
npm run --silent sideload
ok "Tally is registered with Excel"

# ---------------------------------------------------------------- 7. background server
step "Starting the Tally server (and at every login)"
sh scripts/install-service.sh
for _ in $(seq 1 15); do
  curl -fsk --max-time 2 https://localhost:3443/taskpane.html >/dev/null 2>&1 && break; sleep 1
done
curl -fsk --max-time 2 https://localhost:3443/taskpane.html >/dev/null 2>&1 \
  && ok "Server is answering on https://localhost:3443" \
  || fail "The server didn't start. See ~/Library/Logs/Tally/server.log"

# ---------------------------------------------------------------- 8. optional dictation
if [ "$DICTATION" = 1 ]; then
  step "Setting up voice dictation (local Whisper)"
  need_brew whisper-cli whisper-cpp
  if [ -f models/ggml-large-v3-turbo-q5_0.bin ]; then ok "Speech model already downloaded"; else npm run --silent stt-model; ok "Speech model downloaded"; fi
fi

step "Done"
echo "  1. Quit Excel completely (Excel > Quit Excel) and open it again."
echo "  2. Open a workbook and click Tally at the right end of the Home tab."
echo "  If the button isn't there: Insert > Add-ins > My Add-ins > Developer Add-ins > Tally."
[ "$DICTATION" = 1 ] || echo "  Voice dictation is optional: run ./install.sh --with-dictation in $DIR to add it."
echo "  To remove Tally: $DIR/uninstall.sh"
}

main
