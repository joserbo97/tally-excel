# Roadmap

Where Tally stands (2026-11-21) and what would make it better, most important first.

| Area | Now | Goal |
|---|---|---|
| Accuracy | 97.5% on everyday tasks; reconciliation now 3 of 3 in 18–63 s (was 1 of 5, up to 12 min) | every hard task passes 3 of 3 |
| Install | one line, but untested on a fresh Mac | tested on a clean macOS account |
| Evidence | no outside users yet | 5 usability sessions done |
| Reach | Mac with Apple silicon only | Windows |
| Speed | seconds to minutes per task | quick answers for quick questions |

## 1. Reconciliation (done)

Prompt rules and a thinking cap didn't help, because the model still compared rows in its head or built dozens of helper formulas one call at a time. Tally now has a `compare_lists` tool that does the matching in code in one step. It writes a report sheet listing missing keys, value differences and duplicates, with live XLOOKUP amounts, so the model only has to choose the two lists and explain the result.

- Tests: `scripts/test-compare.mjs`, plus a reference solution in `npm run bench:selftest`.
- Benchmark (2026-11-22): **3 of 3 in 18–63 s**, down from 1 of 5 clean runs that took up to 12 minutes. In an earlier round, 2 of 3 runs guessed the bank range was as long as the ledger's (`Bank!A1:C61`) and missed its last row, so the tool now extends a range when the list continues below it, and says so.
- The core suite ran once afterwards: 26 of 27. The one failure, `formulas-totals`, summed the wrong rows for one total (601,000 instead of 661,000) and didn't use the new tool.

## 2. Installer (hardened; a clean-account test is still to do)

- `curl | bash` now runs only after the whole script has downloaded, and no command can swallow the rest of it from stdin.
- Homebrew is found even when a fresh install isn't on `PATH` yet.
- `--check` lists what the install would set up, so a new user sees the steps, including the password prompt, before anything changes.
- **To do:** run the installer on a new macOS user account (System Settings → Users & Groups → Add User), from a Mac without Homebrew, Node or Ollama.

## 3. Usability sessions

Five 45-minute sessions with people who reconcile, clean or model data in Excel every week, each using their own kind of workbook. Watch for: whether they trust a change without checking it, whether they find the Changed list and Undo, and what they'd never hand to it.

## 4. Windows

Office.js and Ollama both run on Windows, so the panel and agent should carry over. The work is in the installer:

- a PowerShell installer;
- the localhost certificate in the Windows certificate store;
- sideloading through a shared-folder catalog or the registry;
- a scheduled task in place of the LaunchAgent.

Needs a Windows machine with 32 GB of RAM and a GPU for testing.

## 5. Speed

- For read-only questions ("what's the total in F?"), skip thinking on the first step and only think when a tool call fails.
- Measure the time to the first useful action, not just the total.
- Keep a smaller model option for 16–24 GB Macs, labeled with its benchmark score.

## Known weaknesses

- When writing a row that starts with a label, the total can land one column off (CSV import, 2 of 3 runs once).
- The stand-in workbook can't evaluate spill references (`A2#`), so tasks that use them score pessimistically.
