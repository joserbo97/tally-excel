# Roadmap

Where Tally stands (2026-11-21) and what would make it better, most important first.

| Area | Now | Goal |
|---|---|---|
| Accuracy | 26 of 27 everyday tasks per run; the three tasks that slipped now pass 9 of 9; reconciliation 3 of 3 in 18–63 s | every hard task passes 3 of 3 |
| Install | one line, but untested on a fresh Mac | tested on a clean macOS account |
| Evidence | no outside users yet | 5 usability sessions done |
| Reach | Mac with Apple silicon only | Windows |
| Speed | about 4–5 s to the first action, 45–51 s per task on average | quick answers for quick questions |

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

- **Measured (2026-11-22):** the benchmark now records the time to the first action and how much the model thinks. Median time to the first action is 5.3 s, and the average task takes 51 s.
- **Tried:** skipping thinking on the first step (`quickFirstStep`, `--quick-first` in the benchmark). It produced 40% less thinking text, 45 s per task and 4.4 s to the first action, with the same pass rate. That's a single run each, so it's still off by default. Turn it on once a repeat run confirms it.
- Quick questions are not slow because of thinking: the 4 analysis tasks took 32 s on average either way. The time goes into reading the sheet and writing the answer.
- Keep a smaller model option for 16–24 GB Macs, labeled with its benchmark score.

## Fixed on 2026-11-22

- **A total written over the last data row.** `formulas-totals` once put its Total row on top of the Andes row. `write_range` now reports every non-empty cell it replaces (`replaced_cells`, with the old content), and the panel shows it under the tool card.
- **A "Total" label above its sum, in the same column.** `clean-text-numbers` once wrote "Total" in B8 and the SUM in B9, then told the user A8/B8. `write_range` now returns a `layout_warning` with the fix. A "Total" column header above per-row formulas isn't flagged.
- **An unreadable tool call ended the turn.** Ollama rejected a malformed call ("XML syntax error…") after the work was done, and the run crashed. The agent now asks the model to repeat the call, up to 2 times.
- After these fixes, the three affected tasks passed 9 of 9 runs.

## Known weaknesses

- The label-in-the-wrong-column slip is now flagged, but not yet seen fixed by the model in a run where it happened (in the 9 re-runs, the model didn't make the slip).
- The stand-in workbook can't evaluate spill references (`A2#`), so tasks that use them score pessimistically.
