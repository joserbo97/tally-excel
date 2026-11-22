# Benchmark (core): qwen3.8:latest (think=low)

2026-11-22 00:46 · 27 task runs · avg 40s per task

**Pass rate: 96%** (all checks passed) · average check score 99%

| Category | Passed |
|---|---|
| analysis | 4/4 |
| formulas | 4/5 |
| fix | 3/3 |
| cleanup | 5/5 |
| modeling | 3/3 |
| visual | 2/2 |
| files | 4/4 |
| multi-sheet | 1/1 |

| Difficulty | Passed |
|---|---|
| ★ | 8/9 |
| ★★ | 16/16 |
| ★★★ | 2/2 |

| Task | Result | Score | Time | Steps | Failed checks |
|---|---|---|---|---|---|
| analysis-top-region | ✅ pass | 100% | 21s | 2 | — |
| analysis-category-average | ✅ pass | 100% | 26s | 3 | — |
| analysis-find-errors | ✅ pass | 100% | 22s | 3 | — |
| analysis-fastest-growth | ✅ pass | 100% | 20s | 2 | — |
| formulas-totals | ❌ fail | 83% | 58s | 4 | Total row: Q1 = 661,000 |
| formulas-growth-percent | ✅ pass | 100% | 22s | 4 | — |
| formulas-lookup | ✅ pass | 100% | 45s | 5 | — |
| formulas-sumifs-summary | ✅ pass | 100% | 40s | 6 | — |
| formulas-commission-tiers | ✅ pass | 100% | 57s | 4 | — |
| fix-div0 | ✅ pass | 100% | 23s | 3 | — |
| fix-sum-range | ✅ pass | 100% | 16s | 3 | — |
| fix-hardcoded-totals | ✅ pass | 100% | 21s | 3 | — |
| clean-split-names | ✅ pass | 100% | 55s | 4 | — |
| clean-text-numbers | ✅ pass | 100% | 49s | 4 | — |
| clean-dedupe | ✅ pass | 100% | 118s | 13 | — |
| clean-sort | ✅ pass | 100% | 21s | 4 | — |
| clean-standardize | ✅ pass | 100% | 18s | 3 | — |
| model-loan | ✅ pass | 100% | 41s | 3 | — |
| model-projection | ✅ pass | 100% | 55s | 3 | — |
| model-breakeven | ✅ pass | 100% | 43s | 3 | — |
| visual-table-chart | ✅ pass | 100% | 32s | 4 | — |
| visual-format | ✅ pass | 100% | 26s | 3 | — |
| file-csv-import | ✅ pass | 100% | 36s | 4 | — |
| file-pdf-invoice | ✅ pass | 100% | 38s | 4 | — |
| file-pptx-chart | ✅ pass | 100% | 57s | 5 | — |
| file-word-table | ✅ pass | 100% | 30s | 4 | — |
| multisheet-quarter | ✅ pass | 100% | 96s | 6 | — |
