# Benchmark (core): qwen3.8:latest (think=low)

2026-11-22 02:44 · 27 task runs · avg 51s per task · median 5.3s to first action

**Pass rate: 96%** (all checks passed) · average check score 98%

| Category | Passed |
|---|---|
| analysis | 4/4 |
| formulas | 5/5 |
| fix | 3/3 |
| cleanup | 4/5 |
| modeling | 3/3 |
| visual | 2/2 |
| files | 4/4 |
| multi-sheet | 1/1 |

| Difficulty | Passed |
|---|---|
| ★ | 9/9 |
| ★★ | 15/16 |
| ★★★ | 2/2 |

| Task | Result | Score | Time | First action | Steps | Failed checks |
|---|---|---|---|---|---|---|
| analysis-top-region | ✅ pass | 100% | 40s | 22.1s | 2 | — |
| analysis-category-average | ✅ pass | 100% | 21s | 3.2s | 2 | — |
| analysis-find-errors | ✅ pass | 100% | 25s | 3.4s | 3 | — |
| analysis-fastest-growth | ✅ pass | 100% | 43s | 8.7s | 2 | — |
| formulas-totals | ✅ pass | 100% | 105s | 7.9s | 5 | — |
| formulas-growth-percent | ✅ pass | 100% | 35s | 8s | 4 | — |
| formulas-lookup | ✅ pass | 100% | 102s | 8.2s | 9 | — |
| formulas-sumifs-summary | ✅ pass | 100% | 48s | 3.5s | 8 | — |
| formulas-commission-tiers | ✅ pass | 100% | 63s | 4.1s | 3 | — |
| fix-div0 | ✅ pass | 100% | 20s | 3.4s | 3 | — |
| fix-sum-range | ✅ pass | 100% | 24s | 4.3s | 3 | — |
| fix-hardcoded-totals | ✅ pass | 100% | 29s | 3.6s | 3 | — |
| clean-split-names | ✅ pass | 100% | 64s | 4.2s | 5 | — |
| clean-text-numbers | ❌ fail | 50% | 35s | 4.5s | 4 | total = 3,697.50; total is a formula |
| clean-dedupe | ✅ pass | 100% | 169s | 3.9s | 8 | — |
| clean-sort | ✅ pass | 100% | 19s | 4.6s | 4 | — |
| clean-standardize | ✅ pass | 100% | 23s | 5.3s | 3 | — |
| model-loan | ✅ pass | 100% | 49s | 15.9s | 3 | — |
| model-projection | ✅ pass | 100% | 77s | 44s | 3 | — |
| model-breakeven | ✅ pass | 100% | 47s | 10.8s | 3 | — |
| visual-table-chart | ✅ pass | 100% | 32s | 4.7s | 5 | — |
| visual-format | ✅ pass | 100% | 28s | 4.1s | 3 | — |
| file-csv-import | ✅ pass | 100% | 49s | 14.5s | 6 | — |
| file-pdf-invoice | ✅ pass | 100% | 48s | 11.4s | 4 | — |
| file-pptx-chart | ✅ pass | 100% | 58s | 13.7s | 4 | — |
| file-word-table | ✅ pass | 100% | 38s | 9.6s | 4 | — |
| multisheet-quarter | ✅ pass | 100% | 92s | 6.5s | 5 | — |
