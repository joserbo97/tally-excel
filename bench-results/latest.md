# Benchmark (all): qwen3.8:latest (think=low)

2026-11-16 15:39 · 8 task runs · avg 179s per task

**Pass rate: 75%** (all checks passed) · average check score 83% · 1 errors/timeouts

| Category | Passed |
|---|---|
| analysis | 4/4 |
| cleanup | 1/1 |
| files | 1/1 |
| hard | 0/2 |

| Difficulty | Passed |
|---|---|
| ★ | 2/2 |
| ★★ | 4/4 |
| ★★★ | 0/2 |

| Task | Result | Score | Time | Steps | Failed checks |
|---|---|---|---|---|---|
| analysis-top-region | ✅ pass | 100% | 19s | 2 | — |
| analysis-category-average | ✅ pass | 100% | 38s | 2 | — |
| analysis-find-errors | ✅ pass | 100% | 88s | 3 | — |
| analysis-fastest-growth | ✅ pass | 100% | 51s | 2 | — |
| clean-dedupe | ✅ pass | 100% | 85s | 6 | — |
| file-csv-import | ✅ pass | 100% | 38s | 4 | — |
| hard-amortization-schedule | 💥 error | 14% | 720s | 0 | 180 monthly rows; final balance ≈ 0; month 1 interest = 1,080; total interest ≈ 114,855; schedule is live; still pays off |
| hard-large-data-question | ❌ fail | 50% | 393s | 12 | top 3 = Matcha, Coffee, Cocoa in order; Matcha March revenue ≈ 19,350 |
