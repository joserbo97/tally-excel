# Benchmark fixtures

`file-pptx-chart` needs `budget-deck.pptx` here: a PowerPoint deck whose slide 5 holds a native chart of revenue by source. The deck used in development is a third-party document and isn't included in this repository, so the benchmark skips that task when the file is missing.

To run it, put any deck with a revenue-by-source chart on slide 5 at `scripts/bench/fixtures/budget-deck.pptx` and adjust the expected values in the task's checker (`scripts/bench/tasks.mjs`, `file-pptx-chart`).
