# Evidence ledger: PPT 98-percent fidelity

## Stage 0 — baseline lock
- **Generated from:** `benchmarks/reports/ppt-sample-screenshot-report.json`
- **Target slides:** `sample1 slide1`, `sample5 slide2`, `sample6 slide1`
- **Acceptance policy:** later stages are accepted only if they improve score or materially improve vision with only bounded metric loss per consensus plan.

| Slide | Baseline score | Diff artifact | Primary hotspot class |
|---|---:|---|---|
| sample1 / 1 | 85.20 | `benchmarks/reports/ppt-sample-diffs/sample1/sample.001.png` | geometry + text/logo composition |
| sample5 / 2 | 81.14 | `benchmarks/reports/ppt-sample-diffs/sample5/sample.002.png` | rocket/ring geometry + agenda stroke/text layout |
| sample6 / 1 | 89.88 | `benchmarks/reports/ppt-sample-diffs/sample6/sample.001.png` | circular text layout + dashed ring + scene ordering |

## Baseline notes
- `sample1/1`: decorative machinery geometry and skyline layering remain visibly off; title/logo composition still depends on preview heuristics.
- `sample5/2`: the rocket/ring silhouette and agenda card borders are still farther from the reference than the target threshold allows.
- `sample6/1`: closest of the three, but circular text layout and dashed ring fidelity still block 98%.

## Next stage intent
1. Introduce a scene-renderer path behind fallback.
2. Keep the current preview path as baseline until the new path wins on all three target slides.
3. Update this ledger after every accepted/reverted stage.
