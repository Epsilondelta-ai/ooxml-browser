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

## Stage 1 — scene renderer scaffold
- **Status:** partial scaffold kept, default-on attempt rejected
- **What landed:** a real `scene-svg` PPT render mode exists behind opt-in (`?pptxRenderer=scene-svg`) while metadata mode stays default/fallback.
- **Why default-on was rejected:** the first default-on screenshot loop produced severe visual regressions despite infrastructure progress.
- **Rejected attempt evidence:**
  - `sample1/1`: `85.20 -> 85.74` but vision regressed badly (white placeholder blocks / broken machinery silhouette)
  - `sample5/2`: `81.14 -> 82.37` but vision regressed badly (rocket/ring collapse and black block artifacts)
  - `sample6/1`: `89.88 -> 88.87` with text/layout degradation
- **Decision:** revert default usage, keep scaffold hidden behind fallback, and continue toward inheritance/text + geometry completion before re-enabling it for evidence runs.

## Stage 2 — placeholder/layout/master inheritance + text defaults
- **Status:** accepted
- **What changed:** parser now prefers layout placeholder matches over conflicting master placeholders and merges list-style/default-run placeholder text defaults into slide placeholders.
- **Verification:** `npm test -- pptx-shape-transform.test.ts`, `npm test -- render-and-browser.test.ts pptx-shape-transform.test.ts`, `npm run typecheck`, `npm run build --workspace @ooxml/example-basic`, `PPT_SAMPLE_SCREENSHOT_PORT=4278 npm run quality:ppt-sample-screenshots`
- **Evidence:**
  - `sample1/1`: `85.20 -> 85.54` (essentially neutral)
  - `sample5/2`: `81.14 -> 81.93` (improved)
  - `sample6/1`: `89.88 -> 89.76` (small drop, but subtitle layout moved materially closer to centered multi-line behavior)
- **Decision:** accepted under mixed-evidence rule because score loss stayed bounded and the text/layout hotspot class improved visibly on `sample6/1` while `sample5/2` improved numerically.
