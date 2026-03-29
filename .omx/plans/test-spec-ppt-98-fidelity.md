# Test Specification: PPT 98-percent fidelity push

## Evidence contract
Target slides:
- `sample1` slide 1
- `sample5` slide 2
- `sample6` slide 1

Each iteration must record:
- score before / after
- diff hotspot notes
- vision verdict
- accepted or reverted decision

## Verification layers
### Unit
- placeholder/layout/master inheritance merges
- text style / paragraph default resolution
- theme font and color transform resolution
- geometry primitive/path conversion
- fill/stroke semantics including dash and opacity

### Integration
- target-slide semantic snapshots or equivalent structured assertions
- renderer output assertions for geometry/text/stroke metadata
- example consumes renderer output correctly

### Visual
- `npm run quality:ppt-sample-screenshots`
- extract scores only for the target slides
- preserve diff hotspot notes for those three slides

### Vision
For each target slide, record:
- structural match
- text layout match
- geometry silhouette match
- color/stroke match
- remaining hotspot list

## Stage gates
- Stage 0: baseline ledger reproducible
- Stage 1: scene renderer path renders target slides end-to-end
- Stage 2: text/inheritance hotspots materially shrink
- Stage 3: geometry hotspots materially shrink
- Stage 4: fill/stroke/theme hotspots materially shrink
- Stage 5: heuristics reduced without regression
- Stage 6: all targets 98%+

## Final verification
- `npm test`
- `npm run typecheck`
- `npm run lint`
- `npm run build`
- `git diff --check`
- architect review
- verifier review
