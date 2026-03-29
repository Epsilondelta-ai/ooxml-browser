# Consensus Plan: PPT 98-percent fidelity push

## RALPLAN-DR summary
### Principles
1. 98% evidence requires renderer-accurate semantics, not just preview cosmetics.
2. Parser semantics and renderer architecture must evolve together.
3. Only target-slide evidence decides whether a stage is accepted.
4. Remove heuristics only after semantic replacement exists.
5. Use the lightest renderer that can still cross the 98% threshold.

### Decision drivers
1. Current scores remain far below 98% on two of three target slides.
2. The existing preview still depends on example heuristics, limiting trustworthy convergence.
3. SVG/Canvas freedom removes the previous architecture constraint.

### Viable options
- **A. Continue incremental semantic HTML + heuristic tuning**
  - Pros: smallest local diffs
  - Cons: unlikely to reach 98% reliably
- **B. Introduce PPT scene renderer (SVG-first, Canvas/hybrid if needed)** — chosen
  - Pros: direct path to geometry/text/stroke fidelity, cleaner heuristic removal
  - Cons: larger architectural change
- **C. Build a target-slide-specific custom renderer**
  - Pros: fastest path to better screenshots
  - Cons: violates parser-grounded reusable semantics

### Decision
Choose **Option B**. Build a PPT scene renderer and drive hotspot reduction through parser-grounded semantics.

### Consequences
- Near-term complexity increases.
- Example-specific logic can shrink over time instead of accumulating.
- Evidence becomes more trustworthy because it comes from the real render path.

## Exact staged implementation order
1. Evidence contract / baseline lock
2. Scene renderer introduction (SVG-first behind a fallback flag)
3. Placeholder/layout/master inheritance + text engine
4. Geometry engine
5. Fill / stroke / theme color semantics
6. Overlay / heuristic removal
7. Hotspot loop to 98%+

### Stage ownership boundaries
- **Stage 1**: `tools/generate-ppt-sample-screenshot-report.mjs`, `benchmarks/reports/**`, `.omx/plans/**`
- **Stage 2**: `packages/render/**`, `examples/basic/**`, `tests/render-and-browser.test.ts`
- **Stage 3**: `packages/pptx/src/parser.ts`, `packages/pptx/src/model.ts`, `tests/pptx-inheritance.test.ts`, `tests/pptx-serializer-roundtrip.test.ts`
- **Stage 4**: `packages/pptx/src/parser.ts`, `packages/render/**`, `tests/pptx-shape-transform.test.ts`
- **Stage 5**: `packages/pptx/src/parser.ts`, `packages/pptx/src/model.ts`, `packages/render/**`, style/line/theme tests
- **Stage 6**: `examples/basic/**`, `tests/render-and-browser.test.ts`
- **Stage 7**: only the hotspot owner files identified in the ledger

## Ralph staffing guidance
- **Implementation lane:** renderer/parser work (`executor`, high reasoning)
- **Evidence lane:** screenshots, diff mining, regression triage (`debugger`/`test-engineer`)
- **Sign-off lane:** `architect`, then `verifier`

## Available-agent-types roster
- `planner`, `architect`, `critic`, `executor`, `debugger`, `test-engineer`, `verifier`, `writer`

## Accept / revert rules
- Any accepted stage must either improve score or materially improve vision verdict without unacceptable score loss.
- A mixed result is acceptable only when the worst score drop is <= 0.30 and the vision verdict records a concrete semantic improvement that removes a hotspot class.
- If the worst target slide stagnates for 3 accepted attempts, escalate architecture rather than stacking more heuristics.
- Revert any stage whose net evidence is negative.
- Keep the current PPT preview/render path as a fallback until the new renderer beats it on all 3 target slides.

## Architecture checkpoints
- **After Stage 2**: if no target slide improves by at least 1.0 score point or hotspot severity clearly drops, escalate from SVG-first to Canvas/hybrid evaluation.
- **After Stage 3**: explicitly decide whether text/layout or geometry is now the dominant blocker before Stage 4 continues.

## Verification path
1. Baseline evidence ledger
2. `target-hotspots-ppt-98-fidelity.md` update (top hotspots, owner, suspected semantic cause, planned stage)
3. Stage-local tests
4. Build preview surface
5. Screenshot run
6. Target-slide score extraction
7. Vision verdict
8. Accept/revert decision
9. Final repo verification
10. Architect + verifier approval
