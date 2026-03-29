# PRD: PPT 98-percent fidelity push

## Metadata
- **Project:** PPT 98-percent fidelity push
- **Slug:** ppt-98-fidelity
- **Mode:** ralplan consensus, deliberate
- **Grounding snapshot:** `.omx/context/ppt-98-fidelity-20260329T134845Z.md`
- **Target slides:** `sample1` slide 1, `sample5` slide 2, `sample6` slide 1
- **Target evidence:** 98%+ each, with 100% as the stretch goal

## Product objective
Upgrade the PPT rendering path until the target slides render close enough to their reference exports that screenshot evidence reaches 98%+ on all three target slides, using parser-grounded semantics and any renderer technology necessary.

## Decision
We will not treat the current preview shell as the final rendering architecture. The system may use SVG, Canvas, or a hybrid scene renderer if that is the shortest path to 98%+ evidence.

## Core principles
1. Parser-grounded semantics before example styling.
2. Target-slide evidence is the source of truth.
3. Replace heuristics only when a semantic renderer path supersedes them.
4. Prefer scene-accurate rendering over incremental cosmetic patching once partial fixes stop compounding meaningfully.
5. Every stage must be measurable with screenshot + diff + vision evidence.

## Non-goals
- Corpus-wide 98% parity for all sample decks in this phase.
- Pixel-perfect Office equivalence for unsupported effects beyond the target slides.
- Preserving existing example heuristics if a stronger renderer makes them redundant.

## Success criteria
1. `sample1/1`, `sample5/2`, `sample6/1` all score 98%+.
2. Vision review reports no major semantic mismatches on those slides.
3. Parser/render changes remain reusable and not deck-hardcoded.
4. Remaining overlay heuristics are either removed or explicitly justified as temporary debt.
5. Final repo verification passes.

## Staged implementation order
### Stage 0 — evidence contract / baseline lock
- Freeze baseline scores, references, diff hotspots, and vision rubric.
- Add a durable evidence ledger for the target slides.

### Stage 1 — PPT scene renderer introduction
- Introduce a PPT-focused scene renderer (SVG-first, Canvas/hybrid if needed).
- Make the example surface consume renderer output rather than synthesizing layout from heuristics.

### Stage 2 — placeholder/layout/master inheritance + text engine
- Complete placeholder fallback and inheritance merging.
- Implement text box semantics: paragraph defaults, line spacing, anchoring, font resolution, and wrapping behavior.

### Stage 3 — custom/preset geometry engine
- Promote custom and preset geometry to renderer-native shapes rather than approximation-only HTML.
- Support adjustment values, multi-path geometry, and proper fill/stroke composition.

### Stage 4 — fill / stroke / theme color semantics
- Finish scheme color transforms, alpha, gradients, dash/width/cap/join behavior, and fill/stroke fallback chains.

### Stage 5 — overlay / heuristic removal
- Remove or narrow preview-only heuristics that have semantic replacements.
- Keep only explicitly justified temporary heuristics.

### Stage 6 — hotspot chasing loop to 98%+
- Use diff hotspots and vision verdicts to drive the next semantic fix until all three targets reach 98%+.

## Module boundaries
- **Parser / semantic extraction:** `packages/pptx/src/parser.ts`, `packages/pptx/src/model.ts`
- **Renderer core:** `packages/render/src/pptx.ts`
- **Preview host / fallback shell:** `examples/basic/src/main.ts`, `playground/**`
- **Evidence tooling:** `tools/generate-ppt-sample-screenshot-report.mjs`, `benchmarks/reports/**`
- **Tests / fixtures:** `tests/pptx-*.test.ts`, `tests/render-and-browser.test.ts`, `tests/fixture-builders.ts`

## Accept / revert rules
- Accept a stage only if screenshot evidence is neutral-or-better overall and vision review confirms the semantics improved or remained stable.
- Revert immediately if a stage introduces target-slide regression without unlocking a stronger next-stage semantic path.
- Commit each accepted stage with Lore protocol.

## Escalation conditions
- If 3 consecutive hotspot iterations fail to improve the worst target slide by >=0.5 evidence points, escalate to a renderer-architecture change.
- If text layout remains the dominant hotspot after inheritance work, prioritize scene-text layout over additional geometry work.
- If SVG scene fidelity plateaus below 98%, escalate to Canvas or a hybrid pipeline.

## Execution gates
- Introduce the new scene renderer behind a flag/fallback until it beats the current path on all target slides.
- After Stage 2, if no target slide improves materially, evaluate Canvas or hybrid rendering instead of continuing SVG-only by inertia.
- Maintain a per-slide hotspot ledger in `.omx/plans/target-hotspots-ppt-98-fidelity.md`.
- Mixed evidence is acceptable only when score movement is small and vision clearly confirms semantic improvement.
