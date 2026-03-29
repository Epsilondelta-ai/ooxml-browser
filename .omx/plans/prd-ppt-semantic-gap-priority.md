# PRD: PPT semantic gap priority pass

## Metadata
- **Project:** PPT semantic gap priority pass
- **Slug:** ppt-semantic-gap-priority
- **Mode:** ralplan consensus
- **Grounding snapshot:** `.omx/context/ppt-semantic-gap-priority-20260329T120751Z.md`
- **Target slides:** `sample1` slide 1, `sample5` slide 2, `sample6` slide 1
- **Baseline screenshot scores:** sample1/1 `85.20`, sample5/2 `81.14`, sample6/1 `89.84`

## Product objective
Improve PPT rendering by closing the dominant OOXML semantic gaps in parser-grounded order, so the renderer becomes meaningfully more faithful because it understands slide semantics better, not because the example app hardcodes deck-specific styling.

## Problem statement
The current PPT pipeline already parses and displays shapes, fills, text, and inherited content, but the highest remaining fidelity losses for the target slides come from incomplete semantic interpretation:
1. group transforms are composed approximately,
2. custom/preset geometry is projected through limited path extraction,
3. placeholder/layout/master inheritance is shallow,
4. theme color transforms are incomplete,
5. line/fill/stroke semantics are simplified,
6. example-level overlay hacks still compensate for parser/render gaps.

## Goals
1. Capture a clean baseline and keep commits free of accidental screenshot-artifact churn.
2. Follow the requested priority order strictly.
2. Prefer parser/model/render semantic fixes over example-only CSS or DOM hacks.
3. After every priority step, run screenshot verification and vision review for the target slides.
4. If a step regresses target-slide quality, revert that step immediately and continue to the next semantic gap.
5. Commit each meaningful stage with Lore-protocol commits while keeping the working tree clean between stages.

## Non-goals
- Pixel-perfect parity for arbitrary PPTX decks.
- Broad SmartArt/effects/animation support in this pass.
- Solving PPT fidelity via external rasterization or deck-specific overlays.

## Strict implementation order
### 1. Group transform accuracy
- Correct nested group offset/scale/rotation/flip composition.
- Ensure child transforms and geometry bounds are normalized through the full ancestry chain.

### 2. Custom/preset geometry fidelity
- Improve extraction and projection of custom geometry paths and preset geometry intent.
- Preserve enough shape semantics for the renderer to draw closer to the authored figure.

### 3. Placeholder/layout/master inheritance
- Deepen inheritance matching beyond exact placeholder type/index pairs where OOXML semantics require fallback.
- Promote layout/master text style, fills, lines, and transforms only when the slide instance omits them.

### 4. Theme color transforms
- Support scheme-based color transforms more faithfully, including luminance/tint/shade style modifiers used by target slides.
- Ensure fill, line, and text colors resolve consistently from the effective theme context.

### 5. Line/fill/stroke semantics
- Improve line width/default handling and fill semantics so rendered strokes/fills better match OOXML intent.
- Preserve parser-grounded data needed by the example surface to render these semantics without hardcoded sample heuristics.

### 6. Overlay hack minimization
- Remove or narrow example-only PPT overlays/classes that become redundant after semantic fixes.
- Keep only generic rendering affordances that remain necessary across samples.

## Acceptance criteria
1. A clean baseline is captured before source edits and commits do not accidentally bundle transient screenshot churn.
2. The implementation attempts priorities 1→6 in order and records verification after each step.
3. Each accepted step is parser-grounded and lands only if it does not regress the target-slide screenshots/vision review.
4. `npm run quality:ppt-sample-screenshots` is run after each step, with evidence pulled for `sample1/1`, `sample5/2`, and `sample6/1`.
5. Final code keeps or improves baseline quality on all three target slides, with at least one slide materially improved by accepted semantic changes.
6. Final repo verification passes: `npm test`, `npm run typecheck`, `npm run lint`, `npm run build`, `git diff --check`.
7. Final architect review approves the completed work.

## Likely codebase touchpoints
- `packages/pptx/src/model.ts`
- `packages/pptx/src/parser.ts`
- `packages/render/src/pptx.ts`
- `examples/basic/src/main.ts`
- `tools/generate-ppt-sample-screenshot-report.mjs`
- `tests/pptx-*.test.ts`
- `tests/render-and-browser.test.ts`

## Risks
- Screenshot evidence may fluctuate slightly because artifact generation rewrites diff/report outputs.
- Target slides may improve through different semantic gaps than expected, so rollback discipline matters.
- Some visible mismatches may still come from example rendering limitations that cannot be removed until parser semantics improve.

## Definition of done
This pass is done when:
1. plan artifacts reflect the strict semantic order,
2. accepted fixes were attempted in that order with rollback discipline,
3. target-slide screenshot + vision evidence is captured after each step,
4. final repo verification passes,
5. architect verification approves the result,
6. Ralph state is cleanly completed/cancelled.
