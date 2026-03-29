# PRD: PPT samples render fidelity

## Metadata
- **Project:** PPT samples render fidelity
- **Slug:** ppt-samples-render-fidelity
- **Mode:** ralplan consensus, deliberate
- **Grounding snapshot:** `.omx/context/ppt-samples-render-fidelity-20260329T031015Z.md`
- **Corpus:** `~/Desktop/ppt-samples/sample{1..6}`

## Product objective
Improve the repository's PPTX parsing and browser rendering so paired sample presentations under `~/Desktop/ppt-samples/` render materially closer to their exported PNG references, not just as semantic/debug projections.

## Why this phase exists
The current PPTX flow can parse representative shapes and render semantic HTML, but the produced slides still look like inspection surfaces instead of presentation exports. The sample corpus exposes concrete fidelity gaps in:
- root presentation detection (`sample6` currently resolves as `unknown`)
- title-slide and content-slide composition
- background/fill/shape primitive rendering
- theme-aware typography and color application
- image-heavy and diagram-heavy slide presentation
- end-to-end screenshot-to-reference verification

## Goals
1. Parse all six sample decks into non-empty PPTX documents, including `sample6`.
2. Introduce a PPTX render path that supports export-like slide presentation rather than metadata-heavy debug blocks.
3. Preserve and render key visual semantics: slide background, theme/text colors, shape fills/lines, image placement, and placeholder-driven layout heuristics.
4. Add a deterministic visual verification harness that compares rendered screenshots against the paired PNG exports for the sample corpus.
5. Keep browser example/playground and docs aligned with the stronger PPTX preview surface.

## Non-goals
- Full Microsoft PowerPoint pixel parity for arbitrary decks.
- Implementing every DrawingML effect, animation, or SmartArt primitive in this wave.
- Replacing semantic render HTML for DOCX/XLSX.

## Acceptance criteria
1. `sample1` through `sample6` all parse as PPTX with expected slide counts > 0.
2. A screenshot harness can render each sample slide and compare it with its paired PNG export.
3. Title slides, image-driven slides, and card/list slides from the sample corpus score materially better than the current baseline in visual review.
4. The browser example can load the sample decks and present slide-first navigation with media-backed previews.
5. PPTX-focused tests cover root detection, media/background fidelity, and screenshot/reference bookkeeping.

## Corpus facts
- `sample1`–`sample4`: 48 PNG exports each
- `sample5`: 37 PNG exports
- `sample6`: 36 PNG exports, but currently parses as 0 slides because package root detection picks `/docProps/app.xml`
- `sample2` has several image-driven slides; `sample1`, `sample4`, and `sample5` lean heavily on vector/text composition

## Likely codebase touchpoints
- `packages/core/src/opc.ts`
- `packages/pptx/src/parser.ts`
- `packages/pptx/src/model.ts`
- `packages/render/src/pptx.ts`
- `examples/basic/src/main.ts`
- `playground/src/main.ts`
- `tests/**`
- `tools/**` or a new screenshot/reference harness under `tools/` + `fixtures/`

## Risks
- Theme/layout semantics may be absent or inconsistent in template decks, forcing heuristic fallback.
- Visual verification can become flaky without a pinned browser/runtime environment.
- Shape/vector fidelity may require incremental support across backgrounds, autoshapes, and text blocks rather than one-shot parity.

## Definition of done
This phase is done when:
1. all six sample decks parse correctly,
2. a screenshot/reference harness exists and is green for the declared sample corpus baseline,
3. the PPTX example/playground behaves like a slide viewer instead of an inspection shell,
4. fresh verification, architect review, and verifier review all pass.
