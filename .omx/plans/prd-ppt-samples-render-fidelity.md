# PRD: PPT sample render fidelity

## Metadata
- **Project:** PPT sample render fidelity
- **Slug:** ppt-samples-render-fidelity
- **Mode:** ralplan consensus, deliberate
- **Grounding snapshot:** `.omx/context/ppt-samples-render-fidelity-20260329T031015Z.md`
- **Reference corpus:** `~/Desktop/ppt-samples/sample{1..6}/sample.pptx` with paired `sample.*.png`

## Product objective
Upgrade the PPTX parsing and rendering stack so the browser preview can render the paired sample corpus much closer to the exported PNG images, not just as semantic/debug HTML.

## Why this phase exists
Current PPTX rendering is useful as a semantic inspection surface, but it is still far from the reference deck exports. The six-sample corpus exposes the real gap: theme/background recovery, image part rendering, shape layering, title/content slide composition, and unsupported-package detection all need deeper treatment to approach slide-export fidelity.

## Goals
1. Turn the paired `pptx/png` corpus into a repeatable fidelity harness.
2. Fix PPTX package detection so unsupported parses like `sample6` are classified correctly and recoverable when possible.
3. Deepen PPTX parsing for slide background, fill/line, placeholder/layout/master/theme, and image/media metadata needed for visual rendering.
4. Replace the current mostly-semantic PPTX HTML with a stronger slide scene projection that can recover title slides, content slides, image slides, and divider slides.
5. Add corpus-backed screenshot/reference comparison so visual fidelity is measured, not guessed.
6. Keep example/playground/product surfaces aligned with the improved renderer.

## Non-goals
- Full PowerPoint pixel parity for every animation/effect.
- Editing every shape property during this phase.
- Rebuilding the whole renderer around a new framework or dependency stack.

## Corpus facts
- `sample1`–`sample4`: 48 slides each, 16:9 deck size `12192000 x 6858000`
- `sample5`: 37 slides, 16:9-ish deck size `9144000 x 5143500`
- `sample6`: currently opens as `officeDocumentKind: unknown` with `rootDocumentUri: /docProps/app.xml`
- `sample2` contains at least 4 image shapes; `sample3` contains at least 1 image shape; the others are text/vector heavy
- Layout names are mostly absent in parsed output, so slide appearance likely depends on direct shape/theme/background content rather than currently recovered layout labels

## Quality targets
### Visual fidelity
- Title, subtitle, footer, and section-break slides should preserve placement, scale hierarchy, and dominant background treatments.
- Image-heavy slides should display embedded raster content at correct relative position/scale.
- Content slides should preserve card/column structure, separators, accent bars, and text blocks closely enough to be recognizable at a glance.

### Parser fidelity
- Detect the true PPTX root reliably, including packages that currently fall back to `/docProps/app.xml`.
- Preserve enough slide/theme/background/shape metadata to drive scene rendering.

### Product surfaces
- `examples/basic` becomes a credible corpus previewer, including slide navigation and presentation-focused mode.
- The playground remains useful for inspection/editing without regressing existing flows.

### Verification
- Every sample slide has a renderable browser preview.
- Screenshot/reference comparison runs against the corpus and reports mismatches per slide.
- Build/tests/lint/typecheck stay green throughout.

## Architecture rules
1. PPTX slide rendering should follow the repo's documented HTML/SVG/Canvas hybrid strategy rather than stuffing everything into text-only cards.
2. Example-specific heuristics are allowed only as thin product-surface affordances; durable scene semantics must live in parser/render packages.
3. The corpus harness must separate parse failures, unsupported features, and visual mismatches.
4. Sample6 root-detection/parsing must be addressed before declaring corpus coverage complete.

## Execution stages
### Stage 0 — corpus grounding + harness design
- inventory all six sample folders
- map each slide index to its reference PNG
- add sample6 root-detection diagnosis note
- define screenshot/reference comparison workflow and thresholds

### Stage 1 — PPTX package detection + corpus ingestion foundation
- fix root-document detection for sample6-class files
- add corpus manifest/loader for local sample decks
- add render driver that can select slide N from a PPTX and capture a browser screenshot deterministically

### Stage 2 — parser depth for slide visuals
- parse/recover slide background and dominant fill data
- improve placeholder/title/body/background/image shape metadata
- recover enough scene data for divider/title/content/image slide families

### Stage 3 — renderer scene upgrade
- introduce stronger PPTX slide scene projection for positioning, image rendering, and slide-family styling
- reduce reliance on debug text/chrome in presentation mode
- keep semantic/debug path available separately

### Stage 4 — corpus-backed visual verification
- add screenshot/reference diff workflow for sample1–sample6
- persist per-slide verdicts, mismatch summaries, and regression evidence
- define thresholds and tolerated differences by slide family

### Stage 5 — hardening + product-surface alignment
- align `examples/basic`, `playground`, docs, and quality scripts with the new PPTX renderer
- fix worst remaining outliers from the sample corpus
- prepare final architect/verifier evidence

## Acceptance criteria
1. All six sample folders are recognized by the corpus harness, with sample6 no longer silently classified as `unknown` without diagnosis.
2. The browser preview can navigate and render every slide in the sample decks.
3. Embedded image slides render actual media content, not text placeholders.
4. Title/divider/content/image slide families have measurable visual improvement against their PNG references.
5. A repeatable screenshot/reference comparison command exists and stores per-slide evidence.
6. `npm test`, `npm run typecheck`, `npm run lint`, and affected build flows pass.

## Risks
- **Theme/layout data may be sparse** → Mitigation: fall back to direct shape/background parsing and scene heuristics, but keep heuristics isolated.
- **Font differences may dominate visual diffs** → Mitigation: pin local screenshot environment and classify typography-only diffs separately.
- **Sample6 may expose broader root-detection bugs** → Mitigation: fix detection first and add a regression fixture for the failure mode.
- **Scene rendering may sprawl into full presentation-engine work** → Mitigation: limit scope to sample-driven slide families and measurable corpus wins.

## Definition of done
This phase is done when:
1. the six-sample corpus is wired into repeatable visual verification,
2. sample6 is either supported or explicitly diagnosed with a tested fallback path,
3. PPTX preview output is materially closer to the paired PNG exports across the corpus,
4. examples/docs/verification surfaces reflect the new behavior,
5. architect and verifier both approve with fresh evidence.
