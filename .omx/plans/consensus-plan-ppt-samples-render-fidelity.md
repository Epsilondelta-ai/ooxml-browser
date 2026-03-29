# Consensus Plan: PPT samples render fidelity

## RALPLAN-DR summary
### Principles
1. Corpus-backed rendering truth beats aesthetic guesswork.
2. Fix package/parsing correctness before tuning slide cosmetics.
3. Add reusable PPTX scene semantics in the core renderer; keep example-only heuristics thin.
4. Visual progress must be measurable with deterministic screenshot evidence.
5. Prefer staged coverage of dominant slide archetypes over shallow support for every shape effect.

### Decision drivers
1. `sample6` currently fails at root PPTX detection, blocking trustworthy corpus-wide work.
2. The current PPTX renderer is semantic/debug-oriented and lacks export-like scene composition.
3. The paired PNG corpus makes screenshot-based verification practical and necessary.

### Viable options
- **A. Core PPTX scene-rendering upgrade + screenshot harness** — chosen
  - Pros: improves reusable library behavior, supports corpus-wide progress, gives measurable outcomes
  - Cons: larger change set across parser/core/example/tests
- **B. Example-only CSS/DOM heuristics over current semantic renderer** — rejected
  - Pros: quick visible gains in the example
  - Cons: does not fix library-level rendering semantics, brittle across six sample decks
- **C. External rasterization bridge (PowerPoint/LibreOffice export) as runtime dependency** — rejected
  - Pros: closer screenshots faster
  - Cons: violates browser-first library direction and avoids solving parsing/render fidelity

### Decision
Choose **Option A**: fix parsing/root detection, add richer PPTX scene semantics in the renderer, and validate them with a screenshot/reference harness.

### Pre-mortem
1. Screenshot tests are too flaky to trust.
2. Renderer improvements overfit one sample family and regress others.
3. Package/root fixes land, but slide-level visual semantics remain too shallow for the sample corpus.

### Expanded test plan
- **Unit:** root detection, background/theme extraction, shape/image projection helpers
- **Integration:** sample corpus parse counts, slide navigation, image-part rendering, viewer mode
- **E2E/visual:** Playwright screenshot capture against paired PNG exports for declared slide subset, then expand
- **Observability:** per-slide comparison manifest with parser diagnostics and verdict notes

## Stage order
### Stage 0 — corpus grounding + harness design
Deliverables:
- sample inventory and slide-count evidence
- screenshot/reference manifest format
- viewport/browser contract for deterministic captures

### Stage 1 — package/root correctness
Deliverables:
- fix PPTX root detection so `sample6` parses as PPTX
- tests covering unusual package part ordering/root selection

### Stage 2 — renderer scene foundation
Deliverables:
- richer PPTX render model for slide background, positioned shapes, image parts, and text blocks
- reduce debug/meta chrome in preview surfaces

### Stage 3 — archetype fidelity slices
Deliverables:
- title/hero slide fidelity
- image-centric slide fidelity
- card/list/content slide fidelity
- representative diagram-heavy slide handling

### Stage 4 — screenshot verification pipeline
Deliverables:
- sample corpus screenshot runner
- slide/reference manifest and report artifacts
- documented comparison workflow

### Stage 5 — product surface alignment
Deliverables:
- browser example + playground slide-viewer updates
- docs updates for PPTX visual capabilities and limits

### Stage 6 — hardening exit
Deliverables:
- fresh corpus run
- architect review
- verifier review
- deslop pass on changed files

## ADR
### Decision
Build a corpus-backed PPTX scene-rendering upgrade with screenshot verification.

### Drivers
- current semantic rendering is insufficient for exported-slide resemblance
- sample6 parse failure blocks trust in corpus-wide work
- paired PNGs make visual verification possible now

### Alternatives considered
- example-only heuristics
- external rasterization bridge

### Why chosen
This is the smallest path that improves the actual library and creates a durable verification lane.

### Consequences
- More upfront work in parser/renderer/test harness
- Better foundation for future PPTX fidelity phases
- Visual verification becomes part of the repo contract

### Follow-ups
- extend beyond declared sample subset once harness stabilizes
- progressively support more DrawingML effects/fills/shape families

## Available-agent-types roster
- `planner`, `architect`, `critic`, `executor`, `debugger`, `test-engineer`, `verifier`, `researcher`, `writer`, `code-reviewer`

## Ralph lanes
- **Lane A:** corpus + root-detection correctness (`packages/core`, `packages/pptx`, tests)
- **Lane B:** PPTX renderer scene semantics (`packages/render`, `packages/pptx`)
- **Lane C:** screenshot/reference harness + reports (`tools`, `tests`, possible fixtures/artifacts)
- **Lane D:** product surfaces/docs (`examples/basic`, `playground`, docs)
- **Lane E:** sign-off and consolidation

## Suggested reasoning by lane
- Lane A: high
- Lane B: high
- Lane C: medium-high
- Lane D: medium
- Lane E: high

## Verification path
1. parse counts for sample1-6, including sample6 root fix
2. screenshot/reference captures for declared subset
3. browser example/playground verification
4. full repo verification commands
5. architect review
6. verifier review
