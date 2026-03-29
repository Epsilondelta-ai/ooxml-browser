# Consensus Plan: PPT sample render fidelity

## RALPLAN-DR summary
### Principles
1. Corpus truth beats intuition.
2. Scene semantics belong in parser/render packages; product chrome stays thin.
3. Visual verification must be repeatable and slide-indexed.
4. Fix unsupported package detection before polishing supported slides.
5. Improve slide-family fidelity in bounded, sample-driven waves.

### Decision drivers
1. The paired `pptx/png` corpus provides direct evidence of real user expectations.
2. Current PPTX output is still a semantic/debug renderer, not a slide-export-quality view.
3. Sample6 shows a structural parser/root-detection gap that blocks honest corpus coverage.

### Viable options
- **A. Corpus-first scene upgrade with parser + visual harness** — chosen
  - Pros: directly targets the reference images, creates durable verification, fixes sample6-class detection
  - Cons: larger up-front harness work before all visual wins are visible
- **B. Example-only heuristics on top of the current renderer** — rejected
  - Pros: fast local improvements
  - Cons: fragile, hard to verify, and insufficient for six heterogeneous sample decks
- **C. Full renderer rewrite before harnessing the corpus** — rejected
  - Pros: cleaner long-term architecture if unlimited scope
  - Cons: too broad, high risk, and poorly bounded for the current repo

### Decision
Choose **Option A**. Build a corpus-backed PPTX fidelity lane that starts with package detection and screenshot/reference verification, then deepens parser/render semantics around the actual sample slide families.

### Pre-mortem
1. We polish the example UI while the core renderer still cannot reproduce the sample slide families.
2. Sample6 stays unsupported and corrupts the final coverage story.
3. Visual diffs are noisy/unactionable because the screenshot environment is not pinned.

### Expanded test plan
- **Unit:** root detection, slide background/fill/theme helpers, media URL resolution, scene-layout helpers
- **Integration:** sample deck open/render/navigate flows, parser+renderer coordination for representative slides
- **E2E:** browser screenshot capture for representative slides per sample family
- **Observability:** slide-indexed verdict JSON, mismatch summaries, unsupported-feature reports

## Stage order
### Stage 0 — corpus grounding + verification architecture
- inventory sample corpus
- define slide-index mapping and artifact layout
- capture sample6 diagnosis

### Stage 1 — package detection + corpus loader
- fix root detection / opener for sample6-class files
- add sample corpus manifest/loader utilities

### Stage 2 — parser depth for visual semantics
- background/fill/placeholder/image/theme extraction
- transform and layer metadata hardening

### Stage 3 — PPTX scene renderer upgrade
- title/content/divider/image slide-family rendering
- actual image media projection
- presentation-mode navigation/rendering surface

### Stage 4 — visual harness + reports
- screenshot capture command
- reference comparison artifacts and thresholds
- mismatch triage reports

### Stage 5 — hardening + surface sync
- docs/playground/example alignment
- residual outlier fixes
- final sign-off evidence

## ADR
### Decision
Adopt a corpus-first PPTX scene-fidelity plan driven by the local `ppt-samples` reference deck exports.

### Drivers
- Real reference images exist for every slide.
- Current renderer lacks enough scene semantics.
- Unsupported root detection blocks honest coverage today.

### Alternatives considered
- Example-only styling pass
- Full renderer rewrite before harnessing the corpus

### Why chosen
It is the smallest path that is still testable, durable, and directly tied to the user's real reference decks.

### Consequences
- We must invest in visual harness tooling now.
- Parser/render changes will likely touch core PPTX semantics, not just UI chrome.
- Completion will be measured by corpus evidence, not by subjective spot checks alone.

### Follow-ups
- Expand the harness to more PPTX corpora after this wave.
- Consider moving stable scene semantics out of example-only code into reusable renderer modules.

## Risks and mitigations
- **Noisy screenshots:** pin capture environment and store structured verdicts.
- **Theme/layout sparsity:** recover direct shape/background semantics and use bounded heuristics where necessary.
- **Scope sprawl:** prioritize representative slide families and sample-driven mismatches only.
- **Parser edge cases:** land sample6 detection fix with a regression test before broader polishing.

## Available-agent-types roster
- `planner`, `architect`, `critic`, `executor`, `debugger`, `test-engineer`, `verifier`, `code-reviewer`, `writer`, `researcher`, `vision`

## Ralph follow-up staffing guidance
- **Lane A — corpus loader + diagnostics:** `executor` / `debugger` (high)
- **Lane B — parser semantics:** `executor` / `architect` (high)
- **Lane C — scene renderer + example/playground:** `executor` / `designer` / `vision` (high)
- **Lane D — visual harness + reports:** `test-engineer` / `verifier` (high)
- **Lane E — docs and evidence surfaces:** `writer` / `verifier` (medium)

## Team launch hints
- Conservative Ralph path: keep implementation sequential but use parallel sidecar review for parser lane vs visual-harness lane.
- If a team burst is needed later, split by lanes A-D and return to Ralph for final verification.

## Team verification path
1. corpus loader recognizes all sample folders and slide counts
2. sample6 handling is explicit and regression-tested
3. representative slides from each sample family render and screenshot successfully
4. verdict report is generated
5. Ralph runs final end-to-end verification and sign-off
