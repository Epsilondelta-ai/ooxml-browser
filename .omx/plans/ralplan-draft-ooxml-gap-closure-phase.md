# Deliberate RALPLAN Draft: OOXML Gap Closure Phase

## Document metadata

- **Project:** OOXML gap closure phase
- **Slug:** ooxml-gap-closure-phase
- **Mode:** ralplan consensus draft, deliberate
- **Grounding snapshot:** `.omx/context/ooxml-gap-closure-20260328T050715Z.md`
- **Baseline inputs reused:** `docs/index.md`, `docs/design/*.md`, `docs/quality/verification.md`, `.omx/plans/prd-frontend-ooxml-library.md`, `.omx/plans/test-spec-frontend-ooxml-library.md`, `.omx/plans/consensus-plan-frontend-ooxml-library.md`
- **Current codebase evidence:** single-file-per-package baseline under `packages/{core,docx,xlsx,pptx,render,editor,serializer,browser,devtools}/src`, fixture-light repo (`fixtures/README.md`), and coarse test coverage concentrated in `tests/*.test.ts`

## Intent

Advance the repository from a representative OOXML parse/render/edit/save baseline to a materially stronger Microsoft Office-meaningful next phase by deepening source-preserving models, expanding serializer-safe coverage, broadening docx/xlsx/pptx fidelity, and institutionalizing corpus-driven interoperability and quality gates.

---

## 1. RALPLAN-DR summary

### Principles
1. **Preservation before regeneration:** untouched package parts, relationships, unknown markup, and dormant compatibility branches must survive unless an edit explicitly targets them.
2. **Shared resolved contracts, format-owned source graphs:** share OPC/XML/diagnostics/transaction infrastructure plus resolved subsystem contracts (styles/themes/tables/drawings/annotations), while keeping docx/xlsx/pptx source semantics and provenance graphs format-owned.
3. **Round-trip evidence over feature claims:** every new fidelity slice must land with no-op/minimal-edit round-trip, reopen, and regression evidence.
4. **Stage by Office-meaningful leverage:** prioritize changes that unblock multiple features or materially improve real Office reopen/render/edit behavior.
5. **Product surfaces are part of the feature:** playground, examples, benchmarks, compatibility matrix, and diagnostics evolve with the core.

### Decision drivers
1. **Current architecture is intentionally representative, but too lossy in typed projections and serializer rewrites** for deeper Office fidelity.
2. **The repo is still structurally shallow** (mostly single-file package entrypoints, sparse fixtures, coarse tests), so gap closure must first create safer module boundaries and evidence surfaces.
3. **The goal is next-phase Office-meaningful improvement without rebooting architecture or narrowing scope**, so the plan must extend the baseline incrementally and verifiably.

### Viable options

#### Option A — Preservation-core first, then format expansion (**chosen**)
- **Shape:** first deepen source-preserving/package mutation infrastructure and test corpus, then expand docx/xlsx/pptx fidelity on top.
- **Pros:** reduces destructive serializer risk, supports all three formats, creates reusable verification rails.
- **Cons:** slower visible feature wins in the first stage; requires discipline to avoid over-abstracting.

#### Option B — Format-first gap closure by product priority
- **Shape:** drive docx/xlsx/pptx feature expansion immediately and retrofit preservation/serializer depth as regressions appear.
- **Pros:** faster visible capability gains per format.
- **Cons:** high risk of divergent models, destructive rewrites, and repeated rework across packages.

#### Option C — Corpus-first black-box interop push
- **Shape:** massively expand fixtures/reopen validation first, using current architecture until failures force redesign.
- **Pros:** fast discovery of real-world breakage and priority signals.
- **Cons:** exposes gaps well, but does not by itself create safe implementation boundaries; can produce noisy failure backlogs without actionable architecture slices.

### Explicit choice
Choose **Option A: preservation-core first, then format expansion**. It best fits the current repo state: the shallow package structure and serializer-heavy rewrites mean deeper fidelity work is too risky unless the source-preserving/package-mutation layer and corpus-backed verification are strengthened first.

---

## 2. Pre-mortem (3 scenarios)

### Scenario 1 — “Feature breadth rises, but round-trip fidelity regresses”
- **Failure mode:** docx/xlsx/pptx gain more parsed fields and rendered output, but serializer rewrites drop unknown nodes, reorder parts, or break Office reopen.
- **Why it could happen:** current serializer paths rebuild major XML parts from normalized models (`packages/serializer/src/index.ts`), while format models omit many source details.
- **Early warning signs:** snapshot churn in untouched parts, relationship ID instability, growing reopen failures on no-op/minimal-edit fixtures.
- **Mitigation:** Stage 1 must introduce source spans/token retention, package mutation helpers, untouched-part diff checks, and explicit “patch-vs-regenerate” rules before broader feature expansion.

### Scenario 2 — “Shared abstractions flatten format semantics”
- **Failure mode:** a generic shared model obscures important docx numbering/sections, xlsx formula/reference/style semantics, or pptx master/layout inheritance.
- **Why it could happen:** pressure to unify everything in current minimal package surfaces.
- **Early warning signs:** repeated escape-hatch fields, cross-format conditional logic inside shared modules, or serializer/renderers depending on weak generic types.
- **Mitigation:** architecture rules must pin shared boundaries to OPC/XML/diagnostics/asset/theme/transaction primitives only; format semantic graphs remain owned inside `packages/docx`, `packages/xlsx`, and `packages/pptx`.

### Scenario 3 — “Corpus and verification overhead outruns delivery”
- **Failure mode:** the team adds large fixture/interop/perf ambitions that slow iteration and produce brittle CI before the core slices stabilize.
- **Why it could happen:** the repo currently has very light fixture infrastructure and no explicit manifest-driven corpus workflow.
- **Early warning signs:** large unchecked binary drops, flaky visual/perf baselines, or stages blocked by non-actionable goldens.
- **Mitigation:** introduce manifest-driven corpus families, tiered gates (micro fixtures -> representative fixtures -> interop corpus), and require stage-specific evidence rather than full-matrix success from day one.

---

## 3. Expanded test plan

### Unit
- `packages/core/src/opc/**`: relationship resolution, content-type updates, part graph mutation, URI normalization, safety budgets.
- `packages/core/src/xml/**`: token tape preservation, namespace/prefix retention, `mc:AlternateContent` active+dormant branch capture, source-span addressing.
- `packages/docx/src/**`: style inheritance, numbering resolution, section/header/footer mapping, comment/revision anchors, table grid normalization, drawing/equation preservation.
- `packages/xlsx/src/**`: styles/numFmt resolution, sharedStrings vs inlineStr, formula tokenization/reference rewrites, merges/frozen panes/names/tables/comments parsing.
- `packages/pptx/src/**`: master/layout/theme inheritance, notes/comments/media relationships, shape property normalization, timing metadata preservation.
- `packages/editor/src/**`: semantic transactions, scoped invalidation hints, undo/redo, relationship-safe asset/part mutations.
- `packages/serializer/src/**`: patch planners, untouched-part preservation, relationship/content-type mutation correctness, deterministic output ordering.

### Integration
- Parse -> semantic graph -> serialize no-op for representative docx/xlsx/pptx fixtures.
- Minimal edits that touch one semantic surface while asserting untouched-part preservation elsewhere.
- Cross-part flows: docx headers/comments/footnotes; xlsx workbook/sharedStrings/styles/tables/drawings; pptx slides/masters/layouts/notes/comments/media.
- Renderer/editor integration asserting derived view model updates only for affected scopes.
- Worker/browser package integration for large fixture parse and render model generation.

### E2E
- **DOCX:** open Office-authored document with styles/numbering/comments/table/header/footer, edit body text + comment + section-adjacent content, save, reopen, and compare diagnostics.
- **XLSX:** open workbook with formulas/styles/merged cells/frozen panes/table, edit values and formulas, save, reopen, and verify sharedStrings/defined names/table integrity.
- **PPTX:** open deck with master/layout/theme/notes/media, edit text and notes, duplicate/reorder slide, save, reopen, and verify master/layout linkage remains intact.
- Playground automation: fixture picker, diagnostics panel, round-trip download, perf counters visible for representative files.

### Observability
- Structured diagnostics snapshot tests by fixture manifest.
- Parse/render/edit/serialize timings emitted per fixture family.
- Mutation traces: which parts changed, why, and whether patch-vs-regenerate policy was used.
- Interop matrix generation as a machine-readable artifact (`docs/quality` + `.omx/plans`/future CI outputs).

### Performance
- Microbenchmarks for XML/token parsing and serializer patching.
- Representative corpus budgets for docx open/render, xlsx sheet activation/edit/serialize, pptx slide open/render/save.
- Memory ceilings for large worksheet/document/slide decks.
- Regression policy: warn on budget drift in early stages, hard-fail after Stage 4 stabilization.

---

## 4. Execution-ready staged plan

## Stage 0 — Planning artifact refresh and repo decomposition map

### Goals
- Turn this draft into updated PRD/test-spec/consensus artifacts.
- Produce a concrete decomposition map from current single-file package entrypoints to target module ownership.

### Deliverables
- Updated `.omx/plans/prd-ooxml-gap-closure-phase.md`
- Updated `.omx/plans/test-spec-ooxml-gap-closure-phase.md`
- Updated `.omx/plans/consensus-plan-ooxml-gap-closure-phase.md`
- Module ownership map appendix under `.omx/plans/`
- Stage-0 budget sheet defining initial provenance-memory smoke thresholds and benchmark gate thresholds
- Stage-0 reopen-evidence policy appendix defining parser CI vs Office/LibreOffice attestation requirements by stage

### Gate
- Planning artifacts explicitly align on stage order, architecture rules, verification gates, and staffing lanes.

## Stage 1 — Preservation-core and module-boundary foundation

### Architecture rules
1. **No format package may directly rewrite unrelated XML parts.** All part mutations flow through shared package mutation helpers.
2. **Every semantic node edited in higher layers must retain source provenance or part ownership metadata.**
3. **Shared modules own resolved contracts, not source semantics:** OPC/package graph, XML/source preservation, diagnostics, asset/theme/annotation primitives, transaction infrastructure, and shared resolved subsystem contracts consumed by render/editor/devtools.
4. **Format semantics stay format-owned:** docx sections/numbering/revisions; xlsx formulas/styles/grid semantics; pptx master/layout/timing semantics.
5. **Patch when scoped, regenerate only when unavoidable, and log why.**

### Stage 1 boundary ADR
- **Decision:** adopt *format-owned source graphs, shared resolved contracts*.
- **Meaning:**
  - `packages/{docx,xlsx,pptx}` own raw-to-semantic source graphs, provenance spans, and format-only invariants.
  - shared modules own normalized contracts for resolved styles/themes/tables/drawings/annotations/assets that render/editor/devtools can consume without duplicating logic three times.
- **Why chosen:** this preserves format richness while still preventing duplication in rendering, diagnostics, and editor surfaces.
- **Constraint:** no shared module may become the source-of-truth for docx numbering, xlsx formula graphs, or pptx master/layout semantics.

### Module boundaries
- `packages/core/src/opc/`: package graph, relationships, content types, mutation helpers.
- `packages/core/src/xml/`: tokenizer, token tape, namespace handling, source spans, compatibility handling.
- `packages/core/src/model/`: shared primitives only (diagnostics, stable ids, raw asset/theme/annotation carriers).
- `packages/core/src/contracts/`: **shared resolved contracts** for styles/themes/tables/drawings/annotations consumed by render/editor/devtools.
- `packages/core/src/serialization/`: low-level deterministic writer and patch utilities.
- `packages/{docx,xlsx,pptx}/src/model/`: format semantic graphs and provenance-bearing adapters.
- `packages/{docx,xlsx,pptx}/src/parser/`: part parsers by concern.
- `packages/{docx,xlsx,pptx}/src/resolve/`: styles/themes/relationships/inheritance.
- `packages/{docx,xlsx,pptx}/src/edit/`: format-specific mutations atop shared editor core.
- `packages/render/src/{docx,xlsx,pptx}/`: view-model builders and HTML/SVG projections.
- `packages/editor/src/`: transaction core, invalidation, undo/redo, format adapter APIs.
- `packages/serializer/src/{docx,xlsx,pptx}/`: format-specific patch planners and writers.

### Work
- Split current monolithic entrypoints into directories with stable re-export boundaries.
- Introduce provenance-bearing semantic node types and source-preserving token/part references.
- Establish `packages/core/src/contracts/` as the single home for shared resolved contracts.
- Add package mutation helpers for parts, relationships, content types, and asset payloads.
- Add fixture manifest schema and corpus tooling skeleton.

### Verification gate
- Existing tests remain green.
- New unit tests cover mutation helpers and source-preserving XML behavior.
- Structural smoke tests confirm re-export compatibility for existing public APIs.
- At least one **persisted representative fixture per format** lands under `fixtures/{docx,xlsx,pptx}/representative/` with manifest metadata.
- Stage-1 no-op and minimal-edit checks assert **untouched-part / untouched-relationship preservation** for those representative fixtures before Stage 2/3/4 open.
- Add one **memory/worker-safe provenance smoke** proving retained token/provenance structures can be materialized and transferred or summarized within Stage-0 thresholds (initial target: representative fixture provenance payload <= 25 MB peak retained data and <= 250 ms worker transfer/summary time on the local benchmark harness).

## Stage 2 — DOCX fidelity expansion

### Parser / resolver expansion
- Style inheritance graph (paragraph/run/table) with defaults and linked styles.
- Numbering model: abstract numbering + instance overrides + level text/format.
- Sections with header/footer linkage and page settings.
- Comments/revisions anchors and metadata.
- Table properties/grid spans and drawing/image/equation preservation hooks.

### Render / editor / serializer expansion
- Render computed styles, numbering labels, section-aware header/footer surfacing, richer table output, comment anchors.
- Editor mutations for paragraph/run styles, list operations, comments, header/footer-safe text edits.
- Serializer patch planners for document/comments/header/footer/numbering/styles/footnotes-endnotes parts with untouched-part preservation.

### Verification gate
- Representative docx fixtures for styles/numbering/headers/comments/tables/equations pass unit+integration coverage.
- No-op + minimal-edit **parser reopen** checks pass in CI for declared representative Office-authored docs; required Office/LibreOffice attestation manifests are present for the stage-owned fixture set.
- Visual and layout checks run through the docx render snapshot harness defined in Stage 0.
- Playground/example exposes docx diagnostics and round-trip for the new fixture family.

## Stage 3 — XLSX fidelity expansion

### Parser / resolver expansion
- Workbook style table, number formats, fills/fonts/borders alignment.
- Formula/reference parsing and rewrite helpers.
- Merges, frozen panes, defined names, tables, comments, drawings, chart references.
- SharedStrings + inlineStr + direct value preservation rules.

### Render / editor / serializer expansion
- Grid/view model for merges, frozen panes, style formatting, formula display/value separation.
- Editor mutations for cell value/formula/style, sheet order, merges, table-safe edits, defined-name-safe rewrites.
- Serializer support for workbook, worksheets, styles, sharedStrings, tables, comments, drawings, and relationship-safe reference updates.

### Verification gate
- Representative xlsx fixtures pass parser reopen in CI after no-op/minimal edit; required Office/LibreOffice attestation manifests are present for declared fixtures.
- Formula/reference rewrite tests cover row/column insert-like edits and table/name preservation.
- Benchmarks track workbook open, active sheet render, edit, and save timings against Stage-0 thresholds.

## Stage 4 — PPTX fidelity expansion

### Parser / resolver expansion
- Master/layout/theme inheritance chain.
- Notes/comments/media and richer relationship graph coverage.
- Shape taxonomy beyond plain text shapes.
- Timing/animation preservation as opaque+structured metadata.

### Render / editor / serializer expansion
- Slide scene graph respects master/layout/theme-derived defaults.
- Notes/comments/media presentation surfaces in render/playground.
- Editor mutations for text, slide order/duplication, notes, selected shape properties.
- Serializer preserves master/layout/timing relationships and untouched media parts.

### Verification gate
- Representative pptx fixtures for master/layout/notes/comments/media survive round trip and parser reopen in CI; required Office/LibreOffice attestation manifests are present for declared fixtures.
- Render snapshots show inherited text/theme defaults on curated decks through the stage-owned visual harness.
- Notes/slide edit E2E passes with reopen validation.

## Stage 5 — Shared serializer depth, interop corpus, and product/quality surfaces

### Shared expansion
- Patch planner rules by part type and mutation category.
- Interop corpus families under the canonical tree: `fixtures/shared/{opc,xml,security}`, `fixtures/{docx,xlsx,pptx}/{micro,representative,interop,perf}`, and `fixtures/manifests/{docx,xlsx,pptx,shared}`.
- Example app + playground automation for fixture loading, diagnostics, diff summaries, and download/reopen loops.
- Compatibility matrix and benchmark report generation under `docs/quality` and `benchmarks/`.
- Docs sync work updates `docs/quality/verification.md` and `fixtures/README.md` to the same canonical fixture/manifests layout before the stage closes.

### Verification gate
- Interop matrix records parse/render/edit/serialize/reopen status for representative Office + LibreOffice fixtures.
- `npm test`, `npm run typecheck`, `npm run lint`, `npm run build`, `npm run bench` pass with expanded suites.
- Architect review confirms shared/format boundary integrity; verifier review confirms evidence sufficiency.

## Stage 6 — MS Office-meaningful hardening exit

### Definition of done
The next phase is done when all of the following are true:
1. **Preservation depth:** untouched parts, relationships, and unknown safe markup survive representative no-op/minimal-edit workflows.
2. **DOCX:** styles/numbering/sections/headers-comments/tables/drawings-equations have representative parse-render-edit-serialize support with explicit diagnostics for remaining unsupported edges.
3. **XLSX:** styles/number formats/formulas/merges/frozen panes/names/tables/comments/charts have representative parse-render-edit-serialize support with safe rewrite rules.
4. **PPTX:** master/layout/theme inheritance, notes/comments/media, richer shapes, and timing preservation have representative support.
5. **Product surfaces:** examples/playground/benchmarks expose the supported fidelity slices and diagnostics.
6. **Evidence:** corpus-backed unit/integration/e2e/perf/interop gates pass, plus architect + verifier sign-off.

---

## 5. Fixture / corpus / interop plan

### Corpus tiers
1. **Micro fixtures:** hand-authored single-feature XML/package fixtures for parser/serializer edge cases.
2. **Representative fixtures:** Office-authored sample docs per feature cluster.
3. **Interop fixtures:** Office + LibreOffice saved documents with manifest expectations and tolerated diffs.
4. **Stress fixtures:** large docs/sheets/decks for perf and memory.
5. **Security fixtures:** malformed ZIP/XML/relationship/content-type cases.

### Proposed directory shape
- `fixtures/shared/{opc,xml,security}`
- `fixtures/docx/{micro,representative,interop,perf}`
- `fixtures/xlsx/{micro,representative,interop,perf}`
- `fixtures/pptx/{micro,representative,interop,perf}`
- `fixtures/manifests/{docx,xlsx,pptx,shared}`

This is the **canonical fixture/manifests layout** for the phase; all stage references and docs must use this tree.

### Manifest minimum fields
- fixture id
- source application + version
- format + feature tags
- expected diagnostics
- supported edit operations
- reopen expectations (Office / LibreOffice / parser)
- visual baseline references
- perf class
- tolerated diffs / known gaps

### Interop workflow
- Add a machine-readable matrix generator.
- Track each representative fixture through parse, render, edit, serialize, parser reopen, Office reopen, LibreOffice reopen.
- Fail hard only on declared-stage fixtures; keep future fixtures visible as known gaps until their stage opens.

### Reopen validation workflow
- **Automated path:** parser reopen always runs in CI for no-op/minimal-edit fixtures and records machine-readable status.
- **Interop attestation path:** Office/LibreOffice reopen results are tracked in `fixtures/manifests/**` with either automated evidence links or manual attestation records.
- **Stage policy:**
  - Stages 1-2: parser reopen is mandatory; Office/LibreOffice attestation required only for the declared representative seed fixtures.
  - Stages 3-4: parser reopen remains mandatory; attestation expands to all stage-owned representative fixtures.
  - Stages 5-6: parser reopen + attestation are both release-blocking for declared interop fixtures.

---

## 6. Ralph staffing lanes and available-agent-types roster

## Available-agent-types roster
- `planner` — artifact refresh, sequencing, milestone/gate alignment.
- `architect` — boundary review, preservation strategy review, stage gate design.
- `executor` — primary implementation for core, format, render, editor, serializer slices.
- `debugger` — regression isolation for reopen, serializer, and fidelity failures.
- `test-engineer` — fixture manifests, integration/E2E/perf/interop harnesses.
- `verifier` — evidence review and completion validation.
- `critic` / `code-reviewer` — plan pressure test, milestone review, anti-slop review.
- `researcher` — targeted standards/interop/source-reference lookup when Office behavior is ambiguous.
- `dependency-expert` — only if low-level XML/ZIP/interop tooling decisions become necessary.
- `writer` — docs/compatibility matrix/report polish during later stages.

## Suggested Ralph lanes

### Lane A — Preservation core (reasoning: high)
- Ownership: `packages/core`, shared fixture schema/tooling, package mutation helpers.
- Preferred agents: `executor`, `architect`, `test-engineer`.

### Lane B — DOCX fidelity (reasoning: high)
- Ownership: `packages/docx`, docx portions of `packages/render`, `packages/editor`, `packages/serializer`, `fixtures/docx`.
- Preferred agents: `executor`, `debugger`, `test-engineer`.

### Lane C — XLSX fidelity (reasoning: high)
- Ownership: `packages/xlsx`, xlsx render/editor/serializer surfaces, `fixtures/xlsx`, perf focus for workbook/grid flows.
- Preferred agents: `executor`, `debugger`, `test-engineer`.

### Lane D — PPTX fidelity (reasoning: high)
- Ownership: `packages/pptx`, pptx render/editor/serializer surfaces, `fixtures/pptx`.
- Preferred agents: `executor`, `debugger`, `test-engineer`.

### Lane E — Product/quality surfaces (reasoning: medium-high)
- Ownership: `examples`, `playground`, `benchmarks`, `docs/quality`, interop matrix generation.
- Preferred agents: `executor`, `writer`, `test-engineer`, `verifier`.

### Lane F — Sign-off and consolidation (reasoning: high)
- Ownership: end-to-end verification, stage evidence, architecture review, risk burndown.
- Preferred agents: `architect`, `verifier`, optional `critic`.

## Ralph execution guidance
- Prefer sequential stage progression with bounded intra-stage subagents rather than full parallel swarm at the beginning.
- Do not open all format lanes before Stage 1 boundary/preservation work lands.
- After Stage 1, Ralph can run targeted parallel sublanes for DOCX/XLSX/PPTX while keeping serializer/core ownership centralized.
- Final consolidation path: feature lane -> shared serializer/core check -> test/perf/interop evidence -> architect review -> verifier review.

## Team / launch hints for follow-up consensus artifacts
- **Ralph-first path:** Stage 1 sequentially, then split Stages 2-4 into bounded format sublanes with shared serializer oversight.
- **Team path:** 4-5 lane split only after Stage 1 package boundaries are merged: Core, DOCX, XLSX, PPTX, Verification/Product.
- **Verification path:** after any team burst, return to Ralph for final integration, full command suite, architect review, and verifier sign-off.

---

## Acceptance criteria for updated PRD / test-spec / consensus artifacts

### PRD must add
- explicit Office-meaningful next-phase targets by format
- preservation-depth requirements
- module-boundary rules
- product-surface obligations
- stage-based release criteria

### Test spec must add
- manifest-driven corpus workflow
- untouched-part preservation assertions
- interop reopen matrix expectations
- observability/perf reporting requirements
- stage-specific evidence thresholds

### Consensus plan must add
- chosen option rationale from this draft
- pre-mortem scenarios and mitigations
- stage gates and staffing lanes
- explicit architect/verifier review expectations

