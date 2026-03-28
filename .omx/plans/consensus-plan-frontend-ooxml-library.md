# Consensus Plan: Frontend OOXML Library

## Planning metadata
- Mode: `$ralplan --deliberate`
- Grounding snapshot: `.omx/context/frontend-ooxml-library-20260328T034425Z.md`
- Supporting docs: `docs/index.md`
- Status: APPROVED

## RALPLAN-DR summary

### Principles
1. **Package-first truth**: treat OOXML as an OPC relationship graph before format-specific semantics.
2. **Round-trip preservation over lossy simplification**: preserve unknown markup, compatibility branches, and untouched binary parts.
3. **Browser-first execution**: public APIs and examples must run in frontend environments with worker-safe offloading.
4. **Shared model, specialized views**: parser, renderer, editor, and serializer share a coherent IR while view/layout layers remain format-specialized.
5. **Verification is product work**: fixtures, benchmarks, examples, and diagnostics ship alongside features.

### Decision drivers
1. High-fidelity parse/edit/save across docx/xlsx/pptx without Node-only assumptions.
2. Maintainable modular architecture that can scale to advanced OOXML features.
3. Continuous verification against corpus, interop, performance, and round-trip regressions.

### Viable options

#### Option A — Package graph + source-preserving XML + shared OOXML IR + format-specialized render/edit modules (**chosen**)
Pros:
- best fit for round-trip preservation and extension markup handling
- shared services reduce duplication across docx/xlsx/pptx
- supports layered verification and future collaboration/worker features
Cons:
- larger upfront architecture investment
- requires careful IR boundaries to avoid leaky abstractions

#### Option B — Independent per-format pipelines with minimal shared abstraction
Pros:
- faster local optimization for each format
- simpler mental model in earliest stages
Cons:
- duplicates OPC/XML logic
- hurts consistency for shared subsystems and public APIs
- makes cross-format tooling/examples/devtools harder

#### Option C — Render-first lossy normalization into HTML/Canvas-native models
Pros:
- can accelerate viewer scenarios
- browser rendering integration seems direct initially
Cons:
- breaks round-trip preservation goals
- weak editor/serializer coherence
- poor support for unsupported/unknown OOXML structures

Choice rationale:
- Option A is the only option that satisfies full-scope parse/render/edit/serialize requirements while keeping interoperability and verification credible.

### Pre-mortem
1. **Word layout drift**: browser pagination/floating-layout fidelity diverges from Office; mitigation is a dedicated layout abstraction, fixture corpus, and explicit fidelity tiers.
2. **Spreadsheet complexity explosion**: formulas/styles/virtualization interact unpredictably; mitigation is storage-vs-display model separation, pluggable recalc, and mandatory virtualization.
3. **Serializer regressions**: edits break untouched OOXML content; mitigation is source-preserving XML/token tape, patch-based writes, and no-op/minimal-edit round-trip tests from early stages.

### Expanded test plan
- **Unit**: OPC/XML/IR/helpers/transactions
- **Integration**: package -> parse -> serialize -> reopen across formats
- **E2E**: browser examples and playground edit/save workflows
- **Observability**: diagnostics snapshots, serializer diff summaries, benchmark JSON, worker timing traces
- **Performance**: parse/render/edit/save latency and memory budgets on representative corpus

## Planner draft

### Product goals
- deliver a production-meaningful frontend OOXML library with browser-first parse/render/edit/serialize support
- establish composable packages, shared IR, worker protocols, examples, playground, benchmarks, and docs
- ensure core round-trip fidelity and verification infrastructure are present before claiming completion

### Quality targets
- package fidelity, semantic fidelity, visual fidelity, behavioral fidelity, and round-trip preservation tracked separately
- security-first ingestion for untrusted documents
- compatibility tracking across Office and LibreOffice corpora
- stable APIs with diagnostics and devtools surfaces

### Architecture

#### Monorepo layout
- `packages/core`
- `packages/opc`
- `packages/xml`
- `packages/ir`
- `packages/serializer`
- `packages/docx`
- `packages/xlsx`
- `packages/pptx`
- `packages/render-core`
- `packages/render-docx`
- `packages/render-xlsx`
- `packages/render-pptx`
- `packages/editor-core`
- `packages/editor-docx`
- `packages/editor-xlsx`
- `packages/editor-pptx`
- `packages/react`
- `packages/devtools`
- `packages/worker`
- `examples/`
- `playground/`
- `fixtures/`
- `benchmarks/`
- `tests/`

#### Internal model stack
1. `PackageGraph`
2. source-preserving XML/token tape
3. normalized `OfficeDocumentModel`
4. derived render/layout projections
5. transaction/invalidation/serialization patch layer

### Parser design
- shared ingest pipeline: blob/file -> zip security gate -> package graph -> XML tokenizer -> format normalizers -> IR
- lazy parse heavy parts; worker support for large documents
- expose diagnostics and degraded-mode handling

### Renderer design
- docx: HTML/CSS paginated + SVG overlays
- xlsx: virtualized HTML grid + Canvas/SVG overlays
- pptx: SVG/HTML hybrid scene graph
- shared asset/style/theme/layout helpers

### Editor design
- semantic transaction API with undo/redo
- format-specific selection models on top of shared transaction/invalidation system
- patch-based serialization with dependency updates
- collaboration hooks at transaction layer

### Serializer design
- patch touched parts where possible
- preserve untouched unknown content
- deterministic relationship/content-type/shared-string/style rewrites
- no-op and minimal-edit round-trip tests mandatory

### Format-specific strategy

#### DOCX
- prioritize stories, paragraphs/runs, styles, numbering, sections, tables, comments, headers/footers, tracked changes, drawings
- derive page and continuous views
- expose story-aware editor APIs

#### XLSX
- prioritize workbook/sheets/shared strings/styles/formulas/defined names/merges/frozen panes/drawings/tables
- build virtualized grid and formula-preserving editor path
- keep calculation engine pluggable

#### PPTX
- prioritize presentation/slides/masters/layouts/themes/notes/comments/shapes/media/timing preservation
- build slide scene graph with placeholder-aware editing
- preserve timing/animation metadata even where playback is partial

### Corpus / fixture / benchmark strategy
- create micro fixtures per feature and real-world mixed corpora
- add interop and security corpora early
- maintain benchmark harness for parse/render/edit/save/memory

### Docs / examples / playground strategy
- docs stay synchronized with implementation
- examples for open/render/edit/save per format plus package inspector
- playground supports drag/drop, diagnostics, package explorer, and save-back flow

### Staged implementation order
1. monorepo/tooling scaffolding + package boundaries + docs links
2. shared core (`core`, `opc`, `xml`, `ir`, `serializer`) + fixture harness
3. docx parser + serializer + smoke render path + tests
4. xlsx parser + serializer + virtual grid path + tests
5. pptx parser + serializer + slide render path + tests
6. shared editor core + per-format editing adapters
7. examples + playground + devtools + benchmark harness
8. compatibility hardening, regression fixes, and final verification

### Parallelizable task decomposition
- lane A: monorepo/core/OPC/XML/IR scaffolding
- lane B: fixture corpus + test harness + benchmark harness
- lane C: format-specific adapters (docx/xlsx/pptx) once core contracts exist
- lane D: render packages once parse IR contracts stabilize
- lane E: editor/devtools/examples/playground after first parse/render loops land

### Verification criteria
- build/typecheck/lint clean
- unit/integration/browser tests pass
- affected-file diagnostics clean
- round-trip tests pass for representative corpus
- examples/playground validated
- docs match implementation

### Definition of done
- parse/render/edit/serialize flows exist for docx/xlsx/pptx
- fixtures/tests/examples/benchmarks/docs are present and coherent
- no unresolved TODOs or known broken states remain
- final architect/verifier review approves evidence

## Architect review

### Strongest steelman antithesis
A single shared IR could become too generic, leading to awkward abstractions that hide critical format differences and slow delivery. In particular, docx pagination, xlsx formula/grid semantics, and pptx scene-graph behavior may deserve deeper format-native models than a shared abstraction can comfortably support.

### Tradeoff tensions
1. **Shared IR vs format-native precision**: too much generalization hurts fidelity; too little sharing hurts maintainability.
2. **Patch-based serialization vs deterministic regeneration**: patching preserves fidelity, but some tables/indexes are easier to regenerate.
3. **Browser-first ergonomics vs full Office compatibility**: frontend safety/performance constraints will sometimes cap exact parity.

### Synthesis
- keep shared IR limited to genuinely shared primitives and package services
- retain format-native subtrees and adapters for areas with high semantic divergence
- use hybrid serialization: patch untouched/opaque content, regenerate known dependent index structures deterministically

Architect verdict: APPROVE WITH SYNTHESIS

## Critic evaluation

Checks:
- principle/option consistency: pass
- alternatives fairly considered: pass
- risk mitigation clarity: pass
- acceptance criteria and verification steps testable: pass
- deliberate-mode pre-mortem and expanded test plan present: pass

Critic verdict: APPROVE

## ADR

### Decision
Adopt a browser-first monorepo architecture centered on OPC package parsing, source-preserving XML, shared OOXML IR, and format-specialized render/edit modules.

### Drivers
- required full-scope docx/xlsx/pptx support
- round-trip preservation
- need for reusable shared subsystems and verification tooling

### Alternatives considered
- independent per-format stacks
- render-first lossy normalization

### Why chosen
It best satisfies fidelity, maintainability, extensibility, and verification requirements simultaneously.

### Consequences
- more up-front architectural work
- stronger long-term coherence across packages and APIs
- requires disciplined IR boundary design and fixture-first verification

### Follow-ups
- scaffold monorepo and core contracts first
- implement fixture harness before deep format work
- gate later edits with round-trip and render evidence

## Available-agent-types roster for execution
- `planner`: stage sequencing and recovery planning
- `architect`: design guardrails and stage review
- `critic`: quality challenge and approval gate
- `executor`: primary implementation lane
- `test-engineer`: fixtures, test harness, e2e, benchmarks
- `debugger`: regression triage and failure isolation
- `verifier`: final evidence collection and completion checks
- `researcher`: targeted standards/interop follow-up as needed
- `code-reviewer` / `security-reviewer`: optional hardening reviews

## Suggested staffing lanes for later `$ralph`
- **Implementation lane:** `executor` for scaffolding/core/format modules
- **Evidence/regression lane:** `test-engineer` + `debugger` for fixtures, failing cases, perf traces
- **Sign-off lane:** `architect` then `verifier` for stage-end approval

Suggested reasoning levels by lane:
- implementation lane: high
- evidence lane: medium/high
- sign-off lane: high

## Team launch hints
- `$team` or `omx team` should split work into core, fixtures/tests, and format-specific lanes only after core contracts are defined.
- For initial execution, `$ralph` is preferred so the early architecture, scaffolding, and verification loops remain tightly sequenced.
- If team mode is later used, route final integration/fidelity verification back through `$ralph` before declaring completion.

## Team -> Ralph verification path
1. Team lanes land bounded working increments with local verification.
2. Ralph re-integrates the combined branch, reruns full build/test/typecheck/diagnostics.
3. Ralph collects fresh round-trip, example/playground, and benchmark evidence.
4. Architect/verifier sign off the integrated result.
