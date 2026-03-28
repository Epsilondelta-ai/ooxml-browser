# Consensus Execution Plan: Frontend OOXML Library

## RALPLAN-DR summary

### Principles
1. Browser-first runtime and API design.
2. Shared round-trip-friendly IR across parse/render/edit/save flows.
3. Relationship-driven package traversal and preservation-first serialization.
4. Corpus-backed fidelity, compatibility, and performance verification.
5. Extensible architecture for advanced subsystems without fragmenting the core model.

### Decision drivers
1. **Round-trip fidelity** must survive unsupported and extension markup.
2. **Cross-format coherence** is required to keep docx/xlsx/pptx from becoming three disconnected products.
3. **Frontend performance and safety** must be designed in from the start.

### Viable options

#### Option A — Shared core with format adapters (**chosen**)
- **Pros:** maximizes reuse; enables one package graph, XML, IR, serializer, worker, and verification stack; makes multi-format editor/devtools practical.
- **Cons:** requires upfront architecture discipline and careful abstraction to avoid lowest-common-denominator design.

#### Option B — Independent per-format stacks with later convergence
- **Pros:** faster local iteration per format; fewer early abstractions.
- **Cons:** duplicated parsers/serializers/subsystems; expensive later convergence; inconsistent public APIs and verification story.

#### Option C — Viewer-first pipeline with editing added later
- **Pros:** simpler early layout/render focus.
- **Cons:** violates project requirement to deliver a complete parse/render/edit/save library and usually leads to lossy IR that blocks round-trip-safe editing.

### Decision
Choose **Option A**: shared core with format adapters.

### Pre-mortem (deliberate mode)
1. **Layout-fidelity stall:** docx/pptx text metrics and pagination diverge too far from Office.
   - Mitigation: build corpus-driven calibration loops, separate semantic fidelity from visual fidelity, and isolate layout engine contracts early.
2. **IR brittleness:** source-preserving requirements are compromised by overly normalized ASTs.
   - Mitigation: source token tape + semantic AST dual model, serializer patch pipeline, unknown-markup regression fixtures.
3. **Performance collapse on large files:** worksheet/document parsing and render indexing exceed browser budgets.
   - Mitigation: lazy part loading, worker offload, virtualization, and perf corpus gates from the first scaffolding stage.

### Expanded testing strategy
- **Unit:** OPC/XML/IR/style/theme/parser/serializer/transaction primitives.
- **Integration:** parse-render-edit-serialize loops by format and subsystem.
- **E2E:** browser open/edit/save flows for representative fixtures.
- **Observability:** diagnostics, perf traces, cache invalidation signals, worker progress/cancellation.
- **Performance:** corpus-based latency and memory thresholds in CI/local benchmarks.

## ADR

### Decision
Use a monorepo with shared core packages (`core`, `opc`, `xml`, `ir`, `serializer`, `render-core`, `editor-core`, `worker`) and format adapters (`docx`, `xlsx`, `pptx`, renderer/editor companions), plus product surfaces (`react`, `devtools`, `examples`, `playground`, `bench`).

### Drivers
- Shared OOXML/OPC mechanics across all formats.
- Need for one coherent round-trip-safe editor model.
- Requirement to deliver docs/examples/playground/benchmarks together with code.

### Alternatives considered
- independent per-format stacks
- viewer-first delivery
- server-centric architecture with browser wrappers

### Why chosen
It best satisfies browser-first, multi-format, fidelity, and extensibility requirements with manageable long-term complexity.

### Consequences
- More upfront scaffolding and IR work.
- Better long-term reuse, testing consistency, and API clarity.
- Easier worker, devtools, and verification integration.

### Follow-ups
- validate abstraction boundaries during initial scaffolding
- keep format-specific escape hatches where semantics genuinely diverge
- measure package size/runtime costs continuously

## Architecture and module boundaries

### Layered stack
1. **`@ooxml/core`** — diagnostics, errors, ids, utility types.
2. **`@ooxml/opc`** — ZIP/OPC package reader, content types, relationships, package graph.
3. **`@ooxml/xml`** — tokenizer, namespace registry, source-preserving token tape, writer.
4. **`@ooxml/ir`** — normalized shared IR + format container types.
5. **`@ooxml/{docx,xlsx,pptx}`** — format parsers, resolvers, edit adapters.
6. **`@ooxml/render-*`** — view model/layout/render backends.
7. **`@ooxml/editor-*`** — selection, transactions, mutation commands.
8. **`@ooxml/serializer`** — patch planner, deterministic writers, package rebuilder.
9. **`@ooxml/worker`** — worker protocol + task runners.
10. **product surfaces** — React adapter, devtools, examples, playground, benchmarks.

### Internal IR / AST strategy
- **Source layer:** part-scoped XML token tape preserving unknown markup, prefixes, `mc:*`, and writer-relevant trivia.
- **Semantic AST:** typed OOXML node trees per part.
- **Document IR:** shared cross-format primitives (story/sheet/slide containers, text/table/drawing/style/annotation/asset nodes).
- **Derived indexes:** layout caches, dependency graphs, selection maps, relationship reverse index.

### Parser design
- central-directory ZIP inspection and safety budgets
- relationship-driven traversal
- lazy part parse by access pattern
- SAX/token streaming for large parts
- structured diagnostics and degraded open modes

### Renderer design
- docx: HTML/CSS text flow + SVG overlays + page/continuous projections
- xlsx: virtualized HTML grid + Canvas/SVG overlays + print projection
- pptx: SVG/HTML hybrid slide scene graph + notes/sorter projections

### Editor design
- semantic transactions with reversible operations
- format-specific selection models atop shared transaction core
- undo/redo at transaction granularity
- clipboard bridges and structural edits
- collaboration extension hooks at transaction layer

### Serializer design
- part patching when possible, deterministic package rewrite when required
- preserve untouched parts/relationships/content types as much as possible
- stable ids/relationship ids where feasible

## Format-specific execution strategy

### DOCX lane
- story registry, styles, numbering, sections, comments, notes, revisions, headers/footers, drawings, equations
- paginated + continuous layout
- revision-aware editor actions and serializer

### XLSX lane
- workbook/sheet/shared strings/styles/formulas/tables/comments/drawings/charts
- virtualized grid, frozen panes, sheet management
- formula parser and reference rewrite support; recalculation engine pluggable

### PPTX lane
- presentation/slide/master/layout/theme/notes/comments/media/timing metadata
- slide scene graph and text editing
- master/layout-preserving edits and serializer

## Staged implementation order

### Stage A — Workspace and core scaffolding
- monorepo tooling
- shared TS config, build, lint, test harness
- core/opc/xml/ir package skeletons
- docs site/examples/playground/bench placeholders

### Stage B — OPC + XML + IR foundation
- package graph
- XML tokenizer/writer
- namespace + markup compatibility handling
- shared diagnostics
- initial fixtures and unit tests

### Stage C — Format parsers
- docx/xlsx/pptx package parsing and normalized IR projection
- corpus fixtures and round-trip-open integration tests

### Stage D — Renderers
- docx page/continuous renderer
- xlsx grid renderer
- pptx slide renderer
- browser examples and visual tests

### Stage E — Editor core and format editing
- transactions, selection, undo/redo
- docx editing commands
- xlsx cell/sheet editing commands
- pptx text/shape/slide editing commands

### Stage F — Serializer + round-trip hardening
- patch planner and package writer
- no-op/minimal-edit round-trip tests
- interop corpus reopen validation harness

### Stage G — Product surfaces + hardening
- React adapter
- playground and devtools
- benchmark harness
- docs expansion and compatibility matrices

## Milestone gates

### Stage A gate
- workspace builds
- test runner executes
- examples/playground shells boot
- package boundaries documented

### Stage B gate
- OPC/XML/IR unit tests green
- representative OPC/XML fixtures parse successfully
- diagnostics emitted for malformed/security fixtures

### Stage C gate
- representative docx/xlsx/pptx fixtures parse into IR
- no-op serializer smoke tests reopen through parser
- unsupported content preserved with diagnostics

### Stage D gate
- representative document/page, worksheet region, and slide render in browser examples
- visual baselines exist for core fixtures
- render-layer diagnostics/devtools inspectors usable

### Stage E gate
- text/cell/shape editing works in examples
- undo/redo passes integration tests
- selection and clipboard smoke coverage exists per format

### Stage F gate
- no-op and minimal-edit round-trip suites pass for core fixtures
- serializer updates dependent tables/relationships deterministically
- interoperability reopen checks recorded

### Stage G gate
- docs site, playground, examples, and benchmark harness are wired into CI/release workflow
- public API docs and compatibility matrix are published in-repo


## Stage exit criteria

### Stage A exit criteria
- workspace build/test/lint/typecheck commands succeed
- package boundaries exist with documented ownership
- placeholder examples/playground/bench apps boot

### Stage B exit criteria
- OPC/XML/IR foundation passes unit tests
- representative packaging/security fixtures parse with diagnostics
- serializer can round-trip raw XML/token tape for targeted fixtures

### Stage C exit criteria
- representative docx/xlsx/pptx fixtures parse into normalized IR
- no-op round-trip integration tests exist for each format
- parser diagnostics and degraded-open paths are covered

### Stage D exit criteria
- browser examples render representative document/page/grid/slide fixtures
- visual snapshot baselines exist for each format
- perf telemetry captures first render metrics

### Stage E exit criteria
- semantic transactions, selection models, and undo/redo work for core format actions
- browser E2E edit flows pass for each format
- serializer patch invalidation paths are exercised by edit tests

### Stage F exit criteria
- minimal-edit round-trip tests pass for each format
- interop matrix includes Office/LibreOffice reopen checks for representative fixtures
- unsupported/unknown markup preservation tests are green

### Stage G exit criteria
- docs/examples/playground/devtools/bench surfaces are usable end-to-end
- compatibility matrix and benchmark baselines are published in repo
- final architect + verifier sign-off evidence is captured

## Parallelizable task decomposition

### Lane 1 — Core/package/XML
- scaffold workspace
- implement opc/xml/token/diagnostic foundation
- own shared fixtures for packaging and XML

### Lane 2 — Format parsing
- docx parser lane
- xlsx parser lane
- pptx parser lane
- shared IR alignment checkpoints

### Lane 3 — Rendering/product surfaces
- render-core contracts
- format renderers
- examples/playground shell/devtools inspectors

### Lane 4 — Editing/serialization/verification
- transaction core
- serializer pipeline
- corpus tooling, golden tests, benchmarks, CI

## Available-agent-types roster for execution

- `architect` — high-level reviews, boundary validation, tradeoff analysis
- `executor` — implementation lanes (default)
- `debugger` — failure diagnosis and regression isolation
- `test-engineer` — fixture strategy, browser E2E, perf harness
- `verifier` — completion evidence and claim validation
- `code-reviewer` / `critic` — milestone review and simplification pressure
- `researcher` / `dependency-expert` — targeted external package/doc evaluation when necessary

## Suggested staffing guidance

### Ralph execution lanes
- **Implementation lane:** `executor` agents for core/package/XML, format modules, renderers, editor/serializer slices.
- **Evidence/regression lane:** `test-engineer` plus `debugger` for fixtures, CI, and failures.
- **Final sign-off lane:** `architect` then `verifier`, with optional `code-reviewer` for large diffs.

### Team execution hints
- `omx team` / `$team` can split along Core, DOCX, XLSX, PPTX, and Verification surfaces once workspace scaffolding is stable.
- Use shared ownership boundaries by package directory to avoid merge conflicts.
- Team -> Ralph verification path: team builds isolated lanes, then Ralph consolidates, runs end-to-end verification, closes gaps, and performs final architect/verifier sign-off.


## Reasoning and launch guidance

### Suggested reasoning levels by lane
- Core/package/XML: **high**
- Format parsing lanes: **high**
- Rendering lanes: **medium -> high** depending on layout complexity
- Editing/serialization lanes: **high**
- Verification/perf lanes: **medium** for routine evidence, **high** for regressions

### Concrete team launch hints
- `omx team --help` followed by a lane split across Core, DOCX, XLSX, PPTX, Verification once Stage A is merged
- `$team frontend ooxml core/xml/opc lane + docx lane + xlsx lane + pptx lane + verification lane`
- Ralph remains the preferred final consolidation path after any team burst

## Verification criteria

- build/lint/typecheck clean
- unit + integration tests pass
- browser E2E passes for representative fixtures
- parse/render/edit/serialize flows validated for docx/xlsx/pptx
- no-op + minimal-edit round-trip diffs acceptable for supported fixtures
- examples/playground operational
- benchmark suite produces baseline numbers
- docs and feature matrix updated

## Definition of done

The plan is complete when the repo contains a working browser-first OOXML library with package/XML/IR foundations, format parsers, renderers, editors, serializer, tests, fixtures, examples, playground, benchmark harness, and verification evidence showing representative docx/xlsx/pptx round-trip capability.

## Reasoning levels by lane

- **Core/package/XML lane:** high reasoning (`executor` / `architect`)
- **Format parser lanes:** high reasoning (`executor`), medium-high for targeted follow-up fixes (`debugger`)
- **Rendering/product surfaces lane:** medium-high reasoning (`executor`, `designer`/`style-reviewer` as needed)
- **Verification/perf lane:** medium-high reasoning (`test-engineer`, `verifier`)
- **Final sign-off lane:** high reasoning (`architect`, `verifier`, optional `critic`)

## Launch hints

- Sequential Ralph path: `omx ralph` after the plan artifacts are committed, with execution sliced by the stage order above.
- Parallel team path once scaffolding is stable: `omx team` or `$team` with lanes mapped to Core, DOCX, XLSX, PPTX, and Verification.
- Team staffing hint: start with 4-5 workers only after Stage A foundations exist and shared package boundaries are fixed.
- Verification consolidation path: after team delivery, hand back to Ralph for end-to-end lint/typecheck/test/e2e/round-trip/architect-verifier review and final cleanup.


## Architect review (iteration 1)

- **Steelman antithesis:** a single shared-core architecture can become an abstraction trap, forcing unlike formats into weak common models and delaying delivery of format-specific fidelity.
- **Tradeoff tension:** unified IR and serializer consistency vs. format-specific semantic richness; browser-first performance budgets vs. high-fidelity layout ambitions.
- **Synthesis:** keep the shared core limited to truly shared mechanics (OPC, XML, diagnostics, shared primitives, transaction infrastructure) and preserve explicit format-specific escape hatches in docx/xlsx/pptx layers. Pair every shared abstraction with fixture-backed format validations.
- **Verdict:** APPROVE

## Critic review (iteration 1)

- Plan includes principles, drivers, viable options, deliberate-mode pre-mortem, expanded test plan, architecture, staged execution, staffing lanes, verification criteria, and a definition of done.
- Risks are concrete and mitigated through stage exit criteria, corpus strategy, and round-trip-first verification.
- Acceptance criteria are testable at both stage and feature-slice levels via the PRD/test spec and verification gates.
- **Verdict:** APPROVE
