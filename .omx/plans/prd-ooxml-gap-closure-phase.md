# PRD: OOXML Gap Closure Phase

## Metadata
- **Project:** OOXML gap closure phase
- **Slug:** ooxml-gap-closure-phase
- **Mode:** ralplan consensus, deliberate
- **Grounding snapshot:** `.omx/context/ooxml-gap-closure-20260328T050715Z.md`
- **Baseline inputs:** `docs/**`, prior `.omx/plans/*frontend-ooxml-library*.md`, current workspace implementation under `packages/**`

## Product objective
Advance the repository from a representative OOXML parse/render/edit/save baseline toward a substantially more complete Microsoft Office-meaningful frontend library across `.docx`, `.xlsx`, and `.pptx`, while preserving and extending the current browser-first architecture instead of replacing it.

## Why this phase exists
The current baseline proves end-to-end viability, but still relies on shallow typed projections, sparse durable fixtures, and serializer flows that regenerate major XML parts from simplified models. That is sufficient for representative workflows, but insufficient for higher-fidelity Microsoft Office compatibility, richer format semantics, and safer round-trip preservation.

## Goals
1. Deepen source-preserving and package-preserving infrastructure so richer fidelity work can land safely.
2. Expand DOCX fidelity across styles, numbering, sections, comments/revisions, headers/footers, tables, drawings, and equations.
3. Expand XLSX fidelity across styles, number formats, formulas/reference rewriting, merges, frozen panes, defined names, tables, comments, and charts.
4. Expand PPTX fidelity across master/layout/theme inheritance, notes/comments/media breadth, richer shape handling, and timing/animation preservation.
5. Establish a durable fixture/interoperability/benchmark system that validates parse, render, edit, serialize, and reopen behavior.
6. Keep examples, playground, documentation, and quality surfaces aligned with implemented capabilities.

## Non-goals
- Restarting the architecture from scratch.
- Narrowing scope to a single format or MVP.
- Claiming full Office parity without representative evidence.

## Quality targets
### Preservation
- untouched parts, relationships, unknown safe markup, and dormant compatibility branches survive representative no-op/minimal-edit workflows
- patch-scoped mutations are preferred over full regeneration wherever feasible

### Fidelity
- DOCX: representative support for styles, numbering, sections, headers/footers, comments/revisions, tables, and preserved drawing/equation hooks
- XLSX: representative support for styles, number formats, formulas, reference rewrites, merges, frozen panes, names, tables, comments, and chart references
- PPTX: representative support for master/layout/theme inheritance, notes/comments/media, richer shapes, and timing preservation

### Product surfaces
- example and playground flows expose supported fidelity slices and diagnostics
- benchmarks and compatibility matrix are durable repo artifacts, not temporary outputs

### Verification
- corpus-backed unit/integration/E2E/perf/interop gates
- parser reopen in CI for declared fixtures
- staged Office/LibreOffice attestation via fixture manifests

## Architecture rules
1. No format package may directly rewrite unrelated XML parts; all part mutations route through shared package mutation helpers.
2. Every edited semantic node retains source provenance or part ownership metadata.
3. Shared modules own resolved contracts, not source semantics.
4. Format semantics remain format-owned.
5. Patch when scoped, regenerate only when unavoidable, and record why.

## Stage 1 ADR
### Decision
Adopt **format-owned source graphs, shared resolved contracts**.

### Meaning
- `packages/{docx,xlsx,pptx}` own raw-to-semantic source graphs, provenance spans, and format-only invariants.
- shared modules own normalized resolved contracts for styles/themes/tables/drawings/annotations/assets used by render/editor/devtools.

### Why chosen
This preserves format richness while still preventing duplicate resolution logic across product surfaces.

### Constraint
No shared module becomes the source-of-truth for DOCX numbering, XLSX formula graphs, or PPTX master/layout semantics.

## Execution stages
### Stage 0 — Planning refresh + decomposition map
- update PRD, test spec, consensus plan
- produce module ownership map
- define Stage-0 budget sheet and reopen-evidence appendix

### Stage 1 — Preservation core + module-boundary foundation
- decompose monolithic package entrypoints into module directories
- add provenance-bearing semantic node types and source-preserving token/part references
- introduce `packages/core/src/contracts/` for shared resolved contracts
- add package mutation helpers and manifest schema/corpus tooling skeleton

### Stage 2 — DOCX fidelity expansion
- styles, numbering, sections, headers/footers, comments/revisions, tables, drawings/equations preservation hooks

### Stage 3 — XLSX fidelity expansion
- style/numFmt model, formula/reference rewrite helpers, merges/frozen panes/names/tables/comments/charts

### Stage 4 — PPTX fidelity expansion
- master/layout/theme inheritance, notes/comments/media breadth, richer shapes, timing preservation

### Stage 5 — Shared serializer depth + interop/product-quality surfaces
- patch planner rules, canonical fixture corpus, playground/example automation, benchmark + compatibility reports, docs sync

### Stage 6 — MS Office-meaningful hardening exit
- close remaining declared stage gaps, collect release-blocking evidence, architect + verifier sign-off

## Canonical fixture/manifests layout
- `fixtures/shared/{opc,xml,security}`
- `fixtures/docx/{micro,representative,interop,perf}`
- `fixtures/xlsx/{micro,representative,interop,perf}`
- `fixtures/pptx/{micro,representative,interop,perf}`
- `fixtures/manifests/{docx,xlsx,pptx,shared}`

## Definition of done
This phase is done when:
1. preservation depth is demonstrated on representative no-op/minimal-edit workflows,
2. DOCX/XLSX/PPTX stage-owned fidelity slices each have representative parse-render-edit-serialize support,
3. product surfaces expose the implemented capabilities and diagnostics,
4. corpus-backed verification and interop evidence pass,
5. architect and verifier both approve.
