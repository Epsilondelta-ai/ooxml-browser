# Consensus Plan: OOXML Gap Closure Phase

## RALPLAN-DR summary
### Principles
1. Preservation before regeneration.
2. Shared resolved contracts, format-owned source graphs.
3. Round-trip evidence over feature claims.
4. Office-meaningful leverage by stage.
5. Product surfaces evolve with the core.

### Decision drivers
1. Current typed projections and serializer rewrites are too lossy for deeper fidelity.
2. The repo is still structurally shallow and fixture-light.
3. The next phase must extend the current architecture rather than restart it.

### Viable options
- **A. Preservation-core first, then format expansion** — chosen
- **B. Format-first expansion with later preservation retrofit** — rejected for high rework risk
- **C. Corpus-first black-box interop push** — rejected for discovery without enough implementation rails

### Decision
Choose **Option A** because the current serializer-heavy, shallow-model baseline makes safe fidelity expansion too risky without stronger preservation and mutation infrastructure first.

### Pre-mortem
1. Feature breadth rises while round-trip fidelity regresses.
2. Shared abstractions flatten format semantics.
3. Corpus/verification overhead outruns delivery.

### Expanded test plan
- unit, integration, E2E, observability, and performance gates are mandatory
- parser reopen in CI is always-on for declared fixtures
- Office/LibreOffice attestation expands by stage

## Architect review
- **Steelman antithesis:** over-investing in preservation/decomposition before proving semantic shape could stall visible feature gains.
- **Tradeoff tension:** shared resolved contracts vs duplicated format resolution logic; preservation depth vs browser-first performance.
- **Synthesis:** the revised plan now makes the boundary explicit, strengthens Stage 1 proof, and operationalizes reopen evidence.
- **Verdict:** APPROVE

## Critic review
- Option choice matches stage order and repo risk profile.
- Deliberate-mode rigor is adequate.
- Canonical fixture tree, thresholds, and reopen-evidence policy are now explicit.
- **Verdict:** APPROVE

## Stage order
### Stage 0 — planning refresh + decomposition map
Deliverables:
- updated PRD/test spec/consensus plan
- module ownership map appendix
- Stage-0 budget sheet
- Stage-0 reopen-evidence appendix

### Stage 1 — preservation core + module-boundary foundation
Rules:
- no unrelated XML rewrites from format packages
- provenance or part ownership on edited semantic nodes
- shared modules own resolved contracts, not source semantics
- format semantics remain format-owned
- patch when scoped; regenerate only when unavoidable

Boundary ADR:
- `packages/{docx,xlsx,pptx}` own source graphs and invariants
- `packages/core/src/contracts/` owns shared resolved contracts for styles/themes/tables/drawings/annotations/assets

Verification gate:
- existing tests green
- source-preserving/mutation helper unit tests
- representative fixture per format persisted
- untouched-part and untouched-relationship preservation checks
- provenance memory smoke at Stage-0 thresholds

### Stage 2 — DOCX fidelity expansion
Coverage:
- style inheritance, numbering, sections, headers/footers, comments/revisions, tables, drawings/equations hooks
Gate:
- parser reopen in CI for declared fixtures
- Office/LibreOffice attestation manifests for stage-owned seed fixtures
- docx render snapshot harness

### Stage 3 — XLSX fidelity expansion
Coverage:
- styles, numFmt, formulas/reference rewrites, merges, frozen panes, names, tables, comments, chart references
Gate:
- parser reopen in CI for declared fixtures
- rewrite tests + stage-owned attestation manifests
- benchmark thresholds tracked

### Stage 4 — PPTX fidelity expansion
Coverage:
- master/layout/theme inheritance, notes/comments/media breadth, richer shapes, timing preservation
Gate:
- parser reopen in CI for declared fixtures
- stage-owned attestation manifests
- slide render snapshot harness and notes/slide edit E2E

### Stage 5 — shared serializer depth + interop/product-quality surfaces
Coverage:
- patch planner rules
- canonical corpus implementation
- playground/example automation
- benchmark and compatibility report generation
- docs sync for fixture layout and verification docs
Gate:
- interop matrix generated
- repo-wide verification commands green
- architect/verifier evidence sufficient

### Stage 6 — MS Office-meaningful hardening exit
Definition of done:
- preservation depth demonstrated on representative workflows
- DOCX/XLSX/PPTX stage-owned slices have representative support
- product surfaces expose supported capabilities and diagnostics
- corpus-backed verification and interop evidence pass
- architect and verifier approve

## Canonical fixture/manifests tree
- `fixtures/shared/{opc,xml,security}`
- `fixtures/docx/{micro,representative,interop,perf}`
- `fixtures/xlsx/{micro,representative,interop,perf}`
- `fixtures/pptx/{micro,representative,interop,perf}`
- `fixtures/manifests/{docx,xlsx,pptx,shared}`

## Staffing and launch guidance
### Available-agent-types roster
- `planner`, `architect`, `executor`, `debugger`, `test-engineer`, `verifier`, `critic`, `code-reviewer`, `researcher`, `dependency-expert`, `writer`

### Ralph lanes
- **Lane A:** preservation core (`packages/core`, fixture schema/tooling) — reasoning high
- **Lane B:** DOCX fidelity — reasoning high
- **Lane C:** XLSX fidelity — reasoning high
- **Lane D:** PPTX fidelity — reasoning high
- **Lane E:** product/quality surfaces (`examples`, `playground`, `benchmarks`, docs) — reasoning medium-high
- **Lane F:** sign-off and consolidation — reasoning high

### Team hints
- If using team bursts later, split by Lane A-F after Stage 1 boundaries are fixed.
- Final verification still returns to Ralph for end-to-end consolidation.
