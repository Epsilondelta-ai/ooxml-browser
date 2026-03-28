# Test Specification: OOXML Gap Closure Phase

## Verification principles
1. Every new fidelity slice lands with fixture-backed evidence.
2. Round-trip and reopen behavior are primary evidence, not secondary checks.
3. Representative persisted fixtures are required before deeper format stages open.
4. Performance and worker/provenance behavior are foundational gates.
5. Parser reopen, Office/LibreOffice attestation, and docs/examples alignment are all part of completion.

## Test layers
### Unit
- OPC/package mutation helpers
- XML token tape, namespace handling, compatibility branches, source spans
- DOCX style/numbering/section/comment/revision/table/drawing/equation logic
- XLSX styles/numFmt/formula/reference/merge/frozen-pane/name/table/comment logic
- PPTX master/layout/theme/notes/comment/media/shape/timing logic
- editor transaction/invalidation/undo-redo behavior
- serializer patch planner and deterministic writer logic

### Integration
- no-op and minimal-edit parse -> serialize -> parser reopen for representative fixtures
- cross-part flows: docx comments/headers/footers/notes, xlsx workbook/sharedStrings/styles/tables/drawings, pptx masters/layouts/notes/media
- render/editor integration with scoped invalidation
- worker/browser integration for representative fixtures

### E2E
- DOCX open/edit/save/reopen representative documents
- XLSX open/edit/save/reopen representative workbooks
- PPTX open/edit/save/reopen representative decks
- playground automation for fixture loading, diagnostics, round-trip download, and representative edit flows

### Observability
- fixture-manifest-linked diagnostics snapshots
- mutation traces showing changed parts and patch-vs-regenerate decisions
- interop matrix generation
- benchmark result artifacts

### Performance
- microbenchmarks for parse/render/serialize primitives
- representative corpus budgets for open/render/edit/save
- memory and worker-transfer smoke for provenance-bearing structures

## Stage-specific evidence gates
### Stage 1
- persisted representative fixture per format under canonical fixture tree
- untouched-part and untouched-relationship preservation checks
- provenance memory smoke: initial target <= 25 MB retained provenance payload and <= 250 ms worker transfer/summary time on local benchmark harness

### Stage 2
- DOCX parser reopen in CI for declared representative fixtures
- Office/LibreOffice attestation manifests present for stage-owned seed fixtures
- render snapshot harness for docx visual/layout checks

### Stage 3
- XLSX parser reopen in CI for declared fixtures
- formula/reference rewrite tests and attestation manifests for stage-owned fixtures
- workbook benchmark thresholds tracked against Stage-0 sheet

### Stage 4
- PPTX parser reopen in CI for declared fixtures
- notes/comments/media/master-layout representative fixtures with attestation manifests
- slide render snapshot harness and notes/slide edit E2E

### Stage 5-6
- parser reopen + attestation both release-blocking for declared interop fixtures
- benchmark suite and compatibility matrix published in repo
- docs/quality and fixture layout docs synced

## Reopen evidence policy
- **Automated path:** parser reopen always runs in CI for no-op/minimal-edit fixtures.
- **Interop attestation path:** Office/LibreOffice reopen results are recorded in `fixtures/manifests/**` with automated evidence links or manual attestation records.
- **Stage policy:**
  - Stages 1-2: parser reopen mandatory; attestation required only for declared representative seed fixtures
  - Stages 3-4: parser reopen mandatory; attestation expands to all stage-owned representative fixtures
  - Stages 5-6: parser reopen + attestation both release-blocking for declared interop fixtures

## Canonical fixture tree
- `fixtures/shared/{opc,xml,security}`
- `fixtures/docx/{micro,representative,interop,perf}`
- `fixtures/xlsx/{micro,representative,interop,perf}`
- `fixtures/pptx/{micro,representative,interop,perf}`
- `fixtures/manifests/{docx,xlsx,pptx,shared}`
