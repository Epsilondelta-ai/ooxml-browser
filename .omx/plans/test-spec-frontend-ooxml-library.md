# Test Specification: Frontend OOXML Library

## Strategy summary
The test plan covers unit, integration, browser e2e, golden, round-trip, interoperability, security, and performance verification. Every implemented subsystem must map to at least one fixture-backed assertion path.

## Test layers

### 1. Unit tests
Scope:
- `@ooxml/core`: ids, diagnostics, errors, range helpers
- `@ooxml/opc`: content types, relationships, path resolution, security guards
- `@ooxml/xml`: tokenizer, namespace registry, source-preserving writeback, markup compatibility retention
- `@ooxml/ir`: node creation, normalization helpers, id mapping
- per-format parse/serialize helpers
- transaction primitives, undo/redo reducers, worker message codecs

Pass criteria:
- deterministic output for representative fixtures
- 0 failed unit suites

### 2. Integration tests
Scope:
- open package -> parse document -> serialize -> reopen
- docx story/style/numbering/section/comment/header-footer/tracked-change flows
- xlsx workbook/worksheet/sharedString/style/formula/sheet-op flows
- pptx slide/master/layout/notes/comment/shape flows
- shared subsystems: theme, styles, assets, charts, annotations

Pass criteria:
- parse/serialize loops reopen without fatal diagnostics
- invariant assertions on content, relationships, and preserved unknown nodes pass

### 3. Browser e2e tests
Scope:
- paginated document viewer/editor example
- spreadsheet grid example
- slide viewer/editor example
- playground upload/inspect/edit/save workflow

Pass criteria:
- user-visible open/edit/save flows complete in browser automation
- core keyboard navigation and accessibility smoke tests pass

### 4. Golden and visual fidelity tests
Scope:
- document page screenshots
- spreadsheet grid screenshots
- slide screenshots
- package graph snapshots
- semantic IR snapshots
- serializer part snapshots and diff reports

Pass criteria:
- diffs within approved thresholds
- intentional diffs reviewed and updated with fixture manifests

### 5. Round-trip preservation tests
Scope:
- no-op save
- targeted edit save
- unknown markup preservation
- unknown relationship preservation
- alternate content preservation
- strict/transitional preservation

Pass criteria:
- unchanged constructs remain byte-equivalent where feasible or semantically equivalent with stable diffs
- reopened documents preserve intended structure and content

### 6. Interoperability tests
Scope:
- Office-authored corpus
- LibreOffice-authored corpus
- mixed-export corpus where applicable

Pass criteria:
- open success rate and allowed warning budget tracked per format
- serializer outputs remain reopenable by this library and fixture validators

### 7. Security tests
Scope:
- path traversal ZIP entries
- oversized/zip bomb packages
- malformed `.rels`
- malformed XML / namespace abuse / depth abuse
- macros/unsupported embedded objects preserved but never executed

Pass criteria:
- unsafe files fail closed with structured diagnostics
- safe degraded mode behavior remains deterministic

### 8. Performance tests
Scope:
- open package time
- parse time per format and per part size bucket
- render first meaningful paint
- edit latency for representative operations
- serialize latency
- memory ceiling under stress fixtures

Pass criteria:
- baseline budgets recorded and enforced in benchmark runs
- regressions reported with budget deltas

## Fixture layout
- `fixtures/opc/`
- `fixtures/docx/`
- `fixtures/xlsx/`
- `fixtures/pptx/`
- `fixtures/interop/`
- `fixtures/security/`
- `fixtures/perf/`

Each fixture must contain:
- source document or generated package assets
- `manifest.json` with expectations, supported warnings, provenance, and tags
- optional golden images or serialized diff files

## Observability requirements
- parser diagnostics snapshots
- serializer diff summaries
- benchmark JSON output
- optional render trace artifacts
- worker timing instrumentation

## Exit gates per implementation stage
1. Scaffolding/core: lint + typecheck + unit tests + diagnostics clean
2. Parser stage: parser units + integration open/reopen + fixture manifests
3. Renderer stage: browser tests + golden baselines + basic perf metrics
4. Editor stage: mutation/undo-redo/serialize/reopen tests
5. Finalization: full matrix including examples/playground + perf + docs parity check
