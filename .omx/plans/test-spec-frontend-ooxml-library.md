# Test Specification: Frontend OOXML Library

## Verification principles

1. Every format capability lands with fixture-backed tests.
2. Round-trip and interoperability evidence matters as much as unit coverage.
3. Browser behavior is verified through automated UI tests for representative flows.
4. Performance budgets are tracked continuously for representative corpora.
5. Diagnostics and degraded-mode behavior are testable outputs.

## Test layers

### 1. Unit tests

#### OPC / XML core
- ZIP entry validation, safety budget enforcement, path normalization
- content type and relationship resolution
- XML tokenization, namespace resolution, writer round-trips
- markup compatibility branch preservation

#### Shared subsystem core
- theme/color transforms
- style inheritance
- numbering resolution
- asset registry operations
- annotation indexing
- serializer patch planner
- transaction and undo/redo primitives

#### Format units
- DOCX paragraph/run/table/section/comment/revision parsers
- XLSX workbook/sheet/cell/style/shared-string/formula/reference parsers
- PPTX slide/master/layout/text/shape/comment/notes parsers

### 2. Integration tests

#### Parse -> IR -> Serialize
- no-op round trips for docx/xlsx/pptx fixtures
- minimal edits on representative documents
- unknown markup and dormant `mc:AlternateContent` branch preservation
- strict/transitional document retention

#### Render model integration
- DOCX layout model creation for sections, lists, tables, comments, drawings
- XLSX grid view model with merges/frozen panes/styles/formulas
- PPTX slide scene graph with master/layout/theme inheritance

#### Editor integration
- transaction application invalidates only affected regions
- undo/redo restores prior document state and serializer output
- copy/paste maintains semantic structure where format supports it

### 3. Browser E2E tests

#### DOCX
- open fixture in page view
- edit text and paragraph style
- insert table/comment
- save and reopen

#### XLSX
- open sheet
- edit cells/formulas/styles
- merge/unmerge and reorder sheets
- save and reopen

#### PPTX
- open slide deck
- edit text, move shape, duplicate slide
- edit notes
- save and reopen

### 4. Golden / fidelity tests

- package graph snapshots
- IR snapshots
- serialized part snapshots (normalized)
- visual snapshots for selected docx pages, worksheet regions, and slides
- diff tolerances documented per fixture

### 5. Performance / observability tests

- parse/open latency by corpus fixture
- first view model paint latency
- edit latency for representative operations
- memory high-water marks
- serialization time and output size
- worker task timing and cancellation behavior

## Corpus strategy

### Corpus families
- `fixtures/opc`: packaging edge cases
- `fixtures/docx`: text, styles, numbering, sections, headers, comments, revisions, drawings, equations
- `fixtures/xlsx`: shared strings, styles, formulas, merges, tables, comments, drawings, charts
- `fixtures/pptx`: masters, layouts, notes, comments, media, charts, timing metadata
- `fixtures/interop`: Office + LibreOffice + alternate producer samples
- `fixtures/security`: malformed zips/XML/relationships/external targets
- `fixtures/perf`: large docs/sheets/decks

### Fixture manifest fields
- id
- source application/version
- feature tags
- expected diagnostics
- supported operations
- reopen expectations
- visual baseline references

## Interoperability matrix

Track each representative fixture against:
- parse status
- render status
- edit status
- serialize status
- reopen in Office
- reopen in LibreOffice
- tolerated differences

## Definition of done for a feature slice

- unit coverage for new parser/editor/serializer logic
- at least one integration round-trip test
- relevant browser E2E coverage or justified temporary gap with follow-up ticketed in plan
- docs/example updated
- perf impact checked against representative fixture

## Tooling expectations

- Vitest for unit/integration
- Playwright for browser E2E and visual capture
- benchmark harness for perf corpus
- lint/typecheck/build as mandatory gates
