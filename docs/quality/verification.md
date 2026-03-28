# Verification Strategy

This plan covers **unit / integration / e2e / golden / fidelity / round-trip / performance** verification layers.

## Test pyramid

### Unit tests
Cover:
- ZIP/OPC parsing
- content type and relationship resolution
- XML tokenizer/writer round-trip
- namespace + markup compatibility handling
- style/theme/number format resolution
- format-specific parsers and serializers for small fixtures
- transaction/undo/redo primitives

### Integration tests
Cover:
- parse -> render model -> serialize for representative documents
- docx story/comment/header/footer/numbering flows
- xlsx workbook/shared-strings/styles/formulas/sheet ops
- pptx slide master/layout/theme/notes/comments flows
- worker/off-main-thread parse and render index generation

### E2E / interaction tests
Cover browser behavior for:
- open document
- render view
- edit content/formatting
- undo/redo
- serialize and reopen
- compare results to expectations

### Golden / fidelity tests
Need snapshot families:
- package graph snapshots
- semantic IR snapshots
- rendered image snapshots
- serialized part diff snapshots

### Round-trip tests
Required modes:
- no-op round trip
- minimal edit round trip
- unknown markup preservation
- external relationship preservation
- strict/transitional preservation

### Performance tests
Bench:
- parse latency
- render latency
- edit latency
- serialization latency
- memory usage

## Corpus strategy

Need multiple corpora:
- hand-authored micro fixtures (single feature isolation)
- real-world mixed documents
- stress docs (large sheets, long documents, many slides/media)
- interop corpus saved by Office, LibreOffice, Google-exported OOXML where applicable
- adversarial security corpus (malformed XML, bad rels, zip bombs, broken content types)

## Fixture strategy

Canonical directory layout:
- `fixtures/shared/{opc,xml,security}`
- `fixtures/docx/{micro,representative,interop,perf}`
- `fixtures/xlsx/{micro,representative,interop,perf}`
- `fixtures/pptx/{micro,representative,interop,perf}`
- `fixtures/manifests/{docx,xlsx,pptx,shared}`

Each fixture should include:
- source file
- manifest.json (features, expectations, provenance, allowed warnings)
- optional rendered goldens
- optional serialized diff expectations

## Interoperability matrix

Track per fixture against:
- parser success
- renderer success
- edit success
- serializer success
- Office reopen result
- LibreOffice reopen result
- known tolerated diffs

## Regression detection

Required CI jobs:
- lint/typecheck/build
- unit + integration
- selective browser e2e
- golden image diffs
- round-trip diff checks
- perf budget checks on representative corpus

## Definition of “verified enough” per milestone

- feature code lands with fixture coverage
- at least one no-op round-trip test for new part types
- affected-file diagnostics are clean
- serializer output re-opens through parser
- docs/examples updated alongside feature behavior changes

## Decisions

- **D-VERIFY-1:** corpus-based verification is mandatory from early stages.
- **D-VERIFY-2:** round-trip diffs are primary evidence, not just unit assertions.
- **D-VERIFY-3:** verification must cover package, semantic, render, and edit layers.

## Open risks

- visual golden tests require stable font/render environments; CI image baselines need pinned browser/runtime configs
