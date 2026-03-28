# Module Ownership Map: OOXML Gap Closure Phase

## Shared core
- `packages/core/src/opc/` — package graph, relationships, content types, mutation helpers
- `packages/core/src/xml/` — tokenizer, token tape, namespace handling, compatibility handling, source spans
- `packages/core/src/model/` — shared primitives only
- `packages/core/src/contracts/` — shared resolved contracts
- `packages/core/src/serialization/` — deterministic writer and low-level patch utilities

## Format ownership
- `packages/docx/src/{model,parser,resolve,edit}` — DOCX source graphs and semantic invariants
- `packages/xlsx/src/{model,parser,resolve,edit}` — XLSX source graphs and semantic invariants
- `packages/pptx/src/{model,parser,resolve,edit}` — PPTX source graphs and semantic invariants

## Downstream consumers
- `packages/render/src/{docx,xlsx,pptx}` — view model + render projections
- `packages/editor/src/` — transaction core and format adapters
- `packages/serializer/src/{docx,xlsx,pptx}` — format patch planners and writers
- `examples/**`, `playground/**`, `benchmarks/**`, `docs/**` — product and quality surfaces
