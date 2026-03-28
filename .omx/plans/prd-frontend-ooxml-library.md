# PRD: Frontend OOXML Library

## Document metadata

- **Project:** OOXML frontend library
- **Slug:** frontend-ooxml-library
- **Planning mode:** ralplan consensus, deliberate depth
- **Grounding snapshot:** `.omx/context/frontend-ooxml-library-20260328T034425Z.md`
- **Baseline docs:** `docs/index.md` and linked reference/design docs

## Product vision

Deliver a browser-first library that can open, inspect, render, edit, and serialize OOXML-based Office documents (`.docx`, `.xlsx`, `.pptx`) with production-meaningful fidelity, strong round-trip preservation, safe untrusted-document handling, and extensible package boundaries for advanced subsystems.

## Primary users

1. **Application developers** embedding Office-like viewing/editing in web apps.
2. **Workflow/product teams** needing browser-side document import/export and light-to-deep editing.
3. **Platform teams** needing package inspection, document transformation, verification, or interoperability tooling.
4. **Power users / internal tooling authors** building document playgrounds, automation, or migration tools.

## Overall product goals

1. Parse OOXML packages into a shared, round-trip-friendly internal representation.
2. Render `.docx`, `.xlsx`, and `.pptx` in browser-first viewports with meaningful fidelity.
3. Support semantic editing with undo/redo and deterministic serialization back to OOXML.
4. Preserve unsupported or unknown-but-safe markup, relationships, and parts wherever possible.
5. Ship examples, devtools, verification corpus tooling, and benchmark surfaces alongside the core packages.

## Non-goals

The project is not limited to a viewer-only surface, server-only runtime, or one-format prototype. Optional future acceleration work may deepen fidelity, but the delivered baseline must already be production-meaningful across parse/render/edit/save flows.

## Quality targets

### Fidelity targets
- **Package fidelity:** preserve untouched package graph structures and unknown safe parts.
- **Semantic fidelity:** preserve document content, structural anchors, styles, and references.
- **Visual fidelity:** render representative fixtures close to Office baselines for supported feature sets.
- **Interaction fidelity:** expose Office-like editing affordances for core operations.
- **Round-trip fidelity:** no-op and small-edit round trips retain unchanged constructs as much as feasible.

### Compatibility targets
- Open representative Office-produced files across all three formats.
- Serialize outputs that reopen in Microsoft Office and LibreOffice for supported feature sets.
- Preserve unsupported features with explicit degraded diagnostics rather than destructive drops.

### Safety targets
- Default threat model is untrusted documents.
- No automatic external resource execution/loading.
- No executable macro/ActiveX/OLE behavior; preserve as opaque attachments.

## Functional requirements

### OPC / package layer
- Read ZIP central directory and enforce safety budgets.
- Parse content types and relationship parts.
- Resolve root office document and traverse package graph.
- Preserve unknown parts and orphan parts with diagnostics.

### Shared XML + IR layer
- Source-preserving XML token tape + semantic AST.
- Namespace registry with strict/transitional handling.
- Markup compatibility preservation and active-branch projection.
- Shared OOXML IR for text/table/drawing/style/theme/annotation/asset primitives.

### Format-specific requirements

#### DOCX
- Parse stories, styles, numbering, sections, headers/footers, comments, footnotes/endnotes, equations, tracked changes, and drawings.
- Render paginated or continuous document view.
- Support text/style/list/table/comment/revision-aware editing.
- Serialize while preserving story boundaries and revisions.

#### XLSX
- Parse workbook, worksheets, shared strings, styles, names, formulas, merges, validations, tables, comments, drawings, and charts.
- Render virtualized grid with frozen panes and overlays.
- Support cell, row/column, sheet, style, formula, and table edits.
- Serialize shared strings/styles/relationships incrementally.

#### PPTX
- Parse presentations, slides, masters, layouts, themes, notes, comments, text shapes, graphic frames, media, and timing metadata.
- Render slide, notes, and thumbnail/sorter projections.
- Support text, shape, ordering, slide, notes, and asset edits.
- Serialize while preserving master/layout relationships and timing metadata.

## Shared subsystem requirements

- Theme, color, and font resolution.
- Style resolution across document families.
- Table abstractions.
- Drawing/scene graph and asset registry.
- Metadata, hyperlink, embedded object preservation.
- Annotation/comment/note systems.
- Equation preservation + editable projection.

## Public product surfaces

- Package open/inspect APIs.
- Parse/render/edit/serialize APIs.
- Worker APIs.
- React adapter.
- Playground with package inspector and render/edit panes.
- Bench harness and fixture tooling.
- Docs site and examples.

## Constraints

- Browser-first architecture.
- Round-trip-friendly model shared across parser/renderer/editor/serializer.
- Add new dependencies only when necessary and documented.
- Commit after every meaningful implementation stage with Lore protocol.
- Verification must include tests, build, diagnostics, and format flow evidence.

## Release criteria

A release-ready 1.0 baseline requires:
- working monorepo build/test/docs/examples/playground/bench infra
- core parse/render/edit/save support for representative docx/xlsx/pptx fixtures
- corpus-backed round-trip tests and diagnostics
- public package APIs with examples and compatibility notes
