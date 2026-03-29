# OOXML Frontend Library Documentation Index

This documentation set is the implementation baseline for a browser-first library that can parse, render, edit, and serialize OOXML-based Microsoft Office documents (`.docx`, `.xlsx`, `.pptx`).

## Goals

- Build a production-meaningful frontend library, not a demo-only parser.
- Preserve OOXML structure well enough for high-fidelity round-tripping.
- Share one coherent internal model across parser, renderer, editor, and serializer.
- Support browser-first execution with worker offloading and progressive/lazy strategies.
- Remain extensible for advanced Office features, interoperability quirks, and collaboration.

## Documentation map

### User guides
- [Using the library](./guides/using-the-library.md)

### Research / source baseline
- [Research source map](./research/sources.md)

### Format and packaging reference
- [OOXML + OPC packaging reference](./reference/opc-packaging.md)
- [WordprocessingML reference](./reference/wordprocessingml.md)
- [SpreadsheetML reference](./reference/spreadsheetml.md)
- [PresentationML reference](./reference/presentationml.md)

### Architecture and subsystem design
- [Shared subsystems](./design/shared-subsystems.md)
- [Parsing architecture](./design/parsing.md)
- [Rendering architecture](./design/rendering.md)
- [Editing architecture](./design/editing.md)
- [API and package architecture](./design/api-architecture.md)

### Quality, verification, and operations
- [Compatibility, performance, i18n, accessibility, and security](./quality/compatibility-performance-security.md)
- [Verification strategy](./quality/verification.md)
- [Operational plan: docs site, examples, playground, releases](./operations/project-operations.md)

## Baseline implementation decisions

1. **Browser-first core**: all public packages must run in browser environments without Node-only assumptions.
2. **Round-trip-friendly IR**: parsing must preserve enough source structure, unknown markup, relationships, and package-level metadata to serialize back to OOXML without avoidable loss.
3. **Shared package graph**: the first normalized layer is the OPC package graph; format-specific models are projected from it instead of bypassing packaging semantics.
4. **Layered model**: raw package/XML -> normalized OOXML IR -> view/layout model -> editor transactions -> serializer patches.
5. **Explicit fidelity tiers**: package preservation, semantic fidelity, visual fidelity, and behavioral fidelity are tracked separately in docs, tests, and benchmarks.
6. **Strict + Transitional awareness**: the library must read both strict and transitional forms and preserve markup compatibility details rather than aggressively normalizing them away.
7. **Security-first ingestion**: untrusted document parsing is the default threat model.

## Major unresolved implementation risks to carry into planning

- Exact layout fidelity for Word pagination, floating drawing anchoring, and Office-specific typography rules will require a dedicated layout engine plus corpus-based tuning.
- Spreadsheet formula calculation is separable from workbook parsing/rendering; formula parsing and dependency graphs are required even when full recalc engine parity is staged.
- Presentation animation fidelity will require a clear split between static slide rendering and timeline playback semantics.
- Markup compatibility / alternate content handling must preserve dormant branches for round-trip even when only one branch is rendered.
- Embedded binary objects and OLE-style payloads require safe unsupported-object policies plus preservation paths.

## Immediate planning implications

- Package the project as a monorepo with independent parser/render/editor/serializer/runtime packages.
- Establish golden corpus fixtures early because compatibility work depends on them.
- Define the IR before large-scale implementation so docx/xlsx/pptx can share services.
- Build verification and benchmark infrastructure alongside core code rather than after feature work.
