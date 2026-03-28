# PRD: Frontend OOXML Library

## Document status
- Status: Approved for execution
- Planning mode: ralplan consensus (`--deliberate` depth)
- Grounding snapshot: `.omx/context/frontend-ooxml-library-20260328T034425Z.md`
- Documentation baseline: `docs/index.md`

## Product vision
Build a browser-first library that can open, parse, render, edit, and serialize OOXML-based Microsoft Office documents (`.docx`, `.xlsx`, `.pptx`) with production-meaningful fidelity and round-trip preservation. The library must expose stable frontend APIs, worker-friendly runtime boundaries, examples, playground tooling, and verification infrastructure that can keep compatibility quality improving over time.

## Primary users
- frontend product teams building document viewers/editors
- internal tools teams that need browser-native Office document inspection or transformation
- collaboration products requiring structured OOXML editing in web environments
- QA/performance teams needing package-graph, fidelity, and benchmark tooling

## Goals
1. Parse `.docx`, `.xlsx`, and `.pptx` through a shared OPC/OOXML stack.
2. Render Word pages, spreadsheet grids, and presentation slides in browser environments.
3. Support semantic editing flows for text, structure, formatting, sheets/slides, and assets.
4. Serialize edited documents back to OOXML with high round-trip preservation.
5. Provide verification, fixtures, examples, benchmarks, and devtools as first-class product surfaces.

## Non-goals
- No scope reduction to an MVP.
- No server-only assumptions in public APIs.
- No destructive flattening of OOXML structure merely to simplify rendering.

## Quality targets
- Browser-first runtime with optional workers.
- Round-trip-safe preservation of unknown parts, relationships, markup compatibility branches, and extensibility content.
- Fidelity tracked across package, semantic, visual, and behavioral dimensions.
- Untrusted-document-safe parse posture.
- Stable monorepo package boundaries with framework-agnostic core.

## User stories

### US-001 Package ingestion and inspection
As an application developer, I want to open OOXML blobs/files and inspect package parts, relationships, and content types so that I can build workflows on a trustworthy package graph.

Acceptance criteria:
- `openPackage` accepts `Blob | File | ArrayBuffer | Uint8Array`.
- Package graph contains parts, relationships, and content-type metadata.
- Unsupported parts are preserved as opaque nodes.
- Security limits detect malformed ZIP/path traversal conditions.

### US-002 Word processing support
As a product team, I want to parse, render, edit, and save `.docx` files so that users can work with rich paginated documents in-browser.

Acceptance criteria:
- Stories, styles, numbering, sections, headers/footers, comments, tracked changes, tables, drawings, and equations are parsed into IR.
- Word page/continuous views render main content and common layout structures.
- Editor supports text/style/table/image/comment/revision-aware document mutations.
- Serializer writes back valid OOXML and preserves unknown content where untouched.

### US-003 Spreadsheet support
As a product team, I want to parse, render, edit, and save `.xlsx` workbooks so that users can inspect and edit sheets in a web grid UI.

Acceptance criteria:
- Workbook/sheet/sharedStrings/styles/theme/formula metadata are parsed.
- Virtualized sheet renderer supports cells, merged regions, formatting, frozen panes, and drawings/charts overlays.
- Editor supports cell/range/sheet mutations, style changes, formulas, and sheet management.
- Serializer updates shared strings, worksheets, workbook relationships, and styles deterministically.

### US-004 Presentation support
As a product team, I want to parse, render, edit, and save `.pptx` decks so that users can work with slides, notes, and assets in-browser.

Acceptance criteria:
- Presentation/slide/master/layout/theme/notes/comments/timing metadata are parsed.
- Slide renderer preserves master/layout inheritance and basic layered scene composition.
- Editor supports text and shape mutations, slide operations, notes/comments, and asset replacement.
- Serializer retains unsupported timing/media metadata when unchanged.

### US-005 Tooling and verification
As a maintainer, I want fixtures, examples, playground, benchmarks, and diagnostics so that the library remains verifiable and shippable.

Acceptance criteria:
- Fixture corpus exists for package, docx, xlsx, pptx, interop, security, and perf scenarios.
- Examples and a playground demonstrate open/render/edit/save flows.
- Benchmark harness captures parse/render/edit/save timing.
- Tests and diagnostics gate releases.

## Product requirements

### Functional requirements
- Shared OPC package parser and writer
- Source-preserving XML layer
- Normalized OOXML IR
- Format-specific parse/render/edit/serialize packages
- Worker offloading protocol
- Devtools inspection surfaces
- Browser examples/playground

### Compatibility requirements
- Read strict and transitional OOXML.
- Preserve unsupported parts and unknown markup.
- Reopen generated files via this library parser.
- Track interoperability across Office and LibreOffice corpora.

### Performance requirements
- Lazy part parsing
- Worker-ready heavy operations
- Virtualized spreadsheet and long-document rendering
- Configurable memory/security limits

### Security requirements
- Zip bomb and malformed XML protection
- No automatic external resource loading
- Safe policy for macros, OLE, ActiveX, and unsupported binary payloads

## Product architecture summary
- monorepo packages under `packages/`
- fixtures under `fixtures/`
- tests under `tests/`
- examples under `examples/`
- interactive playground under `playground/`
- benchmark harness under `benchmarks/`

## Release criteria
- Build, tests, typecheck, lint, diagnostics pass
- Corpus-backed parse/render/edit/serialize flows validated
- Examples/playground verified
- Documentation updated to match delivered APIs

## Definition of done
- Shared ingestion, IR, render, edit, and serialize flows exist for docx/xlsx/pptx
- Example applications and playground demonstrate complete flows
- Benchmarks and fixture matrix exist and run
- Verification evidence is fresh and recorded before final completion
