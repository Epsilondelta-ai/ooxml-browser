# OOXML Frontend Library Workspace

Browser-first workspace for parsing, rendering, editing, and serializing OOXML (`.docx`, `.xlsx`, `.pptx`).

## Package map
- `@ooxml/core` - shared OPC/XML/package graph primitives and package serialization helpers
- `@ooxml/docx` - WordprocessingML parsing
- `@ooxml/xlsx` - SpreadsheetML parsing
- `@ooxml/pptx` - PresentationML parsing
- `@ooxml/render` - HTML-first render projections
- `@ooxml/editor` - transaction-based editing helpers
- `@ooxml/serializer` - format-specific OOXML writers
- `@ooxml/browser` - browser session facade for open/render/edit/save flows
- `@ooxml/devtools` - package/document summary helpers

## Getting started

The easiest entry point is `@ooxml/browser`.

```ts
import { createBrowserSession } from '@ooxml/browser';

const session = await createBrowserSession(file);
const html = session.renderToHtml();
const savedBlob = session.save();
```

From a browser session you can:

- inspect `packageSummary` and `documentSummary`
- render with `renderToHtml()` or `mount()`
- create an editor with `createEditor()`
- save the current document back to a `Blob`

For full usage documentation, package-selection guidance, and lower-level examples, see:

- [`docs/guides/using-the-library.md`](./docs/guides/using-the-library.md)
- [`docs/reference/editing-surface-matrix.md`](./docs/reference/editing-surface-matrix.md)

## Current capability baseline
- open OOXML archives into a relationship-aware package graph
- parse representative docx/xlsx/pptx fixtures into typed models
- render typed models to semantic HTML strings / mounted DOM output
- perform small transactional edits with undo/redo
- serialize edited models back into OOXML archives and re-open them through the parser

## Workspace commands
- `npm test` - unit and round-trip verification
- `npm run typecheck` - TypeScript project references validation
- `npm run lint` - ESLint
- `npm run build` - package + example + playground builds
- `npm run bench` - micro benchmark harness for open/parse/render/serialize timings

## Examples
- `examples/basic` - file-input preview example with page-like DOCX, spreadsheet-like XLSX, slide-like PPTX rendering, summary, optional HTML debug view, and round-trip download
- `playground` - upload, inspect summaries, apply text/metadata edits, preview, and save

## Documentation
- `docs/` contains the research/design baseline
- `docs/reference/editing-surface-matrix.md` tracks the current public editing helpers and persistence expectations
- `.omx/plans/` contains the consensus plan, PRD, test spec, and review artifacts
