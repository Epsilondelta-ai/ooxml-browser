# Using the OOXML frontend library

This guide is the practical entry point for consumers who want to open, inspect, render, edit, and save `.docx`, `.xlsx`, or `.pptx` files with the packages in this repo.

## Choose an entry point

### `@ooxml/browser`
Use this when you want the simplest browser-first flow:

- open an OOXML file from `Blob`, `ArrayBuffer`, or `Uint8Array`
- inspect package and document summaries
- render to HTML
- create an editor
- save the updated document back to a `Blob`

### Lower-level packages
Use the format and subsystem packages when you need more control:

- `@ooxml/core` — OPC/package graph access
- `@ooxml/docx` / `@ooxml/xlsx` / `@ooxml/pptx` — format parsers
- `@ooxml/render` — HTML rendering
- `@ooxml/editor` — semantic edit helpers
- `@ooxml/serializer` — OOXML serialization
- `@ooxml/devtools` — summaries and inspection helpers

## Package map

```text
@ooxml/browser     high-level browser session API
@ooxml/core        package graph + OPC/XML primitives
@ooxml/docx        WordprocessingML parsing/types
@ooxml/xlsx        SpreadsheetML parsing/types
@ooxml/pptx        PresentationML parsing/types
@ooxml/render      HTML render helpers
@ooxml/editor      semantic edit helpers
@ooxml/serializer  save back to OOXML bytes
@ooxml/devtools    package/document summaries
```

## Quick start with `@ooxml/browser`

```ts
import { createBrowserSession } from '@ooxml/browser';

const file = input.files?.[0];
if (!file) throw new Error('Select a .docx, .xlsx, or .pptx file first.');

const session = await createBrowserSession(file);

console.log(session.packageSummary);
console.log(session.documentSummary);

const html = session.renderToHtml();
preview.innerHTML = html;

const savedBlob = session.save();
```

### What a browser session gives you

- `packageGraph` — the relationship-aware OOXML package graph
- `document` — the parsed docx/xlsx/pptx model
- `packageSummary` — package-level counts and content types
- `documentSummary` — format-specific summary data
- `renderToHtml()` — render to an HTML string
- `mount(target)` — render directly into a DOM element
- `createEditor()` — create a mutable editor session
- `save()` — serialize the current document to a `Blob`

## Editing and saving

`createEditor()` gives you an editor instance. Use format-specific helpers from `@ooxml/editor` to make semantic changes.

### Example: edit a spreadsheet cell

```ts
import { createBrowserSession } from '@ooxml/browser';
import { setWorkbookCellValue } from '@ooxml/editor';

const session = await createBrowserSession(file);
const editor = session.createEditor();

if (editor.document.kind === 'xlsx') {
  setWorkbookCellValue(editor, 'Sheet1', 'A1', 'Updated value');
}

const updatedBlob = session.save();
```

### Example: edit Word paragraph text

```ts
import { createBrowserSession } from '@ooxml/browser';
import { replaceDocxParagraphText } from '@ooxml/editor';

const session = await createBrowserSession(file);
const editor = session.createEditor();

if (editor.document.kind === 'docx') {
  replaceDocxParagraphText(editor, 0, 0, 'Updated paragraph text');
}

const updatedBlob = session.save();
```

### Example: edit presentation text

```ts
import { createBrowserSession } from '@ooxml/browser';
import { setPresentationShapeText } from '@ooxml/editor';

const session = await createBrowserSession(file);
const editor = session.createEditor();

if (editor.document.kind === 'pptx') {
  setPresentationShapeText(editor, 0, 0, 'Updated slide text');
}

const updatedBlob = session.save();
```

## Lower-level pipeline

If you do not want the browser session wrapper, you can build the flow yourself:

```ts
import { openPackage } from '@ooxml/core';
import { parseXlsx } from '@ooxml/xlsx';
import { renderOfficeDocumentToHtml } from '@ooxml/render';
import { serializeOfficeDocument } from '@ooxml/serializer';

const graph = await openPackage(file);
const workbook = parseXlsx(graph);

const html = renderOfficeDocumentToHtml(workbook);
const savedBytes = serializeOfficeDocument(workbook);
```

Use this path when you need direct access to package parts, relationships, or format-specific parsing before rendering or editing.

## Inspecting a document

Use `@ooxml/devtools` when you need compact summaries for logs, debugging, or UI sidebars.

```ts
import { openPackage } from '@ooxml/core';
import { inspectOfficeDocument, summarizePackageGraph } from '@ooxml/devtools';
import { parseOfficeDocument } from '@ooxml/browser';

const graph = await openPackage(file);
const document = parseOfficeDocument(graph);

console.log(summarizePackageGraph(graph));
console.log(inspectOfficeDocument(document));
```

## Running the included demo surfaces

### Basic example

```bash
npm run dev --workspace @ooxml/example-basic
```

This example shows the file-input -> open -> summary -> visual preview -> round-trip download flow using `createBrowserSession`. The preview is tuned to feel more like a document page, spreadsheet grid, or slide canvas instead of just exposing the generated markup, while the raw HTML remains available as an optional debug panel.

### Playground

```bash
npm run dev --workspace @ooxml/playground
```

The playground exposes richer inspection, edit, preview, and save flows for all supported formats.

## Current capability baseline

Today the library is strongest at:

- opening OOXML archives into a relationship-aware package graph
- parsing representative `.docx`, `.xlsx`, and `.pptx` fixtures
- rendering parsed models to semantic HTML
- applying focused semantic edits
- serializing edited documents and reopening them through the parser

For the current edit surface, see:

- [Editing surface matrix](../reference/editing-surface-matrix.md)
- [Verification strategy](../quality/verification.md)
- [Interoperability matrix](../quality/interop-matrix.md)

## Notes and limitations

- The public surface is browser-first.
- Rendering is HTML-first, not full Microsoft Office layout parity.
- Support is representative and fixture-backed; consult the verification and interoperability docs for current boundaries.
- Some Office/LibreOffice attestation entries remain recorded as pending even when local parser/render/save verification is green.
