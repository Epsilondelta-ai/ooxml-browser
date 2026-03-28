# API and Package Architecture

## Product architecture goals

- **browser-first architecture** for all public ingestion/render/edit flows
- composable **package / module boundaries**
- optional worker acceleration
- framework-agnostic rendering core with adapters for major frontend frameworks
- testable public API with stable compatibility policy

## Proposed monorepo packages

### Implemented workspace packages
- `@ooxml/core`: shared types, diagnostics, OPC/XML/IR/serialization primitives
- `@ooxml/docx`: WordprocessingML parsing, mutations, and serialization adapters
- `@ooxml/xlsx`: SpreadsheetML parsing, mutations, and serialization adapters
- `@ooxml/pptx`: PresentationML parsing, mutations, and serialization adapters
- `@ooxml/render`: DOM-oriented renderers for pages, grids, and slides
- `@ooxml/editor`: transaction-driven editor surfaces and per-format helpers
- `@ooxml/browser`: browser-first facade, worker-facing entry points, and ergonomic top-level APIs
- `@ooxml/devtools`: inspectors, summaries, and debugging helpers

### Application/workbench packages
- `@ooxml/example-basic`: lightweight browser example workspace
- `@ooxml/playground`: interactive playground for upload, inspect, render, edit, and save flows

### Internal module boundaries within packages
Even where functionality ships inside one workspace package, the code should still preserve clear internal boundaries for:
- OPC packaging
- XML/token handling
- normalized IR
- serialization patches
- render-model projection
- editor transactions
- worker protocols

## Public API design

The public API design must remain explicit about package boundaries, worker-safe calls, and stable testing/debug surfaces.

### Ingestion
```ts
const packageGraph = await openPackage(file);
const doc = await parseOfficeDocument(packageGraph, { mode: 'editable' });
```

### Format-specific convenience
```ts
const wordDoc = await parseDocx(file);
const workbook = await parseXlsx(file, { sheetData: 'lazy' });
const deck = await parsePptx(file);
```

### Rendering
```ts
const renderer = createRenderer(doc, { view: 'page' | 'grid' | 'slide' });
renderer.mount(element);
```

### Editing
```ts
const editor = createEditor(doc);
editor.transaction(tx => {
  tx.insertText(...);
  tx.applyStyle(...);
});
```

### Serialization
```ts
const blob = await serializeOfficeDocument(doc, { format: 'docx' });
```

## Plugin / extension system

Extension points:
- custom part handlers
- unsupported object renderers
- formula engines
- chart renderers
- collaboration backends
- import/export transforms
- diagnostics/reporting sinks

Plugin contract needs:
- capability declaration
- part/namespace/relationship filters
- serializer preservation rules
- optional worker-safe registration path

## Worker interfaces

Worker-eligible tasks:
- ZIP decompression and package parsing
- large part XML parsing
- formula dependency/index building
- page layout precomputation
- benchmark/test corpus evaluation

Need:
- transferable payloads (`ArrayBuffer`, structured-clone-safe IR shards)
- cancellation support
- progress events
- deterministic worker task protocol

## Testing hooks / debug utilities

Required **testing hooks** and debug utilities:
- package graph inspector
- relationship explorer
- style/theme resolver traces
- render tree dump
- serializer diff summary
- fidelity comparison helpers
- devtools bridge surfaces for inspection panels and diagnostics overlays

## Decisions

- **D-API-1:** ship small composable packages, but also expose ergonomic top-level format APIs.
- **D-API-2:** worker support is built into architecture, not bolted on later.
- **D-API-3:** framework adapters wrap a framework-agnostic renderer/editor core.

## Open risks

- too many tiny packages can increase integration friction; keep package boundaries aligned with runtime responsibilities, not theoretical purity
- public IR exposure should be carefully versioned so internals can evolve without breaking consumers unnecessarily
