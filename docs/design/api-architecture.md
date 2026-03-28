# API and Package Architecture

## Product architecture goals

- **browser-first architecture** for all public ingestion/render/edit flows
- composable **package / module boundaries**
- optional worker acceleration
- framework-agnostic rendering core with adapters for major frontend frameworks
- testable public API with stable compatibility policy

## Proposed monorepo packages

### Core packages
- `@ooxml/core`: shared types, errors, diagnostics, namespace registry, ids, utilities
- `@ooxml/opc`: ZIP/OPC reader, package graph, relationships, content types
- `@ooxml/xml`: XML tokenizer, source-preserving AST, writers
- `@ooxml/ir`: normalized OOXML IR and common primitives
- `@ooxml/serializer`: shared serialization utilities and patch pipeline

### Format packages
- `@ooxml/docx`
- `@ooxml/xlsx`
- `@ooxml/pptx`

Each provides:
- parse APIs
- render model projection APIs
- edit adapters
- serializer hooks

### Rendering packages
- `@ooxml/render-core`
- `@ooxml/render-docx`
- `@ooxml/render-xlsx`
- `@ooxml/render-pptx`

### Editor/runtime packages
- `@ooxml/editor-core`
- `@ooxml/editor-docx`
- `@ooxml/editor-xlsx`
- `@ooxml/editor-pptx`
- `@ooxml/worker`

### Integration packages
- `@ooxml/react`
- `@ooxml/vue` (optional adapter later)
- `@ooxml/devtools`
- `@ooxml/bench`
- `@ooxml/examples`

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
