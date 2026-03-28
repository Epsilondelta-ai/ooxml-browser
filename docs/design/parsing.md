# Parsing Architecture

## Objectives

- ingest untrusted OOXML packages safely in the browser
- parse docx/xlsx/pptx through a shared packaging layer
- preserve enough source fidelity for round-trip serialization
- support large documents through lazy/streaming strategies
- surface structured diagnostics instead of opaque failures

## Parsing pipeline

```text
ArrayBuffer/Blob/File
  -> zip reader + security gate
  -> PackageGraph
  -> part classifier + relationship resolver
  -> namespace-aware XML token stream
  -> raw part AST / token tape
  -> format-specific normalizers
  -> shared OOXML IR
  -> derived render/edit indexes
```

## XML parsing strategy

Use a dual representation:
1. **Token tape / source-preserving layer**
   - namespace declarations
   - prefixes
   - element/attribute order when relevant for lossless writing
   - text / CDATA / processing instruction / comments where present
2. **Normalized semantic AST**
   - typed OOXML nodes with canonical namespace identities
   - easy traversal for renderer/editor services

Reason:
- semantic AST alone is too lossy for safe round-trip in the presence of unknown or extension markup.

## Namespace handling

Requirements:
- canonical namespace registry with strict + transitional mappings
- preserve original prefixes in source layer
- resolve `mc:Ignorable`, `AlternateContent`, and extension namespaces
- keep unknown namespace content attached to parent nodes rather than dropping it

## Relationship resolution

Every parser service should resolve cross-part references via a `RelationshipResolver`:
- source URI + relationship ID -> target part/resource
- target URI -> typed part handle
- reverse index for update propagation

## Shared model normalization

Target common IR layers:
- `PackageGraph`
- `OfficeDocumentModel` (`kind: word|spreadsheet|presentation`)
- shared primitives: text run, paragraph-ish container, table, drawing node, style ref, annotation, story/sheet/slide containers
- edit-intent metadata (stable IDs, source spans, relationship references)

## Validation

Validation levels:
1. package validation (ZIP/OPC integrity)
2. schema-shape validation (required roots/required relationships/basic attribute types)
3. semantic validation (dangling refs, duplicate ids, cyclic master/layout issues, broken numbering/theme references)
4. compatibility validation (strict/transitional, markup compatibility branch selection, unsupported object warnings)

Diagnostic shape:
- code
- severity
- part URI
- XPath-ish location / token span
- human message
- recovery action if any

## Recovery / fault tolerance

- continue past local XML subtree errors when tokenization can resynchronize safely
- downgrade unsupported parts to opaque nodes
- keep broken references as unresolved handles so serializer can preserve them or editor can surface repair UI
- allow “open degraded document” mode with feature flags describing what was skipped

## Incremental parsing

Foundational requirements:
- lazy part parsing by relationship reachability
- memoized parsed parts keyed by package version + part URI
- reparse-on-write only for touched parts
- support partial index building (e.g. workbook sheet metadata before all rows are parsed)

## Large-document strategy

- ZIP central directory read first, defer full inflation where possible
- SAX/token-stream parsing for very large XML parts (worksheets, large document.xml, comments, shared strings)
- chunked row/paragraph indexes for viewport-driven access
- offload CPU-heavy parse steps to Web Workers

## Decisions

- **D-PARSE-1:** use source-preserving token tape + semantic AST; do not choose between fidelity and usability.
- **D-PARSE-2:** parsing services are lazy by default and materialize heavy structures on demand.
- **D-PARSE-3:** diagnostics are part of the public API.

## Open risks

- browser DOMParser alone is insufficient for robust fault tolerance and source-preserving round-trip; use a custom XML tokenizer/parser core or a library that exposes low-level events
- alternate content preprocessing must not destroy inactive branches needed for serialization
