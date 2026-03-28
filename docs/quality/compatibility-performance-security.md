# Compatibility, Performance, I18N, Accessibility, and Security

## Compatibility targets

### Office compatibility

This section defines explicit **Microsoft Office compatibility** targets.

Primary goal:
- documents produced by the library should open cleanly in current Microsoft Office desktop and web apps for supported feature sets
- documents imported from Office should preserve package/semantic structure even when browser rendering is not yet pixel-identical

### Alternative-suite compatibility

Target interoperability with:
- LibreOffice
- Apple/iWork import paths where OOXML is accepted
- browser-side consumers that expect standards-conforming OOXML

### Fidelity tiers

1. **Package fidelity**: parts, relationships, content types, unknown payloads preserved
2. **Semantic fidelity**: content and document structure preserved
3. **Visual fidelity**: layout/rendering closely matches Office
4. **Behavioral fidelity**: editing/selection/formatting behavior aligns with user expectations from Office-like tools
5. **Round-trip fidelity**: parse -> edit/no-op -> serialize preserves unchanged constructs as much as possible

This is the project's explicit **round-trip preservation** target.

### Import/export tolerance policy

This is the library's **importer/exporter tolerance policy**.

- accept well-formed strict/transitional OOXML plus common producer quirks
- preserve unsupported but safe content
- surface warnings for degraded rendering/editing
- fail closed only for unsafe or irreparably malformed packages

## Performance strategy

### Streaming and lazy loading
- central-directory-first ZIP inspection
- lazy inflation for large or unused parts
- SAX/token-stream parsing for massive XML parts
- sheet/story/slide indexes loaded on demand

### Virtualization / chunked rendering
- spreadsheet grid virtualization mandatory
- document page virtualization for long Word documents
- slide deck thumbnail virtualization for large presentations

### Worker offloading
- package parse
- XML tokenize/normalize
- formula index building
- pagination precompute batches
- corpus verification runs

### Memory constraints

Need configurable budgets:
- total decompressed bytes
- max simultaneous inflated parts
- cache eviction policies for render/layout indexes
- incremental garbage collection of detached indexes

### Benchmark strategy

Track:
- open time by file type/size
- peak memory during parse/render
- first meaningful paint
- edit latency (typing, style apply, formula entry, slide duplicate)
- serialization time
- round-trip diff size

## Internationalization

Required support targets:
- RTL and mixed-direction paragraphs
- CJK line breaking and glyph fallback
- vertical text where format supports it
- **locale-sensitive formatting** for dates, numbers, currency, and list semantics
- Unicode normalization safety
- font fallback stacks by script category

## Accessibility

Need:
- **accessibility tree** per view mode
- keyboard navigation for document/grid/slide interactions
- **screen reader** friendly text order, labels, and announcements
- focus management across overlays, comments, and embedded objects
- semantic export for tables, headings, lists, and annotations where possible

## Security

Threat model: **untrusted document policy** assumes untrusted user-supplied OOXML documents in the browser.

### Ingestion protections
- zip bomb detection
- XML size/depth/entity abuse protections
- namespace explosion limits
- malformed relationship target/path validation

### Resource handling
- **external resource handling** never auto-loads external relationships
- define strict **sanitizer boundaries** for generated HTML/SVG/clipboard payloads
- isolate embedded object handling
- treat macros/activeX/scripts as preserved unsupported payloads

### Unsupported object policy
- preserve opaque content parts and relationships
- expose metadata to UI/devtools
- never execute embedded code

## Decisions

- **D-QUAL-1:** compatibility is measured by fidelity tiers, not a single vague “works/doesn't work” label.
- **D-QUAL-2:** performance work is foundational, not a late optimization pass.
- **D-QUAL-3:** untrusted-document safety is the default runtime mode.

## Open risks

- font availability materially affects fidelity; browser-hosted font substitution reporting is required
- some accessibility semantics must be synthesized from OOXML rather than directly read from it
