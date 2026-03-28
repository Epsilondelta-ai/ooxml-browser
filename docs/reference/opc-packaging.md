# OOXML / OPC Packaging Reference

## Why this layer is first

OOXML documents are OPC packages containing content-type declarations, parts, and relationship parts. Every higher-level parser and serializer decision depends on this layer being correct.

## OPC concepts to implement

### ZIP package structure

OOXML files are ZIP containers whose logical model is richer than a plain directory tree.

Required packaging concepts:
- package root
- `[Content_Types].xml`
- package relationships part (`/_rels/.rels`)
- part-local relationship parts (`<folder>/_rels/<name>.rels`)
- unique part URIs
- content types via default and override declarations

Important implication:
- ZIP entry paths are transport details; traversal should be relationship-driven whenever possible.

### Parts

A part is a named resource inside the package with a content type.

Library requirements:
- preserve part URI exactly
- store content type, size, CRC/integrity metadata when available
- distinguish XML, binary, media, embedded package, font, custom XML, and unknown parts
- support untouched pass-through for unsupported but preservable parts

### Relationships

Relationships connect a source (package or part) to a target (internal part or external resource).

Consumer requirements:
- parse package-level and part-level relationship parts
- preserve `Id`, `Type`, `Target`, `TargetMode`
- resolve relative targets against source part URI
- expose typed navigation APIs (`find main document`, `resolve theme`, `sheet -> drawing`, `slide -> notes`)
- preserve unknown relationship types during round-trip

Serializer requirements:
- maintain relationship ID stability when feasible
- regenerate deterministic `.rels` output when topology changes
- preserve external links while marking them unsafe-by-default for loading/rendering

### Content Types

`[Content_Types].xml` maps extensions and part overrides to MIME-like content types.

Implementation rules:
- parse both default and override maps
- detect content type mismatches between declaration and payload class
- prefer override lookup by full URI, then extension default lookup
- preserve ordering/canonicalization strategy for stable serialization

### Package traversal strategy

Recommended traversal order:
1. decompress central directory with size and compression metadata
2. enforce security limits (entry count, compression ratio, total expanded bytes, path validity)
3. parse `[Content_Types].xml`
4. parse package relationships (`/_rels/.rels`)
5. identify primary office document relationship (`officeDocument`) or equivalent root
6. traverse reachable graph by relationship edges
7. record unreachable/orphan parts for preservation diagnostics
8. lazily parse part payloads on demand

### Package graph data structure

```text
PackageGraph
  parts: Map<PartUri, PackagePartNode>
  relationshipsBySource: Map<SourceUri | "package", Relationship[]>
  contentTypes: { defaults, overrides }
  rootDocument: PartUri | null
  strictMode: boolean
  macroEnabled: boolean
  signatures: PackageSignatureInfo[]
  customProperties: ...
```

## OOXML-specific packaging conventions

Typical roots:
- `.docx`: `/word/document.xml`
- `.xlsx`: `/xl/workbook.xml`
- `.pptx`: `/ppt/presentation.xml`

Common shared parts:
- theme parts
- styles parts
- settings/properties parts
- core/app/custom properties
- media parts
- embedded packages/objects
- custom XML data

## Markup compatibility and extensibility

OOXML uses Markup Compatibility (`mc:*`) and Alternate Content to support different producer/consumer capabilities.

Implementation rules:
- store raw compatibility attributes
- preserve all alternate branches in source IR
- choose one active branch for rendering based on capability profile
- serialize dormant branches unchanged unless explicitly edited
- never strip unknown namespaces merely because the current renderer does not use them

## Strict vs Transitional

The library must:
- detect Strict documents and namespaces
- normalize namespace aliases in parser services without losing original namespace intent
- preserve strict/transitional form at serialization policy level
- support compatibility conversion rules only when explicitly requested

## Fault tolerance / recovery

Tolerance policy:
- malformed package -> fail closed with structured diagnostics
- recoverable relationship/content-type mismatch -> keep package open in degraded mode if referenced payload can still be resolved safely
- unknown parts/relationships -> preserve and surface as opaque nodes
- orphaned parts -> retain unless cleanup is explicitly requested

## Security implications

- reject `../` / absolute path traversal in ZIP entry names
- limit total decompressed bytes, per-entry size, nesting-like explosion patterns, and entry counts to resist zip bombs
- never auto-fetch external relationship targets
- treat macros, activeX, OLE, embedded executables, and scriptable payloads as non-executable preserved attachments

## Decisions

- **D-OPC-1:** the parser's first normalized artifact is `PackageGraph`.
- **D-OPC-2:** traversal is relationship-driven, with path heuristics only as fallback diagnostics.
- **D-OPC-3:** package preservation is a first-class fidelity dimension separate from semantic render fidelity.

## Risks / open issues

- digital signature validation is orthogonal to main rendering/editing flow and should be exposed as metadata + verification utilities before full authoring support
- encrypted/password-protected Office packages require a separate pre-processing layer and are out of scope for initial core parsing unless decryption material is provided externally

## Implementation implications

- core package parser should live in a format-agnostic module
- all format parsers should accept a `PackageGraph`
- serializers should patch the graph rather than regenerate the entire package from scratch wherever possible
