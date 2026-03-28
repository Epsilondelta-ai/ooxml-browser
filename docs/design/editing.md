# Editing Architecture

## Objectives

- allow structural, text, formatting, asset, and layout edits across docx/xlsx/pptx
- preserve round-trip fidelity and source provenance
- provide undo/redo, selection, transactions, clipboard, and collaboration extension points

## Editor data model

Layered approach:
1. **Source-preserving document model**: package graph + raw XML/token preservation
2. **Normalized semantic IR**: editable typed objects
3. **Derived interaction indexes**: cursor maps, selection anchors, layout hit-test indexes, dependency graphs
4. **Transaction log**: reversible operations with semantic intent

## Selection / range / cursor model

### Word
- text positions within story-local text streams
- block selections
- table cell and drawing selections
- annotation/comment anchors

### Spreadsheet
- cell range selections
- discontiguous multi-range selections
- row/column/full-sheet selections
- in-cell text edit sub-selection

### Presentation
- shape selections
- text caret inside shape text body
- multi-shape selections, group selections, z-order handles

Shared abstraction:
- `SelectionSnapshot` + format-specific detail payload

## Mutation API

Mutation APIs should be semantic, not raw-XML-only.

Examples:
- insertText
- splitParagraph
- applyRunStyle / applyParagraphStyle
- insertTable / updateTableGrid
- addImage / replaceAsset
- setCellValue / setFormula / applyCellStyle
- addSlide / duplicateSlide / moveSlide
- acceptRevision / rejectRevision

Every mutation returns:
- transaction record
- affected document scopes
- invalidation hints for layout/render caches
- serialization patch hints

## Transaction model

Requirements:
- explicit begin/commit/rollback boundaries
- nested transactions for composite edits
- **undo / redo** stacks built from inverse operations
- deterministic transaction serialization for collaboration/replay hooks

## Clipboard / paste

Must support:
- internal rich paste preserving semantic structures
- HTML/plaintext/image clipboard bridges
- spreadsheet TSV/CSV grid paste
- PowerPoint-like shape duplication with asset deduplication

## Structural editing

- Word sections, numbering, styles, comments, tracked changes, headers/footers
- Spreadsheet sheet ops, row/column ops, named ranges, table resize, chart source updates
- Presentation slide ops, layout changes, placeholder bindings, master-aware overrides

## Collaboration / conflict extensibility

The editor layer must support **collaborative editing extensibility** without replacing the local transaction engine.

Planned extension points:
- operational transform or CRDT adapters at transaction layer
- document object IDs stable enough for remote merges
- semantic conflict descriptors (same cell changed, same paragraph style changed, same shape transform changed)

Conflict policy:
- expose **conflict resolution** metadata at semantic-object granularity
- allow caller-defined merge policies for text, cell, and shape changes

## Serialization back to OOXML

Write path principles:
- patch only affected parts when possible
- preserve untouched parts byte-for-byte where feasible
- preserve unknown elements/attributes/relationship parts
- regenerate dependent indexes/tables (shared strings, relationship parts, content types, style refs) deterministically

## Decisions

- **D-EDIT-1:** editor operations are semantic transactions, not ad hoc DOM mutation.
- **D-EDIT-2:** undo/redo is transaction-based across all formats.
- **D-EDIT-3:** collaboration support is a transaction-layer extension, not a separate editor model.

## Open risks

- Word tracked-changes editing semantics are subtle and require policy switches (`preserve revisions`, `co-authoring mode`, `accept on normalize` never by default)
- formula/reference rewrite correctness in spreadsheets needs a dedicated dependency/reference engine
