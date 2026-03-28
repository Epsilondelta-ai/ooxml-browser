# WordprocessingML Reference (`.docx`)

## Primary structure

Typical main part tree:
- main document: `/word/document.xml`
- relationships: `/word/_rels/document.xml.rels`
- styles: `/word/styles.xml`
- numbering: `/word/numbering.xml`
- settings: `/word/settings.xml`
- theme: `/word/theme/theme1.xml`
- font table: `/word/fontTable.xml`
- web settings: `/word/webSettings.xml`
- comments: `/word/comments.xml`
- footnotes/endnotes
- headers/footers
- glossary/subdocuments where present

The main story is `w:document > w:body` containing block-level elements such as paragraphs and tables.

## Story model

WordprocessingML is split into stories:
- main body
- headers
- footers
- footnotes
- endnotes
- comments
- text boxes / drawing text containers
- glossary/subdocuments

Editing/rendering implication:
- the library needs one document-level story registry, not just a single body tree.

## Core content structures

### Text
- paragraphs: `w:p`
- runs: `w:r`
- text nodes: `w:t`
- breaks/tab/symbol fields and preserved spaces
- character/paragraph properties with direct formatting + style inheritance

### Sections and pagination
- section properties via `w:sectPr`
- page size, margins, columns, page numbering, line numbering
- section breaks can live on paragraphs and affect downstream layout

### Lists / numbering
- abstract numbering definitions
- concrete numbering instances
- bullet/decimal/hybrid multilevel formats
- restart and override rules

### Tables
- grid definitions, preferred widths, cell merges, border/shading, conditional styles
- nested tables
- row/header repetition and layout constraints

### Fields / references
- simple fields and complex field characters
- hyperlinks/bookmarks
- cross-references, TOC, page numbers, date fields

### Footnotes / endnotes / comments
- anchored from main story via references
- separate story payloads with IDs
- comment ranges and author metadata

### Tracked changes
- insertions/deletions/move ranges/property changes/comment changes
- renderer needs both “final” and “show revisions” modes
- editor model must preserve acceptance/rejection semantics

### Drawings and media
- inline and anchored drawings
- images, charts, SmartArt-like diagrams, text boxes, shapes, VML fallbacks
- relationship-driven target resolution

### Equations
- Office Math (`m:*`) objects in runs/paragraphs
- preserve both semantic math tree and layout-ready presentation fragments

### Headers/footers
- first/even/default variants by section
- field-driven content (page number/date)

### Metadata / settings
- document settings, compatibility settings, proofing/language, track-changes defaults, document variables, content controls (SDTs), custom XML bindings

## WordprocessingML inheritance model

Formatting comes from multiple layers:
1. defaults / document defaults
2. theme font/color mappings
3. paragraph/character/table/list styles
4. numbering-linked styles
5. direct formatting on paragraph/run/table/cell
6. revision overlays and compatibility settings

Implementation requirement:
- compute both **preserved source formatting graph** and **resolved effective formatting**.

## Document-specific rendering concerns

- paragraph line layout with tabs, indents, justification, bidi, widow/orphan control
- page/section pagination
- floating object anchoring and wrap regions
- table auto/fixed layout negotiation
- field display vs field-code rendering
- change tracking overlays and comment anchors

## Editing concerns

- cursor positions must map to paragraph/run/text offsets and non-text inline objects
- structural edit operations: split paragraph, merge runs, apply style, toggle numbering, insert table, change section properties, accept/reject revision
- selection may span multiple stories only through higher-level APIs; low-level ranges stay story-local

## Serialization concerns

- preserve revision IDs, rsids, namespace prefixes where practical
- preserve unknown run/paragraph children and markup compatibility branches
- avoid aggressive run normalization unless explicitly requested
- maintain relationship IDs for media/comments/header/footer references where possible

## Decisions

- **D-DOCX-1:** represent Word content as a `DocumentSet` of stories plus cross-story anchors.
- **D-DOCX-2:** keep field codes and rendered field text distinct.
- **D-DOCX-3:** tracked changes remain first-class in IR; “accepted final text” is a derived projection, not destructive normalization.

## Risks / open issues

- exact Office pagination and floating-object layout will require iterative corpus calibration
- VML fallback handling should preserve markup even if renderer prefers DrawingML
- SDT/custom XML binding behavior differs across Office/LibreOffice and needs corpus-backed tolerance rules
