# Shared Subsystems Design

These subsystems cut across docx/xlsx/pptx and must be implemented once as reusable services.

## Styles

Shared needs:
- hierarchical style resolution
- theme-aware font/color references
- direct formatting overlays
- preservation of source style identifiers and unknown attributes

Strategy:
- maintain source style graphs by format
- expose resolved computed style snapshots through common interfaces
- keep source graphs mutable via transactional editing APIs

## Themes, fonts, colors

Requirements:
- parse DrawingML theme parts and font schemes
- resolve scheme colors to actual RGB + transforms (tint/shade/lum/etc.)
- support browser font fallback and document-specified font families
- expose color resolution profiles for light/dark/high-contrast modes

## Numbering / list systems

- Word abstract numbering + instance overrides
- presentation bullet levels / list styles
- spreadsheet number formats as distinct numeric formatting system

Design decision:
- keep “semantic list structure” separate from “display label formatting”.

## Tables

Need a shared table layout abstraction because tables appear in all three formats, though semantics differ.

Shared capabilities:
- cell grid model
- row/column spans
- border/fill inheritance
- preferred/fixed/auto sizing hints
- nested content containers

## Drawings / shapes / images / charts

Reusable graph:
- binary assets registry
- drawing scene graph nodes
- transforms, fills, lines, effects, text bodies
- chart data bindings to source data/ranges
- image crop/stretch/anchor metadata

## Equations

- parse Office Math and preserve exact source markup
- expose editable math AST + presentational layout tokens
- fallback renderer can use MathML-like projection when feasible while preserving original OOXML math for serialization

## Comments / annotations / notes

Need one annotation framework with format adapters:
- Word comments + change annotations
- spreadsheet comments/notes/threaded comments
- presentation comments + speaker notes

## Metadata / hyperlinks / embedded objects

- core/app/custom properties
- hyperlinks with internal/external targets and visited state hints
- embedded packages/OLE/controls/media with safe preserved wrappers

## Tracked changes / revision-like overlays

Word has the richest native revision system, but the editor core should generalize overlay transactions:
- insertion/deletion markers
- property change overlays
- accept/reject pipelines
- author/time metadata

## Implementation implications

- shared packages should include `theme`, `style`, `drawing`, `asset`, `annotation`, and `metadata` modules
- format-specific packages should adapt shared services instead of re-implementing them
- every shared subsystem must support both **source preservation** and **resolved rendering**

## Risks

- charts require both visual scene graph support and links back to workbook/table data
- font fallback behavior differs significantly across browsers and operating systems; verification must be corpus-based
