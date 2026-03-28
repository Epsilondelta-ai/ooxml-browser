# Rendering Architecture

## Objectives

- render OOXML content in frontend environments with high semantic and visual fidelity
- support multiple view models: page, print, grid, slide, notes, outline
- decouple parsing fidelity from browser-specific layout/render implementation details

## Rendering stack

```text
OOXML IR
  -> resolved style/theme services
  -> layout model
  -> render tree
  -> backend adapters (HTML/CSS, SVG, Canvas, hybrid)
```

## Backend strategy

This renderer family intentionally combines **HTML / CSS / SVG / Canvas** backends instead of forcing one output primitive onto every Office feature.

### Word / rich document rendering
- primary backend: HTML/CSS for text flow and editable content
- support SVG overlays for floating objects, page guides, annotations, and selection chrome
- optional Canvas assist for expensive decorations/highlight layers

### Spreadsheet rendering
- primary backend: virtualized HTML grid for cells + Canvas/SVG overlay layers for selection, autofill, frozen guides, charts, and drawings
- **spreadsheet grid rendering** must support frozen panes, merged cells, formula display, and print-preview projection
- renderer must separate data viewport from overlay scene graph

### Presentation rendering
- primary backend: SVG/HTML hybrid scene graph
- **presentation slide rendering** must preserve master/layout inheritance, shape transforms, and notes/thumbnail projections
- HTML for editable text bodies where native text selection is useful
- SVG for transform-heavy shapes and precise layering
- Canvas optional for raster effects/export acceleration

## Layout abstraction

### Word page model
- document -> sections -> pages -> block fragments -> inline fragments
- pagination engine must consider section properties, headers/footers, page breaks, keep-with-next, widow/orphan, tables, anchored objects
- alternate view: continuous/web layout without page boxes

### Spreadsheet viewport model
- workbook -> sheet -> row/column metrics -> cell viewport slices
- virtualization window + frozen panes + merged-cell expansion rules
- print view derives from page setup and print areas

### Presentation slide model
- presentation -> slide master/layout -> slide -> layered shapes -> notes/handout projections
- coordinate space uses presentation size and shape transforms

## Text layout

Requirements:
- font metrics and fallback handling
- bidi/RTL and CJK line breaking
- tab stops / justification / baseline shifts / superscript/subscript
- presentation text box autofit
- spreadsheet text overflow and rotated text

## Table layout

Need one cross-format table engine with mode adapters:
- Word: paginated table flow with row splitting/header repetition
- Spreadsheet: grid-native cell matrix
- Presentation: fixed-position table shape layout

## Drawing layout

- anchor resolution to paragraph/cell/slide coordinates
- wrap/overlay behavior
- chart layout containers
- image crop, fit, stretch, and transforms

## Print / export support

Deliverables:
- Word print-preview page renderer
- Spreadsheet print area/page-break preview
- Presentation slide export
- HTML/SVG/Canvas-backed raster or PDF export hooks

## Fidelity targets

Track separately:
- semantic fidelity (correct structure/content)
- visual fidelity (layout, colors, fonts, borders, shapes)
- interaction fidelity (selection/edit behaviors)
- round-trip fidelity (serialize without destructive rewrites)

## Decisions

- **D-RENDER-1:** use specialized renderers per document family under a shared layout/render contract.
- **D-RENDER-2:** favor HTML/CSS where browser text editing/selection is strongest; use SVG/Canvas as supporting layers, not the only abstraction.
- **D-RENDER-3:** pagination/grid/slide viewports are distinct projections over shared IR, not separate document models.

## Open risks

- browser typography differences will affect Word and slide text metrics; needs corpus-based calibration
- charts and advanced shape effects may initially require mixed native/fallback rendering while preserving editability metadata
