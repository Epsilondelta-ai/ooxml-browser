# SpreadsheetML Reference (`.xlsx`)

## Primary structure

Typical workbook tree:
- workbook: `/xl/workbook.xml`
- workbook relationships: `/xl/_rels/workbook.xml.rels`
- worksheets: `/xl/worksheets/sheetN.xml`
- shared strings: `/xl/sharedStrings.xml`
- styles: `/xl/styles.xml`
- theme: `/xl/theme/theme1.xml`
- calc chain (legacy/optional): `/xl/calcChain.xml`
- tables, drawings, charts, comments, pivot caches/tables, external links, connections

The workbook part contains sheet metadata. Each sheet is a separate part referenced by relationship ID.

## Workbook model

Core workbook concerns:
- sheet ordering and visibility states
- workbook properties / calculation settings
- defined names / named ranges
- external links and connections
- date system (1900/1904)
- workbook protection and view settings

## Worksheet model

Key structures:
- dimensions / sheet format properties
- row and cell records
- merged cells
- hyperlinks
- auto filters / sorts
- freeze panes / sheet views / selection state
- conditional formatting
- data validation
- print/page setup and margins
- drawings anchored to cells or absolute positions

## Cell content model

Cell value kinds:
- numeric
- shared string index
- inline string / rich text
- boolean
- error
- formula
- blank with style
- date/time represented numerically with formatting semantics

Implementation requirement:
- separate raw stored value from interpreted value and displayed text.

## Shared strings

A workbook may contain a single shared string table that deduplicates strings across sheets.

Implementation requirements:
- support shared strings with rich text runs
- preserve string indices on import
- offer stable string-pool APIs for serializer reuse

## Styles and formatting

Spreadsheet formatting is highly compositional:
- number formats
- fonts
- fills
- borders
- cellXfs / cellStyleXfs
- named cell styles
- differential formats for conditional formatting
- theme-aware colors

Implementation requirement:
- compute effective cell style from style index + row/column defaults + conditional format overlays + theme.

## Formula system

The library must support:
- formula token preservation
- shared formulas / array formulas / dynamic arrays when present
- dependency graph extraction hooks
- cached values and stale recalculation flags

Product decision:
- parsing/rendering/editing must work even if full formula recalculation is provided by an optional module.

## Tables, charts, comments, drawings

- structured tables with header/total rows and table styles
- drawing part references to images/charts/shapes
- threaded vs legacy comments preservation strategy
- chart series formulas targeting workbook ranges

## Spreadsheet rendering concerns

- virtualized infinite-ish grid viewport
- pinned row/column headers
- hidden rows/columns / outline groups
- precise displayed text from number format + locale + formula result + cell width
- merged cell layout, overflow, wrap, rotation, and alignment
- sheet-level canvas layers for drawings, selections, freeze lines, print areas

## Editing concerns

- cell editing formula bar + inline entry
- multi-range selection model
- row/column insertion/deletion with reference rewriting
- copy/paste with relative reference adjustment
- style application, merge/unmerge, validation, filter/sort updates
- workbook operations: add/remove/rename/reorder sheets, manage names/tables/charts

## Serialization concerns

- preserve unknown workbook extensions, external link metadata, calc settings
- update shared string table and style tables incrementally
- keep sheet relationship IDs stable when possible
- rewrite formulas/references only for affected dependency regions

## Decisions

- **D-XLSX-1:** workbook IR separates storage, semantic, and display layers for every cell.
- **D-XLSX-2:** formula parsing is core; formula recalculation engine is pluggable.
- **D-XLSX-3:** grid renderer is virtualized from day one.

## Risks / open issues

- Excel display fidelity depends on locale-sensitive number formats and font metrics
- pivot tables, slicers, and advanced data model features should initially preserve+inspect even before full interactive editing parity
- threaded comments / modern notes require compatibility bridging with legacy comment parts
