# Editing Surface Matrix

This document tracks the current public editing surface exposed by `@ooxml/editor` and the corresponding round-trip expectations that are covered by automated tests.

## DOCX

| Surface | Helper | Persistence expectation | Notes |
| --- | --- | --- | --- |
| Paragraph text | `replaceDocxParagraphText` | Persisted | Text-only single-run paragraphs prefer patch preservation. |
| Story paragraph text | `replaceDocxStoryParagraphText` | Persisted | Supports `document`, `header`, `footer`. |
| Paragraph style | `setDocxParagraphStyle` | Persisted | Forces story rebuild when style metadata changes. |
| Paragraph numbering | `setDocxParagraphNumbering` | Persisted | Forces story rebuild when numbering metadata changes. |
| Paragraph run formatting | `setDocxParagraphRunStyle` | Persisted | Current helper edits bold/italic on an indexed run. |
| Table cell text | `setDocxTableCellText` | Persisted | Story rebuild path handles table mutations. |
| Comment text | `setDocxCommentText` | Persisted | Comment XML patch path preserves unrelated story parts. |
| Comment author | `setDocxCommentAuthor` | Persisted | Comment XML patch path updates `w:author`. |
| Revision metadata | `setDocxRevisionMetadata` | Persisted | Supports tracked-change id/kind/author/date/text edits through the revision-aware serializer path. |
| Section page size/margins | `setDocxSectionLayout` | Persisted | Section changes disable paragraph-only patch fast path. |
| Header/footer reference type | `setDocxSectionReferenceType` | Persisted | Reuses the section rebuild path for `sectPr` reference metadata. |
| Header/footer reference target | `setDocxSectionReferenceTarget` | Persisted | Updates document relationships to point at alternate existing header/footer parts. |

## XLSX

| Surface | Helper | Persistence expectation | Notes |
| --- | --- | --- | --- |
| Cell value | `setWorkbookCellValue` | Persisted | Numeric edits avoid shared-string churn when possible. |
| Cell formula + cached value | `setWorkbookCellFormula` | Persisted | Caller supplies the cached value. |
| Cell style index | `setWorkbookCellStyle` | Persisted | Worksheet patch path updates cell `s` attributes. |
| Row insert | `insertWorkbookRow` | Persisted | Rewrites formulas, defined names, ranges, panes, comments, and tables. |
| Defined name reference | `setWorkbookDefinedNameReference` | Persisted | Workbook patch path preserves workbook root attributes. |
| Worksheet rename | `setWorkbookSheetName` | Persisted | Also rewrites defined-name and in-sheet formula references. |
| Comment text | `setWorksheetCommentText` | Persisted | Uses comment-part patch path when author pool is unchanged. |
| Comment author | `setWorksheetCommentAuthor` | Persisted | Rebuilds comment author pool when needed. |
| Comment create/update | `upsertWorksheetComment` | Persisted | Creates a comments part and worksheet relationship on demand when missing. |
| Table name | `setWorksheetTableName` | Persisted | Table-part serializer keeps `name`/`displayName` aligned. |
| Table range | `setWorksheetTableRange` | Persisted | Table-part patch updates `ref`. |
| Frozen pane | `setWorksheetFrozenPane` | Persisted | Worksheet patch path updates pane attributes. |
| Merged ranges | `setWorksheetMergedRanges` | Persisted | May force worksheet rebuild; rebuild path preserves worksheet root attrs. |

## PPTX

| Surface | Helper | Persistence expectation | Notes |
| --- | --- | --- | --- |
| Shape text | `setPresentationShapeText` | Persisted | Text-only slide edits prefer patch preservation. |
| Shape name | `setPresentationShapeName` | Persisted | Slide serializer preserves renamed shape metadata. |
| Shape placeholder type | `setPresentationShapePlaceholderType` | Persisted | Stored directly in shape metadata. |
| Shape transform | `setPresentationShapeTransform` | Persisted | Rebuild path persists transform values. |
| Image target URI | `setPresentationImageTarget` | Persisted | Serializer updates slide image relationship targets. |
| Slide layout target | `setPresentationSlideLayout` | Persisted | Serializer retargets the slideLayout relationship to an alternate existing layout part. |
| Notes text | `setPresentationNotesText` | Persisted | Creates a notes part on demand when missing, otherwise patches the existing notes part. |
| Comment text | `setPresentationCommentText` | Persisted | Comment-part patch path. |
| Comment author | `setPresentationCommentAuthor` | Persisted | Occurrence-aware attribute patch path. |
| Comment create/update | `addPresentationComment` | Persisted | Creates a comments part and slide relationship on demand when missing. |
| Transition | `setPresentationTransition` | Persisted | Slide metadata rebuild path. |
| Timing nodes | `setPresentationTimingNodes` | Persisted | Parser reads `p:cTn` preset metadata. |
| Presentation size | `setPresentationSize` | Persisted | Serializer patches `p:sldSz` in `presentation.xml`. |

## Gaps that remain outside the public editing surface

- XLSX chart/comment-threaded extensions, print settings, selection state, and relationship-backed media/chart retargeting
- PPTX master/theme reassignment and richer animation graph editing beyond flat timing nodes
- Relationship-safe embedded-object retargeting across all formats

## Maintenance rule

When a new editing helper is added:
1. add or update round-trip tests,
2. update this matrix,
3. refresh verification artifacts if save output changes.
