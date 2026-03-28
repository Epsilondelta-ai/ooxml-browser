# Interoperability Matrix

Generated from `fixtures/manifests/**`, `benchmarks/reports/latest-fixture-results.json`, and `benchmarks/reports/latest-attestation-report.json`.

## DOCX

| Fixture | Tags | Mutation | Parser open | Parser round-trip | Edited round-trip | Part preservation | Changed parts | Office | LibreOffice |
| --- | --- | --- | --- | --- | --- | --- | --- | --- | --- |
| docx-basic | paragraphs, comments, table | comment-text-edit | true | true | true | 2/4 | /word/_rels/document.xml.rels, /word/comments.xml | seed-fixture-attestation-pending | seed-fixture-attestation-pending |
| docx-numbered | numbering, list-rendering | paragraph-text-edit | true | true | true | 3/4 | /word/document.xml | stage-2-attestation-pending | stage-2-attestation-pending |
| docx-revisions | tracked-changes, revisions | paragraph-text-edit | true | true | true | 1/2 | /word/document.xml | stage-2-attestation-pending | stage-2-attestation-pending |
| docx-sectioned | sections, headers, footers | paragraph-text-edit | true | true | true | 4/5 | /word/document.xml | stage-2-attestation-pending | stage-2-attestation-pending |
| docx-styled | styles, style-inheritance | paragraph-text-edit | true | true | true | 3/4 | /word/document.xml | stage-2-attestation-pending | stage-2-attestation-pending |

## PPTX

| Fixture | Tags | Mutation | Parser open | Parser round-trip | Edited round-trip | Part preservation | Changed parts | Office | LibreOffice |
| --- | --- | --- | --- | --- | --- | --- | --- | --- | --- |
| pptx-basic | slide, notes, text-shape | notes-text-edit | true | true | true | 5/6 | /ppt/notesSlides/notesSlide1.xml | seed-fixture-attestation-pending | seed-fixture-attestation-pending |
| pptx-inherited | slide-layout, slide-master, theme, placeholder | notes-text-edit | true | true | true | 10/11 | /ppt/notesSlides/notesSlide1.xml | stage-4-attestation-pending | stage-4-attestation-pending |
| pptx-media-comments | image, comments | comment-text-edit | true | true | true | 6/7 | /ppt/comments/comment1.xml | stage-4-attestation-pending | stage-4-attestation-pending |
| pptx-timed | transition, timing | shape-text-edit | true | true | true | 3/4 | /ppt/slides/slide1.xml | stage-4-attestation-pending | stage-4-attestation-pending |
| pptx-transformed | shape-transform, image-transform | shape-text-edit | true | true | true | 5/6 | /ppt/slides/slide1.xml | stage-4-attestation-pending | stage-4-attestation-pending |

## XLSX

| Fixture | Tags | Mutation | Parser open | Parser round-trip | Edited round-trip | Part preservation | Changed parts | Office | LibreOffice |
| --- | --- | --- | --- | --- | --- | --- | --- | --- | --- |
| xlsx-basic | sharedStrings, formula, worksheet | cell-value-edit | true | true | true | 4/5 | /xl/worksheets/sheet1.xml | seed-fixture-attestation-pending | seed-fixture-attestation-pending |
| xlsx-commented | comments, tables | comment-text-edit | true | true | true | 6/7 | /xl/comments1.xml | stage-3-attestation-pending | stage-3-attestation-pending |
| xlsx-structured | defined-names, merged-cells, frozen-panes, formula-references | cell-value-edit | true | true | true | 3/4 | /xl/worksheets/sheet1.xml | stage-3-attestation-pending | stage-3-attestation-pending |
| xlsx-styled | styles, number-formats | cell-value-edit | true | true | true | 5/6 | /xl/worksheets/sheet1.xml | stage-3-attestation-pending | stage-3-attestation-pending |

