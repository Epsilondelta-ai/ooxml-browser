# Interoperability Matrix

Generated from `fixtures/manifests/**`.

## DOCX

| Fixture | Tags | Parser | Office | LibreOffice | Supported operations |
| --- | --- | --- | --- | --- | --- |
| docx-basic | paragraphs, comments, table | true | seed-fixture-attestation-pending | seed-fixture-attestation-pending | no-op round trip, paragraph text edit |
| docx-numbered | numbering, list-rendering | true | stage-2-attestation-pending | stage-2-attestation-pending | no-op round trip, numbered paragraph render |
| docx-revisions | tracked-changes, revisions | true | stage-2-attestation-pending | stage-2-attestation-pending | no-op round trip, revision render |
| docx-sectioned | sections, headers, footers | true | stage-2-attestation-pending | stage-2-attestation-pending | no-op round trip, header/footer render |
| docx-styled | styles, style-inheritance | true | stage-2-attestation-pending | stage-2-attestation-pending | no-op round trip, styled paragraph render |

## PPTX

| Fixture | Tags | Parser | Office | LibreOffice | Supported operations |
| --- | --- | --- | --- | --- | --- |
| pptx-basic | slide, notes, text-shape | true | seed-fixture-attestation-pending | seed-fixture-attestation-pending | no-op round trip, shape text edit, notes edit |
| pptx-inherited | slide-layout, slide-master, theme, placeholder | true | stage-4-attestation-pending | stage-4-attestation-pending | no-op round trip, inheritance metadata render |
| pptx-media-comments | image, comments | true | stage-4-attestation-pending | stage-4-attestation-pending | no-op round trip, media/comment render |
| pptx-timed | transition, timing | true | stage-4-attestation-pending | stage-4-attestation-pending | no-op round trip, timing metadata render |
| pptx-transformed | shape-transform, image-transform | true | stage-4-attestation-pending | stage-4-attestation-pending | no-op round trip, transform metadata render |

## XLSX

| Fixture | Tags | Parser | Office | LibreOffice | Supported operations |
| --- | --- | --- | --- | --- | --- |
| xlsx-basic | sharedStrings, formula, worksheet | true | seed-fixture-attestation-pending | seed-fixture-attestation-pending | no-op round trip, cell edit |
| xlsx-commented | comments, tables | true | stage-3-attestation-pending | stage-3-attestation-pending | no-op round trip, comment and table render |
| xlsx-structured | defined-names, merged-cells, frozen-panes, formula-references | true | stage-3-attestation-pending | stage-3-attestation-pending | no-op round trip, worksheet structure render |
| xlsx-styled | styles, number-formats | true | stage-3-attestation-pending | stage-3-attestation-pending | no-op round trip, styled numeric render |

