# Research Source Map

This project uses the following primary references as the baseline research set for documentation and implementation planning.

## Standards and official references

1. **ECMA-376 Office Open XML File Formats**
   - https://ecma-international.org/publications-and-standards/standards/ecma-376/
   - Why it matters: canonical high-level standard entry point for OOXML parts, packaging, markup compatibility, and format vocabularies.

2. **Open Packaging Conventions Fundamentals (Microsoft Learn)**
   - https://learn.microsoft.com/en-us/previous-versions/windows/desktop/opc/open-packaging-conventions-overview
   - Why it matters: concise explanation of OPC logical/physical model, parts, and relationships.

3. **About the Open XML SDK for Office (Microsoft Learn)**
   - https://learn.microsoft.com/en-us/office/open-xml/about-the-open-xml-sdk
   - Why it matters: practical summary of package structure, strict/transitional support, and common part organization.

4. **Introduction to markup compatibility (Microsoft Learn)**
   - https://learn.microsoft.com/en-us/office/open-xml/general/introduction-to-markup-compatibility
   - Why it matters: practical overview of `mc:*` behavior, alternate content, and preprocessing expectations.

5. **XML file name extension reference for Office (Microsoft Learn)**
   - https://learn.microsoft.com/en-us/office/compatibility/xml-file-name-extension-reference-for-office
   - Why it matters: strict/transitional and macro-enabled extension distinctions.

## Format-specific structure references

6. **Structure of a WordprocessingML document (Microsoft Learn)**
   - https://learn.microsoft.com/en-us/office/open-xml/word/structure-of-a-wordprocessingml-document
   - Why it matters: minimal and typical `.docx` part structure.

7. **Structure of a SpreadsheetML document (Microsoft Learn)**
   - https://learn.microsoft.com/en-us/office/open-xml/spreadsheet/structure-of-a-spreadsheetml-document
   - Why it matters: workbook/sheet relationship model and typical `.xlsx` structure.

8. **Structure of a PresentationML document (Microsoft Learn)**
   - https://learn.microsoft.com/en-us/office/open-xml/presentation/structure-of-a-presentationml-document
   - Why it matters: presentation/slide master/layout/theme/notes model.

## Shared subsystem / large-document references

9. **Working with the shared string table (Microsoft Learn)**
   - https://learn.microsoft.com/en-us/office/open-xml/spreadsheet/working-with-the-shared-string-table
   - Why it matters: workbook-wide string deduplication and rich text implications.

10. **How to parse and read a large spreadsheet document (Microsoft Learn)**
    - https://learn.microsoft.com/en-us/office/open-xml/spreadsheet/how-to-parse-and-read-a-large-spreadsheet
    - Why it matters: DOM vs SAX tradeoffs, large-part streaming strategy.

11. **Working with slide masters (Microsoft Learn)**
    - https://learn.microsoft.com/en-us/office/open-xml/presentation/working-with-slide-masters
    - Why it matters: theme/text style inheritance on slides.

## How these sources are used in this repo

- The standard/official references inform package rules, markup compatibility, strict/transitional handling, and required/optional part reasoning.
- This repo's design docs intentionally go beyond source summaries: they translate official format rules into frontend-specific architecture choices for parsing, rendering, editing, and serialization.
- Where implementation docs discuss strategy not explicitly prescribed by the standard (for example worker interfaces, editor transactions, virtualization, HTML/CSS rendering, or test harness design), those are project decisions derived from product requirements rather than quotations from the standard.

## Implementation posture derived from the sources

- Treat OOXML as **an OPC package graph first**, not “a zip full of XML files”.
- Resolve document navigation by **relationship type + relationship target**, not by hardcoded paths alone.
- Preserve **markup compatibility branches, unknown extensions, and strict/transitional distinctions** for round-trip safety.
- Prefer **streaming/SAX-style part parsing** for large XML parts and DOM-style trees only where whole-tree editing is necessary.
