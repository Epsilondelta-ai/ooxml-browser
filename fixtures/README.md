# Fixture Corpus

This directory is reserved for durable OOXML corpus fixtures used by unit, integration, interoperability, security, and performance tests.

Current state:
- parser-stage tests use programmatic micro fixtures in `tests/fixture-builders.ts`
- future commits should add persisted fixture manifests and representative Office/LibreOffice samples under `fixtures/docx`, `fixtures/xlsx`, `fixtures/pptx`, `fixtures/interop`, `fixtures/security`, and `fixtures/perf`
