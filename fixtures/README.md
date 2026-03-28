# Fixture Corpus

Canonical fixture layout for the gap-closure phase:

- `fixtures/shared/{opc,xml,security}`
- `fixtures/docx/{micro,representative,interop,perf}`
- `fixtures/xlsx/{micro,representative,interop,perf}`
- `fixtures/pptx/{micro,representative,interop,perf}`
- `fixtures/manifests/{docx,xlsx,pptx,shared}`

Current representative seed fixtures:
- `fixtures/docx/representative/basic.docx`
- `fixtures/xlsx/representative/basic.xlsx`
- `fixtures/pptx/representative/basic.pptx`

Current manifest files:
- `fixtures/manifests/docx/basic.json`
- `fixtures/manifests/xlsx/basic.json`
- `fixtures/manifests/pptx/basic.json`

Policy:
- parser reopen is mandatory in CI for declared fixtures
- Office/LibreOffice attestation begins with representative seed fixtures and expands by stage
