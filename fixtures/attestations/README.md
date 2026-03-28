# Reopen Attestations

Store manual or semi-automated reopen evidence here as one JSON file per fixture.

File naming:
- `fixtures/attestations/<fixture-id>.json`

Shape:
```json
{
  "fixtureId": "docx-basic",
  "office": {
    "status": "pass",
    "validatedAt": "2026-03-28T00:00:00Z",
    "version": "Microsoft 365 2502",
    "notes": "Opened, edited, and resaved without repair dialog"
  },
  "libreOffice": {
    "status": "pass",
    "validatedAt": "2026-03-28T00:10:00Z",
    "version": "LibreOffice 25.2",
    "notes": "Opened and resaved successfully"
  }
}
```

Allowed statuses are project-defined strings such as:
- `pass`
- `fail`
- `pending`
- `seed-fixture-attestation-pending`
- `stage-2-attestation-pending`
- `stage-3-attestation-pending`
- `stage-4-attestation-pending`

Workflow:
1. Open a representative fixture in the target suite.
2. Validate open/render/edit/save/reopen behavior.
3. Save the attestation JSON under this directory.
4. Run `npm run quality:attestations`.
5. Run `npm run quality:artifacts` to refresh the matrix/report surfaces.
