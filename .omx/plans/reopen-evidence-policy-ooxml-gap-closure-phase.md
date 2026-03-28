# Reopen Evidence Policy: OOXML Gap Closure Phase

## Automated path
- Parser reopen always runs in CI for no-op/minimal-edit declared fixtures.

## Attestation path
- Office/LibreOffice reopen evidence is stored in `fixtures/manifests/**` as either automated evidence links or manual attestation records.

## Stage policy
- **Stages 1-2:** parser reopen mandatory; attestation required only for declared representative seed fixtures
- **Stages 3-4:** parser reopen mandatory; attestation expands to all stage-owned representative fixtures
- **Stages 5-6:** parser reopen + attestation both release-blocking for declared interop fixtures
