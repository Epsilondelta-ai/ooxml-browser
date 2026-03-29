# Test Specification: PPT samples render fidelity

## Verification principles
1. The paired PPTX/PNG corpus is the source of truth for this phase.
2. Visual progress must be measured with reproducible screenshot captures, not only semantic assertions.
3. Parsing, rendering, and viewer behavior all need explicit coverage.
4. Sample-specific regressions should be tracked slide-by-slide.

## Test layers
### Unit
- root document detection / package-opening regressions for corpus decks
- PPTX parser extraction of slide size, backgrounds, shape text, transforms, media, and placeholder metadata
- PPTX renderer projection helpers for title/content/image/card slides

### Integration
- open `sample{1..6}/sample.pptx` and verify slide counts match paired PNG counts
- render selected slides from each sample and verify expected DOM markers/image usage
- browser example loads sample decks and supports slide navigation

### Visual regression
- capture screenshots for declared slides from each sample deck using pinned viewport/browser settings
- compare against paired PNG exports
- record per-slide verdicts and hotspot notes for iteration

### Observability
- persist a manifest/report mapping sample folder -> slide number -> screenshot path -> reference path -> verdict
- record parser diagnostics for unsupported shapes/effects encountered during runs

## Minimum declared corpus gate
- sample1: title slide, section break, representative dense content slide
- sample2: title slide, image-heavy slide, mixed text/image slide
- sample3: title slide, representative content slide
- sample4: title slide, representative content slide
- sample5: title slide, representative graphical/content slide
- sample6: title slide plus one later representative slide

## Required commands
- `npm test`
- `npm run typecheck`
- `npm run lint`
- `npm run build`
- any new screenshot/reference command introduced by this phase
- `git diff --check`

## Exit criteria
1. no sample deck parses with the wrong slide count or silently degrades to empty output,
2. declared screenshot comparisons are captured and reviewed,
3. browser example/playground preview changes are covered by tests or deterministic harness evidence,
4. final architect + verifier sign-off approve the evidence set.
