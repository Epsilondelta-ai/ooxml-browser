# Test Specification: PPT sample render fidelity

## Verification principles
1. Screenshot/reference comparison is first-class evidence for this phase.
2. Parser correctness and visual correctness are separate gates.
3. Every corpus-driven rendering change must have an automated regression path.
4. Sample6 root-detection coverage is mandatory because corpus completeness depends on it.

## Test layers
### Unit
- PPTX root-document detection and fallback handling
- slide/background/fill/theme parsing helpers
- image/media part resolution for slide rendering
- slide-family classification helpers used by the renderer
- scene-layout helpers for transform-to-viewport projection

### Integration
- open sample deck -> parse -> render selected slide -> serialize/reopen where applicable
- example presentation mode loads a sample PPTX and navigates slides without throwing
- media-backed slides resolve image parts to browser-renderable URLs

### Visual / golden
- browser screenshot capture per slide for the sample corpus
- reference comparison against paired `sample.###.png`
- per-slide verdict JSON with score, mismatch categories, and actionable diffs
- visual baselines grouped by sample folder and slide index

### E2E
- load a sample PPTX in `examples/basic`
- navigate slides
- verify slide count/active slide state
- capture screenshots for selected representative slides (title, content, image-heavy, section-break)

### Observability
- corpus manifest with slide counts and status
- mismatch report summarizing top visual divergence causes
- unsupported/partial-support report for any remaining slides

## Corpus matrix
- sample1: text/vector heavy, 48 slides
- sample2: includes image-heavy slides, 48 slides
- sample3: mixed image/text, 48 slides
- sample4: text/vector heavy, 48 slides
- sample5: different deck size and style family, 37 slides
- sample6: current root-detection failure case

## Stage-specific evidence gates
### Stage 1
- sample6 root-detection failure reproduced with a regression test
- corpus manifest or loader covers all sample folders and slide counts

### Stage 2
- parser tests cover background/fill/image/placeholder extraction needed by the sample corpus
- representative slides from sample1/sample2/sample5 show richer scene metadata than today

### Stage 3
- presentation renderer displays actual image media for image slides
- title/content slide families render with transform-aware scene placement
- browser example can switch slides without layout collapse

### Stage 4
- screenshot/reference comparison command runs for the sample corpus
- per-slide verdict artifacts written to a durable report location
- pass threshold defined (target 90+, revise below 90)

### Stage 5
- docs/examples/playground updated alongside renderer behavior
- final corpus run shows no parser crashes and a materially improved visual score distribution

## Commands to add or use
- targeted PPTX parser/render tests under `tests/`
- example/build checks
- screenshot/reference harness command (new)
- final repo verification:
  - `npm test`
  - `npm run typecheck`
  - `npm run lint`
  - relevant build/example checks

## Acceptance thresholds
- No sample deck may fail to open without an explicit tested diagnostic path.
- Title/content/image representative slides from each sample family should reach the agreed visual score threshold or have documented residual gaps.
- Renderer regressions in existing PPTX tests are not allowed.
