# Test Specification: PPT semantic gap priority pass

## Verification principles
1. Parser-grounded semantics outrank deck-specific appearance tweaks.
2. Every priority step needs fresh evidence before the next edit cycle.
3. Screenshot metrics and visual inspection are complementary; neither is sufficient alone.
4. Regressions are handled by rollback of the current step, not by silently keeping degraded code.

## Baseline gate
Before stage 1:
1. capture baseline target-slide scores from the current screenshot report,
2. decide whether generated screenshot/diff artifacts are source-of-truth outputs or disposable verification artifacts for the current commit,
3. ensure the working tree is clean except for intentional plan/code changes.

## Priority-step verification loop
For each priority stage (1 through 6):
1. implement the smallest reversible semantic change for that stage,
2. run focused/unit/integration tests for touched parser/render behavior,
3. run `npm run quality:ppt-sample-screenshots`,
4. extract metrics for `sample1` slide 1, `sample5` slide 2, `sample6` slide 1,
5. perform LLM vision review on the screenshot/reference pairs for those target slides,
6. accept the stage only if evidence is neutral-or-better overall; otherwise revert the stage and continue to the next priority.

## Required verification layers
### Unit / parser semantics
- transform composition tests for grouped shapes
- geometry/path parsing tests for custom/preset shapes
- inheritance resolution tests for layout/master placeholders
- theme color resolution tests for scheme/transforms
- fill/line semantics tests for stroke width, opacity, and fallback behavior

### Integration
- target-slide parser snapshots or structured assertions showing expected transforms/colors/inheritance fields
- browser/render tests that ensure effective slide data reaches the DOM/canvas layer

### Visual regression
- `npm run quality:ppt-sample-screenshots`
- record target-slide score deltas versus the baseline:
  - sample1/1: `85.20`
  - sample5/2: `81.14`
  - sample6/1: `89.84`
- preserve screenshot/diff artifacts for accepted stages

### Vision review
- review the reference image and generated screenshot for:
  - shape placement/alignment
  - geometry silhouette fidelity
  - placeholder-driven layout correctness
  - theme color correctness
  - line/fill/stroke appearance
  - reduced reliance on overlay chrome

## Repo-wide exit verification
- `npm test`
- `npm run typecheck`
- `npm run lint`
- `npm run build`
- `git diff --check`
- architect review approval

## Evidence ledger requirements
For each stage record:
- files changed
- tests run
- screenshot scores for target slides
- vision verdict summary
- accepted vs reverted decision
- next stage
