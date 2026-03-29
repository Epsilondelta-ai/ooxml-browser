# Consensus Plan: PPT semantic gap priority pass

## RALPLAN-DR summary
### Principles
1. Fix OOXML semantics before touching example styling.
2. Follow the user-specified priority order strictly.
3. Keep every step small, reversible, and parser-grounded.
4. Require fresh screenshot plus vision evidence before advancing.
5. Prefer rollback and forward progress over defending a regressive patch.

### Decision drivers
1. The target slides already show decent but incomplete fidelity, so incremental semantic gains are realistic and measurable.
2. The current renderer still contains example-specific PPT styling/overlay logic that can mask the real semantic gaps.
3. The parser already exposes partial transform/fill/theme/inheritance data, making parser-grounded improvement the shortest viable path.

### Viable options
- **A. Strict semantic-gap closure in parser→renderer order** — chosen
  - Pros: aligns exactly with the user mandate, improves reusable library behavior, keeps changes attributable
  - Cons: some visible issues may remain until later stages
- **B. Example-surface retuning first, then parser cleanup** — rejected
  - Pros: faster cosmetic wins
  - Cons: violates the task constraint and risks overfitting sample styling
- **C. Mixed opportunistic fixes without strict order** — rejected
  - Pros: may chase the largest immediate score delta
  - Cons: obscures causality and conflicts with the requested priority discipline

### Decision
Choose **Option A**: execute semantic fixes in the exact requested order, verifying and reverting per stage as needed.

### Consequences
- Work will bias toward parser/model correctness even when example-only fixes seem faster.
- Some stages may be reverted if evidence is negative.
- Overlay cleanup is intentionally deferred until semantic foundations improve.

## Exact staged implementation order
### Stage 0 — baseline capture and clean-tree gate
- **Intent:** confirm baseline target-slide scores, record current artifacts, and reset non-source artifact churn before semantic edits/commits.
- **Likely files:** `benchmarks/reports/ppt-sample-screenshot-report.json`, `benchmarks/reports/ppt-sample-screenshots/**`, `benchmarks/reports/ppt-sample-diffs/**`, `.omx/plans/*`.
- **Gate:** baseline ledger captured; generated artifacts excluded or intentionally refreshed before each commit.

### Stage 1 — group transform accuracy
- **Intent:** correct nested group transform propagation so child shapes land in the right position/scale/orientation.
- **Likely files:** `packages/pptx/src/parser.ts`, `packages/pptx/src/model.ts`, `tests/pptx-shape-transform.test.ts`, possibly `tests/render-and-browser.test.ts`.
- **Gate:** target slides must not regress after screenshot + vision review.

### Stage 2 — custom/preset geometry fidelity
- **Intent:** improve shape geometry extraction and renderer projection for authored OOXML geometry.
- **Likely files:** `packages/pptx/src/parser.ts`, `packages/pptx/src/model.ts`, `packages/render/src/pptx.ts`, `examples/basic/src/main.ts`, geometry-focused PPT tests.
- **Gate:** geometry silhouette/placement is neutral-or-better on target slides.

### Stage 3 — placeholder/layout/master inheritance
- **Intent:** deepen inheritance resolution for placeholders, layout/master defaults, and effective slide properties.
- **Likely files:** `packages/pptx/src/parser.ts`, `packages/pptx/src/model.ts`, `tests/pptx-inheritance.test.ts`, `tests/opc-and-parsers.test.ts`, `tests/render-and-browser.test.ts`.
- **Gate:** target-slide text/layout structure improves or stays stable.

### Stage 4 — theme color transforms
- **Intent:** resolve scheme colors and transforms consistently for fills, text, and lines.
- **Likely files:** `packages/pptx/src/parser.ts`, `packages/pptx/src/model.ts`, color-focused tests.
- **Gate:** color fidelity is neutral-or-better in metrics/vision review.

### Stage 5 — line/fill/stroke semantics
- **Intent:** refine width/fill opacity/default stroke behavior using OOXML semantics, not CSS guessing.
- **Likely files:** `packages/pptx/src/parser.ts`, `packages/render/src/pptx.ts`, `examples/basic/src/main.ts`, render tests.
- **Gate:** stroke/fill appearance is neutral-or-better on target slides.

### Stage 6 — overlay hack minimization
- **Intent:** delete or narrow example-only PPT overlay/style hacks once semantic rendering can carry more of the output.
- **Likely files:** `examples/basic/src/main.ts`, `tests/render-and-browser.test.ts`.
- **Gate:** no regression after removal/narrowing.

## Test-spec additions for this workflow
1. Baseline score ledger for the three target slides.
2. Stage-by-stage accept/revert record in the final report.
3. Focused parser/render tests added before or alongside each semantic change.
4. Mandatory screenshot + vision evidence for each stage, not just final output.

## ADR
### Decision
Use strict ordered semantic-gap closure with per-stage verification and rollback.

### Drivers
- user-specified priority order
- parser-grounded constraint
- target-slide evidence already available

### Alternatives considered
- example styling first
- unordered opportunistic fixes

### Why chosen
It best isolates which OOXML semantics actually improve fidelity and prevents hidden regressions.

### Consequences
Progress may be slower, but evidence quality and architectural cleanliness improve.

### Follow-ups
If accepted semantic changes still leave major gaps, open a later phase for renderer-surface improvements outside this strict order.

## Available-agent-types roster
- `executor` — implementation lane for parser/render/test changes
- `debugger` — regression diagnosis when a stage fails verification
- `test-engineer` — verification hardening and focused regression tests
- `architect` — final plan/work review and sign-off
- `verifier` — completion evidence cross-check
- `writer` — final evidence/report polishing if needed

## Ralph staffing guidance
- **Implementation lane:** `executor` with high reasoning for parser/render code changes.
- **Evidence lane:** local shell + `test-engineer`/`debugger` support for screenshot, metrics, and regression triage.
- **Final sign-off lane:** `architect` minimum, `verifier` optional but recommended.

## Team / Ralph handoff guidance
- Prefer **Ralph** for this task because the stages are order-dependent and each next stage depends on the previous stage's accept/revert decision.
- If team mode is later needed, split only into:
  - implementation support on a bounded active stage,
  - evidence collection on that same stage,
  - independent final review.

## Verification path
1. establish/confirm baseline target-slide scores,
2. stage 1 through 6 in order,
3. after each stage run tests + `npm run quality:ppt-sample-screenshots` + vision review,
4. revert failed stages immediately,
5. run final repo-wide verification,
6. obtain architect approval,
7. complete Ralph cleanup.
