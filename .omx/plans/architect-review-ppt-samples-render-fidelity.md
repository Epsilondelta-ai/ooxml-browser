# Architect Review: PPT sample render fidelity

## Steelman antithesis
A corpus-first fidelity program can become a trap: the renderer starts chasing one local sample deck family with ad-hoc heuristics, accumulating brittle special cases without producing a reusable presentation scene model. Under this critique, the right move would be a principled renderer redesign first, not sample-led iteration.

## Tradeoff tension
- **Corpus wins vs reusable architecture:** sample-driven work provides fast measurable progress, but can bias implementation toward deck-specific quirks.
- **Example-surface heuristics vs core semantics:** improving the example is valuable for user trust, but if too much logic lives there, the core renderer remains weak.

## Synthesis
The plan is sound because it explicitly constrains heuristics to thin product surfaces and keeps durable scene semantics in `packages/pptx` and `packages/render`. Starting with corpus/harness work is appropriate here because the repo currently lacks the verification rails needed to know whether a deeper renderer redesign is actually helping.

## Verdict
**APPROVE**
