# Architect Review: Frontend OOXML Library Plan

## Strongest steelman antithesis

A shared-core multi-format architecture may be the wrong first move because Word, Excel, and PowerPoint have materially different semantics, layout engines, and editing models. An aggressive shared IR could devolve into an awkward least-common-denominator abstraction that slows delivery, hides format-specific constraints, and produces a serializer that is too normalized to preserve source fidelity. Under this critique, three format-native stacks with a thin shared OPC/XML base would better protect fidelity and keep architectural boundaries honest.

## Real tradeoff tension

The core tension is **reuse vs fidelity-specific escape hatches**.

- More shared abstractions improve consistency, package reuse, and tooling.
- More format-specific models improve correctness for pagination, formula/reference semantics, and slide/master inheritance.

If the shared layer grows too semantic, round-trip preservation suffers. If it stays too low-level, renderer/editor reuse collapses.

## Synthesis / improvements applied

The current plan is architecturally sound after explicitly reinforcing the following points:

1. **Dual model requirement stays non-negotiable**: source token tape + semantic AST + document IR.
2. **Milestone gates were added** so each phase has objective exit criteria before downstream expansion.
3. **Reasoning/staffing guidance and launch hints were added** to make the plan execution-ready for Ralph or Team follow-up.
4. **Format-specific escape hatches remain explicit** in DOCX/XLSX/PPTX lanes rather than forcing everything into one generic editor or layout primitive.
5. **Round-trip and unsupported-content preservation remain first-class quality goals**, not secondary cleanup items.

## Verdict

**APPROVE**

The plan is execution-ready for a complete browser-first OOXML library, provided implementation keeps validating shared abstractions against format-specific fidelity requirements at every milestone gate.
