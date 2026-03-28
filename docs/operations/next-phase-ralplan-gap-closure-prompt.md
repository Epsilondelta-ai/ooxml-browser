# Next-phase OOXML Gap-Closure Prompt

아래 프롬프트를 그대로 Codex/OMX에 입력하면, 현재 저장소의 부족한 점을 기준으로 `$ralplan --deliberate` 계획 수립부터 `$ralph` 실행까지 이어지는 다음 단계 작업을 시작할 수 있습니다.

```text
$ralplan --deliberate
You are an autonomous Codex/OMX coding agent operating in /Users/juunini/Desktop/code/epsilondelta/ooxml.

Context:
- The repository already contains a browser-first OOXML workspace with docs, consensus plan artifacts, OPC parsing, basic docx/xlsx/pptx parsers, HTML rendering, transaction-based editing, serializer support, example app, playground, and benchmark harness.
- Current baseline is usable for representative parse/render/edit/save flows, but not yet Microsoft Office-complete.
- Existing durable docs live under docs/.
- Existing planning artifacts live under .omx/plans/.
- You must preserve and extend the existing architecture rather than restarting from scratch.

Current gaps to close in the next planning/execution wave:
1. DOCX fidelity gaps
   - deeper style inheritance
   - numbering fidelity
   - headers/footers and section behavior
   - comments/revisions fidelity
   - tables/layout fidelity
   - drawings/images/charts/equations preservation and projection
2. XLSX fidelity gaps
   - styles and number formats
   - formulas/reference rewriting breadth
   - merges/frozen panes/defined names/tables/comments/charts
   - sharedStrings/inlineStr/full round-trip coverage
   - grid/view/print fidelity
3. PPTX fidelity gaps
   - slide master/layout/theme inheritance depth
   - notes/comments/media handling breadth
   - richer shape model
   - timing/animation preservation
   - non-trivial notes/relationship scenarios
4. Shared architecture gaps
   - round-trip/source-preserving model depth vs currently lossy typed projections
   - relationship-safe mutation helpers for more part types
   - broader serializer coverage without destructive regeneration
   - compatibility corpus expansion using durable fixtures
5. Product/quality gaps
   - stronger playground/example automation
   - broader corpus fixtures under fixtures/
   - richer benchmark suite
   - interoperability matrix and regression gates
   - more realistic MS Office / LibreOffice reopen validation

Tasks for ralplan:
- Produce a deliberate consensus plan for closing the above gaps without reducing scope.
- Reuse and extend the existing docs/ and .omx/plans/ baseline.
- Create updated PRD and test-spec artifacts under .omx/plans/ for this gap-closure phase.
- Include:
  - exact staged implementation order
  - module/file ownership boundaries
  - parser/render/editor/serializer expansion plan
  - fixture/corpus plan
  - benchmark/interoperability plan
  - architecture rules for preserving source fidelity while extending typed models
  - explicit verification gates for each stage
  - definition of done for “MS Office-meaningful” next phase
- The plan must be concrete enough for immediate execution via Ralph.

After ralplan approval:
- Execute the full approved plan with $ralph.
- Do not stop halfway.
- Commit after each meaningful stage using the Lore Commit Protocol.
- Keep the working tree clean between stages.
- Update docs, tests, examples, playground, and benchmarks alongside implementation.
- Before final completion, verify with:
  - npm test
  - npm run typecheck
  - npm run lint
  - npm run build
  - npm run bench
  - affected diagnostics clean
  - architect review
  - verifier review

Final objective:
Advance the repository from the current representative OOXML workflow baseline toward a substantially more complete Microsoft Office-compatible frontend library across docx/xlsx/pptx parse, render, edit, and round-trip serialization behavior.
```
