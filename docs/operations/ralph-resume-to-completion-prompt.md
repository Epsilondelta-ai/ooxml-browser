# Ralph Resume-To-Completion Prompt

아래 프롬프트를 그대로 Codex/OMX에 입력하면, 현재 저장소의 기존 Ralph 상태와 계획 아티팩트를 사용해서 재계획 없이 바로 이어서 끝까지 실행하도록 지시할 수 있습니다.

```text
$ralph
Continue from current mode state.

Use the existing Ralph state and plan artifacts only:
- .omx/state/sessions/omx-1774668969490-9klz5c/ralph-state.json
- .omx/plans/consensus-plan-ooxml-gap-closure-phase.md
- .omx/plans/prd-ooxml-gap-closure-phase.md
- .omx/plans/test-spec-ooxml-gap-closure-phase.md

Do not restart ralplan.
Do not re-plan from scratch.
Resume the active approved workflow and carry it through to final completion.

Execution requirements:
- keep going until the full approved plan is complete
- do not stop at intermediate summaries
- commit after every meaningful stage using the Lore Commit Protocol
- keep the working tree clean between stages
- update docs, tests, examples, playground, fixtures, benchmarks, and quality artifacts alongside implementation
- verify every stage before claiming progress
- when the next step is clear, proceed automatically
- if work remains, continue instead of summarizing
- do not pause for confirmation

Final completion gate:
Run all of the following before stopping:
- npm test
- npm run typecheck
- npm run lint
- npm run build
- npm run bench
- npm run quality:fixtures
- npm run quality:attestations
- npm run quality:artifacts
- git diff --check

Then perform final architect sign-off and final verifier sign-off.
Stop only when Stage 6 hardening exit is fully complete and the repository is clean.
```
