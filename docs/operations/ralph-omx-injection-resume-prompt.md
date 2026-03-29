# Ralph OMX Injection Resume Prompt

이전 세션에서 OMX injection 이 빠져 Ralph가 끝까지 지속되지 않았다고 판단될 때, 아래 프롬프트를 그대로 다음 Codex/OMX 세션에 넣으면 됩니다.

- 목적: 기존 OOXML gap-closure 작업을 **재계획 없이 복구/재개**하고, **검증 완료 + `/cancel` 실행 전까지 멈추지 않게** 지시합니다.
- 전제: 저장소 루트는 `/Users/juunini/Desktop/code/epsilondelta/ooxml` 입니다.

```text
$ralph
You are an autonomous Codex/OMX agent operating in /Users/juunini/Desktop/code/epsilondelta/ooxml.

Critical recovery context:
- The prior Ralph session did not persist correctly because OMX injection/context was incomplete.
- This run is a recovery-and-complete run.
- Do not treat this as a fresh task unless repository evidence proves there is no resumable state.

Before doing anything else:
- Read and obey the repository AGENTS.md instructions that govern this workspace.
- Load and honor existing Ralph/OMX artifacts before re-planning.

Primary objective:
Resume and fully complete the unfinished OOXML gap-closure execution work that was being tracked by Ralph, carrying the approved plan through to real completion.

Use existing recovery artifacts first:
- .omx/state/sessions/omx-1774668969490-9klz5c/ralph-state.json
- .omx/state/sessions/omx-1774668969490-9klz5c/ralph-progress.json
- .omx/context/ooxml-gap-closure-20260328T050715Z.md
- .omx/plans/consensus-plan-ooxml-gap-closure-phase.md
- .omx/plans/prd-ooxml-gap-closure-phase.md
- .omx/plans/test-spec-ooxml-gap-closure-phase.md
- docs/operations/ralph-resume-to-completion-prompt.md

Recovery rules:
- Reconstruct missing or stale Ralph state from repository evidence if needed.
- Inspect git status, recent commits, changed files, pending gaps, TODOs, tests, docs, fixtures, benchmarks, and quality artifacts.
- If the saved state says "complete" but repo evidence or fresh verification shows unfinished work, treat the saved completion flag as stale and continue from the last incomplete stage.
- Do not restart ralplan unless the required plan artifacts are actually missing or invalid.
- Prefer continuing the approved workflow over making a new plan.

Execution rules:
- Continue from the last incomplete stage of the OOXML gap-closure phase.
- Keep going automatically through clear next steps.
- Do not stop at interim summaries.
- Do not ask for confirmation on reversible next actions.
- Do not reduce scope.
- Commit after every meaningful stage using the Lore Commit Protocol.
- Keep the working tree clean between stages.
- Update docs, tests, examples, playground, fixtures, benchmarks, and quality artifacts alongside implementation.
- Verify every stage before claiming progress.
- If architect or verifier review fails, fix issues and continue.
- If the hook/system says “The boulder never stops”, continue the iteration.

Known last tracked stage from prior state:
- Stage 6 hardening - charted XLSX representative corpus expansion

Completion gate:
Do not stop until all required work is truly complete and the full Ralph completion checklist is satisfied with fresh evidence.

Minimum final verification required before stopping:
- npm test
- npm run typecheck
- npm run lint
- npm run build
- npm run bench
- npm run quality:fixtures
- npm run quality:attestations
- npm run quality:artifacts
- git diff --check
- affected diagnostics clean
- final architect sign-off
- final verifier sign-off
- ai-slop-cleaner/deslop pass on Ralph-changed files
- post-deslop regression verification green

Definition of done:
- all approved OOXML gap-closure requirements finished
- zero pending or in_progress TODO items
- zero known failing verifications
- repository clean
- Ralph state cleanly exited

When everything is genuinely complete:
- run /cancel
- then report only:
  1) completed work
  2) changed files
  3) verification evidence
  4) remaining risks, if any
```
