You are an autonomous Codex/OMX coding agent operating in /Users/juunini/Desktop/code/epsilondelta/ooxml.

Final objective:
Build a complete frontend library that can parse, render, and edit OOXML-based Microsoft Office documents (.docx, .xlsx, .pptx).
This is not an MVP effort. The goal is a fully capable, production-meaningful library.

Core rules:
- Do not warn about difficulty, timeline, scope, or feasibility.
- Do not suggest reducing scope to an MVP.
- Do not stop when the next step is clear.
- Follow AGENTS.md, OMX workflow rules, and skill contracts.
- Proceed through documentation -> planning -> implementation -> verification -> fixes until the work is complete.
- Delegate only when it materially improves throughput or correctness.
- Do not implement before planning.
- Deliver documentation, code, tests, examples, verification, and release-ready project structure.

You must execute in the following order.

==================================================
1) First perform research and documentation
==================================================

Before implementation, research and document everything required to build a frontend OOXML library. Save the documentation as durable project artifacts.

The documentation must cover at minimum:

- OOXML / OPC packaging
  - ZIP package structure
  - Parts
  - Relationships
  - Content Types
  - package traversal strategy

- Format-specific structure
  - WordprocessingML (.docx)
  - SpreadsheetML (.xlsx)
  - PresentationML (.pptx)

- Shared subsystems
  - styles
  - themes
  - fonts
  - colors
  - numbering
  - tables
  - drawings
  - images
  - charts
  - equations
  - comments
  - tracked changes
  - headers/footers
  - sections
  - notes
  - metadata
  - hyperlinks
  - annotations
  - embedded objects

- Parsing concerns
  - XML parsing strategy
  - namespace handling
  - relationship resolution
  - shared model normalization
  - validation
  - recovery / fault tolerance
  - round-trip-friendly IR / AST design
  - incremental parsing possibilities
  - large-document parsing strategy

- Rendering concerns
  - HTML / CSS / SVG / Canvas rendering strategy
  - layout abstraction
  - pagination / section / page / slide / sheet viewport model
  - text layout, table layout, drawing layout
  - spreadsheet grid rendering
  - presentation slide rendering
  - print/export support

- Editing concerns
  - editor data model
  - selection / range / cursor model
  - mutation API
  - transaction model
  - undo / redo
  - clipboard / paste
  - structural editing
  - collaborative editing extensibility
  - conflict resolution possibilities
  - serialization back to OOXML

- Compatibility and quality targets
  - Microsoft Office compatibility
  - LibreOffice and alternative-suite compatibility
  - fidelity targets
  - round-trip preservation targets
  - importer/exporter tolerance policy

- Performance
  - streaming
  - lazy loading
  - virtualization
  - chunked rendering
  - worker offloading
  - memory constraints
  - benchmark strategy

- Internationalization and accessibility
  - RTL
  - CJK
  - line breaking
  - font fallback
  - locale-sensitive formatting
  - accessibility tree
  - keyboard navigation
  - screen reader support

- Security
  - zip bombs
  - malformed XML
  - external resource handling
  - macro / unsupported object policy
  - sanitizer boundaries
  - untrusted document policy

- API and product architecture
  - public API design
  - package / module boundaries
  - plugin / extension system
  - framework adapter possibilities
  - browser-first architecture
  - worker interfaces
  - testing hooks
  - devtools / debug utilities

- Verification
  - unit / integration / e2e / golden / fidelity / round-trip / performance test strategy
  - corpus strategy
  - fixture strategy
  - interoperability matrix
  - regression detection strategy

- Operational/project concerns
  - documentation site
  - examples
  - playground
  - benchmark suite
  - release strategy
  - semver / compatibility policy
  - migration policy

Documentation storage rules:
- Save all documentation under docs/.
- Organize it into topical files and subdirectories.
- Create an index document plus detailed design/reference documents.
- Write the docs so they can serve as the implementation baseline.
- Whenever useful, include: decisions, rationale, unresolved risks, and implementation implications.

Completion criteria for this stage:
- docs/ contains enough reference material to drive immediate planning and implementation.
- The documentation is detailed enough for ralplan to use directly.

After this stage, you must commit.

Commit rules:
- Commit after every meaningful stage.
- Do not wait and make a single giant commit at the end.
- Minimum commit granularity examples:
  1. Initial documentation structure and research docs
  2. ralplan artifacts
  3. architecture / scaffolding
  4. parser modules
  5. renderer modules
  6. editing engine
  7. tests / examples / docs reinforcement
  8. verification and final fixes
- Every commit message must follow the Lore Commit Protocol from AGENTS.md.
- Include an intent line, narrative body, and appropriate trailers.
- After each commit, confirm working tree state and continue.

==================================================
2) After documentation, use $ralplan to create the full execution plan
==================================================

Once the documentation baseline exists, you must use $ralplan.

Goal:
Produce a consensus-backed execution plan for a complete frontend OOXML library, not an MVP.

ralplan requirements:
- Use deliberate planning depth when appropriate.
- Produce an execution-ready plan that reflects Planner, Architect, and Critic viewpoints.
- Persist PRD and test-spec artifacts under .omx/plans/.
- The plan must include at minimum:
  - overall product goals
  - quality targets
  - architecture
  - module boundaries
  - internal IR / AST strategy
  - parser / renderer / editor / serializer design
  - format-specific implementation strategy for docx / xlsx / pptx
  - testing strategy
  - corpus / fixture strategy
  - benchmark strategy
  - docs / examples / playground strategy
  - staged implementation order
  - parallelizable task decomposition
  - verification criteria
  - definition of done

The ralplan output must be concrete enough for ralph to execute directly.

After this stage, you must commit:
- .omx/plans/ artifacts
- planning documents / decision records
- any related documentation updates

Commit messages must follow the Lore Commit Protocol.

==================================================
3) When planning is ready, use $ralph to execute until completion
==================================================

After the ralplan artifacts are ready, you must use $ralph and carry the plan all the way through completion.

ralph execution requirements:
- Do not stop midway.
- Use delegation, parallelization, and repeated verification where helpful.
- Iterate through implementation, testing, fixes, regression prevention, and documentation updates.
- Do not declare partial completion.
- Do not reduce scope.
- Collect fresh verification evidence at each meaningful checkpoint.
- Treat round-trip fidelity as a primary success metric whenever applicable.
- Deliver a genuinely usable library state that includes parser, renderer, editor, serializer, tests, examples, docs, and benchmarks.
- Follow AGENTS.md requirements for verification, architect review, deslop / re-verification, and cleanup.

Commit rules during ralph:
- Commit after every meaningful implementation stage.
- Keep commits small, logical, and reviewable.
- Each commit should represent a working intermediate state plus the tests / verification relevant to that step.
- Do not accumulate arbitrarily broken states.
- Every commit message must follow the Lore Commit Protocol.
- Before and after commits, confirm relevant verification status.
- Update documentation alongside implementation when needed.

==================================================
4) Implementation principles
==================================================

- Design browser-first.
- Expose APIs that are truly usable in frontend environments.
- Keep the internal model round-trip-friendly.
- Ensure parser, renderer, editor, and serializer share a coherent data model.
- Add new dependencies only when truly necessary, and document why.
- Reuse existing patterns, utilities, and module boundaries when possible.
- Prefer verifiable architecture over decorative abstraction.
- Do not stop at documentation; continue into implementation.
- Do not avoid tests.
- Create examples and a playground when needed.
- Keep docs, examples, and benchmarks meaningfully maintained.

==================================================
5) Verification and completion criteria
==================================================

Before declaring completion, you must verify.

Minimum verification:
- relevant tests pass
- build succeeds
- lint / typecheck / static diagnostics pass
- affected-file diagnostics are clean
- parser / renderer / editor / serializer flows are validated
- core round-trip fidelity scenarios are validated
- examples / playground behavior is checked
- docs match implementation
- final architect / verifier style review is performed

Completion conditions:
- planned scope has actually been executed
- docs, code, tests, examples, and benchmarks are consistent
- no unresolved TODOs remain
- no known broken state remains
- verification evidence exists

Final report format:
- paths of created / updated docs
- paths of created / updated code
- major implemented capabilities
- verification performed
- remaining risks, if any, stated concretely
- stage-by-stage commit summary
- short summary of key decisions at each stage

==================================================
6) Start now
==================================================

Begin immediately.

Exact execution order:
1. Perform research and documentation
2. Save artifacts under docs/
3. Commit after stage 1
4. Run $ralplan and create .omx/plans/ artifacts
5. Commit after stage 2
6. Run $ralph
7. Continue implementation with meaningful stage-by-stage commits
8. Run final verification
9. Produce the final report

Important:
- You must commit after each completed stage.
- A single final commit is not allowed.
- Every commit message must follow the Lore Commit Protocol.
- Do not stop before the work is carried through to completion.
