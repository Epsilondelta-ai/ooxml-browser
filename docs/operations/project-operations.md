# Operational and Project Plan

## Documentation site

Requirements:
- architecture guides
- API reference
- format feature coverage matrix
- compatibility notes
- cookbook/how-to guides
- fixture and benchmarking documentation

Suggested implementation:
- static site generator in repo
- docs sourced from Markdown + generated API docs
- live playground embeds for examples

## Examples

Need curated examples for:
- open and inspect package graph
- render a paginated Word document
- render/edit spreadsheet grid
- render/edit presentation slides
- round-trip serialize after edits
- worker-based large-document parsing
- custom unsupported-part plugin

## Playground

The project should ship a browser playground that can:
- drag/drop `.docx`, `.xlsx`, `.pptx`
- inspect package parts/relationships
- switch between raw/package/render/edit views
- test serializer round-trip
- show diagnostics and fidelity warnings

## Benchmark suite

Need repeatable benchmark harness with:
- corpus selection
- browser automation
- CPU/memory/open/render/edit/save timing capture
- historical trend reporting

## Release strategy

- monorepo changesets/release tooling
- package-level semver
- prerelease channels for major feature waves
- release checklist includes corpus verification and browser matrix validation

## Compatibility policy

- document supported browsers/runtime targets
- explicitly version public APIs and serialized-IR contracts
- maintain feature coverage matrix by format and subsystem

## Migration policy

- provide migration notes when IR/public API changes
- keep serializer compatibility options explicit and versioned
- preserve fixture manifests across major versions and document expected diffs

## Implementation implications

- examples/playground/benchmarks should be in the repo from early implementation stages, not appended later
- docs should be updated in the same PR/commit series as feature changes where possible

## Decisions

- **D-OPS-1:** docs, examples, playground, and benchmark harness are product surfaces, not side projects.
- **D-OPS-2:** every significant capability should have a fixture + example + docs trail.
