# Decision Record 001: Shared Core Architecture for the OOXML Frontend Library

## Status
Accepted during ralplan consensus planning.

## Context
The library must support three OOXML document families, preserve package fidelity, run browser-first, and expose parse/render/edit/serialize APIs without fragmenting into incompatible stacks.

## Decision
Adopt a shared-core monorepo architecture with common OPC/XML/IR/serializer/editor/render foundations and format-specific adapters for docx/xlsx/pptx.

## Drivers
- Shared OPC/XML mechanics.
- Round-trip-safe editing and serialization.
- Consistent verification, docs, examples, and devtools surfaces.

## Alternatives
1. Per-format isolated stacks.
2. Viewer-first delivery with later editing support.
3. Server-centric internals with browser wrappers.

## Why chosen
The shared-core model best aligns with the required complete frontend product surface and minimizes long-term duplication.

## Consequences
- More up-front design work.
- Clearer long-term package boundaries and more reusable verification infrastructure.
- Need to guard against over-abstraction by validating format-specific escape hatches.
