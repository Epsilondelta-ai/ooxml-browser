# PresentationML Reference (`.pptx`)

## Primary structure

Typical presentation tree:
- presentation: `/ppt/presentation.xml`
- slides: `/ppt/slides/slideN.xml`
- slide masters: `/ppt/slideMasters/slideMasterN.xml`
- slide layouts: `/ppt/slideLayouts/slideLayoutN.xml`
- theme: `/ppt/theme/themeN.xml`
- notes slides / notes master
- handout master
- comments
- media / embedded packages / charts / diagrams

The presentation root links slide master list, slide list, notes master list, and handout master.

## Inheritance model

Presentation formatting/inheritance flows across:
1. theme
2. slide master
3. slide layout
4. slide local properties
5. placeholder content overrides

Implementation requirement:
- preserve placeholder identity and master/layout binding so editing does not flatten slides into standalone drawings.

## Core slide model

Each slide contains a common slide payload with shapes such as:
- text shapes
- pictures
- graphic frames (charts/tables/diagrams)
- connectors
- groups
- background references/overrides
- transitions and timing data
- speaker notes association

## Text model

Text in PresentationML uses DrawingML text bodies:
- paragraphs with bullet levels and text styles
- runs with character properties
- placeholder/default styles inherited from master/layout
- autofit / vertical/horizontal alignment rules

## Notes, comments, handouts

- notes slide per slide when present
- notes master for notes-page styling
- comments anchored to slide geometry or logical targets
- handout master preserved even if not rendered in standard slide view

## Media and embedded objects

- images and videos
- audio
- charts / tables / diagrams
- embedded OLE/packages
- hyperlinks and actions

## Animation and timing

A full-featured library should preserve:
- timing trees
- transitions
- trigger/action metadata
- build order and effect parameters

Renderer decision:
- static slide fidelity is mandatory
- animation playback is a layered feature built on preserved timing IR, not on destructive conversion to CSS alone

## Slide rendering concerns

- slide viewport based on presentation size
- layered rendering: background -> master/layout placeholders -> local shapes -> overlay guides/comments/selection
- high-fidelity text box layout and auto-fit
- transform matrices for grouped shapes
- theme and color-map resolution
- notes view and slide sorter views as alternate projections

## Editing concerns

- shape-level selection, resize, rotate, reorder
- placeholder-aware text editing
- master/layout-safe overrides
- slide insertion/duplication/reordering/deletion
- notes/comments editing
- asset replacement without breaking relationship graph

## Serialization concerns

- preserve stable shape IDs where feasible
- retain master/layout/theme references instead of flattening formatting
- preserve timing trees, transitions, and unsupported media metadata unchanged if not edited

## Decisions

- **D-PPTX-1:** master/layout inheritance remains explicit in IR.
- **D-PPTX-2:** slide rendering uses SVG/HTML hybrid composition with a transform-aware scene graph.
- **D-PPTX-3:** animation/timing is preserved in the model even before every effect is executable in-browser.

## Risks / open issues

- Office text auto-fit and shape layout behavior may need empirically tuned compatibility rules
- SmartArt and some diagram constructs may require preserved fallback rendering before native editable support reaches parity
