# Target hotspots: PPT 98-percent fidelity

## Baseline scores
- sample1 slide1: 85.20
- sample5 slide2: 81.14
- sample6 slide1: 89.88

## sample1 slide1
1. Mechanical arm / gear cluster geometry and alignment — likely geometry + transform semantics — owner: `packages/pptx/src/parser.ts`, `packages/render/**` — planned stage: 4
2. Gradient text/logo composition — likely text/layout + fill semantics — owner: `packages/pptx/src/parser.ts`, `packages/render/**` — planned stage: 3/5
3. Decorative skyline/background layering — likely renderer scene ordering — owner: `packages/render/**`, `examples/basic/**` — planned stage: 2/6

## sample5 slide2
1. Rocket + ring silhouette fidelity — likely custom/preset geometry — owner: `packages/pptx/src/parser.ts`, `packages/render/**` — planned stage: 4
2. Agenda block borders/strokes — likely line/stroke semantics — owner: `packages/pptx/src/parser.ts`, `packages/render/**` — planned stage: 5
3. Text box spacing and title alignment — likely inheritance + text layout — owner: `packages/pptx/src/parser.ts`, `packages/render/**` — planned stage: 3

## sample6 slide1
1. Circular text layout and title wrapping — likely placeholder/layout/master + text engine — owner: `packages/pptx/src/parser.ts`, `packages/render/**` — planned stage: 3
2. Dashed badge ring fidelity — likely stroke dash/opacity semantics — owner: `packages/pptx/src/parser.ts`, `packages/render/**` — planned stage: 5
3. Background splatter layering and logo placement — likely scene ordering and image composition — owner: `packages/render/**`, `examples/basic/**` — planned stage: 2/6
