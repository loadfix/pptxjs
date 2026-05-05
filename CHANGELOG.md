# Changelog

All notable changes to `pptx-preview` (loadfix/pptxjs) are tracked here.
The project follows [Semantic Versioning](https://semver.org/); until a
1.x release the minor bump signals user-visible feature additions while
the patch bump signals fixes.

## 0.1.0 — 2026-05-05

Wave 9 consolidation — folds the five post-Wave-4a feature branches
into master and takes the render-assertion conformance matrix from
190/201 to **201/201**. The sibling `feat/w[4-8]a*` branches on the
remote are legacy Wave 4a work and are intentionally left un-merged.

Merged branches (in merge order):

- **`fix/w1-f-conformance-gaps-2026-05-04`** — parser + renderer
  emit queryable `data-*` hooks so corpus selectors resolve:
  `data-placeholder-type` / `data-placeholder-idx` on placeholder
  shapes; `data-kind="sp"` on every shape; `data-shape-type="textbox"`
  when `<p:cNvSpPr txBox="1"/>`; `data-shape-preset` from
  `<a:prstGeom @prst>`; `data-kind="chart"` on chart graphic-frames;
  `<ul>` / `<ol>` coalescing of adjacent bulleted paragraphs with
  `data-bullet-type` on each `<li>`. Closes 10 of 11 corpus gaps.
- **`fix/w9-b-notes-transitions`** — notes-slide rendering
  (`<aside class="pptx-notes" data-notes>` adjacent to each slide,
  hidden behind a top-of-output "Show speaker notes" checkbox); slide
  transitions surfaced as `data-transition-preset="<effect>"` on the
  rendered section. Closes the 11th corpus gap (`slide-notes`).
  Result: 201/201 conformance.
- **`fix/w9-d-a11y-pptxjs`** — a11y landmarks and headings: each
  slide `<section>` gets `role="region"` + `aria-label="Slide N"`;
  `phType="title"` placeholders get `role="heading"` + `aria-level=1`;
  `phType="subTitle"` gets `aria-level=2`.
- **`fix/w9-e-responsive`** — new `responsive?: boolean` render
  option (default `false`). When `true`, each slide is wrapped in a
  `.pptx-slide-container` whose `aspect-ratio` reserves the slide's
  native shape and whose inner `<section>` is `transform: scale(...)`d
  to fit the container width via a `ResizeObserver`. Upscaling is
  capped at 1.0. Example in `examples/responsive.html`.
- **`fix/w9-f-shared-css`** — cross-format shared class nomenclature
  (`.oox-page`, `.oox-paragraph`, `.oox-run`, `.oox-table`,
  `.oox-table-row`, `.oox-table-cell`, `.oox-image`) emitted alongside
  the existing `.pptx-*` classes. Lets corpus manifests write a single
  selector that targets the same logical concept across docxjs / pptxjs
  / xlsxjs output. Existing `.pptx-*` classes are preserved for
  back-compat.

Tests: 205/205 Playwright pass (201 conformance + 4 render-test
specs including the two new responsive tests and one new shared-class
spec).
