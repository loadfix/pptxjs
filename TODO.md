# pptxjs — TODO

Tracked work for this fork. Move entries into the "Done" section below as they ship; link the PR / commit.

## Open

Remaining gaps after waves 1–5. Split into "possible follow-ups" and "deliberate non-goals" — items in the second group need a new conversation before starting, not just an agent.

### Possible follow-ups

- **Richer notes rendering.** `renderNotesBlock` applies the notesMaster's level-1 `defRPr` to a flat `<div>`. Full fidelity would walk the notes slide's shapes and paragraphs through the main renderer pipeline, including body placeholders sized against the notes master's slide size (usually 6.5×10").
- **Handout master rendering.** Loaded by `src/presentation.ts` but never rendered — belongs in a print-handouts feature, not the default renderer.
- **Gradient / pattern / blip strokes on CSS paths.** Wave 5 made SVG strokes paint with full fidelity via `fillToSvgPaint`. CSS `border:` emissions (degenerate lines + plain rectangles without `presetGeom`/`custGeom`) still fall back to an approximated solid because CSS can't take `url(#id)` for a border. Rendering these shapes as inline SVG even when they have no path geometry would close the gap — but changes the emitted DOM for every shape.
- **Tile offset/scale on blipFill tiles.** Tile `tx/ty/sx/sy/flip/algn` currently approximate via CSS `background-position` percentages. True tile positioning requires probing the image's natural pixel size, which needs an `Image` load → `await decode()` dance.
- **Diagonal cell borders inside merged table cells.** Diagonals render per-cell; they don't stretch across merged (`hMerge`/`vMerge`-hidden) cells. Fixable with a table-wide SVG overlay.
- **SmartArt with no drawing cache.** When `diagrams/drawingN.xml` is absent the frame falls back to a `[SmartArt]` placeholder. A proper layout engine that reads `data1.xml` + `layout1.xml` would produce correct visuals.
- **`softEdge` alpha mask.** Currently approximated as `filter: blur()` (blurs content too). Proper implementation needs an SVG `feGaussianBlur` + `feComposite` alpha mask sandwich.
- **`duotone` via `feColorMatrix`.** Currently approximated via a `mix-blend-mode: multiply` overlay. Proper implementation: inline SVG filter mapping luminance to a two-color ramp.
- **biLevel image filter.** Currently approximated via `grayscale(1) brightness() contrast(1000)`. Proper: SVG `feComponentTransfer` with a discrete threshold table.
- **Arrow heads with `none` `type` attribute** — already handled. `stealth` / `diamond` / `oval` markers scale by `w`/`len`; if a deck relies on exact pixel-width arrow sizes they may look slightly off.
- **custGeom `darken` / `lighten` path fill modes.** Currently fall through to `norm`. Proper: apply a CSS `filter: brightness()` to the SVG path based on shape fill luminance.
- **`<p14:sectionLst>` UI affordances.** Sections are parsed into `Presentation.sections` but the demo doesn't render them. A host could offer a jump-to-section drawer.
- **Per-master tableStyles.** `pres.tableStyles` is a single global map. Multi-master decks with different `tableStyles.xml` contents per master aren't supported; this is an OOXML-structural limitation (tableStyles lives at the presentation level, not per-master).

### Deliberate non-goals

Out of scope for pptxjs in its current form. Don't start these without a design conversation first — each would significantly change the library's shape.

- **Animations & transitions** (`<p:timing>`, entrance/exit effects, build order). pptxjs renders a static view; animation playback is its own project.
- **Encrypted / password-protected decks.** Requires the OOXML encryption spec + a crypto path in `OpenXmlPackage`. The Python-side equivalent (in python-pptx) is also out of scope. Users should decrypt externally and re-save.
- **EMF / WMF image rasterization.** Browsers don't render these natively. Rasterizing would require a JS parser (~big) or a server-side round-trip. Currently `mime.ts` returns `application/octet-stream` so the image slot stays empty.
- **Office Math (`<m:oMath>`) / equations.** MathML rendering is a separate problem space. Could be lowered to MathML in supporting browsers as a follow-up, but the conversion isn't trivial.
- **Full OLE object preview.** Embedded workbooks / docs / equations. Rendering the `<p:oleObj><p:pic>` cached image would be straightforward (piggyback on the picture pipeline); rendering the embedded payload itself is separate projects.
- **Linked (not embedded) media.** `r:link` instead of `r:embed`. Resolving external references requires caller-side fetching; out of scope for the package-only renderer.
- **Speaker narration audio / video playback.** The current option set (`renderNotes`, `renderComments`) doesn't extend to timed media.
- **3D shape bevels / `<a:sp3d>`.** CSS `perspective` / 3D transforms would need matrix composition from the OOXML 3D scene — feasible but large, and rare in real decks.

## Done

- Chart and SmartArt graphic-frame fallback rendering (`feat/p8-chart-smartart-fallback`). Charts emit their cached preview image when present, else a `[Chart]` placeholder at the frame position. SmartArt frames expand the sibling `diagrams/drawingN.xml` DrawingML cache through the standard shape pipeline (offset by the frame origin), falling back to a `[SmartArt]` placeholder when no drawing cache exists.
- Hyperlink rendering (P9, Wave 2): external URLs and intra-deck slide jumps on text runs, shapes, and images are wrapped in `<a>` anchors. External links use `target="_blank" rel="noopener noreferrer"`; intra-deck jumps (ppaction hlinkshowjump + slide-to-slide rels) become `#slide-N` fragments. URL scheme whitelist (http/https/mailto/tel + fragments/relatives) rejects `javascript:` and other unsafe targets.
- Graceful malformed-deck handling (Wave 4, A6, `feat/w4a6-errors`). Per-slide try/catch in `Presentation.load` so one bad slide can't poison the rest of the deck — the failed slide keeps its index slot with `Slide.parseError` set, and renders as a red banner inside a correctly-sized section so layout isn't broken. Per-shape try/catch in the `spTree` walkers drops mangled shapes in place with a warning. Missing slide parts are surfaced via the same banner path. New `Options.onSlideError?: (index, err) => void` callback lets hosts hook parse failures. `OpenXmlPackage.load` now prefixes JSZip failures with `[pptxjs]` so non-pptx inputs produce an identifiable rejection. `parseError` text is inserted via `textContent` only — document-derived strings can't inject HTML.
- Graceful malformed-deck handling (Wave 4, A6, `feat/w4a6-errors`). Per-slide try/catch in `Presentation.load`, `Slide.parseError` + red banner rendering, per-shape try/catch in `spTree` walkers, `Options.onSlideError` callback, and a `[pptxjs]` prefix on JSZip failures in `OpenXmlPackage.load`. `parseError` text inserted via `textContent` only.

---

## Conformance gaps (auto-filed from corpus 2026-05-04 overnight run)

The 950-case OOXML conformance corpus run
(`loadfix/ooxml-validate` → `conformance/results/pptxjs/`) surfaced 11
rendering gaps against the pptxjs fork at `c09111d`. Grouped below by
root cause. Each bullet links to the result JSON on GitHub and ends
with an actionable fix hypothesis.

- **Body placeholders do not emit `<ul>` / `<ol>` for bulleted runs.**
  [bullet-list](https://github.com/loadfix/ooxml-validate/blob/master/conformance/results/pptxjs/pptx/bullet-list.json)
  and [bullet-numbered](https://github.com/loadfix/ooxml-validate/blob/master/conformance/results/pptxjs/pptx/bullet-numbered.json)
  each fail four assertions: `*-three-items` expected 3 list items and
  got 0, plus `item-one`/`item-two`/`item-three` selectors for
  `ul li` / `ol li` / `[data-bullet-type='numbered']` / `[data-placeholder-type='body'] li`
  all return empty. Root cause: `renderTextBody` paints each `a:p` as an
  absolutely-positioned `<div>` with a glyph character prefix, never
  grouping consecutive bulleted paragraphs into a real list container.
  Fix: in `renderTextBody`, coalesce adjacent `a:p` with
  `buChar`/`buAutoNum` into a `<ul>` (or `<ol>` when `buAutoNum` is set)
  and emit a `data-bullet-type="bullet|numbered"` attribute on the
  wrapping list. Handles both fixtures with one change.
- **Placeholder-aware selectors find no `[data-placeholder-type]` hooks.**
  [slide-title](https://github.com/loadfix/ooxml-validate/blob/master/conformance/results/pptxjs/pptx/slide-title.json)
  (selector `[data-placeholder-type='ctrTitle'], [data-placeholder-type='title'], .pptx-title, h1, h2`)
  and [slide-with-subtitle](https://github.com/loadfix/ooxml-validate/blob/master/conformance/results/pptxjs/pptx/slide-with-subtitle.json)
  (selector `[data-placeholder-type='subTitle'], .pptx-subtitle, h2`) both
  fail with "No nodes matched ...; cannot check text". `Shape.ph` holds
  the placeholder type but the renderer never writes it to the DOM.
  Fix: in `renderShape`, when `sp.nvSpPr.nvPr.ph` is present, emit
  `data-placeholder-type="<type>"` (and `data-placeholder-idx` if set)
  on the shape's wrapping element — a 2-line change that unblocks five
  downstream conformance cases.
- **Textbox shapes emit no queryable selector.**
  [textbox-plain](https://github.com/loadfix/ooxml-validate/blob/master/conformance/results/pptxjs/pptx/textbox-plain.json)
  and [text-color](https://github.com/loadfix/ooxml-validate/blob/master/conformance/results/pptxjs/pptx/text-color.json)
  both fail on selector
  `[data-shape-type='textbox'], .pptx-textbox, [data-kind='sp']`. Pptxjs
  today paints textboxes as generic shape `<div>`s without marking them
  as textbox-kind. Fix: in `renderShape` set
  `data-shape-type="textbox"` when `sp.nvSpPr.nvSpPr.txBox="1"` and
  unconditionally set `data-kind="sp"` on every shape root element —
  allows downstream consumers to distinguish textboxes from preset
  geometries and lets the conformance harness read the computed colour
  off the correct node for `text-color-computed-red`.
- **Preset shapes lack `data-shape-preset` hooks.**
  [shape-line](https://github.com/loadfix/ooxml-validate/blob/master/conformance/results/pptxjs/pptx/shape-line.json)
  fails `hr, svg line, [data-shape='line'], [data-shape-preset='line']`,
  [shape-oval](https://github.com/loadfix/ooxml-validate/blob/master/conformance/results/pptxjs/pptx/shape-oval.json)
  fails `[data-shape-preset='ellipse'], [data-shape-preset='oval']`, and
  [shape-rectangle](https://github.com/loadfix/ooxml-validate/blob/master/conformance/results/pptxjs/pptx/shape-rectangle.json)
  expects `[data-shape-preset='rect']`. Preset geometry is known inside
  the shape renderer but isn't surfaced. Fix: emit
  `data-shape-preset="<prst>"` on each shape root from the
  already-parsed `a:prstGeom@prst` value — single line in `renderShape`,
  closes all three cases plus future preset-specific conformance tests.
- **Chart graphic frames render without a chart-kind hook.**
  [chart-column](https://github.com/loadfix/ooxml-validate/blob/master/conformance/results/pptxjs/pptx/chart-column.json)
  fails `.chart, [data-chart-type='column'], [data-chart-type='columnClustered'], [data-kind='chart'], svg.chart, .pptx-chart`.
  The cached-preview fallback renderer (see "Chart and SmartArt graphic-
  frame fallback rendering" in the Done list above) inserts an `<img>`
  or `[Chart]` placeholder but tags neither. Fix: set
  `data-kind="chart"` and `data-chart-type="<kind>"` (resolved from
  `c:barChart/@grouping="clustered"` etc.) on the graphic-frame wrapper
  regardless of whether the cached image is present.
- **Notes slides emit a flat `<div>` with no `.pptx-notes` hook.**
  [slide-notes](https://github.com/loadfix/ooxml-validate/blob/master/conformance/results/pptxjs/pptx/slide-notes.json)
  fails `.pptx-notes, [data-notes]`. `renderNotesBlock` (see the
  "Richer notes rendering" entry above) currently produces a bare
  `<div>`. Fix: give the notes block a stable class `pptx-notes`
  and/or `data-notes="true"` attribute — a trivial change independent
  of the richer-rendering follow-up already tracked.
