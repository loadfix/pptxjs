# pptxjs — TODO

Tracked work for this fork. Move entries into the "Done" section below as they ship; link the PR / commit.

## Audit findings 2026-05-05

Captured from a project audit on 2026-05-05. None of these are blockers for 0.1.0; they are housekeeping / drift items.

1. **Triage and delete ~30 legacy remote branches.** All confirmed merged into 0.1.0 via wave-integrated branches, but they clutter the remote. Groups:
   - `feat/p{1..12}` (12 branches — preset-geom, rotation-flips, fills, effects, typography, image-adjust, table-style, chart-smartart-fallback, hyperlinks, fields, a11y, notes-comments). Merged via later waves — **delete**.
   - `feat/w4a{1..6}` (6), `feat/w5a{1..5}` (5), `feat/w6a{1..5}` (5), `feat/w7a{1..3}` (3), `feat/w8a{1..3}` (3). All integrated via the corresponding `wave{N}-integrated` branch — **delete**.
   - `feat/wave{1,2,4,5,6,7,8}-integrated` + `wave1-render-split` + `wave1-types` (9 branches). Merged into master via PRs #2 / #4 / #5 / #6 / #7 — **delete**.
   - `feat/w1-e-conformance-ci` — parked (per the removed-`.github/` memory note). **Leave** (explicit parked).
   - `feat/chart-render-v2`, `feat/emf-svg-conversion` — unclear provenance; `emf-svg-conversion` intersects an explicit non-goal. **Inspect briefly, likely delete.**
   - `fix/w1-f-conformance-gaps-2026-05-04`, `fix/w9-b/d/e/f-*`, `fix/dickinson-sample-bugs` (6 fix branches). All merged into 0.1.0 — **delete**.
2. **Close GitHub issue #8.** `slide-master-title-placeholder.pptx` produces no rendered slide. Filed 2026-05-05 via the W7-D visual-compare run; it is the only open bug. Likely cause: the layout title placeholder inherits prompt text from the master, and the parser trips on an inherited-only placeholder that has no direct text.
3. **Refresh README.** Three documentation drifts:
   - The Options list omits `renderHandouts` and `responsive` (both Wave-9 additions).
   - The Contributing section still says to run `npm run e2e (Karma + Chrome)` — there is no Karma; the suite is Playwright.
   - Drop the "Early scaffolding" status line now that 0.1.0 has shipped with full conformance.
4. **Bound `jszip` version.** Currently declared as `jszip >=3.0.0` (unbounded upper). Bound to `^3`.
5. **Clean up 43 `test-results/` / `playwright-report-interop/` local artefacts.** Gitignored, but worth a periodic sweep.

## Open

Remaining gaps after waves 1–5. Split into "possible follow-ups" and "deliberate non-goals" — items in the second group need a new conversation before starting, not just an agent.

### Possible follow-ups

- **Richer notes rendering.** Wave 9 added the basic `<aside class="pptx-notes">` with paragraphs parsed through the main pipeline (master text-style inheritance resolved against the slide's own master, not the notesMaster). Full fidelity would walk every shape on the notesSlide through the renderer pipeline and size the aside against the notes master's slide size (usually 6.5×10"), so things like page numbers and non-body placeholders render at the right positions.
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
- **Office Math (`<m:oMath>`) / equations.** MathML rendering is a separate problem space. Could be lowered to MathML in supporting browsers as a follow-up, but the conversion isn't trivial.
- **Full OLE object preview.** Embedded workbooks / docs / equations. Rendering the `<p:oleObj><p:pic>` cached image would be straightforward (piggyback on the picture pipeline); rendering the embedded payload itself is separate projects.
- **Linked (not embedded) media.** `r:link` instead of `r:embed`. Resolving external references requires caller-side fetching; out of scope for the package-only renderer.
- **Speaker narration audio / video playback.** The current option set (`renderNotes`, `renderComments`) doesn't extend to timed media.
- **3D shape bevels / `<a:sp3d>`.** CSS `perspective` / 3D transforms would need matrix composition from the OOXML 3D scene — feasible but large, and rare in real decks.

## Done

- EMF / WMF image rasterisation (Wave 9, A1, `feat/emf-svg-conversion`). EMF and WMF blips on slides / layouts / masters are now rasterised to PNG data URLs via the [`emf-converter`](https://www.npmjs.com/package/emf-converter) npm package, lazy-loaded via a dynamic `import()` so decks without any metafiles pay no bundle weight. New `Options.convertEmf` flag (default `true`) gates the feature; `setEmfConverter(mod)` is an escape hatch for UMD consumers who want to preload the converter from a script tag. Known limitations of the underlying library: no font fallback (text falls back to system defaults), bitmap recordings get bilinear-upscaled, and some rare GDI+ records (`EMR+DrawImagePoints`, text transforms) render approximately.
- Chart and SmartArt graphic-frame fallback rendering (`feat/p8-chart-smartart-fallback`). Charts emit their cached preview image when present, else a `[Chart]` placeholder at the frame position. SmartArt frames expand the sibling `diagrams/drawingN.xml` DrawingML cache through the standard shape pipeline (offset by the frame origin), falling back to a `[SmartArt]` placeholder when no drawing cache exists.
- Hyperlink rendering (P9, Wave 2): external URLs and intra-deck slide jumps on text runs, shapes, and images are wrapped in `<a>` anchors. External links use `target="_blank" rel="noopener noreferrer"`; intra-deck jumps (ppaction hlinkshowjump + slide-to-slide rels) become `#slide-N` fragments. URL scheme whitelist (http/https/mailto/tel + fragments/relatives) rejects `javascript:` and other unsafe targets.
- Graceful malformed-deck handling (Wave 4, A6, `feat/w4a6-errors`). Per-slide try/catch in `Presentation.load` so one bad slide can't poison the rest of the deck — the failed slide keeps its index slot with `Slide.parseError` set, and renders as a red banner inside a correctly-sized section so layout isn't broken. Per-shape try/catch in the `spTree` walkers drops mangled shapes in place with a warning. Missing slide parts are surfaced via the same banner path. New `Options.onSlideError?: (index, err) => void` callback lets hosts hook parse failures. `OpenXmlPackage.load` now prefixes JSZip failures with `[pptxjs]` so non-pptx inputs produce an identifiable rejection. `parseError` text is inserted via `textContent` only — document-derived strings can't inject HTML.
- Graceful malformed-deck handling (Wave 4, A6, `feat/w4a6-errors`). Per-slide try/catch in `Presentation.load`, `Slide.parseError` + red banner rendering, per-shape try/catch in `spTree` walkers, `Options.onSlideError` callback, and a `[pptxjs]` prefix on JSZip failures in `OpenXmlPackage.load`. `parseError` text inserted via `textContent` only.
- Notes-slide rendering (Wave 9, `fix/w9-b-notes-transitions`). Each slide's companion `ppt/notesSlides/notesSlide*.xml` is parsed through the main paragraph pipeline and emitted as an `<aside class="pptx-notes" data-notes>` adjacent to the slide section. Hidden by default via CSS; a `"Show speaker notes"` checkbox at the top of the rendered output reveals every aside using `:has(:checked)`.
- Slide-transition queryability (Wave 9, `fix/w9-b-notes-transitions`). Each slide `<section>` now carries `data-transition-preset="<effect>"` derived from the first direct child of `<p:transition>` (fade/push/wipe/split/reveal/randomBar/zoom/cut/flythrough), so downstream `data-*` selectors resolve.
- Conformance-hook fill-in for 11 gaps surfaced by the 2026-05-04 corpus run (Wave 9, `fix/w9-b-notes-transitions`). Addresses every item under "Conformance gaps" below: `data-placeholder-type` / `data-placeholder-idx` on placeholder shapes, `data-kind="sp"` on non-placeholder shapes (with `data-shape-type="textbox"` when `<p:cNvSpPr txBox="1"/>`), `data-shape-preset` from `<a:prstGeom @prst>`, `data-kind="chart"` on chart graphic-frames, `<ul>` / `<ol>` coalescing of adjacent bulleted paragraphs with `data-bullet-type` on each `<li>`, shape-level `color` propagation from uniform-colour runs, ARIA `role="region"` + `aria-label` on slide sections, and `role="heading"` + `id` on title / subTitle placeholders. Result: 201/201 render assertions pass, up from 190/201.

---

## Conformance gaps (auto-filed from corpus 2026-05-04 overnight run)

The 11 gaps against `c09111d` that this section tracked are all
resolved in `fix/w9-b-notes-transitions` (see the last three "Done"
entries above). A fresh corpus run may surface new gaps — file them
here as a new section.
