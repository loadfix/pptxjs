import type { Presentation } from '../presentation';
import {
	parseStaticShapes,
	emptyPlaceholderMap,
	SlideParseContext,
} from '../presentation-parser';
import { emuToPx } from './geom';
import { renderShapeLike } from './dispatch';

// Emit the handout-master "print layout" pages — one `<section class=
// "pptx-handout-page">` after the slide list when `Options.renderHandouts`
// is on. Each page renders:
//   - The handoutMaster's static chrome (header/footer/date/slideNum
//     placeholders plus any non-placeholder decoration) at full size, so
//     the author-supplied template frames its page.
//   - A 3-row × 2-column grid of the deck's slides as scaled-down clones
//     of the already-rendered slide sections. The thumbnails are laid out
//     via CSS grid rather than absolute EMU positions because OOXML
//     doesn't actually specify where the six thumbnails go on the page —
//     that's a PowerPoint print-time concern, not a handoutMaster feature.
//     The handoutMaster carries header/footer placeholder slots only; the
//     thumbnail grid is synthesised by the consumer (PowerPoint at print
//     time, us at render time).
//
// The handout page is sized from `presentation.notesSize` because OOXML
// uses the same `<p:notesSz>` element for both notes and handout canvas
// dimensions. Typical value is 7.5"×10" portrait (Letter-adjacent).
//
// Each page packs `HANDOUT_SLIDES_PER_PAGE` slides (6 = 3×2); a deck with
// N slides produces ceil(N / 6) pages, each repeating the master chrome.
const HANDOUT_ROWS = 3;
const HANDOUT_COLS = 2;
const HANDOUT_SLIDES_PER_PAGE = HANDOUT_ROWS * HANDOUT_COLS;

export function renderHandoutPages(
	presentation: Presentation,
	renderedSlides: HTMLElement[],
	cls: string,
): HTMLElement[] {
	const master = presentation.handoutMaster;
	if (!master || !master.doc) return [];

	const pageW = emuToPx(presentation.notesSize.cx);
	const pageH = emuToPx(presentation.notesSize.cy);

	// Parse the handoutMaster's static shapes on demand. We pass an empty
	// embedUrls map — most handout masters carry only text placeholders
	// (date/header/footer/slideNum) with no images. If the master *does*
	// reference images, those blips won't resolve here; that's acceptable
	// for a first-pass implementation since the PowerPoint default handout
	// master has no embedded graphics.
	const ctx: SlideParseContext = {
		layout: emptyPlaceholderMap(),
		master: master.placeholders,
		masterTextStyles: master.textStyles,
		theme: master.theme,
		clrMap: master.clrMap,
		embedUrls: new Map(),
	};
	const chromeShapes = parseStaticShapes(master.doc, ctx);

	const pages: HTMLElement[] = [];
	const slideCount = renderedSlides.length;
	const pageCount = Math.max(1, Math.ceil(slideCount / HANDOUT_SLIDES_PER_PAGE));

	// Field substitution on the handout chrome: when the master carries
	// sldNum/ftr/hdr/dt placeholders, their fields resolve against a
	// synthetic "first slide on the page" so that e.g. the page number
	// shown in the footer advances per handout page. We pick the first
	// slide of the page as the field source; if the deck has zero slides
	// we emit no handouts at all (nothing to tile).
	if (slideCount === 0) return [];

	for (let pageIdx = 0; pageIdx < pageCount; pageIdx++) {
		const startSlide = pageIdx * HANDOUT_SLIDES_PER_PAGE;
		const pageSlides = renderedSlides.slice(startSlide, startSlide + HANDOUT_SLIDES_PER_PAGE);
		const firstSlide = presentation.slides[startSlide];

		const section = document.createElement("section");
		section.className = `${cls}-handout-page`;
		section.dataset.handoutPage = String(pageIdx);
		section.style.width = `${pageW}px`;
		section.style.height = `${pageH}px`;

		// Chrome first (behind), thumbnails on top.
		if (firstSlide) {
			const fieldCtx = { slide: firstSlide, firstSlideNum: presentation.firstSlideNum + startSlide };
			for (const shape of chromeShapes) {
				const node = renderShapeLike(
					shape,
					cls,
					presentation.embedUrls,
					presentation.tableStyles,
					// Handout chrome has no hyperlinks of its own.
					new Map(),
					fieldCtx,
				);
				if (node) section.appendChild(node);
			}
		}

		// Thumbnail grid overlay. Positioned inside the page with CSS
		// grid; the `.pptx-handout-page` stylesheet sizes each cell and
		// the cloned slide within it is scaled to fit.
		const grid = document.createElement("div");
		grid.className = `${cls}-handout-grid`;
		for (const slideEl of pageSlides) {
			const cell = document.createElement("div");
			cell.className = `${cls}-handout-cell`;
			// Deep-clone the already-rendered slide section. The clone
			// keeps the slide's inline sizes (matching the deck's sldSz)
			// and absolutely-positioned shapes; the outer .pptx-handout-
			// -thumb wrapper applies a CSS `transform: scale(...)` via
			// the stylesheet so the full-size markup shrinks into the
			// grid cell.
			const clone = slideEl.cloneNode(true) as HTMLElement;
			// Drop the slide's own bottom margin so thumbnails sit flush
			// against the grid cell. The cloned classes still inherit
			// the main slide rule's box styling otherwise.
			clone.style.margin = "0";
			// Make absolute positioning for inner shapes relative to the
			// cloned slide (it already is via the .pptx-slide rule, but
			// cloned nodes lose their id and we already have position:
			// relative from the base rule — leave it alone).
			const thumb = document.createElement("div");
			thumb.className = `${cls}-handout-thumb`;
			thumb.appendChild(clone);
			cell.appendChild(thumb);
			grid.appendChild(cell);
		}
		section.appendChild(grid);

		pages.push(section);
	}

	return pages;
}
