import type { Options } from '../pptx-preview';
import type { Presentation } from '../presentation';
import { emuToPx } from './geom';
import { makeStyleNode } from './style';
import { applyBackground } from './background';
import { renderShapeLike } from './dispatch';
import { renderNotesBlock, renderCommentMarkers } from './notes';
import { renderHandoutPages } from './handout';
import { wrapResponsive } from './responsive';

export class HtmlRenderer {
	async render(presentation: Presentation, options: Options): Promise<Node[]> {
		const out: Node[] = [];
		// Collect the rendered slide sections so the handout pass can clone
		// them into its thumbnail grid. We still push them into `out` in the
		// main flow — the handout tiles are *copies*, not replacements.
		const slideSections: HTMLElement[] = [];

		const slideW = emuToPx(presentation.slideSize.cx);
		const slideH = emuToPx(presentation.slideSize.cy);
		out.push(makeStyleNode(options.className, slideW, slideH, options.responsive === true));

		const embedUrls = presentation.embedUrls;
		for (const slide of presentation.slides) {
			const section = document.createElement("section");
			section.className = `${options.className}-slide`;
			section.dataset.slideIndex = String(slide.index);
			// Intra-deck hyperlinks (`#slide-N`) use this id as their target.
			section.id = `slide-${slide.index}`;
			// Accessibility landmark: each slide is a region so screen-reader
			// users can skim the deck. The aria-label is 1-based (matching
			// how PowerPoint displays slide numbers) and stays ASCII — the
			// slide's own title, when present, is authored content we can't
			// sanitise here, so we use a synthetic name to avoid attribute
			// injection. A title-shape carrying role="heading" below
			// provides the richer per-slide heading outline for AT.
			section.setAttribute("role", "region");
			section.setAttribute("aria-label", `Slide ${slide.index + 1}`);
			// Slide-level parse failure: skip the normal shape pipeline and
			// emit a visible error banner so the reader sees which slide
			// failed. The section is still sized by the style block so
			// surrounding layout isn't disturbed.
			if (slide.parseError !== null) {
				const banner = document.createElement("div");
				banner.setAttribute(
					"style",
					"padding: 1rem; background: #fee; border: 1px solid #c00; color: #c00;",
				);
				const strong = document.createElement("strong");
				// textContent (not innerHTML) — parseError comes from
				// user-supplied PPTX content and must never be interpreted as
				// HTML. The " " and the text after it are appended as
				// separate text nodes for the same reason.
				strong.textContent = `Slide ${slide.index + 1} could not be rendered:`;
				banner.appendChild(strong);
				banner.appendChild(document.createTextNode(" "));
				banner.appendChild(document.createTextNode(slide.parseError));
				section.appendChild(banner);
				slideSections.push(section);
				if (options.responsive) {
					out.push(wrapResponsive(section, options.className, slideW, slideH));
				} else {
					out.push(section);
				}
				continue;
			}
			applyBackground(section, slide);
			const fieldCtx = { slide, firstSlideNum: presentation.firstSlideNum };
			for (const shape of slide.shapes) {
				const node = renderShapeLike(shape, options.className, embedUrls, presentation.tableStyles, slide.hyperlinkUrls, fieldCtx);
				if (node) section.appendChild(node);
			}
			if (options.renderComments) {
				const markers = renderCommentMarkers(slide, options.className);
				if (markers) section.appendChild(markers);
			}
			// Responsive mode: wrap each slide in a fluid container so the
			// slide scales down on narrow viewports. The container reserves
			// height via aspect-ratio; the <section> inside is transformed
			// by attachResponsive after mount — pure CSS can't know the
			// container width.
			slideSections.push(section);
			if (options.responsive) {
				out.push(wrapResponsive(section, options.className, slideW, slideH));
			} else {
				out.push(section);
			}
			if (options.renderNotes) {
				const notesEl = renderNotesBlock(slide, options.className, presentation, embedUrls);
				if (notesEl) out.push(notesEl);
			}
		}

		if (options.renderHandouts) {
			const pages = renderHandoutPages(presentation, slideSections, options.className);
			for (const p of pages) out.push(p);
		}

		return out;
	}
}
