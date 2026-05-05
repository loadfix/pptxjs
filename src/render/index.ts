import type { Options } from '../pptx-preview';
import type { Presentation } from '../presentation';
import { emuToPx } from './geom';
import { makeStyleNode } from './style';
import { applyBackground } from './background';
import { renderShapeLike } from './dispatch';
import { renderNotesBlock, renderCommentMarkers } from './notes';

export class HtmlRenderer {
	async render(presentation: Presentation, options: Options): Promise<Node[]> {
		const out: Node[] = [];

		const slideW = emuToPx(presentation.slideSize.cx);
		const slideH = emuToPx(presentation.slideSize.cy);
		out.push(makeStyleNode(options.className, slideW, slideH));

		const embedUrls = presentation.embedUrls;
		for (const slide of presentation.slides) {
			const section = document.createElement("section");
			section.className = `${options.className}-slide`;
			section.dataset.slideIndex = String(slide.index);
			// Intra-deck hyperlinks (`#slide-N`) use this id as their target.
			section.id = `slide-${slide.index}`;
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
				out.push(section);
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
			out.push(section);
			if (options.renderNotes) {
				const notesEl = renderNotesBlock(slide, options.className, presentation, embedUrls);
				if (notesEl) out.push(notesEl);
			}
		}

		return out;
	}
}
