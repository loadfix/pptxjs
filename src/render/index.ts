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

		for (const slide of presentation.slides) {
			const section = document.createElement("section");
			section.className = `${options.className}-slide`;
			section.dataset.slideIndex = String(slide.index);
			applyBackground(section, slide);
			for (const shape of slide.shapes) {
				section.appendChild(renderShapeLike(shape, options.className));
			}
			if (options.renderComments) {
				const markers = renderCommentMarkers(slide, options.className);
				if (markers) section.appendChild(markers);
			}
			out.push(section);
			if (options.renderNotes) {
				const notesEl = renderNotesBlock(slide, options.className);
				if (notesEl) out.push(notesEl);
			}
		}

		return out;
	}
}
