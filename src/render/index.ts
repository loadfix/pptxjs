import type { Options } from '../pptx-preview';
import type { Presentation } from '../presentation';
import { emuToPx } from './geom';
import { makeStyleNode } from './style';
import { applyBackground } from './background';
import { renderShapeLike } from './dispatch';

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
				section.appendChild(renderShapeLike(shape, options.className, presentation.tableStyles));
			}
			out.push(section);
		}

		return out;
	}
}
