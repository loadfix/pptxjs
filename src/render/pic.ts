import type { PicShape } from '../presentation-parser';
import { positionStyle } from './geom';

export function renderPic(pic: PicShape, cls: string): HTMLElement {
	const wrap = document.createElement("div");
	wrap.className = `${cls}-pic`;
	Object.assign(wrap.style, positionStyle(pic.x, pic.y, pic.cx, pic.cy));
	if (pic.src) {
		const img = document.createElement("img");
		img.src = pic.src;
		// The DOCX-side render guidance warns against putting document-derived
		// strings into innerHTML/CSS; `alt` via setAttribute is safe (the
		// browser HTML-encodes attribute values).
		if (pic.alt) img.setAttribute("alt", pic.alt);
		img.style.width = "100%";
		img.style.height = "100%";
		img.style.objectFit = "fill";
		wrap.appendChild(img);
	}
	return wrap;
}
